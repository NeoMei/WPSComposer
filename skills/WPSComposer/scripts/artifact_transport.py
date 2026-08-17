"""Validation and atomic publication for WPS-produced artifacts."""

from __future__ import annotations

import multiprocessing
import os
from pathlib import Path
import sys
import tempfile
import time
from typing import Callable, Iterable
import zipfile
from xml.etree import ElementTree


class ArtifactValidationError(RuntimeError):
    """Raised when a staged or published artifact is structurally invalid."""


class ArtifactTransportError(RuntimeError):
    """Stable failure raised by a particular artifact transport stage."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


_OFFICE_MEMBERS = {
    "docx": "word/document.xml",
    "xlsx": "xl/workbook.xml",
    "pptx": "ppt/presentation.xml",
}
_COPY_CHUNK_BYTES = 1024 * 1024


def _require_deadline(deadline: float | None) -> None:
    if deadline is not None and time.monotonic() >= deadline:
        raise TimeoutError("Artifact transport deadline expired")


def copy_stream_before_deadline(incoming, outgoing, deadline: float | None) -> None:
    while True:
        _require_deadline(deadline)
        block = incoming.read(_COPY_CHUNK_BYTES)
        _require_deadline(deadline)
        if not block:
            return
        outgoing.write(block)
        _require_deadline(deadline)


def copy_file_before_deadline(
    source: Path, target: Path, *, deadline: float
) -> Path:
    """Stage one file cooperatively and remove partial output on timeout."""
    source_path = Path(source).expanduser().resolve()
    target_path = Path(target).expanduser().resolve()
    if source_path == target_path:
        raise ValueError("Artifact source and target must be different paths")
    _require_deadline(deadline)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    _require_deadline(deadline)
    created = False
    try:
        with source_path.open("rb") as incoming, target_path.open("xb") as outgoing:
            created = True
            copy_stream_before_deadline(incoming, outgoing, deadline)
            _require_deadline(deadline)
            outgoing.flush()
            _require_deadline(deadline)
            os.fsync(outgoing.fileno())
            _require_deadline(deadline)
        os.chmod(target_path, 0o600)
        _require_deadline(deadline)
        return target_path
    except BaseException:
        if created:
            target_path.unlink(missing_ok=True)
        raise


def _validation_process_entry(connection, validator, path: Path) -> None:
    try:
        validator(path)
    except ArtifactValidationError as exc:
        payload = ("artifact", str(exc))
    except BaseException as exc:
        payload = ("runtime", type(exc).__name__, str(exc))
    else:
        payload = ("ok",)
    try:
        connection.send(payload)
    finally:
        connection.close()


def validate_before_deadline(
    validator: Callable[[Path], None], path: Path, deadline: float
) -> None:
    """Run one read-only validator within the remaining public budget.

    Validation may involve vendor parsers and compressed-package traversal
    that cannot be interrupted cooperatively. A spawned worker avoids forking
    the live multi-threaded bridge process and is available on Windows too.
    """
    _require_deadline(deadline)
    context = multiprocessing.get_context("spawn")
    receiver, sender = context.Pipe(duplex=False)
    worker = context.Process(
        target=_validation_process_entry,
        args=(sender, validator, Path(path)),
        name="wpscomposer-artifact-validation",
        daemon=True,
    )
    worker.start()
    sender.close()
    try:
        budget = max(0.0, deadline - time.monotonic())
        if budget <= 0 or not receiver.poll(budget):
            worker.terminate()
            worker.join(timeout=0.25)
            if worker.is_alive():
                worker.kill()
                worker.join(timeout=0.25)
            raise TimeoutError("Artifact validation deadline expired")
        payload = receiver.recv()
        worker.join(timeout=0.25)
        if worker.is_alive():
            worker.terminate()
            worker.join(timeout=0.25)
        status = payload[0]
        if status == "artifact":
            raise ArtifactValidationError(payload[1])
        if status == "runtime":
            raise RuntimeError(
                f"Artifact validator raised {payload[1]}: {payload[2]}"
            )
        _require_deadline(deadline)
    finally:
        receiver.close()
        if worker.is_alive():
            worker.terminate()
            worker.join(timeout=0.25)


def _require_regular_file(path: Path, format_name: str) -> Path:
    target = Path(path).expanduser().resolve()
    try:
        if not target.is_file():
            raise ArtifactValidationError(
                f"{format_name.upper()} artifact is missing: {target}"
            )
        if target.stat().st_size < 256:
            # floor catches empty/truncated output; small valid PDFs are ~400+ bytes
            raise ArtifactValidationError(
                f"{format_name.upper()} artifact is too small: {target}"
            )
    except FileNotFoundError as exc:
        # the file vanished between is_file() and stat() (async replace)
        raise ArtifactValidationError(
            f"{format_name.upper()} artifact is missing: {target}"
        ) from exc
    return target


def validate_pdf(path: Path) -> None:
    """Require a parseable PDF with a trailer, root, and page tree."""
    target = _require_regular_file(path, "pdf")
    try:
        data = target.read_bytes()
    except FileNotFoundError as exc:
        raise ArtifactValidationError(f"PDF artifact is missing: {target}") from exc
    if not data.startswith(b"%PDF-"):
        raise ArtifactValidationError(f"Invalid PDF signature: {target}")
    tail = data[-4096:].rstrip()
    if not tail.endswith(b"%%EOF") or b"startxref" not in tail:
        raise ArtifactValidationError(f"Invalid PDF structure: {target}")

    try:
        from pypdf import PdfReader
    except ImportError:
        # Dependency-free structural fallback for core installs.  Verify the
        # cross-reference target and the minimum catalog/page-tree contract.
        try:
            startxref = int(tail.rsplit(b"startxref", 1)[1].splitlines()[1])
            marker = data[startxref:startxref + 32].lstrip()
        except (IndexError, ValueError):
            raise ArtifactValidationError(f"Invalid PDF structure: {target}")
        if not (marker.startswith(b"xref") or b"/Type /XRef" in marker):
            raise ArtifactValidationError(f"Invalid PDF structure: {target}")
        if b"/Root" not in tail or b"/Type /Catalog" not in data:
            raise ArtifactValidationError(f"Invalid PDF structure: {target}")
        if b"/Type /Pages" not in data:
            raise ArtifactValidationError(f"Invalid PDF structure: {target}")
        return

    try:
        reader = PdfReader(str(target), strict=True)
        reader.trailer["/Root"]
        len(reader.pages)
    except Exception as exc:
        raise ArtifactValidationError(f"Invalid PDF structure: {target}") from exc


def validate_office_package(path: Path, format_name: str) -> None:
    """Require a valid OOXML ZIP containing its format-defining member."""
    normalized = str(format_name).lower().lstrip(".")
    try:
        expected_member = _OFFICE_MEMBERS[normalized]
    except KeyError as exc:
        raise ValueError(f"Unsupported Office package format: {format_name}") from exc
    target = _require_regular_file(path, normalized)
    try:
        with zipfile.ZipFile(target) as package:
            corrupt_member = package.testzip()
            if corrupt_member is not None:
                raise ArtifactValidationError(
                    f"Corrupt {normalized.upper()} package member: {corrupt_member}"
                )
            members = set(package.namelist())
            if expected_member not in members:
                raise ArtifactValidationError(
                    f"{normalized.upper()} package is missing {expected_member}: {target}"
                )
            if "[Content_Types].xml" not in members:
                raise ArtifactValidationError(
                    f"{normalized.upper()} package is missing [Content_Types].xml: {target}"
                )
            defining_xml = package.read(expected_member)
            content_types_xml = package.read("[Content_Types].xml")
    except (OSError, zipfile.BadZipFile) as exc:
        raise ArtifactValidationError(
            f"Invalid {normalized.upper()} ZIP package: {target}"
        ) from exc
    try:
        ElementTree.fromstring(content_types_xml)
        ElementTree.fromstring(defining_xml)
    except ElementTree.ParseError as exc:
        raise ArtifactValidationError(
            f"Invalid {normalized.upper()} XML (malformed XML): {target}"
        ) from exc


def publish_artifact(
    staged: Path,
    destination: Path,
    *,
    overwrite: bool,
    validator: Callable[[Path], None],
    deadline: float | None = None,
) -> Path:
    """Validate, copy locally, fsync, atomically replace, and revalidate."""
    source = Path(staged).expanduser().resolve()
    target = Path(destination).expanduser().resolve()
    try:
        _require_deadline(deadline)
        validator(source)
        _require_deadline(deadline)
    except ArtifactValidationError as exc:
        raise ArtifactTransportError(
            "STAGED_ARTIFACT_INVALID", str(exc)
        ) from exc
    _require_deadline(deadline)
    target.parent.mkdir(parents=True, exist_ok=True)
    _require_deadline(deadline)
    if target.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {target}")

    temporary: Path | None = None
    backup: Path | None = None
    try:
        try:
            with tempfile.NamedTemporaryFile(
                mode="wb",
                dir=target.parent,
                prefix=".wpscomposer-",
                suffix=".tmp",
                delete=False,
            ) as stream:
                temporary = Path(stream.name)
                with source.open("rb") as incoming:
                    copy_stream_before_deadline(incoming, stream, deadline)
                _require_deadline(deadline)
                stream.flush()
                _require_deadline(deadline)
                os.fsync(stream.fileno())
                _require_deadline(deadline)
        except (OSError, TimeoutError) as exc:
            raise ArtifactTransportError(
                "ARTIFACT_PUBLISH_FAILED", str(exc)
            ) from exc

        try:
            _require_deadline(deadline)
            validator(temporary)
            _require_deadline(deadline)
        except (ArtifactValidationError, TimeoutError) as exc:
            raise ArtifactTransportError(
                "ARTIFACT_PUBLISH_FAILED", str(exc)
            ) from exc
        try:
            if overwrite:
                if target.exists():
                    try:
                        with tempfile.NamedTemporaryFile(
                            mode="wb",
                            dir=target.parent,
                            prefix=".wpscomposer-backup-",
                            suffix=".tmp",
                            delete=False,
                        ) as stream:
                            backup = Path(stream.name)
                            with target.open("rb") as existing:
                                copy_stream_before_deadline(existing, stream, deadline)
                            _require_deadline(deadline)
                            stream.flush()
                            _require_deadline(deadline)
                            os.fsync(stream.fileno())
                            _require_deadline(deadline)
                    except (OSError, TimeoutError) as exc:
                        raise ArtifactTransportError(
                            "ARTIFACT_PUBLISH_FAILED", str(exc)
                        ) from exc
                _require_deadline(deadline)
                os.replace(temporary, target)
            else:
                _require_deadline(deadline)
                os.link(temporary, target)
                temporary.unlink()
        except FileExistsError:
            raise
        except OSError as exc:
            raise ArtifactTransportError(
                "ARTIFACT_PUBLISH_FAILED", str(exc)
            ) from exc
        temporary = None
        try:
            validator(target)
            _require_deadline(deadline)
        except BaseException as exc:
            try:
                if backup is not None:
                    os.replace(backup, target)
                    backup = None
                else:
                    target.unlink(missing_ok=True)
            except BaseException as restore_exc:
                recovery_path = backup
                # Ownership transfers to the operator: never delete the only
                # recovery copy after a failed rollback.
                backup = None
                recovery = (
                    f"; previous artifact retained at {recovery_path}"
                    if recovery_path is not None
                    else ""
                )
                rollback_error = ArtifactTransportError(
                    "ARTIFACT_ROLLBACK_FAILED", f"{restore_exc}{recovery}"
                )
                if not isinstance(exc, Exception):
                    # Keep process-control semantics at the top of the chain
                    # while retaining both rollback diagnostics and recovery.
                    rollback_error.__cause__ = restore_exc
                    raise exc from rollback_error
                raise rollback_error from restore_exc
            if not isinstance(exc, Exception):
                raise
            raise ArtifactTransportError(
                "FINAL_ARTIFACT_INVALID", str(exc)
            ) from exc
        if backup is not None:
            backup.unlink(missing_ok=True)
            backup = None
        return target
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
        if backup is not None:
            backup.unlink(missing_ok=True)


def publish_artifact_group(
    artifacts: Iterable[
        tuple[Path, Path, bool, Callable[[Path], None]]
    ],
) -> list[Path]:
    """Publish a validated artifact set with best-effort whole-set rollback.

    Every staged artifact is copied and validated in its destination directory
    before any destination is changed. Publication is sequential because files
    cannot be atomically replaced as a set; if a later publish or final
    validation fails, all earlier destinations are restored from local backups.
    """
    entries = [
        (Path(staged).expanduser().resolve(),
         Path(destination).expanduser().resolve(), bool(overwrite), validator)
        for staged, destination, overwrite, validator in artifacts
    ]
    targets = [destination for _staged, destination, _overwrite, _validator in entries]
    if len(set(targets)) != len(targets):
        raise ValueError("Artifact group destinations must be unique")

    prepared: list[Path] = []
    backups: dict[Path, Path] = {}
    published: list[Path] = []
    try:
        for staged, target, overwrite, validator in entries:
            try:
                validator(staged)
            except ArtifactValidationError as exc:
                raise ArtifactTransportError(
                    "STAGED_ARTIFACT_INVALID", str(exc)
                ) from exc
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists() and not overwrite:
                raise FileExistsError(f"Output already exists: {target}")
            try:
                with tempfile.NamedTemporaryFile(
                    mode="wb",
                    dir=target.parent,
                    prefix=".wpscomposer-group-publish-",
                    suffix=".tmp",
                    delete=False,
                ) as stream:
                    local = Path(stream.name)
                    with staged.open("rb") as incoming:
                        copy_stream_before_deadline(incoming, stream, None)
                    stream.flush()
                    os.fsync(stream.fileno())
                validator(local)
            except ArtifactValidationError as exc:
                local.unlink(missing_ok=True)
                raise ArtifactTransportError(
                    "ARTIFACT_PUBLISH_FAILED", str(exc)
                ) from exc
            except OSError as exc:
                if "local" in locals():
                    local.unlink(missing_ok=True)
                raise ArtifactTransportError(
                    "ARTIFACT_PUBLISH_FAILED", str(exc)
                ) from exc
            prepared.append(local)

        for _staged, target, overwrite, _validator in entries:
            if overwrite and target.exists():
                try:
                    with tempfile.NamedTemporaryFile(
                        mode="wb",
                        dir=target.parent,
                        prefix=".wpscomposer-group-backup-",
                        suffix=".tmp",
                        delete=False,
                    ) as stream:
                        backup = Path(stream.name)
                        # Own the backup as soon as it exists.  A partial copy,
                        # flush, or fsync failure must not orphan old document
                        # bytes in the destination directory.
                        backups[target] = backup
                        with target.open("rb") as existing:
                            copy_stream_before_deadline(existing, stream, None)
                        stream.flush()
                        os.fsync(stream.fileno())
                except OSError as exc:
                    raise ArtifactTransportError(
                        "ARTIFACT_PUBLISH_FAILED", str(exc)
                    ) from exc
        for local, (_staged, target, overwrite, _validator) in zip(prepared, entries):
            try:
                if overwrite:
                    os.replace(local, target)
                else:
                    os.link(local, target)
                    # The link makes the destination externally visible.  Own
                    # it for rollback before any later cleanup can fail.
                    published.append(target)
                    local.unlink()
            except FileExistsError:
                raise
            except OSError as exc:
                raise ArtifactTransportError(
                    "ARTIFACT_PUBLISH_FAILED", str(exc)
                ) from exc
            if overwrite:
                published.append(target)

        for _staged, target, _overwrite, validator in entries:
            try:
                validator(target)
            except BaseException as exc:
                raise ArtifactTransportError(
                    "FINAL_ARTIFACT_INVALID", str(exc)
                ) from exc
        return targets
    except BaseException as publish_exc:
        rollback_errors = []
        for target in reversed(published):
            backup = backups.pop(target, None)
            try:
                if backup is None:
                    target.unlink(missing_ok=True)
                else:
                    os.replace(backup, target)
            except BaseException as restore_exc:
                rollback_errors.append((target, backup, restore_exc))
        if rollback_errors:
            retained = ", ".join(
                str(backup if backup is not None else target)
                for target, backup, _exc in rollback_errors
            )
            raise ArtifactTransportError(
                "ARTIFACT_ROLLBACK_FAILED",
                f"{publish_exc}; recovery artifacts retained at {retained}",
            ) from publish_exc
        raise
    finally:
        active_error = sys.exc_info()[1]
        cleanup_errors = []
        for owned in [*prepared, *backups.values()]:
            try:
                owned.unlink(missing_ok=True)
            except OSError as cleanup_exc:
                cleanup_errors.append((owned, cleanup_exc))
        if cleanup_errors:
            retained = ", ".join(str(path) for path, _exc in cleanup_errors)
            detail = "; ".join(str(exc) for _path, exc in cleanup_errors)
            cleanup_error = ArtifactTransportError(
                "ARTIFACT_CLEANUP_FAILED",
                f"{detail}; recovery artifacts retained at {retained}",
            )
            if active_error is not None and not isinstance(active_error, Exception):
                raise active_error from cleanup_error
            raise cleanup_error from active_error
