"""Validation and atomic publication for WPS-produced artifacts."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import importlib
import multiprocessing
from multiprocessing.reduction import ForkingPickler
import os
from pathlib import Path
import pickle
import stat
import sys
import tempfile
import threading
import time
from typing import Any, Callable, Iterable, Union
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
_ROLLBACK_STATE_BUDGET_SECONDS = 1.0
_SPAWN_MAIN_LOCK = threading.Lock()
_MISSING = object()


@dataclass(frozen=True)
class ArtifactState:
    """Observed state of one logical publication path."""

    logical_path: Path
    resolved_path: Path
    exists: bool
    device: int | None = None
    inode: int | None = None
    size: int | None = None
    modified_ns: int | None = None
    sha256: str | None = None
    symlink_target: str | None = None


def snapshot_artifact_state(
    path: Path, *, deadline: float | None = None
) -> ArtifactState:
    """Capture the regular-file bytes currently bound to a logical path."""
    _require_deadline(deadline)
    logical = Path(path).expanduser().absolute()
    try:
        path_stat = logical.lstat()
    except FileNotFoundError:
        return ArtifactState(logical, logical.resolve(strict=False), False)
    symlink_target = None
    if stat.S_ISLNK(path_stat.st_mode):
        symlink_target = os.readlink(logical)
        resolved = logical.resolve(strict=True)
        target_stat = resolved.stat()
    else:
        resolved = logical.resolve(strict=True)
        target_stat = path_stat
    if not stat.S_ISREG(target_stat.st_mode):
        raise ValueError(f"Artifact destination must be a regular file path: {logical}")
    digest = hashlib.sha256()
    with resolved.open("rb") as stream:
        opened_before = os.fstat(stream.fileno())
        while True:
            _require_deadline(deadline)
            block = stream.read(_COPY_CHUNK_BYTES)
            _require_deadline(deadline)
            if not block:
                break
            digest.update(block)
        opened_after = os.fstat(stream.fileno())
    _require_deadline(deadline)
    stable = (
        opened_before.st_dev == target_stat.st_dev
        and opened_before.st_ino == target_stat.st_ino
        and opened_before.st_size == target_stat.st_size
        and opened_before.st_mtime_ns == target_stat.st_mtime_ns
        and opened_after.st_size == opened_before.st_size
        and opened_after.st_mtime_ns == opened_before.st_mtime_ns
    )
    try:
        final_lstat = logical.lstat()
        stable = stable and stat.S_ISLNK(final_lstat.st_mode) == bool(symlink_target)
        if symlink_target is not None:
            stable = (
                stable
                and os.readlink(logical) == symlink_target
                and logical.resolve(strict=True) == resolved
            )
        else:
            stable = (
                stable
                and final_lstat.st_dev == path_stat.st_dev
                and final_lstat.st_ino == path_stat.st_ino
            )
    except OSError:
        stable = False
    if not stable:
        raise ArtifactTransportError(
            "ARTIFACT_STATE_UNSTABLE",
            f"Artifact changed while its publication state was captured: {logical}",
        )
    return ArtifactState(
        logical,
        resolved,
        True,
        opened_before.st_dev,
        opened_before.st_ino,
        opened_before.st_size,
        opened_before.st_mtime_ns,
        digest.hexdigest(),
        symlink_target,
    )


def _artifact_state_matches(
    expected: ArtifactState, *, deadline: float | None = None
) -> bool:
    try:
        current = snapshot_artifact_state(expected.logical_path, deadline=deadline)
    except (OSError, ValueError, ArtifactTransportError):
        return False
    return current == expected


def _require_artifact_state(
    expected: ArtifactState | None, target: Path, *, deadline: float | None = None
) -> None:
    if expected is None:
        return
    if not isinstance(expected, ArtifactState):
        raise TypeError("expected_destination must be an ArtifactState")
    if expected.resolved_path != target:
        raise ValueError("Expected destination state does not match publication target")
    if not _artifact_state_matches(expected, deadline=deadline):
        raise ArtifactTransportError(
            "ARTIFACT_DESTINATION_CHANGED",
            f"Artifact destination changed since publication preflight: {expected.logical_path}",
        )


def _matches_published_state(expected: ArtifactState, current: ArtifactState) -> bool:
    """Compare the exact regular-file identity and bytes prepared for publish."""
    return (
        current.exists
        and current.device == expected.device
        and current.inode == expected.inode
        and current.size == expected.size
        and current.modified_ns == expected.modified_ns
        and current.sha256 == expected.sha256
    )


def _publication_still_owned(
    expected: ArtifactState, target: Path, *, deadline: float | None
) -> bool:
    """Prove that rollback still owns the bytes at *target*.

    Validators and concurrent editors may mutate a destination after publish.
    Identity metadata cannot distinguish an in-place same-size write, so
    rollback is allowed only after a stable content snapshot taken within the
    caller's bounded recovery budget.
    """
    try:
        current = snapshot_artifact_state(target, deadline=deadline)
    except (OSError, ValueError, ArtifactTransportError, TimeoutError):
        return False
    return _matches_published_state(expected, current)


def _rollback_state_deadline() -> float:
    """Give recovery its own bounded state-verification budget."""
    return time.monotonic() + _ROLLBACK_STATE_BUDGET_SECONDS


def _require_plain_serializable(value: Any, *, path: str = "arguments") -> None:
    if value is None or isinstance(value, (bool, int, float, str)):
        return
    if isinstance(value, (list, tuple)):
        for index, item in enumerate(value):
            _require_plain_serializable(item, path=f"{path}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            if not isinstance(key, str):
                raise TypeError(
                    f"ValidatorSpec {path} must use string dictionary keys"
                )
            _require_plain_serializable(item, path=f"{path}.{key}")
        return
    raise TypeError(
        f"ValidatorSpec {path} must contain only plain serializable values"
    )


@dataclass(frozen=True)
class ValidatorSpec:
    """Spawn-safe reference to a top-level validator and plain arguments."""

    identifier: str
    arguments: tuple[Any, ...] = ()

    def __post_init__(self) -> None:
        module_name, separator, function_name = self.identifier.partition(":")
        if (
            not separator
            or not module_name
            or not function_name.isidentifier()
            or any(not part.isidentifier() for part in module_name.split("."))
        ):
            raise ValueError(
                "ValidatorSpec identifier must be 'module:top_level_function'"
            )
        arguments = tuple(self.arguments)
        _require_plain_serializable(arguments)
        object.__setattr__(self, "arguments", arguments)

    @classmethod
    def from_callable(cls, validator: Callable, *arguments: Any) -> "ValidatorSpec":
        module_name = getattr(validator, "__module__", "")
        function_name = getattr(validator, "__qualname__", "")
        if not module_name or not function_name or "." in function_name:
            raise TypeError("ValidatorSpec requires a top-level callable")
        return cls(f"{module_name}:{function_name}", tuple(arguments))

    def resolve(self) -> Callable:
        module_name, function_name = self.identifier.split(":", 1)
        validator = getattr(importlib.import_module(module_name), function_name)
        if not callable(validator):
            raise TypeError(f"ValidatorSpec target is not callable: {self.identifier}")
        return validator

    def __call__(self, path: Path) -> None:
        self.resolve()(Path(path), *self.arguments)


Validator = Union[Callable[[Path], None], ValidatorSpec]


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


def _validation_process_entry(connection, validator: Validator, path: Path) -> None:
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
    validator: Validator, path: Path, deadline: float
) -> None:
    """Run one read-only validator within the remaining public budget.

    Validation may involve vendor parsers and compressed-package traversal
    that cannot be interrupted cooperatively. A spawned worker avoids forking
    the live multi-threaded bridge process and is available on Windows too.
    """
    _require_deadline(deadline)
    if not callable(validator):
        raise TypeError("validator must be a picklable callable or ValidatorSpec")
    try:
        ForkingPickler.dumps((validator, Path(path)))
    except (pickle.PickleError, AttributeError, TypeError, ValueError) as exc:
        raise TypeError(
            "validator must be a picklable callable or ValidatorSpec"
        ) from exc
    context = multiprocessing.get_context("spawn")
    receiver, sender = context.Pipe(duplex=False)
    worker = None
    try:
        worker = context.Process(
            target=_validation_process_entry,
            args=(sender, validator, Path(path)),
            name="wpscomposer-artifact-validation",
            daemon=True,
        )
        # ``spawn`` reloads ``__main__.__file__``. Interactive Python uses a
        # synthetic path such as ``<stdin>``, which makes an otherwise valid
        # public API call fail before the validator starts. Point preparation
        # at a no-op real module only for the synchronous start handshake.
        main_module = sys.modules.get("__main__")
        with _SPAWN_MAIN_LOCK:
            original_main_file = (
                getattr(main_module, "__file__", _MISSING)
                if main_module is not None else _MISSING
            )
            current_main_file = (
                original_main_file if isinstance(original_main_file, str) else ""
            )
            patched_main = bool(
                main_module is not None and not os.path.isfile(current_main_file)
            )
            if patched_main:
                main_module.__file__ = str(
                    Path(__file__).with_name("_spawn_bootstrap.py")
                )
            try:
                worker.start()
            finally:
                if patched_main:
                    if original_main_file is _MISSING:
                        delattr(main_module, "__file__")
                    else:
                        main_module.__file__ = original_main_file
        sender.close()
        budget = max(0.0, deadline - time.monotonic())
        if budget <= 0 or not receiver.poll(budget):
            worker.terminate()
            worker.join(timeout=0.25)
            if worker.is_alive():
                worker.kill()
                worker.join(timeout=0.25)
            raise TimeoutError("Artifact validation deadline expired")
        try:
            payload = receiver.recv()
        except EOFError:
            raise RuntimeError("Artifact validator worker exited without a result") from None
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
        sender.close()
        receiver.close()
        if worker is not None and worker.is_alive():
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
    expected_destination: ArtifactState | None = None,
) -> Path:
    """Validate, copy locally, fsync, atomically replace, and revalidate."""
    source = Path(staged).expanduser().resolve()
    requested_target = Path(destination).expanduser().absolute()
    if expected_destination is None:
        target = requested_target.resolve()
    elif requested_target in {
        expected_destination.logical_path,
        expected_destination.resolved_path,
    }:
        target = expected_destination.resolved_path
    else:
        target = requested_target
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
            published_state = snapshot_artifact_state(temporary, deadline=deadline)
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
                _require_artifact_state(expected_destination, target, deadline=deadline)
                os.replace(temporary, target)
            else:
                _require_deadline(deadline)
                _require_artifact_state(expected_destination, target, deadline=deadline)
                os.link(temporary, target)
                # The link makes the target visible. From this point onward it
                # is an owned publication and every failure must take the same
                # final-validation rollback path.
        except FileExistsError:
            raise
        except OSError as exc:
            raise ArtifactTransportError(
                "ARTIFACT_PUBLISH_FAILED", str(exc)
            ) from exc
        if overwrite:
            temporary = None
        foreign_after_validation = False
        try:
            if not overwrite:
                temporary.unlink()
                temporary = None
            validator(target)
            _require_deadline(deadline)
            try:
                final_state = snapshot_artifact_state(target, deadline=deadline)
            except TimeoutError:
                raise
            except (OSError, ValueError, ArtifactTransportError) as state_exc:
                foreign_after_validation = True
                raise ArtifactTransportError(
                    "FINAL_ARTIFACT_CHANGED",
                    f"Published destination changed during final verification: {target}",
                ) from state_exc
            if not _matches_published_state(published_state, final_state):
                foreign_after_validation = True
                raise ArtifactTransportError(
                    "FINAL_ARTIFACT_CHANGED",
                    f"Published destination changed during final verification: {target}",
                )
        except BaseException as exc:
            if foreign_after_validation:
                recovery_path = backup
                backup = None
                recovery = (
                    f"; previous artifact retained at {recovery_path}"
                    if recovery_path is not None else ""
                )
                raise ArtifactTransportError(
                    "ARTIFACT_ROLLBACK_FAILED",
                    "Published destination changed after final validation" + recovery,
                ) from exc
            rollback_error: ArtifactTransportError | None = None
            try:
                if not _publication_still_owned(
                    published_state, target, deadline=_rollback_state_deadline()
                ):
                    recovery_path = backup
                    backup = None
                    recovery = (
                        f"; previous artifact retained at {recovery_path}"
                        if recovery_path is not None else ""
                    )
                    raise ArtifactTransportError(
                        "ARTIFACT_ROLLBACK_FAILED",
                        "Published destination changed before rollback" + recovery,
                    )
                if backup is not None:
                    os.replace(backup, target)
                    backup = None
                else:
                    target.unlink(missing_ok=True)
            except BaseException as restore_exc:
                if (
                    isinstance(restore_exc, ArtifactTransportError)
                    and restore_exc.code == "ARTIFACT_ROLLBACK_FAILED"
                ):
                    rollback_error = restore_exc
                else:
                    recovery_path = backup
                    recovery = (
                        f"; previous artifact retained at {recovery_path}"
                        if recovery_path is not None
                        else ""
                    )
                    rollback_error = ArtifactTransportError(
                        "ARTIFACT_ROLLBACK_FAILED", f"{restore_exc}{recovery}"
                    )
                    rollback_error.__cause__ = restore_exc
                # Ownership transfers to the operator: never delete the only
                # recovery copy after a failed or unproved rollback.
                backup = None
            if rollback_error is not None:
                if not isinstance(exc, Exception):
                    # Keep process-control semantics at the top of the chain
                    # while retaining rollback diagnostics and recovery.
                    raise exc from rollback_error
                raise rollback_error from exc
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
        | tuple[Path, Path, bool, Callable[[Path], None], ArtifactState | None]
    ],
    *,
    deadline: float | None = None,
) -> list[Path]:
    """Publish a validated artifact set with best-effort whole-set rollback.

    Every staged artifact is copied and validated in its destination directory
    before any destination is changed. Publication is sequential because files
    cannot be atomically replaced as a set; if a later publish or final
    validation fails, all earlier destinations are restored from local backups.
    """
    _require_deadline(deadline)
    entries = []
    for entry in artifacts:
        if len(entry) == 4:
            staged, destination, overwrite, validator = entry
            expected = None
        elif len(entry) == 5:
            staged, destination, overwrite, validator, expected = entry
        else:
            raise ValueError("Artifact group entries require four or five fields")
        requested = Path(destination).expanduser().absolute()
        target = (
            requested.resolve()
            if expected is None
            else expected.resolved_path
            if requested in {expected.logical_path, expected.resolved_path}
            else requested
        )
        entries.append((
            Path(staged).expanduser().resolve(),
            target,
            bool(overwrite),
            validator,
            expected,
        ))
    targets = [destination for _staged, destination, _overwrite, _validator, _expected in entries]
    if len(set(targets)) != len(targets):
        raise ValueError("Artifact group destinations must be unique")

    prepared: list[Path] = []
    backups: dict[Path, Path] = {}
    published: list[Path] = []
    foreign_publications: set[Path] = set()
    recovery_deadline: float | None = None
    try:
        for staged, target, overwrite, validator, _expected in entries:
            try:
                _require_deadline(deadline)
                validator(staged)
                _require_deadline(deadline)
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
                    prepared.append(local)
                    with staged.open("rb") as incoming:
                        copy_stream_before_deadline(incoming, stream, deadline)
                    stream.flush()
                    _require_deadline(deadline)
                    os.fsync(stream.fileno())
                    _require_deadline(deadline)
                validator(local)
                _require_deadline(deadline)
            except ArtifactValidationError as exc:
                local.unlink(missing_ok=True)
                raise ArtifactTransportError(
                    "ARTIFACT_PUBLISH_FAILED", str(exc)
                ) from exc
            except TimeoutError:
                raise
            except OSError as exc:
                if "local" in locals():
                    local.unlink(missing_ok=True)
                raise ArtifactTransportError(
                    "ARTIFACT_PUBLISH_FAILED", str(exc)
                ) from exc

        for _staged, target, overwrite, _validator, _expected in entries:
            _require_deadline(deadline)
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
                            copy_stream_before_deadline(existing, stream, deadline)
                        stream.flush()
                        _require_deadline(deadline)
                        os.fsync(stream.fileno())
                        _require_deadline(deadline)
                except TimeoutError:
                    raise
                except OSError as exc:
                    raise ArtifactTransportError(
                        "ARTIFACT_PUBLISH_FAILED", str(exc)
                    ) from exc
        # A changed later destination must be rejected before the first member
        # becomes visible. Each member is checked again at its own replacement
        # boundary because this multi-file protocol cannot provide filesystem CAS.
        for _staged, target, _overwrite, _validator, expected in entries:
            _require_artifact_state(expected, target, deadline=deadline)
        published_states: dict[Path, ArtifactState] = {}
        for local, (_staged, target, overwrite, _validator, expected) in zip(prepared, entries):
            _require_deadline(deadline)
            try:
                published_states[target] = snapshot_artifact_state(
                    local, deadline=deadline
                )
                _require_artifact_state(expected, target, deadline=deadline)
                if overwrite:
                    os.replace(local, target)
                    published.append(target)
                else:
                    os.link(local, target)
                    # The link makes the destination externally visible.  Own
                    # it for rollback before any later cleanup can fail.
                    published.append(target)
                    local.unlink()
            except FileExistsError:
                raise
            except TimeoutError:
                raise
            except OSError as exc:
                raise ArtifactTransportError(
                    "ARTIFACT_PUBLISH_FAILED", str(exc)
                ) from exc
        final_validation_error: BaseException | None = None
        for _staged, target, _overwrite, validator, _expected in entries:
            try:
                _require_deadline(deadline)
                validator(target)
                _require_deadline(deadline)
            except BaseException as exc:
                final_validation_error = exc
                recovery_deadline = _rollback_state_deadline()
                break
        final_state_errors = []
        state_deadline = (
            recovery_deadline
            if final_validation_error is not None
            else deadline
        )
        for target in targets:
            try:
                current_state = snapshot_artifact_state(
                    target, deadline=state_deadline
                )
            except (OSError, ValueError, ArtifactTransportError, TimeoutError) as state_exc:
                foreign_publications.add(target)
                final_state_errors.append((target, state_exc))
                continue
            if not _matches_published_state(published_states[target], current_state):
                foreign_publications.add(target)
        if (
            final_validation_error is not None
            and not isinstance(final_validation_error, Exception)
        ):
            raise final_validation_error
        if foreign_publications:
            changed = ", ".join(str(target) for target in sorted(
                foreign_publications, key=str
            ))
            error = ArtifactTransportError(
                "FINAL_ARTIFACT_CHANGED",
                f"Published destinations changed during final verification: {changed}",
            )
            if final_state_errors:
                raise error from final_state_errors[0][1]
            raise error
        if final_validation_error is not None:
            raise ArtifactTransportError(
                "FINAL_ARTIFACT_INVALID", str(final_validation_error)
            ) from final_validation_error
        return targets
    except BaseException as publish_exc:
        rollback_errors = []
        rollback_deadline = recovery_deadline or _rollback_state_deadline()
        for target in reversed(published):
            backup = backups.pop(target, None)
            try:
                if (
                    target in foreign_publications
                    or not _publication_still_owned(
                        published_states[target],
                        target,
                        deadline=rollback_deadline,
                    )
                ):
                    raise ArtifactTransportError(
                        "ARTIFACT_ROLLBACK_FAILED",
                        "Published destination changed before group rollback",
                    )
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
            rollback_error = ArtifactTransportError(
                "ARTIFACT_ROLLBACK_FAILED",
                f"{publish_exc}; recovery artifacts retained at {retained}",
            )
            if not isinstance(publish_exc, Exception):
                raise publish_exc from rollback_error
            raise rollback_error from publish_exc
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
