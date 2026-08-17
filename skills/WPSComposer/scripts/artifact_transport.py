"""Validation and atomic publication for WPS-produced artifacts."""

from __future__ import annotations

import os
from pathlib import Path
import shutil
import tempfile
from typing import Callable
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
) -> Path:
    """Validate, copy locally, fsync, atomically replace, and revalidate."""
    source = Path(staged).expanduser().resolve()
    target = Path(destination).expanduser().resolve()
    try:
        validator(source)
    except ArtifactValidationError as exc:
        raise ArtifactTransportError(
            "STAGED_ARTIFACT_INVALID", str(exc)
        ) from exc
    target.parent.mkdir(parents=True, exist_ok=True)
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
                    shutil.copyfileobj(incoming, stream)
                stream.flush()
                os.fsync(stream.fileno())
        except OSError as exc:
            raise ArtifactTransportError(
                "ARTIFACT_PUBLISH_FAILED", str(exc)
            ) from exc

        try:
            validator(temporary)
        except ArtifactValidationError as exc:
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
                                shutil.copyfileobj(existing, stream)
                            stream.flush()
                            os.fsync(stream.fileno())
                    except OSError as exc:
                        raise ArtifactTransportError(
                            "ARTIFACT_PUBLISH_FAILED", str(exc)
                        ) from exc
                os.replace(temporary, target)
            else:
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
        except ArtifactValidationError as exc:
            try:
                if backup is not None:
                    os.replace(backup, target)
                    backup = None
                else:
                    target.unlink(missing_ok=True)
            except OSError as restore_exc:
                raise ArtifactTransportError(
                    "ARTIFACT_PUBLISH_FAILED", str(restore_exc)
                ) from restore_exc
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
