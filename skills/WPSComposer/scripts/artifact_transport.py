"""Validation and atomic publication for WPS-produced artifacts."""

from __future__ import annotations

import os
from pathlib import Path
import shutil
import tempfile
from typing import Callable
import zipfile


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
    """Require a non-trivial file with a PDF signature."""
    target = _require_regular_file(path, "pdf")
    try:
        with target.open("rb") as stream:
            if not stream.read(5).startswith(b"%PDF-"):
                raise ArtifactValidationError(f"Invalid PDF signature: {target}")
    except FileNotFoundError as exc:
        raise ArtifactValidationError(f"PDF artifact is missing: {target}") from exc


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
    except (OSError, zipfile.BadZipFile) as exc:
        raise ArtifactValidationError(
            f"Invalid {normalized.upper()} ZIP package: {target}"
        ) from exc
    if expected_member not in members:
        raise ArtifactValidationError(
            f"{normalized.upper()} package is missing {expected_member}: {target}"
        )


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
            raise ArtifactTransportError(
                "FINAL_ARTIFACT_INVALID", str(exc)
            ) from exc
        return target
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
