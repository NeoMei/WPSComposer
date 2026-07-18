from __future__ import annotations

import os
from pathlib import Path
import zipfile

import pytest

from skills.WPSComposer.scripts.artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    publish_artifact,
    validate_office_package,
    validate_pdf,
)


def _write_pdf(path: Path, payload: bytes = b"x" * 2048) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"%PDF-1.7\n" + payload)
    return path


def _write_package(path: Path, member: str) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        package.writestr(member, "<root />" + "x" * 2048)
    return path


def test_validate_pdf_requires_signature_and_minimum_size(tmp_path: Path):
    validate_pdf(_write_pdf(tmp_path / "valid.pdf"))

    bad = tmp_path / "bad.pdf"
    bad.write_bytes(b"not-a-pdf" + b"x" * 2048)
    with pytest.raises(ArtifactValidationError, match="Invalid PDF signature"):
        validate_pdf(bad)


@pytest.mark.parametrize(
    ("format_name", "member"),
    [
        ("docx", "word/document.xml"),
        ("xlsx", "xl/workbook.xml"),
        ("pptx", "ppt/presentation.xml"),
    ],
)
def test_validate_office_package_requires_expected_member(
    tmp_path: Path, format_name: str, member: str
):
    valid = _write_package(tmp_path / f"valid.{format_name}", member)
    validate_office_package(valid, format_name)

    incomplete = _write_package(
        tmp_path / f"incomplete.{format_name}", "custom/missing.xml"
    )
    with pytest.raises(ArtifactValidationError, match=member):
        validate_office_package(incomplete, format_name)


def test_publish_validates_staged_temporary_and_final_copies(
    tmp_path: Path,
):
    staged = _write_pdf(tmp_path / "stage" / "result.pdf")
    destination = tmp_path / "output" / "result.pdf"
    validated: list[Path] = []

    def validator(path: Path) -> None:
        target = Path(path).resolve()
        validated.append(target)
        validate_pdf(target)

    result = publish_artifact(
        staged,
        destination,
        overwrite=False,
        validator=validator,
    )

    assert result == destination.resolve()
    assert validated[0] == staged.resolve()
    assert validated[-1] == destination.resolve()
    assert len(validated) == 3
    assert validated[1].parent == destination.parent.resolve()
    assert validated[1].name.startswith(".wpscomposer-")
    assert not validated[1].exists()
    assert result.read_bytes() == staged.read_bytes()


def test_publish_refuses_existing_output_without_overwrite(tmp_path: Path):
    staged = _write_pdf(tmp_path / "stage.pdf")
    destination = _write_pdf(tmp_path / "result.pdf", b"old" * 1024)
    original = destination.read_bytes()

    with pytest.raises(FileExistsError, match="Output already exists"):
        publish_artifact(
            staged,
            destination,
            overwrite=False,
            validator=validate_pdf,
        )

    assert destination.read_bytes() == original


def test_publish_overwrites_atomically_when_requested(tmp_path: Path):
    staged = _write_pdf(tmp_path / "stage.pdf", b"new" * 1024)
    destination = _write_pdf(tmp_path / "result.pdf", b"old" * 1024)

    result = publish_artifact(
        staged,
        destination,
        overwrite=True,
        validator=validate_pdf,
    )

    assert result.read_bytes() == staged.read_bytes()
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))


def test_publish_validation_failure_leaves_no_partial_destination(
    tmp_path: Path,
):
    staged = tmp_path / "bad.pdf"
    staged.write_bytes(b"broken")
    destination = tmp_path / "result.pdf"

    with pytest.raises(ArtifactTransportError) as caught:
        publish_artifact(
            staged,
            destination,
            overwrite=False,
            validator=validate_pdf,
        )

    assert caught.value.code == "STAGED_ARTIFACT_INVALID"
    assert not destination.exists()
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))


def test_publish_uses_os_replace_in_destination_directory(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    staged = _write_pdf(tmp_path / "staging" / "result.pdf")
    destination = tmp_path / "external" / "result.pdf"
    calls: list[tuple[Path, Path]] = []
    original_replace = os.replace

    def observed_replace(source, target):
        calls.append((Path(source).resolve(), Path(target).resolve()))
        original_replace(source, target)

    monkeypatch.setattr(os, "replace", observed_replace)

    publish_artifact(
        staged,
        destination,
        overwrite=True,
        validator=validate_pdf,
    )

    assert len(calls) == 1
    assert calls[0][0].parent == destination.parent.resolve()
    assert calls[0][1] == destination.resolve()


def test_publish_does_not_clobber_target_created_during_no_overwrite_race(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    staged = _write_pdf(tmp_path / "stage.pdf", b"staged" * 1024)
    destination = tmp_path / "result.pdf"
    intruder = b"%PDF-1.7\n" + b"intruder" * 1024
    original_link = os.link

    def raced_link(source, target):
        Path(target).write_bytes(intruder)
        return original_link(source, target)

    monkeypatch.setattr(os, "link", raced_link)

    with pytest.raises(FileExistsError):
        publish_artifact(
            staged,
            destination,
            overwrite=False,
            validator=validate_pdf,
        )

    assert destination.read_bytes() == intruder
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))
