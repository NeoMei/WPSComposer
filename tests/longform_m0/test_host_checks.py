from __future__ import annotations

import math
import os
from pathlib import Path
from types import MappingProxyType
import zipfile

import pytest
from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, NumberObject

from skills.WPSComposer.scripts.longform_m0 import host_checks


def write_docx(path: Path) -> Path:
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        package.writestr(
            "word/document.xml",
            "<document>" + ("x" * 2048) + "</document>",
        )
    return path


def write_pdf(path: Path) -> Path:
    writer = PdfWriter()
    page = writer.add_blank_page(width=612, height=792)
    page.cropbox.lower_left = (10, 20)
    page.cropbox.upper_right = (602, 772)
    page[NameObject("/Rotate")] = NumberObject(90)
    font = DictionaryObject(
        {
            NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"),
            NameObject("/BaseFont"): NameObject("/Helvetica"),
        }
    )
    font_ref = writer._add_object(font)
    page[NameObject("/Resources")] = DictionaryObject(
        {
            NameObject("/Font"): DictionaryObject(
                {NameObject("/F1"): font_ref}
            )
        }
    )
    with path.open("wb") as stream:
        writer.write(stream)
    return path


def test_dependency_failure_happens_before_executor(monkeypatch):
    called = False

    def executor():
        nonlocal called
        called = True

    monkeypatch.setattr(
        host_checks,
        "_find_missing_dependencies",
        lambda: ["pypdf"],
    )

    with pytest.raises(RuntimeError, match="pypdf"):
        host_checks.run_after_dependency_gate(executor)

    assert called is False


def test_dependency_gate_accepts_installed_core_packages():
    assert host_checks.require_probe_dependencies() == (
        "pypdf",
        "pdfplumber",
        "Pillow",
    )


# mode-bit assertions apply to POSIX filesystems only
@pytest.mark.skipif(os.name != "posix", reason="POSIX file-mode assertions")
def test_prepare_evidence_directory_creates_private_empty_directory(tmp_path: Path):
    output = tmp_path / "evidence"

    prepared = host_checks.prepare_evidence_directory(output)

    assert prepared == output.resolve()
    assert prepared.is_dir()
    assert prepared.stat().st_mode & 0o777 == 0o700


@pytest.mark.skipif(os.name != "posix", reason="POSIX file-mode assertions")
def test_prepare_evidence_directory_rejects_prior_entries(tmp_path: Path):
    output = tmp_path / "evidence"
    output.mkdir()
    os.chmod(output, 0o755)
    (output / "old.pdf").write_bytes(b"old")

    with pytest.raises(FileExistsError, match="not empty"):
        host_checks.prepare_evidence_directory(output)

    assert output.stat().st_mode & 0o777 == 0o755


def test_validate_native_artifacts_returns_hashes_and_pdf_snapshot(tmp_path: Path):
    docx = write_docx(tmp_path / "probe.docx")
    pdf = write_pdf(tmp_path / "probe.pdf")

    result = host_checks.validate_native_artifacts(docx, pdf)

    assert result["docx"]["name"] == "probe.docx"
    assert len(result["docx"]["sha256"]) == 64
    assert result["pdf"]["name"] == "probe.pdf"
    assert result["pdf"]["snapshot"]["pageCount"] == 1
    page = result["pdf"]["snapshot"]["pages"][0]
    assert page == {
        "physicalPage": 1,
        "mediaBox": [0.0, 0.0, 612.0, 792.0],
        "cropBox": [10.0, 20.0, 602.0, 772.0],
        "rotation": 90,
        "fonts": ["Helvetica"],
        "characterBounds": None,
    }
    assert result["pdf"]["snapshot"]["fonts"] == ["Helvetica"]


def test_validate_native_artifacts_rejects_malformed_files(tmp_path: Path):
    docx = tmp_path / "bad.docx"
    pdf = tmp_path / "bad.pdf"
    docx.write_bytes(b"not-a-zip" * 64)
    pdf.write_bytes(b"not-a-pdf" * 64)

    with pytest.raises(RuntimeError, match="Invalid DOCX ZIP package"):
        host_checks.validate_native_artifacts(docx, pdf)


def test_snapshot_pdf_uses_bounds_but_never_returns_text(monkeypatch, tmp_path: Path):
    pdf = write_pdf(tmp_path / "probe.pdf")

    class FakePage:
        chars = [
            {"text": "secret", "x0": 10, "x1": 20, "top": 30, "bottom": 40},
            {"text": "body", "x0": 5, "x1": 25, "top": 20, "bottom": 45},
        ]

    class FakePdf:
        pages = [FakePage()]

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, traceback):
            return None

    monkeypatch.setattr(host_checks.pdfplumber, "open", lambda _: FakePdf())

    snapshot = host_checks.snapshot_pdf(pdf)

    assert snapshot["pages"][0]["characterBounds"] == {
        "count": 2,
        "bbox": [5.0, 20.0, 25.0, 45.0],
    }
    assert "secret" not in repr(snapshot)
    assert "body" not in repr(snapshot)


def test_snapshot_pdf_markers_returns_only_page_and_bbox(monkeypatch, tmp_path: Path):
    pdf = write_pdf(tmp_path / "probe.pdf")

    class FakePage:
        chars = [
            {"text": "X", "x0": 70, "x1": 76, "top": 72, "bottom": 83},
            {"text": "Y", "x0": 76, "x1": 82, "top": 72, "bottom": 83},
            {"text": "5", "x0": 82, "x1": 88, "top": 72, "bottom": 83},
        ]
        rects = [
            {"x0": 70, "x1": 190, "top": 70, "bottom": 106},
        ]

    class FakePdf:
        pages = [FakePage()]

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, traceback):
            return None

    monkeypatch.setattr(host_checks.pdfplumber, "open", lambda _: FakePdf())

    snapshot = host_checks.snapshot_pdf_markers(pdf, {"shape": "XY5"})

    assert snapshot == {
        "shape": {
            "physicalPage": 1,
            "bbox": [70.0, 72.0, 88.0, 83.0],
            "frameBBox": [70.0, 70.0, 190.0, 106.0],
        }
    }
    assert "XY5" not in repr(snapshot)


@pytest.mark.parametrize(
    "value",
    [
        {"text": "document body"},
        {"bookmarkMap": {"external": "internal"}},
        {"field_hash": "a" * 64},
        {"nested": {"sourcePath": "/private/source.docx"}},
        {"path": "/Users/alice/private.docx"},
        {"path": "C:\\Users\\alice\\private.docx"},
    ],
)
def test_privacy_validation_rejects_sensitive_keys_and_paths(value):
    with pytest.raises(ValueError, match="private evidence"):
        host_checks.validate_evidence_privacy(value)


def test_privacy_validation_rejects_explicit_runtime_roots(tmp_path: Path):
    staging = tmp_path / "session"

    with pytest.raises(ValueError, match="private evidence"):
        host_checks.validate_evidence_privacy(
            {"message": f"failed under {staging}/probe.docx"},
            forbidden_roots=(staging,),
        )


def test_privacy_validation_accepts_finite_structural_snapshot():
    value = {
        "pageCount": 2,
        "pages": [{"rotation": 0, "cropBox": [0.0, 0.0, 612.0, 792.0]}],
        "fonts": ["Songti SC"],
        "confidence": "high",
    }

    assert host_checks.validate_evidence_privacy(value) is None


def test_privacy_validation_accepts_immutable_contract_mappings():
    value = MappingProxyType(
        {"metrics": MappingProxyType({"pageCount": 1, "fonts": ("Times",)})}
    )

    assert host_checks.validate_evidence_privacy(value) is None


def test_privacy_validation_rejects_non_finite_numbers():
    with pytest.raises(ValueError, match="finite"):
        host_checks.validate_evidence_privacy({"coordinate": math.inf})
