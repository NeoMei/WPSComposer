from __future__ import annotations

from pathlib import Path

import pytest

from skills.WPSComposer.scripts import conversion
from skills.WPSComposer.scripts.conversion import (
    ConversionError,
    convert_to_pdf,
)


def _write_pdf(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"%PDF-1.7\n" + b"x" * 2048)
    return path


@pytest.mark.parametrize(
    ("suffix", "component"),
    [
        ("doc", "writer"),
        ("DOCX", "writer"),
        ("xls", "spreadsheet"),
        ("XLSX", "spreadsheet"),
        ("ppt", "presentation"),
        ("PPTX", "presentation"),
    ],
)
def test_convert_routes_all_supported_suffixes(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    suffix: str,
    component: str,
):
    source = tmp_path / f"input.{suffix}"
    source.write_bytes(b"source")
    observed = []

    def backend(request):
        observed.append(request)
        return _write_pdf(request.output)

    monkeypatch.setattr(
        conversion, "_select_backend", lambda request: ("test-backend", backend)
    )

    result = convert_to_pdf(str(source))

    assert result == str((tmp_path / "input.pdf").resolve())
    assert observed[0].source == source.resolve()
    assert observed[0].component == component
    assert observed[0].overwrite is False


def test_convert_accepts_explicit_pdf_and_overwrite(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "input.docx"
    source.write_bytes(b"source")
    output = tmp_path / "exports" / "renamed.PDF"
    output.parent.mkdir()
    _write_pdf(output)
    observed = []

    def backend(request):
        observed.append(request)
        return _write_pdf(request.output)

    monkeypatch.setattr(
        conversion, "_select_backend", lambda request: ("test-backend", backend)
    )

    assert convert_to_pdf(str(source), str(output), overwrite=True) == str(
        output.resolve()
    )
    assert observed[0].overwrite is True


def test_convert_rejects_missing_source_before_backend(tmp_path: Path):
    with pytest.raises(FileNotFoundError, match="Source file not found"):
        convert_to_pdf(str(tmp_path / "missing.docx"))


@pytest.mark.parametrize("suffix", ["docm", "xlsm", "pptm", "txt", "pdf"])
def test_convert_rejects_unsupported_source_suffix(tmp_path: Path, suffix: str):
    source = tmp_path / f"input.{suffix}"
    source.write_bytes(b"source")
    with pytest.raises(ValueError, match="Unsupported source format"):
        convert_to_pdf(str(source))


def test_convert_requires_pdf_output_suffix(tmp_path: Path):
    source = tmp_path / "input.docx"
    source.write_bytes(b"source")
    with pytest.raises(ValueError, match="Output path must end in .pdf"):
        convert_to_pdf(str(source), str(tmp_path / "output.docx"))


def test_convert_refuses_existing_output_without_overwrite(tmp_path: Path):
    source = tmp_path / "input.docx"
    source.write_bytes(b"source")
    output = _write_pdf(tmp_path / "input.pdf")
    with pytest.raises(FileExistsError, match="Output already exists"):
        convert_to_pdf(str(source), str(output))


def test_conversion_error_exposes_stable_fields(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "input.docx"
    source.write_bytes(b"source")

    def failed_backend(request):
        raise RuntimeError("boom")

    monkeypatch.setattr(
        conversion,
        "_select_backend",
        lambda request: ("mac-wps-jsapi", failed_backend),
    )

    with pytest.raises(ConversionError) as caught:
        convert_to_pdf(str(source))

    error = caught.value
    assert error.code == "CONVERSION_FAILED"
    assert error.source == str(source.resolve())
    assert error.component == "writer"
    assert error.backend == "mac-wps-jsapi"
    assert error.message == "boom"


def test_conversion_preserves_backend_conversion_error(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "input.xlsx"
    source.write_bytes(b"source")
    expected = ConversionError(
        code="PASSWORD_REQUIRED",
        source=str(source.resolve()),
        component="spreadsheet",
        backend="mac-wps-jsapi",
        message="password required",
    )

    def failed_backend(request):
        raise expected

    monkeypatch.setattr(
        conversion,
        "_select_backend",
        lambda request: ("mac-wps-jsapi", failed_backend),
    )

    with pytest.raises(ConversionError) as caught:
        convert_to_pdf(str(source))
    assert caught.value is expected


def test_conversion_preserves_public_path_errors_from_backend(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = tmp_path / "input.pptx"
    source.write_bytes(b"source")

    def raced_backend(request):
        raise FileExistsError(f"Output already exists: {request.output}")

    monkeypatch.setattr(
        conversion,
        "_select_backend",
        lambda request: ("test-backend", raced_backend),
    )

    with pytest.raises(FileExistsError, match="Output already exists"):
        convert_to_pdf(str(source))
