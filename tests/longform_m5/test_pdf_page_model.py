from __future__ import annotations

from pathlib import Path

from pypdf import PdfReader, PdfWriter
import pytest

from skills.WPSComposer.scripts.longform.pdf_quality import load_pdf_pages
from tests._pdf_fixture import write_minimal_pdf


def _rotated_cropped_pdf(path: Path) -> Path:
    source = write_minimal_pdf(path.with_name("source.pdf"))
    reader = PdfReader(str(source), strict=True)
    page = reader.pages[0]
    page.cropbox.lower_left = (36, 72)
    page.cropbox.upper_right = (576, 720)
    page.rotate(90)
    writer = PdfWriter()
    writer.add_page(page)
    with path.open("wb") as stream:
        writer.write(stream)
    return path


def test_load_pdf_pages_normalizes_rotation_and_crop_to_visible_points(tmp_path):
    pages = load_pdf_pages(_rotated_cropped_pdf(tmp_path / "rotated.pdf"))
    assert len(pages) == 1
    page = pages[0]
    assert page.physical_page == 1
    assert page.rotation == 90
    assert page.media_box == (0.0, 0.0, 612.0, 792.0)
    assert page.crop_box == (36.0, 72.0, 576.0, 720.0)
    assert page.width == pytest.approx(648.0)
    assert page.height == pytest.approx(540.0)
    assert page.glyph_bounds == ()


def test_load_pdf_pages_accepts_a_valid_minimal_pdf(tmp_path):
    path = write_minimal_pdf(tmp_path / "plain.pdf")
    pages = load_pdf_pages(path)
    assert pages[0].width == 612.0
    assert pages[0].height == 792.0
    assert pages[0].object_bounds == ()


def test_load_pdf_pages_rejects_corrupt_pdf_without_leaking_path(tmp_path):
    path = tmp_path / "private-customer.pdf"
    path.write_bytes(b"%PDF-1.7\ncorrupt")
    with pytest.raises(ValueError, match="PDF quality input is invalid") as exc:
        load_pdf_pages(path)
    assert str(path) not in str(exc.value)


def test_load_pdf_pages_rejects_parser_page_count_disagreement(tmp_path, monkeypatch):
    path = write_minimal_pdf(tmp_path / "plain.pdf")

    class EmptyPdf:
        pages = ()

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return None

    monkeypatch.setattr(
        "skills.WPSComposer.scripts.longform.pdf_quality.pdfplumber.open",
        lambda unused: EmptyPdf(),
    )
    with pytest.raises(ValueError, match="parsers disagree"):
        load_pdf_pages(path)
