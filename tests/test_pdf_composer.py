from __future__ import annotations

import pytest

pypdf = pytest.importorskip("pypdf")
pytest.importorskip("reportlab")
pytest.importorskip("pdfplumber")

from pypdf import PdfWriter

from skills.WPSComposer.scripts.pdf import PdfComposer


def _make_pdf(path, width=612, height=792):
    w = PdfWriter()
    w.add_blank_page(width=width, height=height)
    with open(path, "wb") as fh:
        w.write(fh)
    return str(path)


def test_merge_and_page_count(tmp_path):
    a = _make_pdf(tmp_path / "a.pdf")
    b = _make_pdf(tmp_path / "b.pdf")
    out = PdfComposer.merge([a, b], tmp_path / "out.pdf")
    assert PdfComposer.page_count(out) == 2


def test_split_rotate_extract_pages(tmp_path):
    src = _make_pdf(tmp_path / "src.pdf")
    pages = PdfComposer.split(src, tmp_path / "pages")
    assert len(pages) == 1
    merged = PdfComposer.extract_pages(src, [1], tmp_path / "one.pdf")
    assert PdfComposer.page_count(merged) == 1
    rotated = PdfComposer.rotate(src, 90, tmp_path / "rot.pdf")
    assert PdfComposer.page_count(rotated) == 1


def test_watermark_mixed_page_sizes(tmp_path):
    w = PdfWriter()
    w.add_blank_page(width=612, height=792)
    w.add_blank_page(width=1190, height=842)
    src = tmp_path / "mixed.pdf"
    with open(src, "wb") as fh:
        w.write(fh)
    out = PdfComposer.add_text_watermark(src, "机密", tmp_path / "wm.pdf")
    assert PdfComposer.page_count(out) == 2
    text = PdfComposer.extract_text(out)
    assert isinstance(text, str)
