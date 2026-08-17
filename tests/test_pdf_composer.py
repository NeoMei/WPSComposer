from __future__ import annotations

from pathlib import Path

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


@pytest.mark.parametrize("page", [0, -1, True, 2])
def test_extract_pages_rejects_invalid_one_based_page_numbers(tmp_path, page):
    src = _make_pdf(tmp_path / "src.pdf")

    with pytest.raises(ValueError, match="page index"):
        PdfComposer.extract_pages(src, [page], tmp_path / "out.pdf")


@pytest.mark.parametrize("page", [0, -1, True, 2])
def test_extract_text_rejects_invalid_one_based_page_numbers(tmp_path, page):
    src = _make_pdf(tmp_path / "src.pdf")

    with pytest.raises(ValueError, match="page index"):
        PdfComposer.extract_text(src, pages=[page])


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


@pytest.mark.parametrize(
    "operation",
    [
        lambda src, out: PdfComposer.merge([src], out, overwrite=True),
        lambda src, out: PdfComposer.extract_pages(src, [1], out, overwrite=True),
        lambda src, out: PdfComposer.rotate(src, 90, out, overwrite=True),
        lambda src, out: PdfComposer.add_text_watermark(
            src, "CONFIDENTIAL", out, overwrite=True
        ),
    ],
)
def test_single_pdf_writes_use_staging_and_preserve_existing_on_write_failure(
    tmp_path, monkeypatch, operation
):
    src = _make_pdf(tmp_path / "source.pdf")
    output = tmp_path / "approved.pdf"
    output.write_bytes(b"APPROVED-PDF")
    real_write = PdfWriter.write
    write_paths = []

    def partial_then_fail(self, stream):
        write_paths.append(Path(stream.name).resolve())
        stream.write(b"PARTIAL")
        raise OSError("disk full")

    monkeypatch.setattr(PdfWriter, "write", partial_then_fail)
    with pytest.raises(OSError, match="disk full"):
        operation(src, output)
    monkeypatch.setattr(PdfWriter, "write", real_write)

    assert output.read_bytes() == b"APPROVED-PDF"
    assert len(write_paths) == 1
    assert write_paths[0].parent == output.parent.resolve()
    assert write_paths[0] != output.resolve()
    assert not write_paths[0].exists()


def test_pdf_output_refuses_existing_file_without_explicit_overwrite(tmp_path):
    src = _make_pdf(tmp_path / "source.pdf")
    output = tmp_path / "approved.pdf"
    output.write_bytes(b"APPROVED-PDF")

    with pytest.raises(FileExistsError, match="Output already exists"):
        PdfComposer.rotate(src, 90, output)

    assert output.read_bytes() == b"APPROVED-PDF"


def test_split_stages_entire_set_before_publish_and_preserves_existing_on_failure(
    tmp_path, monkeypatch
):
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    writer.add_blank_page(width=612, height=792)
    source = tmp_path / "source.pdf"
    with source.open("wb") as stream:
        writer.write(stream)
    writer.close()
    output_dir = tmp_path / "pages"
    output_dir.mkdir()
    first = output_dir / "part-001.pdf"
    second = output_dir / "part-002.pdf"
    first.write_bytes(b"APPROVED-ONE")
    second.write_bytes(b"APPROVED-TWO")
    originals = (first.read_bytes(), second.read_bytes())
    real_write = PdfWriter.write
    writes = 0

    def fail_second_stage(self, stream):
        nonlocal writes
        writes += 1
        if writes == 2:
            stream.write(b"PARTIAL")
            raise OSError("second page failed")
        return real_write(self, stream)

    monkeypatch.setattr(PdfWriter, "write", fail_second_stage)

    with pytest.raises(OSError, match="second page failed"):
        PdfComposer.split(source, output_dir, stem="part", overwrite=True)

    assert (first.read_bytes(), second.read_bytes()) == originals
    assert not list(output_dir.glob(".wpscomposer-pdf-*"))


def test_pdf_reader_and_writer_close_when_page_copy_fails(tmp_path, monkeypatch):
    from pypdf import PdfReader

    source = _make_pdf(tmp_path / "source.pdf")
    closed = {"reader": 0, "writer": 0}
    real_reader_close = PdfReader.close
    real_writer_close = PdfWriter.close

    def close_reader(self):
        closed["reader"] += 1
        return real_reader_close(self)

    def close_writer(self):
        closed["writer"] += 1
        return real_writer_close(self)

    monkeypatch.setattr(PdfReader, "close", close_reader)
    monkeypatch.setattr(PdfWriter, "close", close_writer)
    monkeypatch.setattr(
        PdfWriter, "add_page", lambda self, page: (_ for _ in ()).throw(
            RuntimeError("page copy failed")
        )
    )

    with pytest.raises(RuntimeError, match="page copy failed"):
        PdfComposer.rotate(source, 90, tmp_path / "out.pdf")

    assert closed == {"reader": 1, "writer": 1}


def test_split_closes_reader_when_output_preflight_fails(tmp_path, monkeypatch):
    from pypdf import PdfReader

    source = _make_pdf(tmp_path / "source.pdf")
    output_dir = tmp_path / "pages"
    output_dir.mkdir()
    (output_dir / "source-001.pdf").write_bytes(b"APPROVED")
    closed = []
    real_close = PdfReader.close

    def observed_close(self):
        closed.append(True)
        return real_close(self)

    monkeypatch.setattr(PdfReader, "close", observed_close)

    with pytest.raises(FileExistsError):
        PdfComposer.split(source, output_dir)

    assert closed == [True]
