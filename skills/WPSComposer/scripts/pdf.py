"""PdfComposer - edit existing PDFs via pypdf + pdfplumber.

Merge, split, rotate, extract text, and watermark.
No COM dependency - cross-platform pure Python.
"""

from __future__ import annotations

import os
from pathlib import Path
import tempfile

from .artifact_transport import (
    publish_artifact,
    publish_artifact_group,
    validate_pdf,
)


def _pdf_abs(path):
    return os.path.abspath(path)


def _validate_page_indices(page_indices, page_count):
    indices = list(page_indices)
    for index in indices:
        if (
            isinstance(index, bool)
            or not isinstance(index, int)
            or index < 1
            or index > page_count
        ):
            raise ValueError(
                f"page index must be an integer from 1 to {page_count}: {index!r}"
            )
    return indices


def _pdf_target(path, overwrite):
    target = Path(path).expanduser().resolve()
    if target.suffix.lower() != ".pdf":
        raise ValueError("PDF output path must use the '.pdf' suffix")
    if target.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    return target


def _write_pdf_stage(writer, target):
    descriptor, temporary_name = tempfile.mkstemp(
        dir=target.parent,
        prefix=".wpscomposer-pdf-",
        suffix=".pdf",
    )
    os.close(descriptor)
    staged = Path(temporary_name)
    try:
        with staged.open("wb") as stream:
            writer.write(stream)
        validate_pdf(staged)
        return staged
    except BaseException:
        staged.unlink(missing_ok=True)
        raise
    finally:
        writer.close()


def _write_and_publish_pdf(writer, output_path, overwrite):
    target = _pdf_target(output_path, overwrite)
    staged = _write_pdf_stage(writer, target)
    try:
        return str(publish_artifact(
            staged,
            target,
            overwrite=overwrite,
            validator=validate_pdf,
        ))
    finally:
        staged.unlink(missing_ok=True)


# ===========================================================================
# WPS exposes KPDF.Application but its Dispatch blocks headless (GUI RPC),
# so PDF *editing* uses pypdf/pdfplumber instead. PDF *generation* still goes
# through the three WPS composers above (export_pdf).

class PdfComposer:
    """Edit existing PDFs: merge, split, rotate, extract text/watermark.

    All methods are static-ish helpers operating on file paths; no COM host
    needed, cross-platform.
    """

    @staticmethod
    def merge(input_paths, output_path, *, overwrite=False):
        """Concatenate multiple PDFs into one. Returns output path."""
        from pypdf import PdfWriter
        _pdf_target(output_path, overwrite)
        w = PdfWriter()
        try:
            for p in input_paths:
                w.append(_pdf_abs(p))
        except BaseException:
            w.close()
            raise
        return _write_and_publish_pdf(w, output_path, overwrite)

    @staticmethod
    def split(input_path, output_dir, stem=None, *, overwrite=False):
        """Split a PDF into one-page PDFs. Returns list of output paths."""
        from pypdf import PdfReader, PdfWriter
        import os
        src = _pdf_abs(input_path)
        odir = _pdf_abs(output_dir)
        os.makedirs(odir, exist_ok=True)
        base = stem or os.path.splitext(os.path.basename(src))[0]
        reader = PdfReader(src)
        staged = []
        try:
            targets = [
                _pdf_target(
                    os.path.join(odir, f"{base}-{i:03d}.pdf"), overwrite
                )
                for i in range(1, len(reader.pages) + 1)
            ]
            for page, target in zip(reader.pages, targets):
                writer = PdfWriter()
                try:
                    writer.add_page(page)
                except BaseException:
                    writer.close()
                    raise
                staged.append(_write_pdf_stage(writer, target))
            published = publish_artifact_group([
                (stage, target, overwrite, validate_pdf)
                for stage, target in zip(staged, targets)
            ])
            return [str(path) for path in published]
        finally:
            for stage in staged:
                stage.unlink(missing_ok=True)
            reader.close()

    @staticmethod
    def extract_pages(input_path, page_indices, output_path, *, overwrite=False):
        """Keep only 1-based page indices into a new PDF. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        _pdf_target(output_path, overwrite)
        reader = PdfReader(_pdf_abs(input_path))
        try:
            page_indices = _validate_page_indices(page_indices, len(reader.pages))
            w = PdfWriter()
            try:
                for idx in page_indices:
                    w.add_page(reader.pages[idx - 1])
            except BaseException:
                w.close()
                raise
            return _write_and_publish_pdf(w, output_path, overwrite)
        finally:
            reader.close()

    @staticmethod
    def rotate(input_path, angle, output_path, *, overwrite=False):
        """Rotate all pages. angle in {90, 180, 270}. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        _pdf_target(output_path, overwrite)
        reader = PdfReader(_pdf_abs(input_path))
        try:
            w = PdfWriter()
            try:
                for page in reader.pages:
                    page.rotate(angle)
                    w.add_page(page)
            except BaseException:
                w.close()
                raise
            return _write_and_publish_pdf(w, output_path, overwrite)
        finally:
            reader.close()

    @staticmethod
    def extract_text(input_path, pages=None):
        """Extract text. pages=None for all, else 1-based list. Returns str."""
        import pdfplumber
        with pdfplumber.open(_pdf_abs(input_path)) as pdf:
            idxs = (
                _validate_page_indices(pages, len(pdf.pages))
                if pages is not None
                else range(1, len(pdf.pages) + 1)
            )
            chunks = []
            for i in idxs:
                txt = pdf.pages[i - 1].extract_text() or ""
                chunks.append(txt)
            return "\n".join(chunks)

    @staticmethod
    def page_count(input_path):
        """Return number of pages."""
        from pypdf import PdfReader
        reader = PdfReader(_pdf_abs(input_path))
        try:
            return len(reader.pages)
        finally:
            reader.close()

    @staticmethod
    def _watermark_page(text, pw, ph, fontsize, opacity, angle):
        from pypdf import PdfReader
        from reportlab.pdfgen import canvas
        from reportlab.lib.colors import Color
        import io
        font = "Helvetica"
        if any(ord(ch) > 0x2E7F for ch in text):
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.cidfonts import UnicodeCIDFont
            try:
                pdfmetrics.getFont("STSong-Light")
            except KeyError:
                pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
            font = "STSong-Light"
        wm_buf = io.BytesIO()
        c = canvas.Canvas(wm_buf, pagesize=(pw, ph))
        c.saveState()
        c.setFont(font, fontsize)
        try:
            c.setFillColor(Color(0.5, 0.5, 0.5, alpha=opacity))
        except Exception:
            c.setFillGray(0.5)
        c.translate(pw / 2, ph / 2)
        c.rotate(angle)
        c.drawCentredString(0, 0, text)
        c.restoreState()
        c.showPage()
        c.save()
        wm_buf.seek(0)
        return PdfReader(wm_buf).pages[0]

    @staticmethod
    def add_text_watermark(input_path, text, output_path,
                            fontsize=50, opacity=0.15, angle=45,
                            *, overwrite=False):
        """Stamp a diagonal text watermark on every page. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        src = _pdf_abs(input_path)
        _pdf_target(output_path, overwrite)
        reader = PdfReader(src)
        try:
            w = PdfWriter()
            try:
                for page in reader.pages:
                    pw = float(page.mediabox.width)
                    ph = float(page.mediabox.height)
                    watermark = PdfComposer._watermark_page(
                        text, pw, ph, fontsize, opacity, angle
                    )
                    w.add_page(page)
                    w.pages[-1].merge_page(watermark)
            except BaseException:
                w.close()
                raise
            return _write_and_publish_pdf(w, output_path, overwrite)
        finally:
            reader.close()
