"""PdfComposer - edit existing PDFs via pypdf + pdfplumber.

Merge, split, rotate, extract text, and watermark.
No COM dependency - cross-platform pure Python.
"""

from __future__ import annotations

import os


def _pdf_abs(path):
    return os.path.abspath(path)


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
    def merge(input_paths, output_path):
        """Concatenate multiple PDFs into one. Returns output path."""
        from pypdf import PdfWriter
        w = PdfWriter()
        for p in input_paths:
            w.append(_pdf_abs(p))
        out = _pdf_abs(output_path)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out

    @staticmethod
    def split(input_path, output_dir, stem=None):
        """Split a PDF into one-page PDFs. Returns list of output paths."""
        from pypdf import PdfReader, PdfWriter
        import os
        src = _pdf_abs(input_path)
        odir = _pdf_abs(output_dir)
        os.makedirs(odir, exist_ok=True)
        base = stem or os.path.splitext(os.path.basename(src))[0]
        reader = PdfReader(src)
        out = []
        for i, page in enumerate(reader.pages, start=1):
            w = PdfWriter()
            w.add_page(page)
            p = os.path.join(odir, f"{base}-{i:03d}.pdf")
            with open(p, "wb") as fh:
                w.write(fh)
            w.close()
            out.append(p)
        return out

    @staticmethod
    def extract_pages(input_path, page_indices, output_path):
        """Keep only 1-based page indices into a new PDF. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(_pdf_abs(input_path))
        w = PdfWriter()
        for idx in page_indices:
            w.add_page(reader.pages[idx - 1])
        out = _pdf_abs(output_path)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out

    @staticmethod
    def rotate(input_path, angle, output_path):
        """Rotate all pages. angle in {90, 180, 270}. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(_pdf_abs(input_path))
        w = PdfWriter()
        for page in reader.pages:
            page.rotate(angle)
            w.add_page(page)
        out = _pdf_abs(output_path)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out

    @staticmethod
    def extract_text(input_path, pages=None):
        """Extract text. pages=None for all, else 1-based list. Returns str."""
        import pdfplumber
        with pdfplumber.open(_pdf_abs(input_path)) as pdf:
            idxs = pages if pages is not None else range(1, len(pdf.pages) + 1)
            chunks = []
            for i in idxs:
                txt = pdf.pages[i - 1].extract_text() or ""
                chunks.append(txt)
            return "\n".join(chunks)

    @staticmethod
    def page_count(input_path):
        """Return number of pages."""
        from pypdf import PdfReader
        return len(PdfReader(_pdf_abs(input_path)).pages)

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
                            fontsize=50, opacity=0.15, angle=45):
        """Stamp a diagonal text watermark on every page. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        src = _pdf_abs(input_path)
        out = _pdf_abs(output_path)
        reader = PdfReader(src)
        w = PdfWriter()
        for page in reader.pages:
            pw = float(page.mediabox.width)
            ph = float(page.mediabox.height)
            page.merge_page(
                PdfComposer._watermark_page(text, pw, ph, fontsize, opacity, angle)
            )
            w.add_page(page)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out
