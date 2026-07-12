"""PdfComposer - edit existing PDFs via pypdf + pdfplumber.

Merge, split, rotate, extract text, and watermark.
No COM dependency - cross-platform pure Python.
"""

from __future__ import annotations

from ._dispatch import _abs


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
            w.append(_abs(p))
        out = _abs(output_path)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out

    @staticmethod
    def split(input_path, output_dir, stem=None):
        """Split a PDF into one-page PDFs. Returns list of output paths."""
        from pypdf import PdfReader, PdfWriter
        import os
        src = _abs(input_path)
        odir = _abs(output_dir)
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
        reader = PdfReader(_abs(input_path))
        w = PdfWriter()
        for idx in page_indices:
            w.add_page(reader.pages[idx - 1])
        out = _abs(output_path)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out

    @staticmethod
    def rotate(input_path, angle, output_path):
        """Rotate all pages. angle in {90, 180, 270}. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(_abs(input_path))
        w = PdfWriter()
        for page in reader.pages:
            page.rotate(angle)
            w.add_page(page)
        out = _abs(output_path)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out

    @staticmethod
    def extract_text(input_path, pages=None):
        """Extract text. pages=None for all, else 1-based list. Returns str."""
        import pdfplumber
        with pdfplumber.open(_abs(input_path)) as pdf:
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
        return len(PdfReader(_abs(input_path)).pages)

    @staticmethod
    def add_text_watermark(input_path, text, output_path,
                            fontsize=50, opacity=0.15, angle=45):
        """Stamp a diagonal text watermark on every page. Returns output path."""
        from pypdf import PdfReader, PdfWriter
        from reportlab.pdfgen import canvas
        from reportlab.lib.colors import Color
        import io
        src = _abs(input_path)
        out = _abs(output_path)
        reader = PdfReader(src)
        first = reader.pages[0]
        pw = float(first.mediabox.width)
        ph = float(first.mediabox.height)
        wm_buf = io.BytesIO()
        c = canvas.Canvas(wm_buf, pagesize=(pw, ph))
        c.saveState()
        c.setFont("Helvetica", fontsize)
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
        wm_page = PdfReader(wm_buf).pages[0]
        w = PdfWriter()
        for page in reader.pages:
            page.merge_page(wm_page)
            w.add_page(page)
        with open(out, "wb") as fh:
            w.write(fh)
        w.close()
        return out
