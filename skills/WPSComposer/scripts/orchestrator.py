"""Document orchestrator — unified entry point for MD to WPS generation.

Usage::

    from orchestrator import generate

    generate("report.md", format="docx", preset="academic", output="report.docx")
    generate("slides.md", format="pptx", preset="business", output="slides.pptx")
    generate("data.md",  format="xlsx", output="data.xlsx")
    generate("report.md", format="pdf",  output="report.pdf")
"""

from __future__ import annotations

import os
from typing import Optional

from .md_parser import parse_file, parse
from .document_model import StructuredDocument
from .design_presets import get_preset, list_presets, DesignPreset


def generate(
    source: str,
    format: str = "docx",
    preset: Optional[str] = None,
    output: Optional[str] = None,
    source_is_text: bool = False,
) -> str:
    """Generate a beautifully formatted document from Markdown.

    This is the single entry point for all WPSComposer document generation.
    Parses Markdown, applies an optional design preset, and renders to the
    requested format via the appropriate WPS Composer.

    Args:
        source: Path to a .md file, or raw Markdown text if
                ``source_is_text=True``.
        format: Output format — ``"docx"``, ``"pptx"``, ``"xlsx"``, or ``"pdf"``.
        preset: Name of a design preset (``"academic"``, ``"business"``,
                ``"consultant"``, ``"tech"``) or ``None`` for defaults.
        output: Output file path.  Auto-generated from source name + format
                if omitted.
        source_is_text: Treat ``source`` as raw Markdown text instead of a
                        file path.

    Returns:
        Absolute path to the generated file.

    Raises:
        ValueError: Unknown format or preset name.
        FileNotFoundError: Source file not found.
    """
    # Validate format
    format = format.lower().strip()
    if format not in ("docx", "pptx", "xlsx", "pdf"):
        raise ValueError(
            f"Unknown format '{format}'. Use: docx, pptx, xlsx, or pdf."
        )

    # Parse Markdown
    if source_is_text:
        doc = parse(source)
        base_name = "document"
    else:
        if not os.path.isfile(source):
            raise FileNotFoundError(f"Source file not found: {source}")
        doc = parse_file(source)
        base_name = os.path.splitext(os.path.basename(source))[0]

    # Resolve preset
    design_preset = None
    if preset:
        try:
            design_preset = get_preset(preset)
        except KeyError:
            available = ", ".join(list_presets())
            raise ValueError(
                f"Unknown preset '{preset}'. Available: {available}"
            )

    # Determine output path
    if output is None:
        output = f"{base_name}.{format}"
    output = os.path.abspath(output)

    # Route to renderer
    if format == "docx":
        from .renderers.writer_renderer import render as _render
        return _render(doc, output, preset=design_preset)

    elif format == "pptx":
        from .renderers.slide_renderer import render as _render
        return _render(doc, output, preset=design_preset)

    elif format == "xlsx":
        from .renderers.sheet_renderer import render as _render
        return _render(doc, output, preset=design_preset)

    elif format == "pdf":
        from .renderers.writer_renderer import render as _writer_render
        from .writer import WriterComposer
        import tempfile

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_docx = os.path.join(tmpdir, f"{base_name}.docx")
            _writer_render(doc, tmp_docx, preset=design_preset)
            with WriterComposer() as w:
                w._doc.Close(False)
                w._doc = w._app.Documents.Open(tmp_docx)
                return w.export_pdf(output)

    raise ValueError(f"Format '{format}' not implemented.")


def list_formats() -> list:
    """Return supported output formats."""
    return ["docx", "pptx", "xlsx", "pdf"]


def list_available_presets() -> list:
    """Return available design preset names."""
    return list_presets()
