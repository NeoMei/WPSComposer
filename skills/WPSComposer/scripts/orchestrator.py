"""Document orchestrator — unified entry point for MD to WPS generation.

Usage::

    from orchestrator import generate

    generate("report.md", format="docx", preset="academic", output="report.docx")
    generate("slides.md", format="pptx", preset="business", output="slides.pptx")
    generate("data.md",  format="xlsx", output="data.xlsx")
    generate("report.md", format="pdf",  output="report.pdf")

With plugins::

    generate("notes.md", format="pdf", plugins=["excalidraw"])
"""

from __future__ import annotations

import os
from pathlib import Path
import sys
from typing import List, Optional

from .md_parser import parse_file, parse
from .document_model import StructuredDocument
from .design_presets import get_preset, list_presets, DesignPreset
from .macos_probe.generation import GenerationError, generate_macos
from .plugins import run_plugins


def generate(
    source: str,
    format: str = "docx",
    preset: Optional[str] = None,
    output: Optional[str] = None,
    source_is_text: bool = False,
    plugins: Optional[List[str]] = None,
    timeout: float = 600,
    overwrite: bool = False,
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
        plugins: List of plugin names to run before parsing.
                 Available: ``"excalidraw"`` (renders .excalidraw.md to PNG).
        timeout: Timeout in seconds for WPS generation (default: 600).
        overwrite: If True, overwrite existing output file.

    Returns:
        Absolute path to the generated file.

    Raises:
        ValueError: Unknown format or preset name.
        FileNotFoundError: Source file not found.
        FileExistsError: Output file already exists (unless overwrite=True).
    """
    # Validate format
    format = format.lower().strip()
    if format not in ("docx", "pptx", "xlsx", "pdf"):
        raise ValueError(
            f"Unknown format '{format}'. Use: docx, pptx, xlsx, or pdf."
        )

    # Parse Markdown (with optional plugin preprocessing)
    if source_is_text:
        content = source
        base_dir = os.getcwd()
        base_name = "document"
    else:
        if not os.path.isfile(source):
            raise FileNotFoundError(f"Source file not found: {source}")
        with open(source, "r", encoding="utf-8") as f:
            content = f.read()
        base_dir = os.path.dirname(os.path.abspath(source))
        base_name = os.path.splitext(os.path.basename(source))[0]

    # Run plugins before parsing
    if plugins:
        content = run_plugins(content, base_dir, plugins)

    # Parse the (possibly modified) content
    doc = parse(content, base_dir=base_dir)

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
    output_path = Path(output).expanduser().resolve()
    if output_path.suffix.lower() != f".{format}":
        raise ValueError(
            f"Output extension must match requested format '.{format}'."
        )
    if output_path.exists() and not overwrite:
        raise FileExistsError(
            f"Output already exists: {output_path}. "
            f"Use overwrite=True to replace it."
        )
    if output_path.exists() and overwrite:
        output_path.unlink()
    output = str(output_path)

    # Route to renderer
    if sys.platform == "darwin":
        return str(generate_macos(
            doc, format, output_path, design_preset, timeout=timeout
        ))
    if sys.platform != "win32":
        component = {
            "docx": "writer",
            "pdf": "writer",
            "xlsx": "spreadsheet",
            "pptx": "presentation",
        }[format]
        raise GenerationError(
            code="MACOS_CAPABILITY_UNAVAILABLE",
            output=output,
            component=component,
            backend="unsupported-platform",
            message=f"WPS generation is unavailable on platform {sys.platform}",
        )
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
