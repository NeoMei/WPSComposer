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
import warnings

from .md_parser import parse_file, parse
from .document_model import StructuredDocument
from .design_presets import get_preset, list_presets, DesignPreset
from .macos_probe.generation import GenerationError, generate_macos
from .plugins import run_plugins
from .artifact_transport import (
    publish_artifact,
    validate_office_package,
    validate_pdf,
)
from .heading_numbering import detect_numbering_scheme
from .presentation import present_artifact, validate_open_result
from .office_engines import resolve_engine, validate_engine, validate_timeout, EngineUnavailableError, com_engine


def _generate_longform_outcome(build, format_name, output, timeout, overwrite, *, engine="wps"):
    """Private indirection keeps platform runtime imports lazy and testable."""
    from .longform.platform_runtime import generate_longform

    return generate_longform(
        build,
        format_name=format_name,
        output=output,
        timeout=timeout,
        overwrite=overwrite,
        engine=engine,
    )


def _return_artifact(path: Path, *, open_result: bool, engine: str = "wps") -> str:
    artifact = Path(path).expanduser().resolve()
    if open_result:
        try:
            if artifact.suffix.lower() in {".docx", ".xlsx", ".pptx"}:
                present_artifact(artifact, engine=engine)
            else:
                present_artifact(artifact)
        except Exception as exc:
            warnings.warn(
                f"Final artifact was published but could not be opened: {exc}",
                RuntimeWarning,
                stacklevel=2,
            )
    return str(artifact)


def generate(
    source: str,
    format: str = "docx",
    preset: Optional[str] = None,
    output: Optional[str] = None,
    source_is_text: bool = False,
    plugins: Optional[List[str]] = None,
    timeout: float = 600,
    overwrite: bool = False,
    *,
    open_result: bool = False,
    engine: str = "wps",
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
        timeout: Positive finite generation timeout in seconds (default: 600).
        overwrite: If True, overwrite existing output file.
        open_result: If True, ask the desktop default application to open the
                     finalized artifact after generation cleanup completes.
                     Native MS Word DOCX results open explicitly in Word.
        engine: Native engine: "wps" (default), "msoffice" (Word DOCX/PDF),
                or "auto" (installed WPS first, then Word; pinned per task).

    Returns:
        Absolute path to the generated file.

    Raises:
        ValueError: Unknown format or preset name.
        FileNotFoundError: Source file not found.
        FileExistsError: Output file already exists (unless overwrite=True).
    """
    validate_open_result(open_result)
    validate_engine(engine)
    validate_timeout(timeout)

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

    selected_engine = resolve_engine(engine, {"docx": "writer", "pdf": "writer", "xlsx": "spreadsheet", "pptx": "presentation"}[format])

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
    output = str(output_path)

    # DOCX/PDF now default to the protocol-v2 long-form engine.  The legacy
    # route remains available only through an explicit frontmatter request.
    if format in {"docx", "pdf"}:
        from .longform.pipeline import build_longform_generation

        longform_build = build_longform_generation(
            content,
            base_dir=base_dir,
            design_preset=preset,
        )
        if selected_engine == "msoffice" and longform_build.semantic.config.layout_engine == "legacy":
            raise EngineUnavailableError("MS Office does not support the legacy layout route")
        if longform_build.semantic.config.layout_engine != "legacy":
            if sys.platform not in {"darwin", "win32"}:
                raise GenerationError(
                    code="MACOS_CAPABILITY_UNAVAILABLE",
                    output=output,
                    component="writer",
                    backend="unsupported-platform",
                    message=(
                        "WPS long-form generation is unavailable on platform "
                        f"{sys.platform}"
                    ),
                )
            options = {"engine": selected_engine} if selected_engine != "wps" else {}
            outcome = _generate_longform_outcome(
                longform_build, format, output_path, timeout, overwrite, **options
            )
            return _return_artifact(Path(outcome.path), open_result=open_result, engine=selected_engine)

    if selected_engine == "msoffice":
        if sys.platform == "darwin":
            from .msoffice.macos_office_runtime import generate as generate_office
        elif sys.platform == "win32":
            from .msoffice.windows_office_runtime import generate as generate_office
        else:
            raise EngineUnavailableError("Native Microsoft Office requires Windows or macOS")
        result = generate_office(doc, format, output_path, design_preset,
                                 timeout=timeout, overwrite=overwrite)
        return _return_artifact(result, open_result=open_result, engine=selected_engine)

    # Route to renderer
    if sys.platform == "darwin":
        result = generate_macos(
            doc,
            format,
            output_path,
            design_preset,
            timeout=timeout,
            overwrite=overwrite,
        )
    elif sys.platform != "win32":
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
    else:
        import tempfile

        output_path.parent.mkdir(parents=True, exist_ok=True)
        with com_engine("wps"), tempfile.TemporaryDirectory(
            dir=output_path.parent, prefix=".wpscomposer-generate-"
        ) as tmpdir:
            staged = Path(tmpdir) / f"artifact.{format}"
            if format == "docx":
                from .renderers.writer_renderer import render as _render

                _render(doc, str(staged), preset=design_preset)
            elif format == "pptx":
                from .renderers.slide_renderer import render as _render

                _render(doc, str(staged), preset=design_preset)
            elif format == "xlsx":
                from .renderers.sheet_renderer import render as _render

                _render(doc, str(staged), preset=design_preset)
            else:
                from .renderers.writer_renderer import render as _writer_render
                from .writer import WriterComposer

                tmp_docx = os.path.join(tmpdir, f"{base_name}.docx")
                _writer_render(doc, tmp_docx, preset=design_preset)
                with WriterComposer() as w:
                    w._doc.Close(False)
                    w._doc = w._app.Documents.Open(tmp_docx)
                    w.export_pdf(str(staged))

            if format == "docx" and detect_numbering_scheme(doc.sections) == "chinese":
                from .numbering_native import apply_native_numbering

                apply_native_numbering(staged)
            validator = (
                validate_pdf
                if format == "pdf"
                else lambda path: validate_office_package(path, format)
            )
            result = publish_artifact(
                staged, output_path, overwrite=overwrite, validator=validator
            )

    # Convert plain-text heading numbers to native Word/WPS multi-level
    # numbering so renumbering stays automatic when headings change.
    if (
        sys.platform == "darwin"
        and format == "docx"
        and output_path.exists()
        and detect_numbering_scheme(doc.sections) == "chinese"
    ):
        try:
            from .numbering_native import apply_native_numbering

            apply_native_numbering(output_path)
        except Exception as exc:  # never block a successful generation
            import warnings

            warnings.warn(
                f"Native heading numbering could not be applied: {exc}"
            )
    return _return_artifact(output_path, open_result=open_result)


def list_formats() -> list:
    """Return supported output formats."""
    return ["docx", "pptx", "xlsx", "pdf"]


def list_available_presets() -> list:
    """Return available design preset names."""
    return list_presets()
