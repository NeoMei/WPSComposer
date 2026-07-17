"""WPS Office COM engine for rich-layout DOCX / PPTX / XLSX generation.

Backward-compatible re-export facade.  All composers and infrastructure
have been moved to individual modules; this module re-exports them so
existing ``from wps_engine import ...`` code continues to work.

New module layout::

    _dispatch.py   COM dispatch, ProgIDs, format constants, WPS path search
    _colors.py     Unified colour model (hex_to_rgb_long, resolve_color)
    _base.py       Enhanced BaseComposer with shared save/export lifecycle
    writer.py      WriterComposer  (WPS Writer  -> docx)
    sheet.py       SheetComposer   (WPS Sheet   -> xlsx)
    slide.py       SlideComposer   (WPS Present -> pptx)
    pdf.py         PdfComposer     (pypdf/pdfplumber -> pdf edit)

Companion modules (importable from this package)::

    design_presets.py   DesignPreset / PRESETS - 4 colour+font presets
    layout_templates.py LayoutTemplate / LAYOUTS / TALK_PRESETS
    quality_checks.py   validate_slide / review_deck / REVIEW_DIMENSIONS
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Infrastructure
# ---------------------------------------------------------------------------
from ._dispatch import (
    WRITER_PROGIDS,
    SHEET_PROGIDS,
    SLIDE_PROGIDS,
    FMT_DOCX,
    FMT_PDF_FROM_DOC,
    FMT_DOC,
    FMT_DOCM,
    FMT_DOTX,
    FMT_DOTM,
    FMT_TXT,
    FMT_HTML,
    FMT_MHTML,
    FMT_RTF,
    FMT_XML,
    FMT_ODT,
    FMT_XPS,
    FMT_XLSX,
    FMT_PDF_FROM_XLS,
    FMT_XLS,
    FMT_XLSM,
    FMT_XLSB,
    FMT_XLTX,
    FMT_XLTM,
    FMT_CSV,
    FMT_TSV,
    FMT_ODS,
    FMT_PPTX,
    FMT_PDF_FROM_PPT,
    FMT_PPT,
    FMT_PPSX,
    FMT_PPTM,
    FMT_POTX,
    FMT_POTM,
    FMT_PPSM,
    FMT_ODP,
    WPS_SEARCH_PATHS,
    WPSUnavailable,
    find_wps_executable,
)

# ---------------------------------------------------------------------------
# Colour utilities (keep old name for backward compat)
# ---------------------------------------------------------------------------
from ._colors import hex_to_rgb_long as _hex_to_rgb_long
from ._colors import hex_to_rgb_long, resolve_color, resolve_color_long

# ---------------------------------------------------------------------------
# Composers
# ---------------------------------------------------------------------------
from .writer import WriterComposer
from .sheet import SheetComposer
from .slide import SlideComposer
from .pdf import PdfComposer
# High-level API (MD -> document generation)
from .orchestrator import generate, list_formats, list_available_presets

# Document model + parser
from .document_model import (
    StructuredDocument, Section, Paragraph, Span,
    ListBlock, TableBlock, CodeBlock, ImageBlock, BlockQuote,
)
from .md_parser import parse, parse_file
from .document_api import (
    open_document, attach_active, inspect, edit, apply_patches,
    snapshot_json, supported_formats,
)


__all__ = [
    # Composers
    "WriterComposer",
    "SheetComposer",
    "SlideComposer",
    "PdfComposer",
    # Infrastructure
    "WPSUnavailable",
    "find_wps_executable",
    "WPS_SEARCH_PATHS",
    # ProgIDs
    "WRITER_PROGIDS",
    "SHEET_PROGIDS",
    "SLIDE_PROGIDS",
    # Format constants
    "FMT_DOCX",
    "FMT_PDF_FROM_DOC",
    "FMT_DOC",
    "FMT_DOCM",
    "FMT_DOTX",
    "FMT_DOTM",
    "FMT_TXT",
    "FMT_HTML",
    "FMT_MHTML",
    "FMT_RTF",
    "FMT_XML",
    "FMT_ODT",
    "FMT_XPS",
    "FMT_XLSX",
    "FMT_PDF_FROM_XLS",
    "FMT_XLS",
    "FMT_XLSM",
    "FMT_XLSB",
    "FMT_XLTX",
    "FMT_XLTM",
    "FMT_CSV",
    "FMT_TSV",
    "FMT_ODS",
    "FMT_PPTX",
    "FMT_PDF_FROM_PPT",
    "FMT_PPT",
    "FMT_PPSX",
    "FMT_PPTM",
    "FMT_POTX",
    "FMT_POTM",
    "FMT_PPSM",
    "FMT_ODP",
    # Colour utilities
    "_hex_to_rgb_long",
    "hex_to_rgb_long",
    "resolve_color",
    "resolve_color_long",
    # High-level generation API
    "generate",
    "list_formats",
    "list_available_presets",
    # Document model and parser
    "StructuredDocument",
    "Section",
    "Paragraph",
    "Span",
    "ListBlock",
    "TableBlock",
    "CodeBlock",
    "ImageBlock",
    "BlockQuote",
    "parse",
    "parse_file",
    # Existing/current document API
    "open_document",
    "attach_active",
    "inspect",
    "edit",
    "apply_patches",
    "snapshot_json",
    "supported_formats",
]
