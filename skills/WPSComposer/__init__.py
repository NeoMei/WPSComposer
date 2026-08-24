"""Stable Python API for WPSComposer."""

from __future__ import annotations

from typing import Any

# Lazily re-export public names so that importing the package does not
# eagerly load WPS/executor modules.  PEP 562 module __getattr__ is used
# to defer imports until an attribute is actually accessed.

_WPS_ENGINE_NAMES: dict[str, str] = {
    # _dispatch
    "WRITER_PROGIDS": "skills.WPSComposer.scripts._dispatch",
    "SHEET_PROGIDS": "skills.WPSComposer.scripts._dispatch",
    "SLIDE_PROGIDS": "skills.WPSComposer.scripts._dispatch",
    "FMT_DOCX": "skills.WPSComposer.scripts._dispatch",
    "FMT_PDF_FROM_DOC": "skills.WPSComposer.scripts._dispatch",
    "FMT_DOC": "skills.WPSComposer.scripts._dispatch",
    "FMT_DOCM": "skills.WPSComposer.scripts._dispatch",
    "FMT_DOTX": "skills.WPSComposer.scripts._dispatch",
    "FMT_DOTM": "skills.WPSComposer.scripts._dispatch",
    "FMT_TXT": "skills.WPSComposer.scripts._dispatch",
    "FMT_HTML": "skills.WPSComposer.scripts._dispatch",
    "FMT_MHTML": "skills.WPSComposer.scripts._dispatch",
    "FMT_RTF": "skills.WPSComposer.scripts._dispatch",
    "FMT_XML": "skills.WPSComposer.scripts._dispatch",
    "FMT_ODT": "skills.WPSComposer.scripts._dispatch",
    "FMT_XPS": "skills.WPSComposer.scripts._dispatch",
    "FMT_XLSX": "skills.WPSComposer.scripts._dispatch",
    "FMT_PDF_FROM_XLS": "skills.WPSComposer.scripts._dispatch",
    "FMT_XLS": "skills.WPSComposer.scripts._dispatch",
    "FMT_XLSM": "skills.WPSComposer.scripts._dispatch",
    "FMT_XLSB": "skills.WPSComposer.scripts._dispatch",
    "FMT_XLTX": "skills.WPSComposer.scripts._dispatch",
    "FMT_XLTM": "skills.WPSComposer.scripts._dispatch",
    "FMT_CSV": "skills.WPSComposer.scripts._dispatch",
    "FMT_TSV": "skills.WPSComposer.scripts._dispatch",
    "FMT_ODS": "skills.WPSComposer.scripts._dispatch",
    "FMT_PPTX": "skills.WPSComposer.scripts._dispatch",
    "FMT_PDF_FROM_PPT": "skills.WPSComposer.scripts._dispatch",
    "FMT_PPT": "skills.WPSComposer.scripts._dispatch",
    "FMT_PPSX": "skills.WPSComposer.scripts._dispatch",
    "FMT_PPTM": "skills.WPSComposer.scripts._dispatch",
    "FMT_POTX": "skills.WPSComposer.scripts._dispatch",
    "FMT_POTM": "skills.WPSComposer.scripts._dispatch",
    "FMT_PPSM": "skills.WPSComposer.scripts._dispatch",
    "FMT_ODP": "skills.WPSComposer.scripts._dispatch",
    "WPSUnavailable": "skills.WPSComposer.scripts._dispatch",
    "find_wps_executable": "skills.WPSComposer.scripts._dispatch",
    "WPS_SEARCH_PATHS": "skills.WPSComposer.scripts._dispatch",
    # _colors
    "hex_to_rgb_long": "skills.WPSComposer.scripts._colors",
    "resolve_color": "skills.WPSComposer.scripts._colors",
    "resolve_color_long": "skills.WPSComposer.scripts._colors",
    # Composers
    "WriterComposer": "skills.WPSComposer.scripts.writer",
    "SheetComposer": "skills.WPSComposer.scripts.sheet",
    "SlideComposer": "skills.WPSComposer.scripts.slide",
    "PdfComposer": "skills.WPSComposer.scripts.pdf",
    # High-level API
    "generate": "skills.WPSComposer.scripts.orchestrator",
    "list_formats": "skills.WPSComposer.scripts.orchestrator",
    "list_available_presets": "skills.WPSComposer.scripts.orchestrator",
    "ConversionError": "skills.WPSComposer.scripts.conversion",
    "convert_to_pdf": "skills.WPSComposer.scripts.conversion",
    "GenerationError": "skills.WPSComposer.scripts.macos_probe.generation",
    # Document model + parser
    "StructuredDocument": "skills.WPSComposer.scripts.document_model",
    "Section": "skills.WPSComposer.scripts.document_model",
    "Paragraph": "skills.WPSComposer.scripts.document_model",
    "Span": "skills.WPSComposer.scripts.document_model",
    "ListBlock": "skills.WPSComposer.scripts.document_model",
    "TableBlock": "skills.WPSComposer.scripts.document_model",
    "CodeBlock": "skills.WPSComposer.scripts.document_model",
    "ImageBlock": "skills.WPSComposer.scripts.document_model",
    "BlockQuote": "skills.WPSComposer.scripts.document_model",
    "HorizontalRule": "skills.WPSComposer.scripts.document_model",
    "TaskList": "skills.WPSComposer.scripts.document_model",
    "parse": "skills.WPSComposer.scripts.md_parser",
    "parse_file": "skills.WPSComposer.scripts.md_parser",
    # Document API
    "open_document": "skills.WPSComposer.scripts.document_api",
    "attach_active": "skills.WPSComposer.scripts.document_api",
    "inspect": "skills.WPSComposer.scripts.document_api",
    "edit": "skills.WPSComposer.scripts.document_api",
    "apply_patches": "skills.WPSComposer.scripts.document_api",
    "apply_ops": "skills.WPSComposer.scripts.document_api",
    "validate_op": "skills.WPSComposer.scripts.document_api",
    "snapshot_json": "skills.WPSComposer.scripts.document_api",
    "supported_formats": "skills.WPSComposer.scripts.document_api",
    "validate_target": "skills.WPSComposer.scripts.document_api",
    "patch_grammar": "skills.WPSComposer.scripts.document_api",
    "snapshot_to_patches": "skills.WPSComposer.scripts.document_api",
    "PatchError": "skills.WPSComposer.scripts.document_api",
    "PATCH_GRAMMAR": "skills.WPSComposer.scripts.document_api",
    "ALL_OPS": "skills.WPSComposer.scripts.document_api",
    "STRUCTURAL_OPS": "skills.WPSComposer.scripts.document_api",
    "INSERT_TYPES": "skills.WPSComposer.scripts.document_api",
    # Plugins
    "list_plugins": "skills.WPSComposer.scripts.plugins",
    "register_plugin": "skills.WPSComposer.scripts.plugins",
}

__all__ = list(_WPS_ENGINE_NAMES.keys())

_CACHE: dict[str, Any] = {}


def __getattr__(name: str) -> Any:
    if name in _CACHE:
        return _CACHE[name]
    module_name = _WPS_ENGINE_NAMES.get(name)
    if module_name is None:
        raise AttributeError(f"module 'skills.WPSComposer' has no attribute '{name}'")
    import importlib

    module = importlib.import_module(module_name)
    value = getattr(module, name)
    _CACHE[name] = value
    return value


def __dir__() -> list[str]:
    return list(_WPS_ENGINE_NAMES.keys())
