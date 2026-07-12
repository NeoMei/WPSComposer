"""Unified conversational API for inspecting and editing WPS documents.

All three COM composers implement the same small contract:

``inspect_document()`` -> JSON-compatible element tree
``inspect_selection()`` -> formatting of the user's current selection
``apply_format_patch(target, **patch)`` -> partial, non-destructive update

This module chooses the correct composer from a file extension or document kind
and adds batch patching for agent workflows.
"""

from __future__ import annotations

import json
import os

from .writer import WriterComposer
from .sheet import SheetComposer
from .slide import SlideComposer


WRITER_EXTENSIONS = {
    ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf",
    ".txt", ".html", ".htm", ".mht", ".mhtml", ".xml", ".odt",
    ".wps", ".wpt",
}
SHEET_EXTENSIONS = {
    ".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm",
    ".csv", ".tsv", ".ods", ".xml", ".html", ".htm", ".et", ".ett",
}
SLIDE_EXTENSIONS = {
    ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".ppsm", ".pot",
    ".potx", ".potm", ".odp", ".dps", ".dpt",
}

_KINDS = {
    "writer": WriterComposer, "word": WriterComposer, "document": WriterComposer,
    "sheet": SheetComposer, "excel": SheetComposer, "spreadsheet": SheetComposer,
    "slide": SlideComposer, "powerpoint": SlideComposer, "presentation": SlideComposer,
}


def composer_for_path(path, kind=None):
    """Return the composer class appropriate for *path*."""
    if kind:
        return composer_for_kind(kind)
    ext = os.path.splitext(os.fspath(path))[1].lower()
    # XML/HTML can be opened by both Writer and Sheet; default to Writer unless
    # the caller supplies kind="sheet".
    if ext in WRITER_EXTENSIONS:
        return WriterComposer
    if ext in SHEET_EXTENSIONS:
        return SheetComposer
    if ext in SLIDE_EXTENSIONS:
        return SlideComposer
    raise ValueError(
        f"Unsupported document extension '{ext}'. Supply kind='writer', "
        "'sheet', or 'slide' for a host-supported format."
    )


def composer_for_kind(kind):
    try:
        return _KINDS[str(kind).lower()]
    except KeyError as exc:
        raise ValueError("kind must be writer, sheet, or slide") from exc


def open_document(path, *, kind=None, read_only=False, visible=False):
    """Open an existing file and return a context-manageable composer."""
    cls = composer_for_path(path, kind)
    return cls.open_document(path, read_only=read_only, visible=visible)


def attach_active(kind=None):
    """Attach to the current WPS/Office document.

    When *kind* is omitted, Writer, Sheet, and Slide are probed in that order.
    Pass a kind when multiple WPS applications are open and the choice matters.
    """
    if kind:
        return composer_for_kind(kind).attach_active()
    errors = []
    for cls in (WriterComposer, SheetComposer, SlideComposer):
        try:
            return cls.attach_active()
        except Exception as exc:
            errors.append(str(exc))
    raise RuntimeError("No active WPS/Office document found: " + " | ".join(errors))


def inspect(path=None, *, kind=None, selection=False, **options):
    """Inspect a file or the current active document and return plain data."""
    if path is None:
        composer = attach_active(kind)
        return composer.inspect_selection() if selection else composer.inspect_document(**options)
    with open_document(path, kind=kind, read_only=True) as composer:
        return composer.inspect_selection() if selection else composer.inspect_document(**options)


def apply_patches(composer, patches, *, stop_on_error=True):
    """Apply ordered ``{"target": ..., ...}`` patches to an open composer."""
    reports = []
    for index, patch in enumerate(patches):
        item = dict(patch)
        target = item.pop("target", None)
        if not target:
            error = "patch requires a target"
            reports.append({"index": index, "ok": False, "error": error})
            if stop_on_error:
                raise ValueError(error)
            continue
        try:
            result = composer.apply_format_patch(target, **item)
            reports.append({
                "index": index,
                "target": target,
                "ok": not result.get("rejected"),
                **result,
            })
        except Exception as exc:
            reports.append({"index": index, "target": target,
                            "ok": False, "error": str(exc)})
            if stop_on_error:
                raise
    return reports


def edit(path=None, *, kind=None, patches, output=None, export_pdf=None,
         visible=False, stop_on_error=True, inspect_after=False):
    """Inspect/edit/save a file or the active document in one agent call.

    If ``path`` and ``output`` are both omitted, the active document is edited
    and saved in place.  Supplying ``output`` keeps the original untouched.
    """
    attached = path is None
    if attached:
        composer = attach_active(kind)
    else:
        composer = open_document(path, kind=kind, read_only=False, visible=visible)
        composer.__enter__()
    try:
        before = composer.inspect_document() if inspect_after else None
        reports = apply_patches(composer, patches, stop_on_error=stop_on_error)
        saved_path = composer.save(output) if output else composer.save_current()
        pdf_path = composer.export_pdf(export_pdf) if export_pdf else None
        after = composer.inspect_document() if inspect_after else None
        return {
            "saved_path": saved_path,
            "pdf_path": pdf_path,
            "patches": reports,
            "before": before,
            "after": after,
        }
    finally:
        composer.close(save_changes=False)


def snapshot_json(snapshot, *, indent=2):
    """Serialize an inspection snapshot for logs or model context."""
    return json.dumps(snapshot, ensure_ascii=False, indent=indent, default=str)


def supported_formats():
    return {
        "writer": sorted(WRITER_EXTENSIONS),
        "sheet": sorted(SHEET_EXTENSIONS),
        "slide": sorted(SLIDE_EXTENSIONS),
    }


__all__ = [
    "open_document", "attach_active", "inspect", "edit", "apply_patches",
    "snapshot_json", "supported_formats", "composer_for_path",
    "composer_for_kind",
]
