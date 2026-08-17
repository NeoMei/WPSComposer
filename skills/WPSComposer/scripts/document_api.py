"""Unified conversational API for inspecting and editing WPS documents.

All three COM composers implement the same small contract:

``inspect_document()`` -> JSON-compatible element tree
``inspect_selection()`` -> formatting of the user's current selection
``apply_format_patch(target, **patch)`` -> partial, non-destructive update (the ``set`` verb)
``apply_structural_op(op)`` -> insert / remove / move / clone element (structural verbs)

This module chooses the correct composer from a file extension or document kind
and adds batch patching for agent workflows.

Agent-friendly additions (borrowed from the officecli contract):

* :func:`validate_target` / :func:`patch_grammar` -- let an agent self-correct
  a bad target without a guess-fail-retry loop.
* :class:`PatchError` and structured per-op reports carrying ``error.code``
  + ``suggestion`` instead of bare tracebacks.
* :func:`apply_ops` / :func:`edit` are atomic by default: any failed op
  blocks the save and returns a structured failure result.
* Structural verbs (``insert`` / ``remove`` / ``move`` / ``clone``) alongside
  the formatting ``set`` verb -- full document surgery, not just formatting.
"""

from __future__ import annotations

import difflib
import json
import os
from pathlib import Path
import re
import tempfile

from .artifact_transport import (
    ArtifactValidationError,
    publish_artifact,
    validate_office_package,
)
from .writer import WriterComposer
from .sheet import SheetComposer
from .slide import SlideComposer


def _com_available():
    """Return True when the Windows COM dispatch path can run here."""
    import platform
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


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


# ---------------------------------------------------------------------------
# Patch target grammar -- single source of truth for validation + help
# ---------------------------------------------------------------------------
# Each entry: (example_form, regex_or_None, element_kind, description).
# regex None marks a literal target ("selection", "presentation").

PATCH_GRAMMAR = {
    "writer": [
        ("selection", None, "selection", "the active selection"),
        ("paragraph:N", r"paragraph:(\d+)", "paragraph",
         "a paragraph by 1-based index"),
        # Stable form: survives structural edits between inspect() and edit().
        # paraId is read from w14:paraId in document.xml
        # (see WriterComposer._read_paraid_map). Hex, case-insensitive.
        ("paragraph:@paraId=HEX", r"paragraph:@paraId=([0-9A-Fa-f]+)", "paragraph",
         "a paragraph by its stable w14:paraId (preferred over positional)"),
        ("range:S-E", r"range:(\d+)-(\d+)", "range",
         "a character range by start,end offsets"),
        ("table:N/cell:R,C", r"table:(\d+)/cell:(\d+),(\d+)", "table_cell",
         "a table cell by table,row,col (1-based)"),
        ("shape:N", r"shape:(\d+)", "shape",
         "a floating shape by 1-based index"),
        ("section:N", r"section:(\d+)", "section",
         "a page-layout section by 1-based index"),
    ],
    "sheet": [
        ("selection", None, "selection", "the active selection"),
        ("sheet:N", r"sheet:(\d+)", "sheet", "a worksheet by 1-based index"),
        ("sheet:N/cell:A1", r"sheet:(\d+)/cell:(.+)", "cell",
         "a cell on a worksheet (leading $ accepted)"),
        ("sheet:N/range:A1:C20", r"sheet:(\d+)/range:(.+)", "range",
         "a range on a worksheet"),
        ("sheet:N/shape:N", r"sheet:(\d+)/shape:(\d+)", "shape",
         "a shape on a worksheet by 1-based index"),
        # Stable forms. shape.Id / shape.Name readback in
        # SheetComposer.inspect_document (shape snapshot).
        ("sheet:N/shape:@id=N", r"sheet:(\d+)/shape:@id=(\d+)", "shape",
         "a shape by its stable COM Id (preferred over positional)"),
        ("sheet:N/shape:@name=NAME", r"sheet:(\d+)/shape:@name=(.+)", "shape",
         "a shape by Name"),
        ("sheet:N/chart:N", r"sheet:(\d+)/chart:(\d+)", "chart",
         "a chart object on a worksheet"),
    ],
    "slide": [
        ("selection", None, "selection", "the active selection"),
        ("presentation", None, "presentation", "the whole presentation"),
        ("slide:N", r"slide:(\d+)", "slide", "a slide by 1-based index"),
        ("slide:N/shape:N", r"slide:(\d+)/shape:(\d+)", "shape",
         "a shape on a slide by 1-based index"),
        # Stable forms. shape.Id / shape.Name readback in
        # SlideComposer._shape_snapshot; resolution in apply_format_patch.
        ("slide:N/shape:@id=N", r"slide:(\d+)/shape:@id=(\d+)", "shape",
         "a shape by its stable Shape.Id (preferred over positional)"),
        ("slide:N/shape:@name=NAME", r"slide:(\d+)/shape:@name=(.+)", "shape",
         "a shape by Name"),
        ("slide:N/shape:N/paragraph:N",
         r"slide:(\d+)/shape:(\d+)/paragraph:(\d+)", "paragraph",
         "a text paragraph inside a shape"),
        ("slide:N/shape:N/paragraph:N/run:N",
         r"slide:(\d+)/shape:(\d+)/paragraph:(\d+)/run:(\d+)", "run",
         "a single run inside a shape paragraph"),
        ("slide:N/shape:N/table/cell:R,C",
         r"slide:(\d+)/shape:(\d+)/table/cell:(\d+),(\d+)", "table_cell",
         "a table cell by shape,row,col (1-based)"),
    ],
}

_KIND_ALIASES = {
    "writer": "writer", "word": "writer", "document": "writer",
    "sheet": "sheet", "excel": "sheet", "spreadsheet": "sheet",
    "slide": "slide", "powerpoint": "slide", "presentation": "slide",
}

_COMPOSER_KIND = {
    WriterComposer: "writer",
    SheetComposer: "sheet",
    SlideComposer: "slide",
}


def _normalize_kind(kind):
    if kind is None:
        return None
    return _KIND_ALIASES.get(str(kind).lower())


def _kind_from_composer(composer):
    for cls in type(composer).__mro__:
        if cls in _COMPOSER_KIND:
            return _COMPOSER_KIND[cls]
    return getattr(composer, "kind", None)


def patch_grammar(kind=None):
    """Return the patch-target grammar as plain data (agent help / discovery).

    With *kind* omitted, returns a dict keyed by kind. With a kind, returns
    that kind's list of ``{form, element, description}`` entries.
    """
    if kind is None:
        return {
            name: [
                {"form": form, "element": element, "description": desc}
                for form, _pat, element, desc in entries
            ]
            for name, entries in PATCH_GRAMMAR.items()
        }
    normalized = _normalize_kind(kind)
    if normalized not in PATCH_GRAMMAR:
        raise ValueError(
            f"Unknown kind {kind!r}; expected writer, sheet, or slide"
        )
    return [
        {"form": form, "element": element, "description": desc}
        for form, _pat, element, desc in PATCH_GRAMMAR[normalized]
    ]


def validate_target(target, kind):
    """Validate a patch target string against the grammar for *kind*.

    Returns ``{"valid": bool, "kind": ..., "element": ..., "form": ...}`` on
    success, or ``{"valid": False, "error": {"code", "message",
    "valid_forms", "closest"}}`` on failure -- the suggestion lets an agent
    self-correct instead of guessing.
    """
    normalized = _normalize_kind(kind)
    if normalized not in PATCH_GRAMMAR:
        return {
            "valid": False,
            "kind": kind,
            "error": {
                "code": "unsupported_kind",
                "message": f"Unknown kind {kind!r}",
                "valid_forms": sorted(PATCH_GRAMMAR),
            },
        }
    for form, pattern, element, _desc in PATCH_GRAMMAR[normalized]:
        if pattern is None:
            matched = target == form
        else:
            matched = re.fullmatch(pattern, target) is not None
        if matched:
            return {
                "valid": True,
                "kind": normalized,
                "element": element,
                "form": form,
            }
    forms = [f for f, _p, _e, _d in PATCH_GRAMMAR[normalized]]
    closest = difflib.get_close_matches(target, forms, n=1, cutoff=0.4)
    return {
        "valid": False,
        "kind": normalized,
        "target": target,
        "error": {
            "code": "invalid_target",
            "message": f"Unsupported {normalized} target: {target!r}",
            "valid_forms": forms,
            "closest": closest[0] if closest else None,
        },
    }


class PatchError(ValueError):
    """Raised when one or more patches fail in atomic mode.

    Carries the full per-patch ``reports`` list and the failing ``errors``
    so callers (or agents) can inspect what went wrong without re-running.
    """

    def __init__(self, reports):
        self.reports = list(reports)
        self.errors = [r for r in self.reports if not r.get("ok")]
        message = "{} of {} patch(es) failed".format(
            len(self.errors), len(self.reports)
        )
        super().__init__(message)


def snapshot_to_patches(snapshot, *, dimensions=("font", "paragraph", "fill")):
    """Convert an :func:`inspect` snapshot into a replayable patch list.

    Walks the snapshot recursively and emits one ``{"target": id, ...}`` patch
    per element that has a valid, stable address plus at least one of the
    requested formatting *dimensions*. The result can be passed to
    :func:`apply_patches` or :func:`edit` to reproduce the captured styling on
    another document ("make this doc look like that one").

    This is the WpsComposer analogue of officecli's ``dump`` -> ``batch``
    replay, scoped to *formatting* (full structural cloning is ``generate()``'s
    job via markdown). Only non-empty dimensions are emitted, so callers can
    pass ``dimensions=("font",)`` to copy just fonts.

    Note: fidelity depends on the snapshot/apply key symmetry for each
    dimension (e.g. ``font_snapshot`` keys vs ``apply_font`` keys); see the
    Windows verification doc for the round-trip check.
    """
    kind = snapshot.get("kind") if isinstance(snapshot, dict) else None
    wanted = tuple(dimensions)
    patches = []
    seen = set()

    # Snapshot keys the host exposes read-only; replaying them always lands
    # in `rejected`. (Fill.Type and ZOrderPosition are read-only in the
    # Office/WPS object models.)
    readonly = {"fill": ("type",), "geometry": ("z_order",)}

    def replayable_id(node, element_id):
        # Stable ids (@paraId/@id/@name) are document-specific; replaying them
        # on another document cannot resolve. Rewrite to the positional form
        # using the element's index when one is available.
        index = node.get("index")
        if not (isinstance(index, int) and index > 0):
            return element_id
        if re.fullmatch(r"paragraph:@paraId=[0-9A-Fa-f]+", element_id):
            return f"paragraph:{index}"
        match = re.fullmatch(
            r"((?:slide|sheet):\d+/shape:)@(?:id|name)=.+", element_id
        )
        if match:
            return f"{match.group(1)}{index}"
        return element_id

    def visit(node):
        if isinstance(node, dict):
            element_id = node.get("id")
            if isinstance(element_id, str):
                element_id = replayable_id(node, element_id)
            if isinstance(element_id, str) and element_id not in seen:
                info = validate_target(element_id, kind) if kind else None
                if info and info.get("valid"):
                    patch = {"target": element_id}
                    for dim in wanted:
                        value = node.get(dim)
                        if isinstance(value, dict) and value:
                            # None means "host reported no value" — nothing to
                            # replay; carrying it would land in `rejected`.
                            value = {k: v for k, v in value.items()
                                     if v is not None
                                     and k not in readonly.get(dim, ())}
                            if value:
                                patch[dim] = value
                    if len(patch) > 1:
                        patches.append(patch)
                        seen.add(element_id)
            for value in node.values():
                visit(value)
        elif isinstance(node, list):
            for value in node:
                visit(value)

    visit(snapshot)
    return patches


# ---------------------------------------------------------------------------
# Structural operations (insert / remove / move / clone) -- the ``ops`` API
# ---------------------------------------------------------------------------
# ``set`` is the formatting verb (apply_format_patch). The four structural
# verbs are dispatched to composer.apply_structural_op(op). The orchestration
# layer (validate_op / apply_ops / edit) is pure-Python and host-agnostic; the
# COM implementations live on each composer.

SET_OPS = {"set"}
STRUCTURAL_OPS = {"insert", "remove", "move", "clone"}
ALL_OPS = SET_OPS | STRUCTURAL_OPS

# Insert-able element types per host (informational; composer is final arbiter).
INSERT_TYPES = {
    "writer": ("paragraph", "heading", "page_break", "table", "image", "textbox"),
    "slide": ("slide", "textbox", "image"),
    "sheet": ("row", "column", "sheet"),
}

# Parent namespaces for insert (extend the patch target grammar). ``body`` is
# the Writer document body; ``presentation``/``sheet:N``/``slide:N`` already
# exist in PATCH_GRAMMAR.
INSERT_PARENTS = {
    "writer": ("body",),
    "slide": ("presentation",),
    "sheet": ("sheet:N",),
}


def _validate_position(position, kind):
    """Position/destination spec for insert/move/clone. Accepts:

    * ``"end"`` | ``"start"``
    * ``{"after": X}`` / ``{"before": X}`` -- X is a target string (Writer
      paragraph, Slide slide:N) OR an int (Sheet sheet number)
    * ``{"index": N}`` -- int (row/column/paragraph index)
    * ``{"slide": N}`` -- int (Slide shape move/clone destination slide)

    The composer is the final arbiter; this only catches obvious garbage and
    validates string anchors against the grammar when *kind* is known.
    """
    if position is None:
        return {"valid": True, "position": "end"}
    if isinstance(position, str) and position in ("end", "start"):
        return {"valid": True, "position": position}
    if isinstance(position, dict) and len(position) == 1:
        key, value = next(iter(position.items()))
        if key in ("index", "slide") and isinstance(value, int):
            return {"valid": True, "position": {key: value}}
        if key in ("after", "before"):
            # Sheet sheet ops use an int sheet number; accept directly.
            if isinstance(value, int):
                return {"valid": True, "position": {key: value}}
            # Writer/Slide use a target string; validate leniently.
            if isinstance(value, str):
                if kind:
                    info = validate_target(value, kind)
                    if not info.get("valid"):
                        return {"valid": False, "error": {
                            "code": "invalid_anchor",
                            "message": f"position.{key} must be a valid {kind} target: {value!r}",
                        }}
                return {"valid": True, "position": {key: value}}
    return {"valid": False, "error": {
        "code": "invalid_position",
        "message": ("position must be 'end', 'start', {'after': X}, "
                    "{'before': X}, {'index': N}, or {'slide': N}"),
    }}


def validate_op(op, kind=None):
    """Validate a single op dict (``set`` / ``insert`` / ``remove`` / ``move``
    / ``clone``) and return ``{"valid": bool, ...}`` with an ``error`` carrying
    ``code`` + ``message`` on failure -- mirrors :func:`validate_target`."""
    if not isinstance(op, dict):
        return {"valid": False, "error": {
            "code": "invalid_op", "message": "op must be a dict"}}
    verb = op.get("op", "set")
    if verb not in ALL_OPS:
        return {"valid": False, "error": {
            "code": "unknown_verb",
            "message": f"unknown op {verb!r}; expected one of {sorted(ALL_OPS)}",
            "valid_verbs": sorted(ALL_OPS),
        }}

    if verb == "set":
        if not op.get("target"):
            return {"valid": False, "error": {
                "code": "missing_target", "message": "set op requires 'target'"}}
        return {"valid": True, "verb": "set", "target": op["target"]}

    if verb == "insert":
        parent = op.get("parent")
        etype = op.get("type")
        if not etype:
            return {"valid": False, "error": {
                "code": "missing_type", "message": "insert op requires 'type'"}}
        if kind and etype not in INSERT_TYPES.get(kind, ()):
            return {"valid": False, "error": {
                "code": "unsupported_type",
                "message": f"{kind} cannot insert type {etype!r}",
                "valid_types": list(INSERT_TYPES.get(kind, ())),
            }}
        pos = _validate_position(op.get("position"), kind)
        if not pos["valid"]:
            return pos
        return {"valid": True, "verb": "insert", "parent": parent,
                "type": etype, "position": pos["position"]}

    # remove / move / clone all need a target
    if not op.get("target"):
        return {"valid": False, "error": {
            "code": "missing_target",
            "message": f"{verb} op requires 'target'"}}
    if verb in ("move", "clone"):
        pos = _validate_position(op.get("to", op.get("position")), kind)
        if not pos["valid"]:
            return pos
    return {"valid": True, "verb": verb, "target": op["target"]}


def apply_ops(composer, ops, *, atomic=True):
    """Apply an ordered list of ops to an open composer.

    Each op is ``{"op": "set"|"insert"|"remove"|"move"|"clone", ...}``. ``set``
    delegates to ``composer.apply_format_patch``; the structural verbs delegate
    to ``composer.apply_structural_op(op)``. Atomic by default: any failure
    raises :class:`PatchError` carrying every report so far.
    """
    kind = _kind_from_composer(composer)
    reports = []

    for index, op in enumerate(ops):
        if not isinstance(op, dict):
            report = {"index": index, "ok": False, "error": {
                "code": "invalid_op", "message": "op must be a dict"}}
            reports.append(report)
            if atomic:
                raise PatchError(reports)
            continue

        verb = op.get("op", "set")
        pre = validate_op(op, kind)
        if not pre.get("valid"):
            err = pre.get("error", {})
            report = {"index": index, "op": verb, "ok": False, "error": err}
            reports.append(report)
            if atomic:
                raise PatchError(reports)
            continue

        try:
            if verb == "set":
                item = {k: v for k, v in op.items() if k != "op"}
                result = composer.apply_format_patch(item.pop("target"), **item)
                accepted = list(result.get("accepted", []))
                rejected = list(result.get("rejected", []))
                report = {"index": index, "op": "set",
                          "target": op.get("target"),
                          "ok": len(rejected) == 0,
                          "accepted": accepted, "rejected": rejected}
                for k, v in result.items():
                    if k not in ("accepted", "rejected"):
                        report[k] = v
            else:
                result = composer.apply_structural_op(op) or {}
                report = {"index": index, "op": verb,
                          "target": op.get("target"),
                          "ok": True, **result}
                if report.get("rejected"):
                    report["ok"] = False
            reports.append(report)
            if not report.get("ok") and atomic:
                raise PatchError(reports)
        except PatchError:
            raise
        except ValueError as exc:
            reports.append(_error_report(index, op.get("target") or op.get("type"), exc, kind))
            if atomic:
                raise PatchError(reports) from exc
        except Exception as exc:
            reports.append({
                "index": index, "op": verb,
                "target": op.get("target"), "ok": False,
                "error": {"code": "apply_failed", "message": str(exc)},
            })
            if atomic:
                raise PatchError(reports) from exc

    return reports


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
    """Inspect a file or the current active document and return plain data.

    On Windows the live COM path is used.  When COM is unavailable (macOS,
    or any host without pywin32) and the file is a presentation format,
    the WPS JSAPI bridge reads structured content through the real WPS
    engine instead of falling back to PDF extraction.
    """
    if path is None:
        composer = attach_active(kind)
        return composer.inspect_selection() if selection else composer.inspect_document(**options)
    if not _com_available():
        from .macos_probe.inspection import INSPECTABLE, inspect_macos, macos_inspection_available
        if macos_inspection_available():
            ext = os.path.splitext(os.fspath(path))[1].lower()
            if ext in INSPECTABLE:
                include_text = options.pop("include_text", True)
                max_shapes = options.pop("max_shapes", None)
                return inspect_macos(
                    os.fspath(path),
                    include_text=bool(include_text),
                    max_shapes=max_shapes,
                )
    with open_document(path, kind=kind, read_only=True) as composer:
        return composer.inspect_selection() if selection else composer.inspect_document(**options)


def apply_patches(composer, patches, *, atomic=True, stop_on_error=None):
    """Backward-compatible wrapper: apply ``set``-only patches.

    Each patch ``{"target": ..., ...}`` is normalised to a ``set`` op and run
    through :func:`apply_ops`. New code should call :func:`apply_ops` directly
    to use the full verb set (``insert`` / ``remove`` / ``move`` / ``clone``).
    """
    if stop_on_error is not None:
        atomic = bool(stop_on_error)
    ops = [{"op": "set", **patch} for patch in patches]
    return apply_ops(composer, ops, atomic=atomic)


def _error_report(index, target, exc, kind):
    """Build a structured error report for a ValueError, enriching target
    mistakes with the grammar suggestion when *kind* is known."""
    text = str(exc)
    looks_like_target = (
        "Unsupported" in text or "target" in text.lower()
        or "Unknown kind" in text
    )
    code = "invalid_target" if looks_like_target else "invalid_value"
    report = {
        "index": index, "target": target, "ok": False,
        "error": {"code": code, "message": text},
    }
    if kind and code == "invalid_target":
        info = validate_target(target, kind)
        err = info.get("error") or {}
        if err:
            report["error"].update({
                k: v for k, v in err.items() if k != "message"
            })
    return report


def _document_family(path, kind=None):
    """Return the Writer/Sheet/Slide family selected for *path*."""
    cls = composer_for_path(path, kind)
    if cls is WriterComposer:
        return "writer"
    if cls is SheetComposer:
        return "sheet"
    if cls is SlideComposer:
        return "slide"
    raise ValueError(f"Unsupported document family for {path!r}")


def _failed_edit_result(reports, *, warnings=None):
    return {
        "ok": False,
        "saved": False,
        "saved_path": None,
        "pdf_path": None,
        "ops": reports,
        "patches": reports,
        "errors": [report for report in reports if not report.get("ok")],
        "warnings": list(warnings or []),
        "before": None,
        "after": None,
    }


def _mutation_primitive_count(value):
    """Count independently applied leaf values in a patch payload."""
    if isinstance(value, dict):
        return sum(_mutation_primitive_count(item) for item in value.values())
    return 1


def _attached_atomic_is_single_primitive(ops):
    if not ops:
        return True
    if len(ops) != 1 or not isinstance(ops[0], dict):
        return False
    op = ops[0]
    if op.get("op", "set") != "set":
        return False
    payload = {
        key: value for key, value in op.items()
        if key not in {"op", "target"}
    }
    return _mutation_primitive_count(payload) <= 1


def _validate_edited_artifact(path):
    target = Path(path).expanduser().resolve()
    suffix = target.suffix.lower()
    if suffix in {".docx", ".xlsx", ".pptx"}:
        validate_office_package(target, suffix[1:])
        return
    try:
        if not target.is_file() or target.stat().st_size == 0:
            raise ArtifactValidationError(
                f"Edited artifact is missing or empty: {target}"
            )
    except FileNotFoundError as exc:
        raise ArtifactValidationError(
            f"Edited artifact is missing: {target}"
        ) from exc


def _save_edited_artifact(composer, destination, *, attached, overwrite):
    """Save beside the destination, validate, then atomically publish."""
    target = Path(destination).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        dir=target.parent,
        prefix=".wpscomposer-edit-",
        suffix=target.suffix,
    )
    os.close(descriptor)
    staged = Path(temporary_name)
    staged.unlink()
    try:
        if attached:
            composer.save_copy(str(staged))
        else:
            composer.save(str(staged))
        published = publish_artifact(
            staged,
            target,
            overwrite=overwrite,
            validator=_validate_edited_artifact,
        )
        return str(published)
    finally:
        staged.unlink(missing_ok=True)


def edit(path=None, *, kind=None, patches=None, ops=None, output=None,
         export_pdf=None, visible=False, atomic=True, stop_on_error=None,
         inspect_after=False, raise_on_error=False, overwrite=False):
    """Inspect/edit/save a file or the active document in one agent call.

    Accepts two equivalent inputs:

    * ``patches`` -- list of ``{"target": ..., ...}`` formatting patches (the
      legacy ``set``-only form; still supported).
    * ``ops`` -- list of ``{"op": "set"|"insert"|"remove"|"move"|"clone", ...}``
      for mixed formatting + structural editing.

    Patches run first, then ops, in one atomic transaction. If *path* and
    *output* are both omitted, the active document is edited and saved in
    place. Supplying *output* keeps the original untouched when opening by
    *path*; in attach mode (*path* omitted) the live document keeps its
    binding and *output* receives a copy (see `BaseComposer.save_copy`).

    Atomic by default: if any op fails, the document is **not** saved and the
    structured result carries ``ok=False`` plus an ``errors`` list. Set
    ``atomic=False`` to keep the old best-effort behaviour (save whatever
    succeeded). ``raise_on_error=True`` re-raises the :class:`PatchError` for
    callers that prefer exceptions over structured results. A distinct
    *output* must not already exist unless ``overwrite=True``; publication is
    destination-local and atomic. Output formats must remain in the source
    document family.

    Atomic composite edits are rejected before mutation when attaching to a
    live document because WPS/Office exposes no reliable rollback boundary
    there. Only a single ``set`` operation containing at most one leaf property
    is provably one mutation primitive. Use ``atomic=False`` explicitly for
    best effort, or edit a file-backed staging copy.

    .. note:: structural ops (insert/remove/move/clone) shift positional ids
       of later siblings. Address subsequent ops by stable id
       (``@paraId`` / ``@id``) or re-inspect between batches.
    """
    if stop_on_error is not None:
        atomic = bool(stop_on_error)

    combined = []
    if patches:
        combined.extend({"op": "set", **patch} for patch in patches)
    if ops:
        combined.extend(ops)

    attached = path is None
    if attached and output is not None:
        output_path = os.path.abspath(os.fspath(output))
        if os.path.exists(output_path) and not overwrite:
            raise FileExistsError(f"Output already exists: {output_path}")

    if path is not None and output is not None:
        source_family = _document_family(path, kind)
        output_family = _document_family(output)
        if output_family != source_family:
            raise ValueError(
                "edit output must use the same document family as the source "
                f"({source_family}); got {output_family}"
            )
        source_path = os.path.abspath(os.fspath(path))
        output_path = os.path.abspath(os.fspath(output))
        if output_path != source_path and os.path.exists(output_path) and not overwrite:
            raise FileExistsError(f"Output already exists: {output_path}")

    if attached and atomic and not _attached_atomic_is_single_primitive(combined):
        reports = [{
            "index": 0,
            "ok": False,
            "error": {
                "code": "atomic_attached_batch_unsupported",
                "message": (
                    "Atomic composite edits of an attached live document are "
                    "unsupported because the host has no reliable rollback"
                ),
            },
        }]
        if raise_on_error:
            raise PatchError(reports)
        return _failed_edit_result(reports)

    # macOS / non-COM host: route presentation edits through the JSAPI bridge.
    # Only the ``set`` verb (formatting patches) is supported on macOS;
    # structural ops still require Windows COM.
    if path is not None and not _com_available():
        ext = os.path.splitext(os.fspath(path))[1].lower()
        from .macos_probe.inspection import EDITABLE, edit_macos, macos_inspection_available
        inspection_available = macos_inspection_available()
        if (
            inspection_available
            and _document_family(path, kind) == "slide"
            and ext not in EDITABLE
        ):
            raise ValueError(
                f"macOS editing supports only '.pptx' input; got {ext!r}"
            )
        if inspection_available and ext in EDITABLE:
            set_entries = [
                (index, item) for index, item in enumerate(combined)
                if item.get("op") == "set"
            ]
            dropped_entries = [
                (index, item) for index, item in enumerate(combined)
                if item.get("op") != "set"
            ]
            warnings = (
                [f"{len(dropped_entries)} structural op(s) dropped: "
                 "insert/remove/move/clone require Windows COM. "
                 "Only 'set' patches are applied on macOS."]
                if dropped_entries else []
            )
            dropped_reports = [
                {
                    "index": index,
                    "op": item.get("op"),
                    "target": item.get("target"),
                    "ok": False,
                    "error": {
                        "code": "unsupported_operation",
                        "message": (
                            f"macOS presentation editing does not support "
                            f"{item.get('op')!r} operations"
                        ),
                    },
                }
                for index, item in dropped_entries
            ]
            if dropped_reports and (atomic or raise_on_error):
                if raise_on_error:
                    raise PatchError(dropped_reports)
                return _failed_edit_result(dropped_reports, warnings=warnings)
            if set_entries:
                bridge_patches = [
                    {k: v for k, v in p.items() if k != "op"}
                    for _index, p in set_entries
                ]
                result = edit_macos(
                    os.fspath(path), bridge_patches, output=output,
                    atomic=atomic,
                    raise_on_error=raise_on_error,
                    overwrite=overwrite,
                )
                set_reports = list(result.get("patches", []))
                for report, (index, _item) in zip(set_reports, set_entries):
                    report.setdefault("index", index)
                    report.setdefault("op", "set")
                    if report.get("rejected"):
                        report["ok"] = False
                reports = sorted(
                    set_reports + dropped_reports,
                    key=lambda report: report.get("index", len(combined)),
                )
                errors = [report for report in reports if not report.get("ok")]
                if errors and raise_on_error:
                    raise PatchError(reports)
                return {
                    "ok": not errors,
                    "saved": bool(result.get("saved", result.get("path"))),
                    "saved_path": result.get("path"),
                    "pdf_path": None,
                    "ops": reports,
                    "patches": reports,
                    "errors": errors,
                    "warnings": warnings,
                    "before": None,
                    "after": None,
                }
            # Nothing to do: no patches, or only structural ops.
            if dropped_reports:
                return _failed_edit_result(dropped_reports, warnings=warnings)
            return {
                "ok": not warnings,
                "saved": False,
                "saved_path": None,
                "pdf_path": None,
                "ops": [],
                "patches": [],
                "errors": [],
                "warnings": warnings,
                "before": None,
                "after": None,
            }

    if attached:
        composer = attach_active(kind)
    else:
        composer = open_document(path, kind=kind, read_only=False, visible=visible)
        try:
            composer.__enter__()
        except Exception:
            composer.close(save_changes=False)
            raise
    try:
        if attached and output is not None:
            source_family = _normalize_kind(_kind_from_composer(composer))
            output_family = _document_family(output)
            if source_family and output_family != source_family:
                raise ValueError(
                    "edit output must use the same document family as the "
                    f"attached {source_family} document; got {output_family}"
                )
        before = composer.inspect_document() if inspect_after else None
        op_failed = False
        try:
            reports = apply_ops(composer, combined, atomic=atomic)
        except PatchError as exc:
            reports = exc.reports
            op_failed = True
            if raise_on_error:
                raise

        errors = [report for report in reports if not report.get("ok")]
        had_failures = bool(errors)
        if had_failures and raise_on_error:
            raise PatchError(reports)
        if had_failures and atomic:
            saved_path = None
            pdf_path = None
        else:
            if attached and output:
                # SaveAs on an attached document would rebind the user's live
                # document to the new path; save_copy avoids that (WPS
                # Writer/Presentation SaveCopyAs is broken — see _base.py).
                saved_path = _save_edited_artifact(
                    composer,
                    output,
                    attached=True,
                    overwrite=overwrite,
                )
            elif not attached and Path(path).expanduser().is_file():
                destination = output if output is not None else path
                in_place = (
                    Path(destination).expanduser().resolve()
                    == Path(path).expanduser().resolve()
                )
                saved_path = _save_edited_artifact(
                    composer,
                    destination,
                    attached=False,
                    overwrite=overwrite or in_place,
                )
            else:
                saved_path = composer.save(output) if output else composer.save_current()
            pdf_path = composer.export_pdf(export_pdf) if export_pdf else None
        after = composer.inspect_document() if inspect_after else None
        return {
            "ok": not had_failures,
            "saved": saved_path is not None,
            "saved_path": saved_path,
            "pdf_path": pdf_path,
            "ops": reports,
            "patches": reports,  # back-compat alias; prefer "ops"
            "errors": errors,
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
    "apply_ops", "validate_op",
    "snapshot_json", "supported_formats", "composer_for_path",
    "composer_for_kind",
    "validate_target", "patch_grammar", "snapshot_to_patches",
    "PatchError", "PATCH_GRAMMAR",
    "ALL_OPS", "STRUCTURAL_OPS", "INSERT_TYPES",
]
