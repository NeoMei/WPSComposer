"""Controlled native field descriptors for M3 captions and references.

The semantic layer supplies only resolved object kinds and safe bookmark names.
This module maps those values to a closed vocabulary; it never accepts field
code fragments from document text.
"""

from __future__ import annotations

import re
from typing import Any

from ..document_model import CaptionBinding, CrossReferenceRun


_KINDS: dict[str, tuple[str, str, str]] = {
    "figure": ("WPSC_FIG", "图 ", ""),
    "table": ("WPSC_TAB", "表 ", ""),
    "equation": ("WPSC_EQ", "(", ")"),
}
_BOOKMARK_RE = re.compile(r"^wpsc_(fig|tab|eq)_[0-9a-f]{24}$")


def caption_numbering_descriptor(
    kind: str,
    binding: CaptionBinding,
) -> dict[str, Any]:
    """Return the complete controlled descriptor for one native caption."""
    try:
        sequence_id, prefix, suffix = _KINDS[kind]
    except KeyError as error:
        raise ValueError("unsupported native caption kind") from error
    if binding.mode not in {"global", "chapter"}:
        raise ValueError("caption binding mode must be global or chapter")
    chapter = binding.mode == "chapter"
    if chapter and not binding.chapter_node_id:
        raise ValueError("chapter caption binding requires a chapter node")
    return {
        "mode": binding.mode,
        "sequenceId": sequence_id,
        "chapterStyleLevel": 1 if chapter else None,
        "resetLevel": 1 if chapter else None,
        "prefix": prefix,
        "suffix": suffix,
    }


def cross_reference_descriptor(run: CrossReferenceRun) -> dict[str, Any]:
    """Return a safe REF descriptor for one resolved inline run."""
    if not run.target_node_id or not run.bookmark_name:
        raise ValueError("cross-reference target and bookmark must be resolved")
    canonical_kind = {
        "fig": "figure",
        "tab": "table",
        "eq": "equation",
        "figure": "figure",
        "table": "table",
        "equation": "equation",
    }.get(str(run.target_kind))
    try:
        _sequence_id, prefix, suffix = _KINDS[str(canonical_kind)]
    except KeyError as error:
        raise ValueError("unsupported cross-reference kind") from error
    expected_kind = {"figure": "fig", "table": "tab", "equation": "eq"}[canonical_kind]
    match = _BOOKMARK_RE.fullmatch(run.bookmark_name)
    if match is None or match.group(1) != expected_kind:
        raise ValueError("cross-reference bookmark kind does not match target")
    return {
        "targetNodeId": run.target_node_id,
        "targetKind": canonical_kind,
        "bookmarkName": run.bookmark_name,
        "prefix": prefix,
        "suffix": suffix,
    }


__all__ = ["caption_numbering_descriptor", "cross_reference_descriptor"]
