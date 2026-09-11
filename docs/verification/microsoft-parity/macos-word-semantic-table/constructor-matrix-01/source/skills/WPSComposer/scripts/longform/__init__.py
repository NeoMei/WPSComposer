"""Platform-independent long-form semantic helpers."""

from importlib import import_module
from typing import Any

from .bookmark_ids import (
    BOOKMARK_COLLISION_UNRESOLVED,
    BOOKMARK_ID_INVALID,
    BOOKMARK_KIND_INVALID,
    BookmarkMapResult,
    map_bookmarks,
)
from .frontmatter_parser import (
    FRONTMATTER_INVALID,
    FRONTMATTER_UNCLOSED,
    FrontmatterParseResult,
    parse_frontmatter_document,
)
from .unicode_text import (
    contains_han,
    display_units,
    normalize_visible_text,
    shorten_display_units,
)


_LAZY_EXPORTS = {
    "CaptionBinding": ("..document_model", "CaptionBinding"),
    "CrossReferenceRun": ("..document_model", "CrossReferenceRun"),
    "FigureImageLayout": (".image_policy", "FigureImageLayout"),
    "ImageProfile": (".resources", "ImageProfile"),
    "LongformBuild": (".pipeline", "LongformBuild"),
    "NativeFieldAdapter": (".field_contract", "NativeFieldAdapter"),
    "NativeFieldContractError": (".field_contract", "NativeFieldContractError"),
    "PreparedLongformResource": (".resources", "PreparedLongformResource"),
    "GenerationOutcome": (".quality", "GenerationOutcome"),
    "PageRole": (".quality", "PageRole"),
    "QualityConfidence": (".quality", "QualityConfidence"),
    "QualityEvidence": (".quality", "QualityEvidence"),
    "QualityFinding": (".quality", "QualityFinding"),
    "QualityReport": (".quality", "QualityReport"),
    "QualitySeverity": (".quality", "QualitySeverity"),
    "TableMerge": (".table_policy", "TableMerge"),
    "TablePolicy": (".table_policy", "TablePolicy"),
    "build_longform_generation": (".pipeline", "build_longform_generation"),
    "execute_longform_plan": (".pipeline", "execute_longform_plan"),
}


def __getattr__(name: str) -> Any:
    """Load M3 integration types lazily to preserve parser import purity."""
    target = _LAZY_EXPORTS.get(name)
    if target is None:
        raise AttributeError(name)
    module_name, attribute = target
    value = getattr(import_module(module_name, __name__), attribute)
    globals()[name] = value
    return value

__all__ = [
    "BOOKMARK_COLLISION_UNRESOLVED",
    "BOOKMARK_ID_INVALID",
    "BOOKMARK_KIND_INVALID",
    "BookmarkMapResult",
    "CaptionBinding",
    "CrossReferenceRun",
    "FigureImageLayout",
    "FRONTMATTER_INVALID",
    "FRONTMATTER_UNCLOSED",
    "FrontmatterParseResult",
    "ImageProfile",
    "LongformBuild",
    "NativeFieldAdapter",
    "NativeFieldContractError",
    "PreparedLongformResource",
    "GenerationOutcome",
    "PageRole",
    "QualityConfidence",
    "QualityEvidence",
    "QualityFinding",
    "QualityReport",
    "QualitySeverity",
    "TableMerge",
    "TablePolicy",
    "build_longform_generation",
    "contains_han",
    "display_units",
    "execute_longform_plan",
    "map_bookmarks",
    "normalize_visible_text",
    "parse_frontmatter_document",
    "shorten_display_units",
]
