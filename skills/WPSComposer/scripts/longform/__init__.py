"""Platform-independent long-form semantic helpers."""

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

__all__ = [
    "BOOKMARK_COLLISION_UNRESOLVED",
    "BOOKMARK_ID_INVALID",
    "BOOKMARK_KIND_INVALID",
    "BookmarkMapResult",
    "FRONTMATTER_INVALID",
    "FRONTMATTER_UNCLOSED",
    "FrontmatterParseResult",
    "contains_han",
    "display_units",
    "map_bookmarks",
    "normalize_visible_text",
    "parse_frontmatter_document",
    "shorten_display_units",
]
