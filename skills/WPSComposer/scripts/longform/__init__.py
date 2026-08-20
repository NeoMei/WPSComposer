"""Platform-independent long-form semantic helpers."""

from .frontmatter_parser import (
    FRONTMATTER_INVALID,
    FRONTMATTER_UNCLOSED,
    FrontmatterParseResult,
    parse_frontmatter_document,
)

__all__ = [
    "FRONTMATTER_INVALID",
    "FRONTMATTER_UNCLOSED",
    "FrontmatterParseResult",
    "parse_frontmatter_document",
]
