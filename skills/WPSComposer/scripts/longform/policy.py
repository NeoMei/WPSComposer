"""Long-form layout and typography policy.

Derives deterministic, platform-independent policy data from the resolved
semantic configuration.  The policy is consumed by the long-form plan builder
but never touches platform objects or file paths.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .semantic import LongformConfig
from .unicode_text import shorten_display_units


@dataclass(frozen=True)
class LongformPolicy:
    """Resolved deterministic policy for a long-form document."""

    title: str
    short_title: str
    author: str
    date: str
    header: str
    title_page: bool
    toc: bool
    figure_index: bool
    table_index: bool
    bibliography_include_uncited: bool
    caption_numbering: str
    heading_numbering: str

    page_size: str
    page_margins: dict[str, float]

    body_font: dict[str, str]
    heading_font: dict[str, str]
    latin_font: str
    mono_font: str

    body_size_pt: int
    heading_size_pt: int
    line_spacing: float

    toc_levels: int
    toc_title: str
    figure_index_title: str
    table_index_title: str
    page_roles: tuple[str, ...]
    section_numbering_rules: dict[str, dict[str, Any]]
    header_text: str
    header_overflow_text: str
    footer_text: str
    toc_density: dict[str, Any]

    def to_json(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "short_title": self.short_title,
            "author": self.author,
            "date": self.date,
            "header": self.header,
            "title_page": self.title_page,
            "toc": self.toc,
            "figure_index": self.figure_index,
            "table_index": self.table_index,
            "bibliography_include_uncited": self.bibliography_include_uncited,
            "caption_numbering": self.caption_numbering,
            "heading_numbering": self.heading_numbering,
            "page_size": self.page_size,
            "page_margins": self.page_margins,
            "body_font": self.body_font,
            "heading_font": self.heading_font,
            "latin_font": self.latin_font,
            "mono_font": self.mono_font,
            "body_size_pt": self.body_size_pt,
            "heading_size_pt": self.heading_size_pt,
            "line_spacing": self.line_spacing,
            "toc_levels": self.toc_levels,
            "toc_title": self.toc_title,
            "figure_index_title": self.figure_index_title,
            "table_index_title": self.table_index_title,
            "page_roles": self.page_roles,
            "section_numbering_rules": self.section_numbering_rules,
            "header_text": self.header_text,
            "header_overflow_text": self.header_overflow_text,
            "footer_text": self.footer_text,
            "toc_density": self.toc_density,
        }


def build_policy(config: LongformConfig) -> LongformPolicy:
    """Build a deterministic policy from a resolved semantic configuration."""
    header_text = shorten_display_units(config.header or "", 64)
    header_overflow_text = shorten_display_units(config.header or "", 32)
    toc_density = {
        "min_font_size_pt": {"toc1": 10.5, "toc2": 10.0, "toc3": 10.0},
        "normal_font_size_pt": {"toc1": 11.0, "toc2": 10.5, "toc3": 10.5},
        "min_space_before_pt": {"toc1": 0.0, "toc2": 0.0, "toc3": 0.0},
        "normal_space_before_pt": {"toc1": 3.0, "toc2": 0.0, "toc3": 0.0},
        "min_space_after_pt": {"toc1": 0.0, "toc2": 0.0, "toc3": 0.0},
        "normal_space_after_pt": {"toc1": 0.0, "toc2": 0.0, "toc3": 0.0},
    }
    section_numbering_rules = {
        "cover": {"format": "none", "start_page_number": None, "restart": False},
        "front_matter": {"format": "roman", "start_page_number": 1, "restart": True},
        "body": {"format": "arabic", "start_page_number": 1, "restart": True},
        "landscape": {"format": "continue", "start_page_number": None, "restart": False},
    }
    return LongformPolicy(
        title=config.title,
        short_title=config.short_title,
        author=config.author,
        date=config.date,
        header=config.header,
        title_page=bool(config.title_page),
        toc=bool(config.toc),
        figure_index=bool(config.figure_index),
        table_index=bool(config.table_index),
        bibliography_include_uncited=bool(config.bibliography_include_uncited),
        caption_numbering=config.caption_numbering or "global",
        heading_numbering=config.heading_numbering or "none",
        page_size="A4",
        page_margins={
            "top_mm": 25.4,
            "bottom_mm": 25.4,
            "left_mm": 30.0,
            "right_mm": 25.0,
        },
        body_font={"cjk": "宋体", "latin": "Times New Roman"},
        heading_font={"cjk": "黑体", "latin": "Times New Roman"},
        latin_font="Times New Roman",
        mono_font="Consolas",
        body_size_pt=12,
        heading_size_pt=14,
        line_spacing=1.5,
        toc_levels=3,
        toc_title="目录",
        figure_index_title="图目录",
        table_index_title="表目录",
        page_roles=("cover", "front_matter", "body", "landscape"),
        section_numbering_rules=section_numbering_rules,
        header_text=header_text,
        header_overflow_text=header_overflow_text,
        footer_text="",
        toc_density=toc_density,
    )


__all__ = ["LongformPolicy", "build_policy"]
