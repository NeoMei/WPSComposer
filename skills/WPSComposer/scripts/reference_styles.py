"""Reference Word style definitions.

Style parameters derived from analysis of formal Chinese corporate proposal
documents. Conventions follow GB/T 9704-style government/corporate formatting:

- Body text:     FangSong, 12pt, justified, first-line indent 2 chars
- Headings H1:   SimHei, 16pt, centered, bold
- Headings H2-3: SimHei, 15pt, bold
- Latin text:    Times New Roman
- Line spacing:  1.5 lines, expressed through Word's line-spacing rule

Each style is a dict of properties recognised by WriterComposer.ensure_styles().
"""

from __future__ import annotations
from typing import Dict, Any


# Font families - formal Chinese document convention
BODY_FONT    = "\u4eff\u5b8b"        # FangSong - standard body
HEADING_FONT = "\u9ed1\u4f53"        # SimHei   - all heading levels
LATIN_FONT   = "Times New Roman"     # Latin companion
MONO_FONT    = "Consolas"            # Monospace for code


# Style definitions
STYLES: Dict[str, Dict[str, Any]] = {

    # Document title (cover page)
    "Title": {
        "name": "Title",
        "type": "paragraph",
        "font_name": HEADING_FONT,
        "font_name_ascii": HEADING_FONT,
        "font_size": 22,
        "bold": True,
        "align": 1,
        "space_before": 0,
        "space_after": 12,
        "line_spacing_rule": "one_and_half",
    },

    "Author": {
        "name": "Author",
        "type": "paragraph",
        "font_name": BODY_FONT,
        "font_name_ascii": LATIN_FONT,
        "font_size": 14,
        "align": 1,
        "color": "#000000",
        "space_after": 4,
        "line_spacing_rule": "one_and_half",
    },

    "Date": {
        "name": "Date",
        "type": "paragraph",
        "font_name": BODY_FONT,
        "font_name_ascii": LATIN_FONT,
        "font_size": 14,
        "align": 1,
        "color": "#000000",
        "space_after": 24,
        "line_spacing_rule": "one_and_half",
    },

    # Body text - FangSong 12pt justified indent 2 chars
    "BodyText": {
        "name": "Body Text",
        "type": "paragraph",
        "based_on": "Normal",
        "font_name": BODY_FONT,
        "font_name_ascii": LATIN_FONT,
        "font_size": 12,
        "align": 3,
        "indent_first": 24,
        "line_spacing_rule": "one_and_half",
        "space_before": 0,
        "space_after": 0,
    },

    "FirstParagraph": {
        "name": "First Paragraph",
        "type": "paragraph",
        "based_on": "Body Text",
        "indent_first": 24,
        "space_before": 0,
        "space_after": 0,
        "line_spacing_rule": "one_and_half",
    },

    "Compact": {
        "name": "Compact",
        "type": "paragraph",
        "based_on": "Body Text",
        "space_after": 0,
        "space_before": 0,
        "line_spacing_rule": "single",
        "indent_first": 0,
    },

    "SourceCode": {
        "name": "Source Code",
        "type": "paragraph",
        "font_name": MONO_FONT,
        "font_name_ascii": MONO_FONT,
        "font_size": 9,
        "left_indent": 36,
        "indent_first": 0,
        "line_spacing_rule": "single",
        "space_after": 0,
        "space_before": 0,
        "shading": "#F5F5F5",
    },

    "BlockText": {
        "name": "Block Text",
        "type": "paragraph",
        "based_on": "Body Text",
        "font_size": 11,
        "italic": False,
        "color": "#000000",
        "left_indent": 48,
        "indent_first": 0,
        "left_border": True,
        "border_color": "#999999",
        "line_spacing_rule": "one_and_half",
        "space_after": 8,
    },

    "ImageCaption": {
        "name": "Image Caption",
        "type": "paragraph",
        "font_name": BODY_FONT,
        "font_name_ascii": LATIN_FONT,
        "font_size": 9,
        "italic": False,
        "align": 1,
        "color": "#000000",
        "space_before": 4,
        "space_after": 12,
    },

    "TableCaption": {
        "name": "Table Caption",
        "type": "paragraph",
        "font_name": BODY_FONT,
        "font_name_ascii": LATIN_FONT,
        "font_size": 9,
        "bold": True,
        "align": 1,
        "color": "#000000",
        "space_before": 12,
        "space_after": 4,
    },

    "ListParagraph": {
        "name": "List Paragraph",
        "type": "paragraph",
        "based_on": "Body Text",
        "indent_first": 0,
        "left_indent": 24,
        "line_spacing_rule": "one_and_half",
        "space_after": 3,
    },

    # Compact table text: no inherited body spacing inside cells.
    "TableBody": {
        "name": "Table Body",
        "type": "paragraph",
        "based_on": "Normal",
        "font_name": BODY_FONT,
        "font_name_ascii": LATIN_FONT,
        "font_size": 10,
        "align": 0,
        "indent_first": 0,
        "left_indent": 0,
        "right_indent": 0,
        "line_spacing_rule": "single",
        "space_before": 0,
        "space_after": 0,
    },

    "TableHeader": {
        "name": "Table Header",
        "type": "paragraph",
        "based_on": "Table Body",
        "font_name": BODY_FONT,
        "font_name_ascii": LATIN_FONT,
        "font_size": 10,
        "bold": True,
        "line_spacing_rule": "single",
        "space_before": 0,
        "space_after": 0,
    },
}


# Character styles for inline formatting
CHAR_STYLES: Dict[str, Dict[str, Any]] = {
    "VerbatimChar": {
        "name": "Verbatim Char",
        "type": "character",
        "font_name": MONO_FONT,
        "font_name_ascii": MONO_FONT,
        "font_size": 9,
    },
    "Hyperlink": {
        "name": "Hyperlink",
        "type": "character",
        "color": "#000000",
        "underline": True,
    },
    "Strikeout": {
        "name": "Strikeout",
        "type": "character",
        "strikethrough": True,
    },
}


# Heading style properties (used by add_heading_level / writer_renderer)
HEADING_STYLE_MAP: Dict[int, Dict[str, Any]] = {
    1: {"font_name": HEADING_FONT, "font_name_ascii": HEADING_FONT, "font_size": 16, "bold": True, "align": 1, "space_before": 16, "space_after": 5, "line_spacing_rule": "single", "keep_with_next": True},
    2: {"font_name": HEADING_FONT, "font_name_ascii": HEADING_FONT, "font_size": 15, "bold": True, "align": 0, "space_before": 14, "space_after": 5, "line_spacing_rule": "single", "keep_with_next": True},
    3: {"font_name": HEADING_FONT, "font_name_ascii": HEADING_FONT, "font_size": 15, "bold": True, "align": 0, "space_before": 12, "space_after": 5, "line_spacing_rule": "single", "keep_with_next": True},
    4: {"font_name": HEADING_FONT, "font_name_ascii": HEADING_FONT, "font_size": 14, "bold": True, "align": 0, "space_before": 10, "space_after": 5, "line_spacing_rule": "single", "keep_with_next": True},
    5: {"font_name": HEADING_FONT, "font_name_ascii": HEADING_FONT, "font_size": 14, "bold": True, "align": 0, "space_before": 8, "space_after": 5, "line_spacing_rule": "single", "keep_with_next": True},
    6: {"font_name": HEADING_FONT, "font_name_ascii": HEADING_FONT, "font_size": 12, "bold": True, "align": 0, "space_before": 6, "space_after": 5, "line_spacing_rule": "single", "keep_with_next": True},
}


def get_style(name: str) -> dict:
    if name not in STYLES:
        raise KeyError(f"Unknown style '{name}'. Available: {list(STYLES)}")
    return STYLES[name]


def get_char_style(name: str) -> dict:
    if name not in CHAR_STYLES:
        raise KeyError(f"Unknown char style '{name}'. Available: {list(CHAR_STYLES)}")
    return CHAR_STYLES[name]


def get_heading_style(level: int) -> dict:
    """Get heading style properties for a given level (1-6)."""
    if level not in HEADING_STYLE_MAP:
        level = 6
    return HEADING_STYLE_MAP[level]

