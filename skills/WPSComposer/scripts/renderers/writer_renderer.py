"""Writer renderer -- Style-driven DOCX generation (pandoc-inspired).

Key improvement over inline formatting:
  - All paragraph/character styles defined ONCE via ensure_styles()
  - Paragraphs reference styles by name, not per-paragraph font tweaking
  - "First Paragraph" style: explicit, consistent first-line indent
  - Consistent, themeable, TOC-aware output
"""

from __future__ import annotations

from ..document_model import (
    StructuredDocument, Section, Paragraph as MDParagraph,
    ListBlock, TableBlock, CodeBlock, ImageBlock, BlockQuote,
    HorizontalRule, TaskList,
)
from ..writer import WriterComposer
from .. import reference_styles as RS
from ..heading_numbering import (
    NumberingState, has_manual_numbering, strip_manual_numbering,
    detect_numbering_scheme,
)


# Standard A4 margins (matches formal Chinese doc conventions)
MARGIN_TOP = 72       # 1 inch = 2.54 cm
MARGIN_BOTTOM = 72
MARGIN_LEFT = 90      # ~3.17 cm (Word default)
MARGIN_RIGHT = 90


def render(doc, output_path, preset=None, composer_factory=WriterComposer):
    """Render StructuredDocument to a professionally styled DOCX."""
    with composer_factory() as w:
        _render_into(w, doc, preset)
        return w.save_docx(output_path)


def _render_into(w, doc, preset):
    _configure_document(w, preset)
    _render_title_page(w, doc, preset)
    _render_body(w, doc, preset)
    w.set_page_number_in_footer()
    w.update_fields()


def _configure_document(w, preset):
    w.set_margins(MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT)
    w.ensure_styles(RS.STYLES)
    w.ensure_styles(RS.CHAR_STYLES)
    w.ensure_heading_styles(RS.HEADING_STYLE_MAP)
    if preset:
        _apply_preset_to_styles(w, preset)


def _render_body(w, doc, preset):
    w.insert_toc("\u76ee  \u5f55")
    scheme = detect_numbering_scheme(doc.sections)
    numbering = NumberingState(scheme=scheme)
    first_section = True
    for section in doc.sections:
        if section.level == 1 and not first_section:
            w.add_section()
        first_section = False
        _render_section(w, section, preset, numbering)


def _apply_preset_to_styles(w, preset):
    """Keep Writer heading text black for formal-document output."""
    w.apply_heading_text_color("#000000")


def _render_title_page(w, doc, preset):
    """Title page using named styles."""
    for _ in range(4):
        w.add_paragraph("", size=12)
    if doc.title:
        w.add_styled_paragraph(doc.title, "Title")
    if doc.metadata.get("author"):
        w.add_styled_paragraph(doc.metadata["author"], "Author")
    if doc.metadata.get("date"):
        w.add_styled_paragraph(doc.metadata["date"], "Date")
    w.add_horizontal_line()
    w.add_page_break()


def _render_section(w, section, preset, ns=None):
    """Render a section with intelligent heading numbering.

    ns: NumberingState instance for auto-numbering. If None, no numbering.
    """
    if section.has_heading:
        level = min(section.level, 6)

        # ---- Intelligent numbering ----
        display_text = section.heading

        if ns is not None:
            if has_manual_numbering(section.heading):
                # User wrote their own numbering — keep it as-is
                display_text = section.heading
            else:
                # Auto-generate numbering prefix
                number = ns.advance(level)
                display_text = f"{number} {section.heading}"

        w.add_heading_level(display_text, level=level)

    is_first_elem = True
    for elem in section.elements:
        _render_element(w, elem, preset, is_first_after_heading=is_first_elem)
        is_first_elem = False


def _render_element(w, elem, preset, is_first_after_heading=False):
    """Dispatch element rendering with style awareness."""
    if isinstance(elem, MDParagraph):
        _render_paragraph(w, elem, is_first_after_heading)
    elif isinstance(elem, ListBlock):
        _render_list(w, elem)
    elif isinstance(elem, TaskList):
        _render_task_list(w, elem)
    elif isinstance(elem, TableBlock):
        _render_table(w, elem, preset)
    elif isinstance(elem, CodeBlock):
        _render_code(w, elem)
    elif isinstance(elem, ImageBlock):
        _render_image(w, elem)
    elif isinstance(elem, BlockQuote):
        _render_blockquote(w, elem)
    elif isinstance(elem, HorizontalRule):
        _render_hr(w)


# ---------------------------------------------------------------------------
# Style-driven paragraph rendering
# ---------------------------------------------------------------------------


def _render_paragraph(w, para, is_first_after_heading=False):
    """Render paragraph using named styles.

    Uses "First Paragraph" style (no indent) for the first paragraph
    after a heading, "Body Text" style for subsequent paragraphs.
    """
    if not para.spans:
        return

    # Choose style: First Paragraph after heading, else Body Text
    style_name = "First Paragraph" if is_first_after_heading else "Body Text"

    w.add_rich_paragraph(para.spans, style_name)

# ---------------------------------------------------------------------------
# Lists
# ---------------------------------------------------------------------------


def _render_list(w, lst):
    """Render list using List Paragraph style as base."""
    items = [_spans_to_text(item) for item in lst.items]
    if lst.ordered:
        w.add_numbered_list(items)
    else:
        w.add_bullet_list(items)


def _render_task_list(w, tasklist):
    """Render task list with checkbox glyphs."""
    for text, checked in tasklist.items:
        glyph = "\u2611" if checked else "\u2610"
        w.add_styled_paragraph(f"{glyph}  {text}", "Body Text")


# ---------------------------------------------------------------------------
# Table
# ---------------------------------------------------------------------------


def _render_table(w, table, preset):
    w.add_paragraph("", size=4)
    data = [table.headers] + table.rows
    w.add_table(
        rows=len(data), cols=len(table.headers),
        data=data, shade_header="#D9D9D9", header_color="#000000",
        font_size=10,
        alignments=table.alignments if table.alignments else None,
        banded_rows=True, auto_fit=True, repeat_header=True,
        border_color="#BFBFBF",
    )


# ---------------------------------------------------------------------------
# Code block — uses Source Code style (shaded background)
# ---------------------------------------------------------------------------


def _render_code(w, block):
    """Render code using the Source Code named style."""
    w.add_code_lines(block.code.split("\n"))


# ---------------------------------------------------------------------------
# Image
# ---------------------------------------------------------------------------


def _render_image(w, img):
    try:
        w.add_image_block(
            img.path,
            width=img.width,
            height=img.height,
            max_width=400,
            max_height=500,
            inline=True,
            preserve_aspect=True,
            alt=img.alt,
        )
        if img.alt:
            w.add_styled_paragraph(img.alt, "Image Caption")
    except Exception:
        w.add_styled_paragraph(f"[Image: {img.alt or img.path}]", "Image Caption")


# ---------------------------------------------------------------------------
# Blockquote — uses Block Text style (border + indent)
# ---------------------------------------------------------------------------


def _render_blockquote(w, quote):
    """Render blockquote using Block Text named style."""
    for para in quote.paragraphs:
        text = _spans_to_text(para.spans)
        if text.strip():
            w.add_styled_paragraph(text, "Block Text")


# ---------------------------------------------------------------------------
# Horizontal rule
# ---------------------------------------------------------------------------


def _render_hr(w):
    """Horizontal rule via paragraph bottom border."""
    w.add_paragraph_horizontal_line()


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _spans_to_text(spans):
    return "".join(s.text for s in spans)


def _is_english(doc):
    text = doc.title
    for sec in doc.sections[:3]:
        text += sec.heading
        for elem in sec.elements[:3]:
            if isinstance(elem, MDParagraph):
                text += elem.plain_text
    cjk = sum(1 for c in text if "\u4e00" <= c <= "\u9fff")
    total = len(text.replace(" ", ""))
    return cjk < total * 0.3 if total > 0 else False
