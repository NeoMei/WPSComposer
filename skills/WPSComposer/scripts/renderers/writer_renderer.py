"""Writer renderer -- Style-driven DOCX generation (pandoc-inspired).

Key improvement over inline formatting:
  - All paragraph/character styles defined ONCE via ensure_styles()
  - Paragraphs reference styles by name, not per-paragraph font tweaking
  - "First Paragraph" style: explicit, consistent first-line indent
  - Consistent, themeable, TOC-aware output
"""

from __future__ import annotations

from ..document_model import (
    StructuredDocument, Section, Paragraph as MDParagraph, Span,
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


def render(doc, output_path, preset=None):
    """Render StructuredDocument to a professionally styled DOCX."""
    with WriterComposer() as w:
        # ---- Page setup ----
        w.set_margins(MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT)

        # ---- Define ALL named styles once (pandoc approach) ----
        w.ensure_styles(RS.STYLES)
        w.ensure_styles(RS.CHAR_STYLES)
        w.ensure_heading_styles(RS.HEADING_STYLE_MAP)

        # ---- Apply preset colors to heading styles ----
        if preset:
            _apply_preset_to_styles(w, preset)

        # ---- Title page ----
        _render_title_page(w, doc, preset)

        # ---- TOC ----
        w.insert_toc("\u76ee  \u5f55")

        # ---- Detect numbering scheme and init state ----
        scheme = detect_numbering_scheme(doc.sections)
        ns = NumberingState(scheme=scheme)

        # ---- Body ----
        first_section = True
        for section in doc.sections:
            if section.level == 1 and not first_section:
                w.add_section()
            first_section = False
            _render_section(w, section, preset, ns)

        # ---- Footer + fields ----
        w.set_page_number_in_footer()
        w.update_fields()

        return w.save_docx(output_path)


def _apply_preset_to_styles(w, preset):
    """Keep Writer heading text black for formal-document output."""
    try:
        doc = w.doc
        for level in range(1, 7):
            style = doc.Styles(-(level + 1))
            style.Font.Color = 0
    except Exception:
        pass


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

    # Fast path: plain text
    if len(para.spans) == 1 and _is_plain(para.spans[0]):
        _add_body_paragraph(w, para.spans[0].text, style_name)
        return

    # Mixed formatting: apply style first, then inline overrides
    sel = w.selection
    _apply_body_style(w, style_name)

    for span in para.spans:
        _apply_span(w, span)
        sel.TypeText(span.text)

    sel.TypeParagraph()
    w._reset_selection_to_normal()




def _add_body_paragraph(w, text, style_name):
    """Write a body paragraph with explicit font properties.

    Ensures correct font even if previous paragraph was bold/styled.
    """
    s = w.selection
    _apply_body_style(w, style_name)
    # Re-apply font properties right before typing (WPS may reset after style)
    s.Font.Bold = False
    s.Font.Italic = False
    w._set_font_family(s.Font, RS.BODY_FONT, RS.LATIN_FONT)
    s.Font.Size = 12
    s.Font.Color = 0
    s.TypeText(text)
    s.TypeParagraph()
    w._reset_selection_to_normal()


def _apply_body_style(w, style_name):
    """Apply body text style + explicit formatting overrides."""
    from .. import reference_styles as RS
    sel = w.selection
    # Clear any direct formatting carried over from previous paragraph
    try:
        sel.ClearFormatting()
    except Exception:
        pass
    try:
        sel.Style = sel.Document.Styles(style_name)
    except Exception:
        pass
    # Clear again after style change (WPS may inject formatting)
    try:
        sel.Font.Bold = False
        sel.Font.Italic = False
    except Exception:
        pass
    # Explicit overrides (defend against WPS built-in style limitations)
    props = dict(RS.STYLES["BodyText"])
    if style_name == "First Paragraph":
        props.update(RS.STYLES["FirstParagraph"])
    w._set_font_family(sel.Font, RS.BODY_FONT, RS.LATIN_FONT)
    sel.Font.Size = props.get("font_size", 12)
    sel.Font.Color = 0
    sel.Font.Bold = False
    sel.Font.Italic = False
    sel.ParagraphFormat.Alignment = props.get("align", 3)
    sel.ParagraphFormat.FirstLineIndent = props.get("indent_first", 24)
    sel.ParagraphFormat.LeftIndent = props.get("left_indent", 0)
    sel.ParagraphFormat.RightIndent = props.get("right_indent", 0)
    sel.ParagraphFormat.SpaceBefore = props.get("space_before", 0)
    sel.ParagraphFormat.SpaceAfter = props.get("space_after", 0)
    w._set_line_spacing(
        sel.ParagraphFormat,
        props.get("line_spacing"),
        props.get("line_spacing_rule", "one_and_half"),
    )


def _apply_span(w, span):
    """Apply inline character formatting on top of the paragraph style."""
    sel = w.selection
    if span.code:
        try:
            sel.Style = sel.Document.Styles("Verbatim Char")
        except Exception:
            w._set_font_family(sel.Font, RS.MONO_FONT, RS.MONO_FONT)
            sel.Font.Size = 9
    else:
        w._set_font_family(sel.Font, RS.BODY_FONT, RS.LATIN_FONT)
        sel.Font.Size = 12

    sel.Font.Bold = span.bold
    sel.Font.Italic = span.italic
    sel.Font.StrikeThrough = getattr(span, "strikethrough", False)

    if span.link:
        try:
            sel.Style = sel.Document.Styles("Hyperlink")
        except Exception:
            sel.Font.Underline = 1
        sel.Font.Color = 0
    else:
        sel.Font.Color = 0
        sel.Font.Underline = 0

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
    for line in block.code.split("\n"):
        w.add_styled_paragraph(line if line else " ", "Source Code")
    w.add_paragraph("", size=4)


# ---------------------------------------------------------------------------
# Image
# ---------------------------------------------------------------------------


def _render_image(w, img):
    try:
        shape = w.add_image(
            img.path,
            width=img.width,
            height=img.height,
            max_width=400,
            max_height=500,
            inline=True,
            preserve_aspect=True,
            alt=img.alt,
        )
        try:
            shape.Range.ParagraphFormat.Alignment = 1
            shape.Range.ParagraphFormat.KeepWithNext = -1
            shape.Range.ParagraphFormat.SpaceAfter = 0
        except Exception:
            pass
        w.selection.TypeParagraph()
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
    sel = w.selection
    sel.ParagraphFormat.Alignment = 1
    try:
        sel.ParagraphFormat.Borders(-4).LineStyle = 1
        sel.ParagraphFormat.Borders(-4).LineWidth = 6
        sel.ParagraphFormat.Borders(-4).Color = 0xC0C0C0
    except Exception:
        pass
    sel.TypeText(" ")
    sel.TypeParagraph()
    try:
        sel.ParagraphFormat.Borders(-4).LineStyle = 0
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _spans_to_text(spans):
    return "".join(s.text for s in spans)


def _is_plain(span):
    return not (span.bold or span.italic or span.code
                or span.link or getattr(span, "strikethrough", False))


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
