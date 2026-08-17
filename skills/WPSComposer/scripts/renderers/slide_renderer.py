"""Slide renderer — StructuredDocument to WPS Presentation to PPTX.

H1 sections become section-divider slides; H2 sections become content
slides.  Smart pacing: splits long content across multiple slides.
"""

from __future__ import annotations

from ..document_model import (
    StructuredDocument, Section, Paragraph,
    ListBlock, TableBlock, CodeBlock, ImageBlock, BlockQuote, TaskList,
    MathBlock,
)
from ..slide import SlideComposer
from ..math_render import latex_to_unicode

MAX_BULLETS = 6
MAX_TABLE_ROWS = 8


def render(
    doc: StructuredDocument,
    output_path: str,
    preset=None,
    composer_factory=SlideComposer,
) -> str:
    """Render a StructuredDocument to a styled PPTX file."""
    with composer_factory() as p:
        p.set_slide_size(960, 540)

        if preset:
            try:
                p.apply_design_preset(preset)
            except Exception:
                pass

        # Title slide
        if doc.title:
            subtitle = doc.metadata.get("author", "")
            if doc.metadata.get("date"):
                sep = " | " if subtitle else ""
                subtitle += sep + doc.metadata["date"]
            p.add_title_slide(doc.title, subtitle or None)

        # Content slides
        title_heading_suppressed = False
        for section in doc.sections:
            if section.level == 1:
                is_cover_title = (
                    not title_heading_suppressed
                    and section.has_heading
                    and section.heading == doc.title
                )
                if is_cover_title:
                    title_heading_suppressed = True
                if section.has_heading and not is_cover_title:
                    p.add_section_slide(section.heading)
                _render_section(p, section, preset)
            elif section.level >= 2 or section.elements:
                _render_section(p, section, preset)

        return p.save_pptx(output_path)


def _render_section(p: SlideComposer, section: Section, preset=None):
    """Render a content section, splitting long content across slides."""
    bullets = []
    tables = []
    images = []

    for elem in section.elements:
        if isinstance(elem, Paragraph):
            text = _spans_to_text(elem.spans).strip()
            if text:
                bullets.append(text)
        elif isinstance(elem, ListBlock):
            for item in elem.items:
                text = _spans_to_text(item).strip()
                if text:
                    bullets.append(text)
        elif isinstance(elem, TableBlock):
            tables.append(elem)
        elif isinstance(elem, ImageBlock):
            images.append(elem)
        elif isinstance(elem, CodeBlock):
            code = elem.code.strip()
            if code:
                bullets.append(code)
        elif isinstance(elem, TaskList):
            for text, checked in elem.items:
                glyph = "☑" if checked else "☐"
                bullets.append(f"{glyph} {text}".rstrip())
        elif isinstance(elem, BlockQuote):
            for para in elem.paragraphs:
                text = _spans_to_text(para.spans).strip()
                if text:
                    bullets.append('"' + text + '"')
        elif isinstance(elem, MathBlock):
            text = latex_to_unicode(elem.latex).strip()
            if text:
                bullets.append(text)

    if not bullets and not tables and not images:
        return

    # First bullet slide
    first = bullets[:MAX_BULLETS]
    if first:
        p.add_bullets_slide(
            section.heading if section.has_heading else "",
            first, title_size=28, body_size=18,
        )

    # Overflow bullet slides
    remaining = bullets[MAX_BULLETS:]
    while remaining:
        chunk = remaining[:MAX_BULLETS]
        remaining = remaining[MAX_BULLETS:]
        p.add_bullets_slide("(continued)", chunk, title_size=24, body_size=18)

    # Table slides
    for table in tables:
        remaining_rows = list(table.rows)
        page = 1
        while remaining_rows or (page == 1 and table.headers):
            chunk = remaining_rows[:MAX_TABLE_ROWS]
            remaining_rows = remaining_rows[MAX_TABLE_ROWS:]
            try:
                data = ([table.headers] if table.headers else []) + chunk
                cols = max(len(r) for r in data) if data else 0
                if not cols:
                    break
                rows = len(data)
                title = section.heading if section.has_heading else "Table"
                if page > 1:
                    title += " (continued)"
                p.add_bullets_slide(title, [], title_size=24, body_size=18)
                idx = p.slide_count
                p.add_table(
                    idx, rows, cols,
                    60, 120, 840, 380,
                    data,
                    header_shade=(
                        preset.get_color("primary") if preset else "#4472C4"
                    ),
                    header_font="#FFFFFF", font_size=14,
                )
            except Exception:
                pass
            page += 1

    # Image slides
    for img in images:
        try:
            p.add_blank_slide()
            idx = p.slide_count
            width = img.width
            height = img.height
            if width is None and height is None:
                width = 800
            p.add_image(
                idx,
                img.path,
                80,
                100,
                width=width,
                height=height,
            )
        except Exception:
            pass


def _spans_to_text(spans: list) -> str:
    """Join span texts, converting any math spans to Unicode."""
    parts = []
    for s in spans:
        if getattr(s, "math", ""):
            parts.append(latex_to_unicode(s.math))
        else:
            parts.append(s.text)
    return "".join(parts)
