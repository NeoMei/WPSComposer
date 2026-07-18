"""Slide renderer — StructuredDocument to WPS Presentation to PPTX.

H1 sections become section-divider slides; H2 sections become content
slides.  Smart pacing: splits long content across multiple slides.
"""

from __future__ import annotations

from ..document_model import (
    StructuredDocument, Section, Paragraph,
    ListBlock, TableBlock, CodeBlock, ImageBlock, BlockQuote,
)
from ..slide import SlideComposer

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
        for section in doc.sections:
            if section.level == 1:
                if section.has_heading:
                    p.add_section_slide(section.heading)
            elif section.level >= 2:
                _render_section(p, section, preset)

        return p.save_pptx(output_path)


def _render_section(p: SlideComposer, section: Section, preset=None):
    """Render a content section, splitting long content across slides."""
    bullets = []
    tables = []
    images = []

    for elem in section.elements:
        if isinstance(elem, Paragraph):
            text = elem.plain_text.strip()
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
        elif isinstance(elem, BlockQuote):
            for para in elem.paragraphs:
                text = _spans_to_text(para.spans).strip()
                if text:
                    bullets.append('"' + text + '"')

    if not bullets and not tables and not images:
        return

    # First bullet slide
    first = bullets[:MAX_BULLETS]
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
        try:
            rows = min(len(table.rows) + 1, MAX_TABLE_ROWS + 1)
            data = [table.headers] + table.rows[:MAX_TABLE_ROWS]
            p.add_blank_slide()
            idx = p.slide_count
            p.add_table(
                idx, rows, len(table.headers),
                60, 120, 840, 380,
                data,
                header_shade="#4472C4", header_font="#FFFFFF", font_size=14,
            )
        except Exception:
            pass

    # Image slides
    for img in images:
        try:
            p.add_blank_slide()
            idx = p.slide_count
            p.add_image(idx, img.path, 80, 100, 800, 400)
        except Exception:
            pass


def _spans_to_text(spans: list) -> str:
    return "".join(s.text for s in spans)
