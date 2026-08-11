"""Writer renderer -- Style-driven DOCX generation (pandoc-inspired).

Key improvement over inline formatting:
  - All paragraph/character styles defined ONCE via ensure_styles()
  - Paragraphs reference styles by name, not per-paragraph font tweaking
  - "First Paragraph" style: explicit, consistent first-line indent
  - Consistent, themeable, TOC-aware output
"""

from __future__ import annotations

from ..document_model import (
    StructuredDocument, Section, Span, Paragraph as MDParagraph,
    ListBlock, TableBlock, CodeBlock, ImageBlock, BlockQuote,
    HorizontalRule, TaskList, ExcalidrawBlock, MathBlock,
)
from ..writer import WriterComposer
from .. import reference_styles as RS
from ..heading_numbering import (
    NumberingState, has_manual_numbering, strip_manual_numbering,
    detect_numbering_scheme,
)
from ..math_render import latex_to_unicode, latex_to_png


# Standard A4 margins (matches formal Chinese doc conventions)
MARGIN_TOP = 72       # 1 inch = 2.54 cm
MARGIN_BOTTOM = 72
MARGIN_LEFT = 90      # ~3.17 cm (Word default)
MARGIN_RIGHT = 90

# Vertical space above the cover title block, in points. A4 usable height is
# ~698pt; 260pt places the title just above the vertical centre.
COVER_TOP_SPACING = 260


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
    """Title page using named styles: title, author, date.

    A single exactly-spaced spacer paragraph positions the title block just
    above the vertical centre of the page. The date accepts either the formal
    ``date`` key or the Obsidian ``created`` front-matter key.
    """
    w.add_paragraph(
        "", line_spacing=COVER_TOP_SPACING, line_spacing_rule="exact"
    )
    if doc.title:
        w.add_styled_paragraph(doc.title, "Title")
    if doc.metadata.get("author"):
        w.add_styled_paragraph(doc.metadata["author"], "Author")
    date = doc.metadata.get("date") or doc.metadata.get("created")
    if date:
        w.add_styled_paragraph(date, "Date")
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
    pending: list = []  # (text, style_name) for coalesced plain paragraphs

    def _flush_pending():
        if not pending:
            return
        # One op inserts every coalesced paragraph (\r-separated) and the
        # style applies to the whole range — a large win over per-paragraph
        # WPS API calls on big documents.
        w.add_paragraph("\r".join(text for text, _ in pending), style=pending[0][1])
        pending.clear()

    for elem in section.elements:
        is_plain_paragraph = (
            isinstance(elem, MDParagraph)
            and all(
                not (span.bold or span.italic or span.code or span.link or span.strikethrough or span.math)
                for span in elem.spans
            )
            and elem.align == 0
        )
        if is_plain_paragraph:
            style_name = "First Paragraph" if is_first_elem else "Body Text"
            if pending and pending[0][1] != style_name:
                _flush_pending()
            pending.append((elem.plain_text, style_name))
        else:
            _flush_pending()
            _render_element(w, elem, preset, is_first_after_heading=is_first_elem)
        is_first_elem = False
    _flush_pending()


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
    elif isinstance(elem, ExcalidrawBlock):
        _render_excalidraw(w, elem)
    elif isinstance(elem, MathBlock):
        _render_math_block(w, elem)
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
    Converts inline math spans ($...$) to Unicode text before rendering.
    """
    if not para.spans:
        return

    # Convert math spans to Unicode text spans
    resolved_spans = []
    for s in para.spans:
        if s.math:
            resolved_spans.append(Span(text=latex_to_unicode(s.math), italic=True))
        else:
            resolved_spans.append(s)

    # Choose style: First Paragraph after heading, else Body Text
    style_name = "First Paragraph" if is_first_after_heading else "Body Text"

    w.add_rich_paragraph(resolved_spans, style_name)

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
    data = ([table.headers] if table.headers else []) + table.rows
    cols = max(len(r) for r in data) if data else 0
    if not cols:
        return
    w.add_table(
        rows=len(data), cols=cols,
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
# Excalidraw
# ---------------------------------------------------------------------------


def _render_excalidraw(w, block):
    """Render an Excalidraw diagram to SVG and embed it."""
    try:
        svg_path = _render_excalidraw_to_svg(block.path, block.width, block.height)
        if svg_path:
            w.add_image_block(
                svg_path,
                width=block.width,
                height=block.height,
                max_width=600,
                max_height=500,
                inline=True,
                preserve_aspect=True,
                alt=block.alt,
            )
            if block.alt:
                w.add_styled_paragraph(block.alt, "Image Caption")
        else:
            w.add_styled_paragraph(f"[Excalidraw: {block.alt or block.path}]", "Image Caption")
    except Exception as e:
        w.add_styled_paragraph(f"[Excalidraw error: {block.alt or block.path}]", "Image Caption")


def _render_excalidraw_to_svg(excalidraw_path, width=None, height=None):
    """Render an Excalidraw .excalidraw.md file to SVG.
    
    Returns the path to the rendered SVG file, or None on failure.
    """
    import re
    import json
    import tempfile
    import os
    
    try:
        import lzstring
    except ImportError:
        return None
    
    # Read the .excalidraw.md file
    with open(excalidraw_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Extract the compressed JSON from ```compressed-json block
    json_match = re.search(r'```compressed-json\s*\n(.*?)\n```', content, re.DOTALL)
    if not json_match:
        return None
    
    compressed_data = json_match.group(1).strip()
    if not compressed_data:
        return None
    
    # Remove newlines and whitespace from compressed data (LZString doesn't handle them)
    compressed_data = re.sub(r'\s+', '', compressed_data)
    
    # Decompress using LZString
    try:
        lz = lzstring.LZString()
        decompressed = lz.decompressFromBase64(compressed_data)
        if not decompressed:
            return None
        scene_data = json.loads(decompressed)
    except Exception:
        return None
    
    # Render to SVG using simple Python-based renderer
    try:
        svg_content = _render_excalidraw_scene_to_svg(scene_data)
        if not svg_content:
            return None
        
        # Save to temp file
        temp_dir = tempfile.gettempdir()
        svg_path = os.path.join(temp_dir, f"excalidraw_{os.path.basename(excalidraw_path)}.svg")
        with open(svg_path, 'w', encoding='utf-8') as f:
            f.write(svg_content)
        
        return svg_path
    except Exception:
        return None


def _render_excalidraw_scene_to_svg(scene_data):
    """Convert Excalidraw scene data to SVG string."""
    elements = scene_data.get('elements', [])
    if not elements:
        return None
    
    # Calculate bounding box
    min_x = min_y = float('inf')
    max_x = max_y = float('-inf')
    
    for el in elements:
        if el.get('isDeleted'):
            continue
        x, y = el.get('x', 0), el.get('y', 0)
        w, h = el.get('width', 0), el.get('height', 0)
        min_x = min(min_x, x)
        min_y = min(min_y, y)
        max_x = max(max_x, x + w)
        max_y = max(max_y, y + h)
    
    if min_x == float('inf'):
        return None
    
    # Add padding
    padding = 20
    min_x -= padding
    min_y -= padding
    max_x += padding
    max_y += padding
    
    width = max_x - min_x
    height = max_y - min_y
    
    # Build SVG
    svg_parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{min_x} {min_y} {width} {height}" width="{width}" height="{height}">',
        '<rect width="100%" height="100%" fill="white"/>',
        '<defs><marker id="arrowhead" markerWidth="10" markerHeight="10" refX="9" refY="3" orient="auto"><polygon points="0 0, 10 3, 0 6" fill="#000000"/></marker></defs>',
    ]
    
    for el in elements:
        if el.get('isDeleted'):
            continue
        
        el_type = el.get('type', '')
        x = el.get('x', 0)
        y = el.get('y', 0)
        w = el.get('width', 0)
        h = el.get('height', 0)
        stroke = el.get('strokeColor', '#000000')
        fill = el.get('backgroundColor', 'transparent')
        stroke_width = el.get('strokeWidth', 1)
        opacity = el.get('opacity', 100) / 100.0
        text = el.get('text', '')
        font_size = el.get('fontSize', 20)
        roundness = el.get('roundness', {}).get('type') if el.get('roundness') else None
        
        # Common attributes
        attrs = f'fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width}" opacity="{opacity}"'
        
        if el_type == 'rectangle':
            rx = '5' if roundness else '0'
            svg_parts.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" {attrs}/>')
        
        elif el_type == 'ellipse':
            cx, cy = x + w/2, y + h/2
            rx, ry = w/2, h/2
            svg_parts.append(f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" {attrs}/>')
        
        elif el_type == 'diamond':
            points = f"{x+w/2},{y} {x+w},{y+h/2} {x+w/2},{y+h} {x},{y+h/2}"
            svg_parts.append(f'<polygon points="{points}" {attrs}/>')
        
        elif el_type == 'text':
            # Escape XML special characters
            text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            svg_parts.append(f'<text x="{x}" y="{y + font_size}" font-size="{font_size}" fill="{stroke}" font-family="Arial, sans-serif">{text}</text>')
        
        elif el_type == 'line':
            points = el.get('points', [[0, 0], [w, h]])
            if len(points) >= 2:
                path_data = f"M {x + points[0][0]} {y + points[0][1]}"
                for pt in points[1:]:
                    path_data += f" L {x + pt[0]} {y + pt[1]}"
                svg_parts.append(f'<path d="{path_data}" fill="none" {attrs}/>')
        
        elif el_type == 'arrow':
            points = el.get('points', [[0, 0], [w, h]])
            if len(points) >= 2:
                path_data = f"M {x + points[0][0]} {y + points[0][1]}"
                for pt in points[1:]:
                    path_data += f" L {x + pt[0]} {y + pt[1]}"
                svg_parts.append(f'<path d="{path_data}" fill="none" {attrs} marker-end="url(#arrowhead)"/>')
        
        elif el_type == 'freedraw':
            points = el.get('points', [])
            if points:
                path_data = f"M {x + points[0][0]} {y + points[0][1]}"
                for pt in points[1:]:
                    path_data += f" L {x + pt[0]} {y + pt[1]}"
                svg_parts.append(f'<path d="{path_data}" fill="none" {attrs} stroke-linecap="round" stroke-linejoin="round"/>')
    
    svg_parts.append('</svg>')
    
    return '\n'.join(svg_parts)


# ---------------------------------------------------------------------------
# Display math block — render as centered PNG
# ---------------------------------------------------------------------------


def _render_math_block(w, block):
    """Render display math ($$...$$) as a centered PNG image."""
    try:
        png_path = latex_to_png(block.latex)
        w.add_image_block(
            png_path,
            max_width=450,
            max_height=300,
            inline=True,
            preserve_aspect=True,
            alt=block.latex,
        )
    except Exception:
        # Fallback: render as styled text
        w.add_styled_paragraph(latex_to_unicode(block.latex), "Block Text")


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
    """Join span texts, converting any math spans to Unicode."""
    parts = []
    for s in spans:
        if s.math:
            parts.append(latex_to_unicode(s.math))
        else:
            parts.append(s.text)
    return "".join(parts)


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
