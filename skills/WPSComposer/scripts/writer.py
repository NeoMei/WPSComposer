"""WriterComposer - WPS Writer (KWps.Application) to docx.

Rich-layout Word documents with columns, floating shapes, WordArt,
merged tables, TOC, and auto-populated fields.
"""

from __future__ import annotations

import re

from ._dispatch import (
    _abs, FMT_DOCX, FMT_PDF_FROM_DOC, FMT_DOC, FMT_DOCM, FMT_DOTX,
    FMT_DOTM, FMT_TXT, FMT_HTML, FMT_MHTML, FMT_RTF, FMT_XML, FMT_ODT,
    FMT_XPS,
)
from ._colors import hex_to_rgb_long
from ._base import BaseComposer
from .formatting import (
    apply_fill,
    apply_font,
    apply_geometry,
    apply_line,
    apply_paragraph,
    color_hex,
    font_snapshot,
    geometry_snapshot,
    fill_snapshot,
    line_snapshot,
    merge_results,
    page_setup_snapshot,
    paragraph_snapshot,
    safe_get,
    safe_set,
)


# ===========================================================================

class WriterComposer(BaseComposer):
    _progids = ("KWps.Application", "Wps.Application", "Word.Application")
    _doc_type = "writer"
    _native_fmt = FMT_DOCX
    _pdf_fmt = FMT_PDF_FROM_DOC
    _formats_by_extension = {
        ".docx": FMT_DOCX, ".doc": FMT_DOC, ".docm": FMT_DOCM,
        ".dotx": FMT_DOTX, ".dotm": FMT_DOTM, ".txt": FMT_TXT,
        ".html": FMT_HTML, ".htm": FMT_HTML, ".mht": FMT_MHTML,
        ".mhtml": FMT_MHTML, ".rtf": FMT_RTF, ".xml": FMT_XML,
        ".odt": FMT_ODT, ".xps": FMT_XPS,
    }

    @staticmethod
    def _create_doc(app):
        return app.Documents.Add()

    @staticmethod
    def _open_document(app, path, read_only=False):
        return app.Documents.Open(path, False, bool(read_only))

    @staticmethod
    def _active_document(app):
        return app.ActiveDocument



    @property
    def selection(self):
        return self._app.Selection

    @property
    def doc(self):
        return self._doc

    # ================================================================
    # Named style system (pandoc-inspired)
    # ================================================================

    def ensure_styles(self, styles_dict):
        """Create named Word styles from a definitions dict."""
        for key, props in styles_dict.items():
            self._ensure_one_style(key, props)

    def _ensure_one_style(self, key, props):
        """Create or update a single named style."""
        name = props.get("name", key)
        stype = props.get("type", "paragraph")
        is_char = (stype == "character")
        doc = self._doc
        try:
            style = doc.Styles(name)
        except Exception:
            try:
                style = doc.Styles.Add(name, 2 if is_char else 1)
            except Exception:
                return
        try:
            if props.get("font_name"):
                style.Font.Name = props["font_name"]
            if props.get("font_size") is not None:
                style.Font.Size = props["font_size"]
            if props.get("bold") is not None:
                style.Font.Bold = props["bold"]
            if props.get("italic") is not None:
                style.Font.Italic = props["italic"]
            if props.get("underline"):
                style.Font.Underline = 1
            if props.get("strikethrough"):
                style.Font.StrikeThrough = True
            if props.get("color"):
                style.Font.Color = hex_to_rgb_long(props["color"])
            if not is_char:
                pf = style.ParagraphFormat
                if props.get("align") is not None:
                    pf.Alignment = props["align"]
                if props.get("indent_first") is not None:
                    pf.FirstLineIndent = props["indent_first"]
                if props.get("left_indent") is not None:
                    pf.LeftIndent = props["left_indent"]
                if props.get("line_spacing") is not None:
                    pf.LineSpacing = props["line_spacing"]
                if props.get("space_before") is not None:
                    pf.SpaceBefore = props["space_before"]
                if props.get("space_after") is not None:
                    pf.SpaceAfter = props["space_after"]
                if props.get("shading"):
                    try:
                        style.Shading.BackgroundPatternColor = hex_to_rgb_long(props["shading"])
                    except Exception:
                        pass
                if props.get("left_border"):
                    try:
                        pf.Borders(-2).LineStyle = 1
                        pf.Borders(-2).LineWidth = 18
                        pf.Borders(-2).Color = hex_to_rgb_long(props.get("border_color", "#CCCCCC"))
                    except Exception:
                        pass
        except Exception:
            pass

    def add_styled_paragraph(self, text, style_name):
        """Write a paragraph using a named style."""
        s = self.selection
        try:
            s.Style = self._doc.Styles(style_name)
        except Exception:
            pass
        s.TypeText(text)
        s.TypeParagraph()

    # ---- page setup ----
    def set_columns(self, count):
        self._doc.PageSetup.TextColumns.SetCount(count)

    def set_orientation(self, landscape):
        self._doc.PageSetup.Orientation = 1 if landscape else 0

    def set_margins(self, top, bottom, left, right):
        ps = self._doc.PageSetup
        ps.TopMargin, ps.BottomMargin = top, bottom
        ps.LeftMargin, ps.RightMargin = left, right

    def set_page_size(self, width, height):
        ps = self._doc.PageSetup
        ps.PageWidth = width
        ps.PageHeight = height

    # ---- sections (independent page layout, e.g. landscape insert) ----
    def add_section(self):
        self.selection.InsertBreak(2)  # wdSectionBreakNextPage

    # ---- header / footer ----
    def set_header(self, text):
        sec = self._doc.Sections(1)
        sec.Headers(1).Range.Text = text  # wdHeaderFooterPrimary=1

    def set_footer(self, text):
        sec = self._doc.Sections(1)
        sec.Footers(1).Range.Text = text

    def set_page_number_in_footer(self):
        sec = self._doc.Sections(1)
        f = sec.Footers(1).Range
        f.Text = "Page "
        self._doc.Fields.Add(f, 33)  # wdFieldPage=33

    # ---- text ----
    def add_heading_level(self, text, level=1, size=None, color=None,
                          line_spacing=None, space_after=None):
        """Add a heading with proper Word outline level (1-6).

        Args:
            text: Heading text.
            level: 1-6 (H1-H6).
            size: Font size in points.
            color: #RRGGBB hex or BGR int.
            line_spacing: Line spacing override (points).
            space_after: Space after paragraph (points).
        """
        _sizes = {1: 16, 2: 15, 3: 15, 4: 14, 5: 13, 6: 12}
        _bold  = {1: True, 2: True, 3: True, 4: True, 5: True, 6: True}
        _align = {1: 1, 2: 3, 3: 3, 4: 3, 5: 3, 6: 3}  # H1=center, rest=justify
        _fonts = {1: "\u9ed1\u4f53", 2: "\u9ed1\u4f53", 3: "\u9ed1\u4f53",
                  4: "\u9ed1\u4f53", 5: "\u9ed1\u4f53", 6: "\u9ed1\u4f53"}

        fs = size if size is not None else _sizes.get(level, 14)
        fb = _bold.get(level, True)
        fa = _align.get(level, 3)
        ff = _fonts.get(level, "\u9ed1\u4f53")
        style_idx = -(level + 1)

        s = self.selection
        # Apply built-in heading style first
        s.Style = self._doc.Styles(style_idx)
        # Override font properties on the selection (must come after style)
        s.Font.Size = fs
        s.Font.Bold = fb
        s.Font.Name = ff
        s.Font.NameFarEast = ff
        s.ParagraphFormat.Alignment = fa
        if color is not None:
            s.Font.Color = hex_to_rgb_long(color)
        if line_spacing is not None:
            s.ParagraphFormat.LineSpacing = line_spacing
        if space_after is not None:
            s.ParagraphFormat.SpaceAfter = space_after
        s.TypeText(text)
        s.TypeParagraph()

        # Re-apply font on the just-written paragraph to ensure WPS respects it
        try:
            para = s.Paragraphs(1).Previous()
            if para:
                rng = para.Range
                rng.Font.Name = ff
                rng.Font.NameFarEast = ff
                rng.Font.Size = fs
                rng.Font.Bold = fb
        except Exception:
            pass

        # Reset to Normal
        s.Style = self._doc.Styles(-1)
        s.Font.Size = 12
        s.Font.Bold = False
        s.Font.Color = 0

    # Backward-compat aliases
    def add_heading(self, text, size=16, bold=True, color=None):
        return self.add_heading_level(text, level=1, size=size, color=color)

    def add_heading2(self, text, size=14, color=None):
        return self.add_heading_level(text, level=2, size=size, color=color)

    def add_paragraph(self, text, size=None, bold=None, italic=None,
                      color=None, align=None, indent_first=None,
                      line_spacing=None, space_after=None, font_name=None):
        s = self.selection
        if font_name is not None:
            s.Font.Name = font_name
        if size is not None:
            s.Font.Size = size
        if bold is not None:
            s.Font.Bold = bold
        if italic is not None:
            s.Font.Italic = italic
        if color is not None:
            s.Font.Color = hex_to_rgb_long(color)
        if align is not None:
            s.ParagraphFormat.Alignment = align  # 0L,1C,2R,3justify
        if indent_first is not None:
            s.ParagraphFormat.FirstLineIndent = indent_first
        if line_spacing is not None:
            s.ParagraphFormat.LineSpacing = line_spacing
        if space_after is not None:
            s.ParagraphFormat.SpaceAfter = space_after
        s.TypeText(text)
        s.TypeParagraph()
        # Reset
        s.Font.Name = ""
        s.Font.Size = 11
        s.Font.Bold = False
        s.Font.Italic = False
        s.Font.Color = 0
    def add_centered(self, text, size=24, bold=True, color=None):
        self.add_paragraph(text, size=size, bold=bold, color=color, align=1)

    def add_bullet_list(self, items, glyph="\u2022", indent=24):
        """Add a bullet list using explicit glyph prefixes.

        Uses Unicode bullet character + hanging indent for reliable
        rendering across WPS and Word (avoids dependence on
        numbering.xml which may not exist in new documents).

        Args:
            items: List of text strings.
            glyph: Bullet character (default: \u2022 = \u2022).
            indent: Left indent in points.
        """
        s = self.selection
        for it in items:
            s.ParagraphFormat.LeftIndent = indent
            s.ParagraphFormat.FirstLineIndent = -indent  # hanging indent
            s.ParagraphFormat.LineSpacing = 1.3
            s.ParagraphFormat.SpaceAfter = 3
            s.Font.Name = ""
            s.Font.Size = 11
            s.Font.Bold = False
            s.Font.Color = hex_to_rgb_long("#333333")
            # Bullet glyph + tab + text
            s.TypeText(glyph + "\t" + it)
            s.TypeParagraph()
        # Reset
        s.ParagraphFormat.LeftIndent = 0
        s.ParagraphFormat.FirstLineIndent = 0
        s.Font.Color = 0

    def add_numbered_list(self, items, indent=24):
        """Add a numbered list using explicit number prefixes.

        Args:
            items: List of text strings.
            indent: Left indent in points.
        """
        s = self.selection
        for idx, it in enumerate(items, 1):
            s.ParagraphFormat.LeftIndent = indent
            s.ParagraphFormat.FirstLineIndent = -indent  # hanging indent
            s.ParagraphFormat.LineSpacing = 1.3
            s.ParagraphFormat.SpaceAfter = 3
            s.Font.Name = ""
            s.Font.Size = 11
            s.Font.Bold = False
            s.Font.Color = hex_to_rgb_long("#333333")
            # Number + dot + tab + text
            s.TypeText(f"{idx}.\t{it}")
            s.TypeParagraph()
        # Reset
        s.ParagraphFormat.LeftIndent = 0
        s.ParagraphFormat.FirstLineIndent = 0
        s.Font.Color = 0

    # ---- table ----
    def add_table(self, rows, cols, data, shade_header="#4472C4",
                    header_color="#FFFFFF", font_size=10,
                    col_widths=None, alignments=None,
                    banded_rows=True, auto_fit=True,
                    repeat_header=True, border_color="#D0D0D0"):
        """Add a professionally styled table.

        Args:
            rows, cols: Table dimensions.
            data: List of rows, each row a list of cell values.
            shade_header: Header background colour (#RRGGBB).
            header_color: Header text colour.
            font_size: Body font size in points.
            col_widths: Optional list of column widths in points.
            alignments: Optional list of "left"/"center"/"right" per column.
            banded_rows: Enable alternating row shading.
            auto_fit: Auto-fit table to page width.
            repeat_header: Repeat header row on each page.
            border_color: Grid line colour.

        Returns:
            The COM Table object.
        """
        s = self.selection
        tbl = self._doc.Tables.Add(s.Range, rows, cols)

        # ---- Cell population ----
        for r in range(rows):
            for c in range(cols):
                if r >= len(data) or c >= len(data[r]):
                    continue
                cell = tbl.Cell(r + 1, c + 1)
                cell.Range.Text = str(data[r][c])
                cell.Range.Font.Size = font_size

                # Cell padding
                try:
                    cell.TopPadding = 3
                    cell.BottomPadding = 3
                    cell.LeftPadding = 6
                    cell.RightPadding = 6
                except Exception:
                    pass

                # Vertical center
                try:
                    cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter
                except Exception:
                    pass

                if r == 0:
                    # Header row
                    cell.Range.Font.Bold = True
                    cell.Range.Font.Size = font_size + 1
                    if header_color:
                        cell.Range.Font.Color = hex_to_rgb_long(header_color)
                    if shade_header:
                        cell.Shading.BackgroundPatternColor = hex_to_rgb_long(shade_header)
                else:
                    # Alternating row shading
                    if banded_rows and r % 2 == 0:
                        try:
                            cell.Shading.BackgroundPatternColor = 0xF2F2F2  # light gray BGR
                        except Exception:
                            pass

        # ---- Alignment per column ----
        if alignments:
            for c_idx, align in enumerate(alignments):
                if c_idx >= cols:
                    break
                if align not in ("left", "center", "right"):
                    continue
                wd_align = {"left": 0, "center": 1, "right": 2}.get(align, 0)
                try:
                    tbl.Columns(c_idx + 1).Select()
                    s.ParagraphFormat.Alignment = wd_align
                except Exception:
                    pass

        # ---- Column widths ----
        if col_widths:
            for c_idx, w in enumerate(col_widths):
                if c_idx >= cols:
                    break
                try:
                    tbl.Columns(c_idx + 1).Width = w
                except Exception:
                    pass

        # ---- Borders ----
        border_color_rgb = hex_to_rgb_long(border_color)
        try:
            for border_idx in range(-6, 0):  # all wdBorder* constants
                try:
                    tbl.Borders(border_idx).LineStyle = 1      # wdLineStyleSingle
                    tbl.Borders(border_idx).Color = border_color_rgb
                    tbl.Borders(border_idx).LineWidth = 2      # wdLineWidth025pt
                except Exception:
                    pass
        except Exception:
            pass

        # ---- Auto-fit ----
        if auto_fit:
            try:
                tbl.AutoFitBehavior(2)  # wdAutoFitWindow
            except Exception:
                pass

        # ---- Repeat header across pages ----
        if repeat_header and rows > 1:
            try:
                tbl.Rows(1).HeadingFormat = -1  # True
            except Exception:
                pass

        # Move cursor past table
        s.EndKey(6)  # wdStory
        s.InsertBreak(7)  # wdLineBreak
        return tbl

    def add_merged_table(self, data, merges=None, shade_header="#4472C4"):
        """data: list of rows (lists). merges: list of (r1,c1,r2,c2) 1-based."""
        rows = len(data)
        cols = max(len(r) for r in data)
        tbl = self._doc.Tables.Add(self.selection.Range, rows, cols)
        for r in range(rows):
            for c in range(len(data[r])):
                cell = tbl.Cell(r + 1, c + 1)
                cell.Range.Text = str(data[r][c])
                cell.Range.Font.Size = 10
                if r == 0:
                    cell.Range.Font.Bold = True
                    cell.Shading.BackgroundPatternColor = hex_to_rgb_long(shade_header)
        if merges:
            for r1, c1, r2, c2 in merges:
                try:
                    tbl.Cell(r1, c1).Merge(tbl.Cell(r2, c2))
                except Exception:
                    pass
        self.selection.EndKey(6)
        return tbl

    # ---- shapes ----
    def add_floating_textbox(self, text, left, top, width, height,
                             wrap=0, fill_color=None, font_size=11,
                             bold=False):
        shape = self._doc.Shapes.AddTextbox(1, left, top, width, height)
        tf = shape.TextFrame
        tf.TextRange.Text = text
        tf.TextRange.Font.Size = font_size
        tf.TextRange.Font.Bold = bold
        if fill_color:
            try:
                shape.Fill.ForeColor.RGB = hex_to_rgb_long(fill_color)
                shape.Fill.Visible = 1
            except Exception:
                pass
        try:
            shape.WrapFormat.Type = wrap
        except Exception:
            pass
        return shape

    def add_wordart(self, text, left=200, top=400, preset=0):
        return self._doc.Shapes.AddTextEffect(preset, text, "Arial", 36,
                                               False, False, left, top)

    def add_image(self, path, width=None, height=None, wrap=0):
        p = _abs(path)
        if width and height:
            shape = self._doc.Shapes.AddPicture(p, False, True, 0, 0, width, height)
        else:
            shape = self._doc.Shapes.AddPicture(p, False, True)
        try:
            shape.WrapFormat.Type = wrap
        except Exception:
            pass
        return shape

    def add_page_break(self):
        self.selection.InsertBreak(7)

    def add_horizontal_line(self):
        s = self.selection
        s.InlineShapes.AddHorizontalLineStandard()

    def insert_toc(self, title="Table of Contents"):
        s = self.selection
        if title:
            # Use Heading 1 style for TOC title so it doesn't inherit
            # the previous paragraph style (e.g. Date)
            try:
                s.Style = self._doc.Styles(-2)  # wdStyleHeading1
                s.Font.Size = 18
                s.Font.Bold = False
                s.Font.Color = 0
                s.ParagraphFormat.Alignment = 1  # center
                s.ParagraphFormat.LineSpacing = 18.0
                s.ParagraphFormat.SpaceAfter = 5
            except Exception:
                pass
            s.TypeText(title)
            s.TypeParagraph()
            # Reset
            try:
                s.Style = self._doc.Styles(-1)
            except Exception:
                pass
            s.Font.Size = 12
            s.Font.Color = 0
        self._doc.TablesOfContents.Add(s.Range, True, 1, 3)
        self.selection.EndKey(6)
        self.selection.InsertBreak(7)

    def update_fields(self):
        try:
            for i in range(1, self._doc.TablesOfContents.Count + 1):
                self._doc.TablesOfContents.Item(i).Update()
        except Exception:
            pass
        try:
            self._doc.Fields.Update()
        except Exception:
            pass

    # ================================================================
    # Existing-document inspection and element-level editing
    # ================================================================

    def inspect_selection(self):
        """Return the active Writer selection as a JSON-compatible snapshot."""
        rng = self.selection.Range
        return self._range_snapshot(rng, "selection")

    def inspect_document(self, include_text=True, max_elements=None):
        """Describe document structure and formatting for an agent.

        Element ids are stable within the current editing session and can be
        passed to :meth:`apply_format_patch` (for example ``paragraph:3`` or
        ``table:1/cell:2,1``).
        """
        doc = self._doc
        paragraph_count = int(safe_get(doc.Paragraphs, "Count", 0) or 0)
        table_count = int(safe_get(doc.Tables, "Count", 0) or 0)
        shape_count = int(safe_get(doc.Shapes, "Count", 0) or 0)
        section_count = int(safe_get(doc.Sections, "Count", 0) or 0)
        limit = max_elements if max_elements is not None else float("inf")

        paragraphs = []
        for index in range(1, min(paragraph_count, int(limit) if limit != float("inf") else paragraph_count) + 1):
            para = doc.Paragraphs(index)
            snap = self._range_snapshot(para.Range, f"paragraph:{index}")
            snap["index"] = index
            snap["style"] = self._style_name(safe_get(para.Range, "Style"))
            if not include_text:
                snap.pop("text", None)
            paragraphs.append(snap)

        tables = []
        for table_index in range(1, min(table_count, int(limit) if limit != float("inf") else table_count) + 1):
            table = doc.Tables(table_index)
            cells = []
            cell_count = int(safe_get(table.Range.Cells, "Count", 0) or 0)
            for cell_index in range(1, cell_count + 1):
                cell = table.Range.Cells(cell_index)
                row = int(safe_get(cell, "RowIndex", 0) or 0)
                col = int(safe_get(cell, "ColumnIndex", 0) or 0)
                cell_snap = self._range_snapshot(
                    cell.Range, f"table:{table_index}/cell:{row},{col}"
                )
                cell_snap.update({
                    "row": row,
                    "column": col,
                    "vertical_alignment": safe_get(cell, "VerticalAlignment"),
                    "width": safe_get(cell, "Width"),
                    "fill": {
                        "color": color_hex(safe_get(
                            safe_get(cell, "Shading"), "BackgroundPatternColor"
                        ))
                    },
                })
                if not include_text:
                    cell_snap.pop("text", None)
                cells.append(cell_snap)
            tables.append({
                "id": f"table:{table_index}",
                "index": table_index,
                "rows": int(safe_get(table.Rows, "Count", 0) or 0),
                "columns": int(safe_get(table.Columns, "Count", 0) or 0),
                "allow_autofit": safe_get(table, "AllowAutoFit"),
                "cells": cells,
            })

        shapes = []
        for index in range(1, min(shape_count, int(limit) if limit != float("inf") else shape_count) + 1):
            shapes.append(self._shape_snapshot(doc.Shapes(index), index, include_text))

        sections = []
        for index in range(1, section_count + 1):
            section = doc.Sections(index)
            columns = safe_get(safe_get(section.PageSetup, "TextColumns"), "Count")
            sections.append({
                "id": f"section:{index}",
                "index": index,
                "page_setup": page_setup_snapshot(section.PageSetup),
                "columns": columns,
            })

        return {
            "kind": "writer",
            "name": safe_get(doc, "Name"),
            "path": safe_get(doc, "FullName"),
            "saved": safe_get(doc, "Saved"),
            "counts": {
                "paragraphs": paragraph_count,
                "tables": table_count,
                "shapes": shape_count,
                "sections": section_count,
                "inline_shapes": int(safe_get(doc.InlineShapes, "Count", 0) or 0),
                "styles": int(safe_get(doc.Styles, "Count", 0) or 0),
            },
            "sections": sections,
            "paragraphs": paragraphs,
            "tables": tables,
            "shapes": shapes,
        }

    def apply_format_patch(self, target, *, text=None, font=None,
                           paragraph=None, geometry=None, fill=None, line=None,
                           style=None, wrap=None, vertical_alignment=None,
                           page_setup=None, columns=None):
        """Apply a partial formatting patch to an addressed Writer element.

        Supported targets: ``selection``, ``paragraph:N``, ``range:S-E``,
        ``table:N/cell:R,C``, ``shape:N``, and ``section:N``.
        Unspecified formatting is preserved.
        """
        font = font or {}
        paragraph = paragraph or {}
        geometry = geometry or {}
        fill = fill or {}
        line = line or {}

        if target == "selection":
            rng = self.selection.Range
            return self._patch_range(rng, text, font, paragraph, style)

        match = re.fullmatch(r"paragraph:(\d+)", target)
        if match:
            rng = self._doc.Paragraphs(int(match.group(1))).Range
            return self._patch_range(rng, text, font, paragraph, style)

        match = re.fullmatch(r"range:(\d+)-(\d+)", target)
        if match:
            rng = self._doc.Range(int(match.group(1)), int(match.group(2)))
            return self._patch_range(rng, text, font, paragraph, style)

        match = re.fullmatch(r"table:(\d+)/cell:(\d+),(\d+)", target)
        if match:
            table = self._doc.Tables(int(match.group(1)))
            cell = table.Cell(int(match.group(2)), int(match.group(3)))
            result = self._patch_range(cell.Range, text, font, paragraph, style)
            if vertical_alignment is not None:
                ok = safe_set(cell, "VerticalAlignment", vertical_alignment)
                result["accepted" if ok else "rejected"].append("vertical_alignment")
            if fill and "color" in fill:
                shading = safe_get(cell, "Shading")
                ok = shading is not None and safe_set(
                    shading, "BackgroundPatternColor", hex_to_rgb_long(fill["color"])
                )
                result["accepted" if ok else "rejected"].append("fill.color")
            return result

        match = re.fullmatch(r"shape:(\d+)", target)
        if match:
            shape = self._doc.Shapes(int(match.group(1)))
            results = [
                self._prefix_result(apply_geometry(shape, geometry), "geometry"),
                self._prefix_result(apply_fill(shape.Fill, fill), "fill"),
                self._prefix_result(apply_line(shape.Line, line), "line"),
            ]
            text_range = safe_get(safe_get(shape, "TextFrame"), "TextRange")
            if text_range is not None:
                results.append(self._prefix_result(apply_font(text_range.Font, font), "font"))
                if text is not None:
                    ok = safe_set(text_range, "Text", text)
                    results.append({"accepted": ["text"] if ok else [],
                                    "rejected": [] if ok else ["text"]})
            elif text is not None or font:
                results.append({"accepted": [], "rejected": ["text_frame"]})
            if wrap is not None:
                ok = safe_set(safe_get(shape, "WrapFormat"), "Type", wrap)
                results.append({"accepted": ["wrap"] if ok else [],
                                "rejected": [] if ok else ["wrap"]})
            return merge_results(*results)

        match = re.fullmatch(r"section:(\d+)", target)
        if match:
            section = self._doc.Sections(int(match.group(1)))
            accepted, rejected = [], []
            mapping = {
                "orientation": "Orientation", "page_width": "PageWidth",
                "page_height": "PageHeight", "top_margin": "TopMargin",
                "bottom_margin": "BottomMargin", "left_margin": "LeftMargin",
                "right_margin": "RightMargin", "header_distance": "HeaderDistance",
                "footer_distance": "FooterDistance", "gutter": "Gutter",
                "paper_size": "PaperSize",
            }
            for key, value in (page_setup or {}).items():
                ok = key in mapping and safe_set(section.PageSetup, mapping[key], value)
                (accepted if ok else rejected).append(f"page_setup.{key}")
            if columns is not None:
                tc = safe_get(section.PageSetup, "TextColumns")
                try:
                    tc.SetCount(columns)
                    accepted.append("columns")
                except Exception:
                    rejected.append("columns")
            return {"accepted": accepted, "rejected": rejected}

        raise ValueError(f"Unsupported Writer target: {target}")

    def _range_snapshot(self, rng, element_id):
        text = str(safe_get(rng, "Text", "")).rstrip("\r\x07")
        return {
            "id": element_id,
            "text": text,
            "start": safe_get(rng, "Start"),
            "end": safe_get(rng, "End"),
            "font": font_snapshot(safe_get(rng, "Font")),
            "paragraph": paragraph_snapshot(safe_get(rng, "ParagraphFormat")),
        }

    def _shape_snapshot(self, shape, index, include_text):
        text_range = safe_get(safe_get(shape, "TextFrame"), "TextRange")
        result = {
            "id": f"shape:{index}",
            "index": index,
            "name": safe_get(shape, "Name"),
            "type": safe_get(shape, "Type"),
            "geometry": geometry_snapshot(shape),
            "fill": fill_snapshot(safe_get(shape, "Fill")),
            "line": line_snapshot(safe_get(shape, "Line")),
            "wrap": safe_get(safe_get(shape, "WrapFormat"), "Type"),
        }
        if text_range is not None:
            if include_text:
                result["text"] = safe_get(text_range, "Text", "")
            result["font"] = font_snapshot(safe_get(text_range, "Font"))
        return result

    def _patch_range(self, rng, text, font, paragraph, style):
        results = [
            self._prefix_result(apply_font(rng.Font, font), "font"),
            self._prefix_result(apply_paragraph(rng.ParagraphFormat, paragraph), "paragraph"),
        ]
        if text is not None:
            ok = safe_set(rng, "Text", text)
            results.append({"accepted": ["text"] if ok else [],
                            "rejected": [] if ok else ["text"]})
        if style is not None:
            try:
                rng.Style = self._doc.Styles(style)
                ok = True
            except Exception:
                ok = False
            results.append({"accepted": ["style"] if ok else [],
                            "rejected": [] if ok else ["style"]})
        return merge_results(*results)

    @staticmethod
    def _prefix_result(result, prefix):
        return {
            "accepted": [f"{prefix}.{key}" for key in result["accepted"]],
            "rejected": [f"{prefix}.{key}" for key in result["rejected"]],
        }

    @staticmethod
    def _style_name(style):
        if style is None:
            return None
        return str(safe_get(style, "NameLocal", safe_get(style, "Name", style)))

    def save_docx(self, path):
        return self.save(path, FMT_DOCX)

    def export_pdf(self, path):
        p = _abs(path)
        self._doc.ExportAsFixedFormat(p, FMT_PDF_FROM_DOC)
        return p


# ===========================================================================
#  SHEET  (xlsx)
# ===========================================================================

