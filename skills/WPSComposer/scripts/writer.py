"""WriterComposer - WPS Writer (KWps.Application) to docx.

Rich-layout Word documents with columns, floating shapes, WordArt,
merged tables, TOC, and auto-populated fields.
"""

from __future__ import annotations

import re
import unicodedata

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

_LINE_SPACING_RULES = {
    "single": 0,
    "one_and_half": 1,
    "double": 2,
    "at_least": 3,
    "exact": 4,
    "multiple": 5,
}


def _visual_text_width(value):
    """Estimate display width with CJK characters counted as two columns."""
    lines = str(value or "").splitlines() or [""]
    return max(
        sum(2 if unicodedata.east_asian_width(ch) in "WFA" else 1 for ch in line)
        for line in lines
    )


def _content_column_widths(data, cols, available_width):
    """Return content-aware column widths in points.

    A 75th-percentile sample prevents one unusually long cell from consuming
    the whole table.  Bounds keep identifier columns usable while allowing a
    narrative column to take most of the width.
    """
    if cols <= 0:
        return []

    representatives = []
    for col in range(cols):
        samples = [
            _visual_text_width(row[col])
            for row in data
            if col < len(row) and str(row[col]).strip()
        ]
        if not samples:
            representatives.append(4.0)
            continue
        header = samples[0] * 1.15
        samples.sort()
        q75 = samples[round((len(samples) - 1) * 0.75)]
        representatives.append(max(4.0, min(80.0, max(header, q75))))

    weights = [value ** 0.7 for value in representatives]
    minimum = {1: 1.0, 2: 0.24, 3: 0.15, 4: 0.11, 5: 0.085}.get(
        cols, min(0.08, 0.6 / cols)
    )
    maximum = 0.72 if cols == 2 else 0.65
    ratios = [0.0] * cols
    active = set(range(cols))
    remaining = 1.0

    while active:
        total_weight = sum(weights[index] for index in active) or len(active)
        trial = {
            index: remaining * weights[index] / total_weight
            for index in active
        }
        constrained = False
        for index in list(active):
            if trial[index] < minimum:
                ratios[index] = minimum
            elif trial[index] > maximum:
                ratios[index] = maximum
            else:
                continue
            remaining -= ratios[index]
            active.remove(index)
            constrained = True
        if not constrained:
            for index in active:
                ratios[index] = trial[index]
            break

    total = sum(ratios) or 1.0
    return [available_width * ratio / total for ratio in ratios]


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

    def ensure_heading_styles(self, styles_by_level):
        """Configure built-in Heading 1-6 styles from one source of truth."""
        for level, props in styles_by_level.items():
            try:
                style = self._doc.Styles(-(int(level) + 1))
            except Exception:
                continue
            self._configure_style(style, props, is_char=False)

    def apply_heading_text_color(self, color):
        """Apply one text color to the built-in Heading 1-6 styles."""
        value = hex_to_rgb_long(color) if isinstance(color, str) else color
        try:
            for level in range(1, 7):
                self._doc.Styles(-(level + 1)).Font.Color = value
        except Exception:
            pass

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
        self._configure_style(style, props, is_char=is_char)

    def _set_font_family(self, font, east_asia, ascii_name=None):
        """Set every Word font slot so Chinese/Latin runs cannot fall back."""
        if not east_asia:
            return
        ascii_name = ascii_name or east_asia
        safe_set(font, "Name", east_asia)
        safe_set(font, "NameFarEast", east_asia)
        safe_set(font, "NameAscii", ascii_name)
        safe_set(font, "NameOther", ascii_name)
        safe_set(font, "NameBi", ascii_name)

    def _set_line_spacing(self, paragraph_format, value=None, rule=None):
        """Apply semantic Word line spacing without confusing lines and points."""
        if rule is None:
            if value is not None:
                safe_set(paragraph_format, "LineSpacing", value)
            return
        code = _LINE_SPACING_RULES.get(rule, rule)
        safe_set(paragraph_format, "LineSpacingRule", code)
        if code in (3, 4) and value is not None:
            safe_set(paragraph_format, "LineSpacing", value)
        elif code == 5 and value is not None:
            try:
                points = self._app.LinesToPoints(float(value))
            except Exception:
                points = float(value) * 12.0
            safe_set(paragraph_format, "LineSpacing", points)

    def _configure_style(self, style, props, is_char=False):
        """Apply a style definition, including base style and CJK font slots."""
        if props.get("based_on"):
            try:
                style.BaseStyle = self._doc.Styles(props["based_on"])
            except Exception:
                try:
                    style.BaseStyle = props["based_on"]
                except Exception:
                    pass
        try:
            if props.get("font_name"):
                self._set_font_family(
                    style.Font, props["font_name"], props.get("font_name_ascii")
                )
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
                if (props.get("line_spacing") is not None
                        or props.get("line_spacing_rule") is not None):
                    self._set_line_spacing(
                        pf,
                        props.get("line_spacing"),
                        props.get("line_spacing_rule"),
                    )
                if props.get("space_before") is not None:
                    pf.SpaceBefore = props["space_before"]
                if props.get("space_after") is not None:
                    pf.SpaceAfter = props["space_after"]
                if props.get("right_indent") is not None:
                    pf.RightIndent = props["right_indent"]
                if props.get("keep_together") is not None:
                    pf.KeepTogether = props["keep_together"]
                if props.get("keep_with_next") is not None:
                    pf.KeepWithNext = props["keep_with_next"]
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
        self._reset_selection_to_normal()
        try:
            s.Style = self._doc.Styles(style_name)
        except Exception:
            pass
        s.TypeText(text)
        s.TypeParagraph()
        self._reset_selection_to_normal()

    def _apply_body_style(self, style_name):
        """Apply the renderer's defensive body-style formatting."""
        from . import reference_styles as RS

        selection = self.selection
        try:
            selection.ClearFormatting()
        except Exception:
            pass
        try:
            selection.Style = selection.Document.Styles(style_name)
        except Exception:
            pass
        try:
            selection.Font.Bold = False
            selection.Font.Italic = False
        except Exception:
            pass
        props = dict(RS.STYLES["BodyText"])
        if style_name == "First Paragraph":
            props.update(RS.STYLES["FirstParagraph"])
        self._set_font_family(selection.Font, RS.BODY_FONT, RS.LATIN_FONT)
        selection.Font.Size = props.get("font_size", 12)
        selection.Font.Color = 0
        selection.Font.Bold = False
        selection.Font.Italic = False
        selection.ParagraphFormat.Alignment = props.get("align", 3)
        selection.ParagraphFormat.FirstLineIndent = props.get("indent_first", 24)
        selection.ParagraphFormat.LeftIndent = props.get("left_indent", 0)
        selection.ParagraphFormat.RightIndent = props.get("right_indent", 0)
        selection.ParagraphFormat.SpaceBefore = props.get("space_before", 0)
        selection.ParagraphFormat.SpaceAfter = props.get("space_after", 0)
        self._set_line_spacing(
            selection.ParagraphFormat,
            props.get("line_spacing"),
            props.get("line_spacing_rule", "one_and_half"),
        )

    def _apply_span_format(self, span):
        """Apply one renderer span to the active Writer selection."""
        from . import reference_styles as RS

        selection = self.selection
        if span.code:
            try:
                selection.Style = selection.Document.Styles("Verbatim Char")
            except Exception:
                self._set_font_family(selection.Font, RS.MONO_FONT, RS.MONO_FONT)
                selection.Font.Size = 9
        else:
            self._set_font_family(selection.Font, RS.BODY_FONT, RS.LATIN_FONT)
            selection.Font.Size = 12

        selection.Font.Bold = span.bold
        selection.Font.Italic = span.italic
        selection.Font.StrikeThrough = getattr(span, "strikethrough", False)
        if span.link:
            try:
                selection.Style = selection.Document.Styles("Hyperlink")
            except Exception:
                selection.Font.Underline = 1
        else:
            selection.Font.Color = 0
            selection.Font.Underline = 0

    def add_rich_paragraph(self, spans, style_name):
        """Write a renderer paragraph while preserving current COM semantics."""
        from . import reference_styles as RS

        spans = list(spans)
        if len(spans) == 1 and not (
            spans[0].bold
            or spans[0].italic
            or spans[0].code
            or spans[0].link
            or getattr(spans[0], "strikethrough", False)
        ):
            selection = self.selection
            self._apply_body_style(style_name)
            selection.Font.Bold = False
            selection.Font.Italic = False
            self._set_font_family(selection.Font, RS.BODY_FONT, RS.LATIN_FONT)
            selection.Font.Size = 12
            selection.Font.Color = 0
            selection.TypeText(spans[0].text)
            selection.TypeParagraph()
            self._reset_selection_to_normal()
            return

        selection = self.selection
        self._apply_body_style(style_name)
        for span in spans:
            self._apply_span_format(span)
            if span.link:
                start = selection.End
                selection.TypeText(span.text)
                try:
                    anchor = self._doc.Range(start, selection.End)
                    self._doc.Hyperlinks.Add(anchor, span.link)
                except Exception:
                    pass
            else:
                selection.TypeText(span.text)
        selection.TypeParagraph()
        self._reset_selection_to_normal()

    def add_code_lines(self, lines):
        """Write Source Code paragraphs followed by the renderer spacer."""
        for line in lines:
            self.add_styled_paragraph(line if line else " ", "Source Code")
        self.add_paragraph("", size=4)

    def _reset_selection_to_normal(self):
        """Clear carried direct formatting and restore the Normal style."""
        s = self.selection
        try:
            s.ClearFormatting()
        except Exception:
            pass
        try:
            s.Style = self._doc.Styles(-1)
        except Exception:
            pass

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
        f.Collapse(0)  # wdCollapseEnd — Fields.Add replaces a non-collapsed range
        self._doc.Fields.Add(f, 33)  # wdFieldPage=33

    # ---- text ----
    def add_heading_level(self, text, level=1, size=None, color=None,
                          line_spacing=None, space_after=None,
                          line_spacing_rule=None):
        """Add a heading with proper Word outline level (1-6).

        Args:
            text: Heading text.
            level: 1-6 (H1-H6).
            size: Font size in points.
            color: #RRGGBB hex or BGR int.
            line_spacing: Line spacing override. Points unless accompanied by
                ``line_spacing_rule="multiple"``.
            space_after: Space after paragraph (points).
            line_spacing_rule: Semantic rule such as ``single``,
                ``one_and_half``, ``exact``, or ``multiple``.
        """
        from .reference_styles import get_heading_style

        level = min(max(int(level), 1), 6)
        props = get_heading_style(level)
        self.ensure_heading_styles({level: props})
        style_idx = -(level + 1)

        s = self.selection
        self._reset_selection_to_normal()
        s.Style = self._doc.Styles(style_idx)
        # Optional direct overrides remain for backward-compatible low-level use.
        if size is not None:
            s.Font.Size = size
        if color is not None:
            s.Font.Color = hex_to_rgb_long(color)
        if line_spacing is not None or line_spacing_rule is not None:
            self._set_line_spacing(
                s.ParagraphFormat, line_spacing, line_spacing_rule
            )
        if space_after is not None:
            s.ParagraphFormat.SpaceAfter = space_after
        s.TypeText(text)
        s.TypeParagraph()

        # Re-apply the style to the completed paragraph so the paragraph mark
        # cannot retain direct formatting from the following selection.
        try:
            para = s.Paragraphs(1).Previous()
            if para:
                para.Range.Style = self._doc.Styles(style_idx)
                if size is not None:
                    para.Range.Font.Size = size
                if color is not None:
                    para.Range.Font.Color = hex_to_rgb_long(color)
        except Exception:
            pass
        self._reset_selection_to_normal()

    # Backward-compat aliases
    def add_heading(self, text, size=None, bold=True, color=None):
        return self.add_heading_level(text, level=1, size=size, color=color)

    def add_heading2(self, text, size=None, color=None):
        return self.add_heading_level(text, level=2, size=size, color=color)

    def add_paragraph(self, text, size=None, bold=None, italic=None,
                      color=None, align=None, indent_first=None,
                      line_spacing=None, space_after=None, font_name=None,
                      line_spacing_rule=None, font_name_ascii=None,
                      space_before=None):
        s = self.selection
        self._reset_selection_to_normal()
        if font_name is not None:
            self._set_font_family(s.Font, font_name, font_name_ascii)
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
        if line_spacing is not None or line_spacing_rule is not None:
            self._set_line_spacing(
                s.ParagraphFormat, line_spacing, line_spacing_rule
            )
        if space_before is not None:
            s.ParagraphFormat.SpaceBefore = space_before
        if space_after is not None:
            s.ParagraphFormat.SpaceAfter = space_after
        s.TypeText(text)
        s.TypeParagraph()
        self._reset_selection_to_normal()
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
        from .reference_styles import BODY_FONT, LATIN_FONT

        s = self.selection
        for it in items:
            self._reset_selection_to_normal()
            try:
                s.Style = self._doc.Styles("List Paragraph")
            except Exception:
                pass
            s.ParagraphFormat.LeftIndent = indent
            s.ParagraphFormat.FirstLineIndent = -indent  # hanging indent
            try:
                # tab after the glyph must land on the indent, not the
                # default 36pt stop, or first line and wraps misalign
                s.ParagraphFormat.TabStops.Add(indent)
            except Exception:
                pass
            self._set_line_spacing(s.ParagraphFormat, rule="one_and_half")
            s.ParagraphFormat.SpaceBefore = 0
            s.ParagraphFormat.SpaceAfter = 3
            self._set_font_family(s.Font, BODY_FONT, LATIN_FONT)
            s.Font.Size = 12
            s.Font.Bold = False
            s.Font.Color = 0
            # Bullet glyph + tab + text
            s.TypeText(glyph + "\t" + it)
            s.TypeParagraph()
        self._reset_selection_to_normal()

    def add_numbered_list(self, items, indent=24):
        """Add a numbered list using explicit number prefixes.

        Args:
            items: List of text strings.
            indent: Left indent in points.
        """
        from .reference_styles import BODY_FONT, LATIN_FONT

        s = self.selection
        for idx, it in enumerate(items, 1):
            self._reset_selection_to_normal()
            try:
                s.Style = self._doc.Styles("List Paragraph")
            except Exception:
                pass
            s.ParagraphFormat.LeftIndent = indent
            s.ParagraphFormat.FirstLineIndent = -indent  # hanging indent
            try:
                # tab after the glyph must land on the indent, not the
                # default 36pt stop, or first line and wraps misalign
                s.ParagraphFormat.TabStops.Add(indent)
            except Exception:
                pass
            self._set_line_spacing(s.ParagraphFormat, rule="one_and_half")
            s.ParagraphFormat.SpaceBefore = 0
            s.ParagraphFormat.SpaceAfter = 3
            self._set_font_family(s.Font, BODY_FONT, LATIN_FONT)
            s.Font.Size = 12
            s.Font.Bold = False
            s.Font.Color = 0
            # Number + dot + tab + text
            s.TypeText(f"{idx}.\t{it}")
            s.TypeParagraph()
        self._reset_selection_to_normal()

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
            col_widths: Optional list of exact column widths in points.
            alignments: Optional list of "left"/"center"/"right" per column.
            banded_rows: Enable alternating row shading.
            auto_fit: Infer content-aware widths that fill the text area when
                ``col_widths`` is omitted.
            repeat_header: Repeat header row on each page.
            border_color: Grid line colour.

        Returns:
            The COM Table object.
        """
        from .reference_styles import BODY_FONT, LATIN_FONT

        s = self.selection
        self._reset_selection_to_normal()
        tbl = self._doc.Tables.Add(s.Range, rows, cols)

        # ---- Cell population ----
        for r in range(rows):
            for c in range(cols):
                if r >= len(data) or c >= len(data[r]):
                    continue
                cell = tbl.Cell(r + 1, c + 1)
                cell.Range.Text = str(data[r][c])
                style_name = "Table Header" if r == 0 else "Table Body"
                try:
                    cell.Range.Style = self._doc.Styles(style_name)
                except Exception:
                    pass
                self._set_font_family(cell.Range.Font, BODY_FONT, LATIN_FONT)
                cell.Range.Font.Size = font_size
                pf = cell.Range.ParagraphFormat
                pf.FirstLineIndent = 0
                pf.LeftIndent = 0
                pf.RightIndent = 0
                pf.SpaceBefore = 0
                pf.SpaceAfter = 0
                self._set_line_spacing(pf, rule="single")

                # Cell padding
                try:
                    cell.TopPadding = 1.5
                    cell.BottomPadding = 1.5
                    cell.LeftPadding = 4
                    cell.RightPadding = 4
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
                    cell.Range.Font.Size = font_size
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
                for row_idx in range(1, rows + 1):
                    try:
                        tbl.Cell(row_idx, c_idx + 1).Range.ParagraphFormat.Alignment = wd_align
                    except Exception:
                        pass

        # ---- Column widths ----
        widths = list(col_widths or [])
        inferred_widths = False
        page_setup = self._doc.PageSetup
        available_width = max(
            72.0,
            float(page_setup.PageWidth)
            - float(page_setup.LeftMargin)
            - float(page_setup.RightMargin),
        )
        if auto_fit and not widths:
            widths = _content_column_widths(data, cols, available_width)
            inferred_widths = True
        if widths:
            try:
                tbl.AutoFitBehavior(0)  # wdAutoFitFixed
                tbl.AllowAutoFit = False
                if inferred_widths:
                    tbl.PreferredWidthType = 3  # wdPreferredWidthPoints
                    tbl.PreferredWidth = available_width
            except Exception:
                pass
            for c_idx, width in enumerate(widths[:cols]):
                try:
                    tbl.Columns(c_idx + 1).SetWidth(float(width), 0)
                except Exception:
                    try:
                        tbl.Columns(c_idx + 1).Width = float(width)
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

        # Keep rows compact but never clip wrapped content.
        try:
            tbl.Rows.HeightRule = 0  # wdRowHeightAuto
        except Exception:
            pass
        try:
            tbl.Rows.AllowBreakAcrossPages = False
        except Exception:
            for row_index in range(1, rows + 1):
                try:
                    tbl.Rows(row_index).AllowBreakAcrossPages = False
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
        self._reset_selection_to_normal()
        s.TypeParagraph()
        return tbl

    def add_merged_table(self, data, merges=None, shade_header="#4472C4"):
        """data: list of rows (lists). merges: list of (r1,c1,r2,c2) 1-based."""
        if not data:
            return None
        rows = len(data)
        cols = max(len(r) for r in data)
        normalized = [list(row) + [""] * (cols - len(row)) for row in data]
        tbl = self.add_table(
            rows,
            cols,
            normalized,
            shade_header=shade_header,
            font_size=10,
            banded_rows=False,
            auto_fit=True,
        )
        if merges:
            for r1, c1, r2, c2 in merges:
                try:
                    tbl.Cell(r1, c1).Merge(tbl.Cell(r2, c2))
                except Exception:
                    pass
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

    def add_image(self, path, width=None, height=None, wrap=0, *,
                  max_width=None, max_height=None, inline=True,
                  preserve_aspect=True, alt=None):
        """Insert an image and optionally constrain it without distortion.

        Inline images are the default because they remain in the text flow
        across WPS Writer, Microsoft Word and PDF export.  ``width`` and
        ``height`` are target sizes in points; ``max_width``/``max_height``
        only scale down oversized images.
        """
        p = _abs(path)
        if inline:
            shape = self._doc.InlineShapes.AddPicture(
                p, False, True, self.selection.Range
            )
            try:
                shape.LockAspectRatio = -1 if preserve_aspect else 0
            except Exception:
                pass

            natural_width = float(shape.Width)
            natural_height = float(shape.Height)
            if width is not None and height is not None:
                if preserve_aspect and natural_width and natural_height:
                    scale = min(width / natural_width, height / natural_height)
                    shape.Width = natural_width * scale
                    shape.Height = natural_height * scale
                else:
                    shape.Width = width
                    shape.Height = height
            elif width is not None:
                shape.Width = width
            elif height is not None:
                shape.Height = height

            current_width = float(shape.Width)
            current_height = float(shape.Height)
            scales = [1.0]
            if max_width and current_width > max_width:
                scales.append(max_width / current_width)
            if max_height and current_height > max_height:
                scales.append(max_height / current_height)
            scale = min(scales)
            if scale < 1.0:
                shape.Width = current_width * scale
                shape.Height = current_height * scale

            if alt:
                try:
                    shape.AlternativeText = alt
                except Exception:
                    pass
            try:
                end = shape.Range.End
                self.selection.SetRange(end, end)
            except Exception:
                pass
            return shape

        if width and height:
            shape = self._doc.Shapes.AddPicture(
                p, False, True, 0, 0, width, height
            )
        else:
            shape = self._doc.Shapes.AddPicture(p, False, True)
            # honour a single-dimension constraint, keeping aspect ratio
            try:
                if width:
                    shape.LockAspectRatio = -1
                    shape.Width = width
                elif height:
                    shape.LockAspectRatio = -1
                    shape.Height = height
            except Exception:
                pass
        try:
            shape.WrapFormat.Type = wrap
        except Exception:
            pass
        if alt:
            try:
                shape.AlternativeText = alt
            except Exception:
                pass
        return shape

    def add_image_block(self, *args, **kwargs):
        """Insert and finish an image block as the Writer renderer expects."""
        shape = self.add_image(*args, **kwargs)
        try:
            shape.Range.ParagraphFormat.Alignment = 1
            shape.Range.ParagraphFormat.KeepWithNext = -1
            shape.Range.ParagraphFormat.SpaceAfter = 0
        except Exception:
            pass
        self.selection.TypeParagraph()
        return shape

    def add_page_break(self):
        self.selection.InsertBreak(7)

    def add_horizontal_line(self):
        s = self.selection
        s.InlineShapes.AddHorizontalLineStandard()

    def add_paragraph_horizontal_line(self):
        """Add the paragraph-border rule used for Markdown horizontal rules."""
        selection = self.selection
        selection.ParagraphFormat.Alignment = 1
        try:
            border = selection.ParagraphFormat.Borders(-3)  # wdBorderBottom
            border.LineStyle = 1
            border.LineWidth = 6
            border.Color = 0xC0C0C0
        except Exception:
            pass
        selection.TypeText(" ")
        selection.TypeParagraph()
        try:
            selection.ParagraphFormat.Borders(-3).LineStyle = 0
        except Exception:
            pass

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
