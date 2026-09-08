"""WriterComposer - WPS Writer (KWps.Application) to docx.

Rich-layout Word documents with columns, floating shapes, WordArt,
merged tables, TOC, and auto-populated fields.
"""

from __future__ import annotations

import os
import math
import re
import unicodedata
import zipfile
from types import SimpleNamespace
from xml.etree import ElementTree

from ._dispatch import (
    _abs, FMT_DOCX, FMT_PDF_FROM_DOC, FMT_DOC, FMT_DOCM, FMT_DOTX,
    FMT_DOTM, FMT_TXT, FMT_HTML, FMT_MHTML, FMT_RTF, FMT_XML, FMT_ODT,
    FMT_XPS,
)
from ._colors import hex_to_rgb_long
from ._base import BaseComposer
from .longform.privacy import redact_private_text
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

_NATIVE_SEQUENCE_IDS = frozenset({"WPSC_FIG", "WPSC_TAB", "WPSC_EQ"})
_NATIVE_BOOKMARK_RE = re.compile(r"^wpsc_(?:fig|tab|eq)_[0-9a-f]{24}$")
_NATIVE_HEADING_BOOKMARK_RE = re.compile(r"^wpsc_head_[0-9a-f]{24}$")
_PUBLIC_ISSUE_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]{0,63}$")


class NativeWriterObjectError(RuntimeError):
    """A closed, operation-local native rendering failure."""

    def __init__(self, code, message="native object operation failed"):
        super().__init__(message)
        self.code = code


def _caption_field_codes(numbering):
    """Return controlled STYLEREF/SEQ codes from a validated descriptor."""
    sequence_id = numbering.get("sequenceId")
    mode = numbering.get("mode")
    if sequence_id not in _NATIVE_SEQUENCE_IDS or mode not in {"global", "chapter"}:
        raise ValueError("invalid native numbering descriptor")
    if mode == "chapter":
        if numbering.get("chapterStyleLevel") != 1 or numbering.get("resetLevel") != 1:
            raise ValueError("invalid chapter numbering descriptor")
        return (
            None,  # STYLEREF code is built at COM time with the localized name
            f"SEQ {sequence_id} \\* ARABIC \\s 1",
        )
    if numbering.get("chapterStyleLevel") is not None or numbering.get("resetLevel") is not None:
        raise ValueError("invalid global numbering descriptor")
    return (None, f"SEQ {sequence_id} \\* ARABIC")


def _reference_field_code(bookmark_name):
    if not isinstance(bookmark_name, str) or not _NATIVE_BOOKMARK_RE.fullmatch(bookmark_name):
        raise ValueError("invalid native bookmark")
    return f"REF {bookmark_name} \\h"


# ---------------------------------------------------------------------------
# Stable paragraph IDs (w14:paraId)
# ---------------------------------------------------------------------------
# Pure helpers, unit-tested without COM. COM-side wiring
# (WriterComposer._read_paraid_map / _paragraph_index_for_paraid).

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
_PARAID_ATTR = "{%s}paraId" % _W14_NS


def _walk_story(element, out):
    """Append (paraId | None) entries aligned to ``Document.Paragraphs``.

    Main-story ``<w:p>`` elements contribute their ``w14:paraId`` (or ``None``
    when untagged). Each ``<w:tr>`` contributes one extra ``None`` AFTER its
    cells: Word's ``Paragraphs`` collection counts the end-of-row mark as a
    paragraph, but the mark has no ``<w:p>`` in the XML. Text-box stories are
    pruned (not counted by ``Document.Paragraphs``).
    """
    if element.tag == "{%s}txbxContent" % _W_NS:
        return
    if element.tag == "{%s}p" % _W_NS:
        out.append(element.get(_PARAID_ATTR))
        return
    if element.tag == "{%s}tr" % _W_NS:
        for child in element:
            _walk_story(child, out)
        out.append(None)  # end-of-row mark: counted by COM, absent from XML
        return
    for child in element:
        _walk_story(child, out)


def _extract_paraids(document_xml_bytes):
    """Return the ``w14:paraId`` list aligned to ``doc.Paragraphs(index)``.

    The returned list is 0-based: element 0 corresponds to ``Paragraphs(1)``.
    Entries are ``None`` when the paragraph has no ``w14:paraId`` (older docs,
    paragraphs Word did not tag) or is a table end-of-row mark (no ``<w:p>``
    in the XML). This is a pure transform of the ``word/document.xml`` part
    and does not touch COM.
    """
    root = ElementTree.fromstring(document_xml_bytes)
    paraids = []
    _walk_story(root, paraids)
    return paraids


def read_paraids_from_docx(path):
    """Open a .docx zip and return its paragraph-id list (see
    :func:`_extract_paraids`). Returns ``[]`` if the part is missing or the
    file cannot be read. Pure except for file IO; safe to call on a doc that
    is open in WPS (read-only zip access)."""
    try:
        with zipfile.ZipFile(os.fspath(path)) as archive:
            with archive.open("word/document.xml") as part:
                return _extract_paraids(part.read())
    except (OSError, zipfile.BadZipFile, KeyError, ElementTree.ParseError):
        return []


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

    # Long-form plan ops emit camelCase style keys (matching the macOS JSAPI
    # addin); WriterComposer's own renderers emit snake_case. Accept both so
    # Windows M5 generation does not silently drop every style property.
    _STYLE_CAMEL_KEYS = {
        "fontName": "font_name",
        "fontNameAscii": "font_name_ascii",
        "fontSize": "font_size",
        "spaceBefore": "space_before",
        "spaceAfter": "space_after",
        "keepWithNext": "keep_with_next",
        "keepTogether": "keep_together",
        "indentFirst": "indent_first",
        "leftIndent": "left_indent",
        "rightIndent": "right_indent",
        "lineSpacing": "line_spacing",
        "lineSpacingRule": "line_spacing_rule",
    }

    def _configure_style(self, style, props, is_char=False):
        """Apply a style definition, including base style and CJK font slots."""
        if any(camel in props for camel in self._STYLE_CAMEL_KEYS):
            merged = {
                self._STYLE_CAMEL_KEYS.get(key, key): value
                for key, value in props.items()
            }
            props = merged
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
    def _current_section_page_setup(self):
        try:
            sections = self._doc.Sections
            return sections(sections.Count).PageSetup
        except Exception:
            return self._doc.PageSetup

    def _current_section_is_landscape(self):
        try:
            setup = WriterComposer._current_section_page_setup(self)
            return int(setup.Orientation) == 1
        except Exception:
            return False

    def _fit_native_table_to_body(self, table, data):
        if not hasattr(table, "Columns"):
            return
        try:
            setup = WriterComposer._current_section_page_setup(self)
            available_width = max(
                72.0,
                float(setup.PageWidth)
                - float(setup.LeftMargin)
                - float(setup.RightMargin),
            )
        except Exception:
            return
        widths = _content_column_widths(data, len(data[0]), available_width)
        try:
            table.AutoFitBehavior(0)
        except Exception:
            pass
        table.AllowAutoFit = False
        try:
            table.PreferredWidthType = 3
            table.PreferredWidth = available_width
        except Exception:
            pass
        for index, width in enumerate(widths, start=1):
            column = table.Columns(index)
            try:
                column.SetWidth(float(width), 0)
            except Exception:
                column.Width = float(width)

    def add_section(self, landscape=None, *, continuous=False):
        self.selection.InsertBreak(3 if continuous else 2)
        if landscape is not None:
            section = self._doc.Sections(self._doc.Sections.Count)
            section.PageSetup.Orientation = 1 if landscape else 0

    def add_landscape_section_before_pending_heading(self):
        position = getattr(self, "_wpsc_last_heading_start", None)
        if not isinstance(position, int) or position < 0:
            raise NativeWriterObjectError(
                "PAGINATION_SNAPSHOT_FAILED", "pending heading is unavailable"
            )
        try:
            self._doc.Range(position, position).InsertBreak(2)
            self.selection.EndKey(6)
            section = self._doc.Sections(self._doc.Sections.Count)
            section.PageSetup.Orientation = 1
        except Exception:
            raise NativeWriterObjectError(
                "EXECUTION_ABORTED", "landscape heading cohesion failed"
            ) from None

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
                          line_spacing_rule=None, bold=None):
        """Add a heading with proper Word outline level (1-6).

        Args:
            text: Heading text.
            level: 1-6 (H1-H6).
            size: Font size in points.
            bold: Optional direct bold override.
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
        if bold is not None:
            s.Font.Bold = bold
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
                if bold is not None:
                    para.Range.Font.Bold = bold
                if color is not None:
                    para.Range.Font.Color = hex_to_rgb_long(color)
        except Exception:
            pass
        self._reset_selection_to_normal()

    # Backward-compat aliases
    def add_heading(self, text, size=None, bold=True, color=None):
        return self.add_heading_level(
            text, level=1, size=size, bold=bold, color=color
        )

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
        try:
            page_setup = self._doc.Sections(
                self._doc.Sections.Count
            ).PageSetup
        except Exception:
            pass
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
            # Use Body Text (no outline level) for the TOC title so it is
            # not collected by the TOC field itself (Heading 1 would put
            # 目  录 at the top of the TOC) and doesn't inherit the
            # previous paragraph style (e.g. Date)
            try:
                s.Style = self._doc.Styles("Body Text")
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
        toc = self._doc.TablesOfContents.Add(s.Range, True, 1, 3)
        self.selection.EndKey(6)
        self.selection.InsertBreak(7)
        return toc


    # ================================================================
    # Long-form M2 COM primitives
    # ================================================================

    def reset(self):
        """Reset to a fresh document.  No-op when the composer already owns one."""
        pass

    def set_page_role(self, role):
        """Tag the current section with a logical page role.

        Word/WPS has no native role slot, so failures are swallowed to keep
        generation robust on blind-COM runs.
        """
        try:
            section = self._doc.Sections(self._doc.Sections.Count)
            try:
                section.Range.DocumentVariables.Add(
                    "WpsComposerSectionRole_" + str(self._doc.Sections.Count), str(role)
                )
            except Exception:
                pass
        except Exception:
            pass

    def set_document_metadata(self, *, title, author):
        """Set export-facing metadata and clear any host-account identity."""

        properties = getattr(self._doc, "BuiltInDocumentProperties", None)
        if properties is None:
            raise NativeWriterObjectError(
                "CAPABILITY_MISMATCH", "document metadata is unavailable"
            )
        try:
            for name, value in (("Title", title), ("Author", author)):
                try:
                    prop = properties(name)
                except Exception:
                    prop = properties.Item(name)
                prop.Value = str(value or "")
        except Exception:
            raise NativeWriterObjectError(
                "EXECUTION_ABORTED", "document metadata apply failed"
            ) from None

    def set_page_numbering(self, format, start=None, restart=None):
        """Apply page-numbering format to the current section."""
        try:
            section = self._doc.Sections(self._doc.Sections.Count)
            footer = section.Footers(1)
            page_numbers = footer.PageNumbers
            if restart is not None:
                page_numbers.RestartNumberingAtSection = -1 if restart else 0
            if start is not None:
                page_numbers.StartingNumber = int(start)
            # Word/WPS NumberStyle: 0=Arabic, 1=UppercaseRoman, 2=LowercaseRoman
            style_map = {
                "none": 0,
                "roman": 2,
                "arabic": 0,
                "continue": 0,
            }
            if format in style_map:
                page_numbers.NumberStyle = style_map[format]
            if format == "none":
                footer.Range.Text = ""
            else:
                # Ensure a PAGE field exists in the primary footer.
                try:
                    footer.Range.ParagraphFormat.Alignment = 1
                    footer.Range.Collapse(0)
                    footer.Range.Fields.Add(footer.Range, 33)
                except Exception:
                    pass
        except Exception:
            pass

    def set_header_footer(
        self,
        header=None,
        footer=None,
        link_to_previous_header=None,
        link_to_previous_footer=None,
    ):
        """Set header/footer text for the current section."""
        try:
            section = self._doc.Sections(self._doc.Sections.Count)
            if link_to_previous_header is not None:
                try:
                    section.Headers(1).LinkToPrevious = -1 if link_to_previous_header else 0
                except Exception:
                    pass
            if link_to_previous_footer is not None:
                try:
                    section.Footers(1).LinkToPrevious = -1 if link_to_previous_footer else 0
                except Exception:
                    pass
            if header is not None:
                hdr = section.Headers(1)
                hdr.Range.Text = str(header)
                # Centered body header with bottom border line.
                try:
                    hdr.Range.ParagraphFormat.Alignment = 1
                    hdr.Range.ParagraphFormat.Borders(-3).LineStyle = 1
                    hdr.Range.ParagraphFormat.Borders(-3).LineWidth = 6
                    hdr.Range.ParagraphFormat.Borders(-3).Color = 0x000000
                except Exception:
                    pass
            if footer is not None:
                footer_text = str(footer)
                if footer_text != "":
                    section.Footers(1).Range.Text = footer_text
        except Exception:
            pass

    def configure_section(
        self,
        *,
        role=None,
        landscape=None,
        page_size=None,
        margins=None,
        restart_page_numbering=None,
        page_number_format=None,
        start_page_number=None,
        header_text=None,
        footer_text=None,
        link_to_previous_header=None,
        link_to_previous_footer=None,
    ):
        """Insert a new section after the first call and apply M2 skeleton settings."""
        first = not getattr(self, "_first_section_configured", False)
        if not first:
            try:
                self.selection.InsertBreak(2)  # wdSectionBreakNextPage
            except Exception:
                pass
        self._first_section_configured = True

        setup = self._current_section_page_setup()
        if landscape is not None:
            setup.Orientation = 1 if landscape else 0
        if page_size is not None:
            # Page-size strings are mapped at the add-in/protocol layer.
            pass
        if margins is not None:
            setup.TopMargin = margins.get("top", 72)
            setup.BottomMargin = margins.get("bottom", 72)
            setup.LeftMargin = margins.get("left", 90)
            setup.RightMargin = margins.get("right", 90)

        self.set_page_role(role or "body")
        self.set_page_numbering(
            format=page_number_format or "continue",
            start=start_page_number,
            restart=restart_page_numbering,
        )
        self.set_header_footer(
            header=header_text,
            footer=footer_text,
            link_to_previous_header=link_to_previous_header,
            link_to_previous_footer=link_to_previous_footer,
        )

    def insert_toc_with_styles(self, title, density):
        """Insert a TOC and apply density minima to TOC 1/2/3 styles."""
        toc = self.insert_toc(title)
        self._native_fields().append(("doc:toc", "TOC", toc, "index"))
        density = density or {}
        for level, key in enumerate(("toc1", "toc2", "toc3"), start=1):
            style_name = "TOC " + str(level)
            try:
                style = self._doc.Styles(style_name)
            except Exception:
                continue
            try:
                min_font = density.get("minFontSizePt", {}).get(key)
                if min_font is not None:
                    style.Font.Size = float(min_font)
                min_before = density.get("minSpaceBeforePt", {}).get(key)
                if min_before is not None:
                    style.ParagraphFormat.SpaceBefore = float(min_before)
                min_after = density.get("minSpaceAfterPt", {}).get(key)
                if min_after is not None:
                    style.ParagraphFormat.SpaceAfter = float(min_after)
            except Exception:
                pass

    def add_heading_level_native(
        self, text, level, numbering=None, scheme=None, keep_with_next=False,
        bookmark_name=None, sequence_transparent=False,
    ):
        """Add native headings, optionally keeping an unnumbered boundary out of SEQ resets."""
        if sequence_transparent and numbering:
            raise ValueError("sequence-transparent headings cannot request numbering")
        heading_start = self._native_position()
        self.add_heading_level(text, level=level)
        self._wpsc_last_heading_start = heading_start
        if bookmark_name is not None:
            if not _NATIVE_HEADING_BOOKMARK_RE.fullmatch(bookmark_name):
                raise ValueError("invalid native heading bookmark")
            try:
                self._doc.Bookmarks.Add(
                    bookmark_name,
                    self._doc.Range(heading_start, self._native_position()),
                )
            except Exception:
                raise NativeWriterObjectError(
                    "EXECUTION_ABORTED", "heading bookmark creation failed"
                ) from None
        if keep_with_next:
            try:
                paragraph = self.selection.Paragraphs(1).Previous()
                paragraph.Range.ParagraphFormat.KeepWithNext = True
            except Exception:
                raise NativeWriterObjectError(
                    "PAGINATION_SNAPSHOT_FAILED", "heading cohesion failed"
                ) from None
        if sequence_transparent or numbering is False:
            try:
                level_idx = min(max(int(level), 1), 6)
                source_style = self._doc.Styles(-1 - level_idx)
                style_name = (
                    "WPSC Sequence Transparent Heading " if sequence_transparent
                    else "WPSC Unnumbered Heading "
                ) + str(level_idx)
                outline_level = 10 if sequence_transparent else level_idx
                try:
                    unnumbered_style = self._doc.Styles(style_name)
                except Exception:
                    unnumbered_style = self._doc.Styles.Add(style_name, 1)
                # Base on Normal, not Heading N: inheriting the heading style
                # would also inherit its list template and STYLEREF identity.
                # Duplicate the complete appearance to retain heading font,
                # spacing, alignment and cohesion independently of its outline.
                unnumbered_style.BaseStyle = self._doc.Styles(-1)
                unnumbered_style.Font = source_style.Font.Duplicate
                unnumbered_style.ParagraphFormat = source_style.ParagraphFormat.Duplicate
                unnumbered_style.ParagraphFormat.OutlineLevel = outline_level
                heading_range = self._doc.Range(heading_start, self._native_position())
                heading_range.Style = unnumbered_style
                heading_range.ListFormat.RemoveNumbers(1)
                heading_range.ParagraphFormat.OutlineLevel = outline_level
            except Exception:
                raise NativeWriterObjectError(
                    "EXECUTION_ABORTED", "unnumbered heading style failed"
                ) from None
        if not numbering:
            return
        try:
            cache = getattr(self, "_wpsc_heading_templates", None)
            if cache is None:
                cache = {}
                self._wpsc_heading_templates = cache
            normalized_scheme = scheme or "decimal"
            list_template = cache.get(normalized_scheme)
            if list_template is None:
                list_template = self._doc.ListTemplates.Add(True)
                formats = (
                    ("第%1章", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4")
                    if normalized_scheme == "chinese-formal"
                    else ("%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4")
                )
                for index, number_format in enumerate(formats, start=1):
                    descriptor = list_template.ListLevels(index)
                    descriptor.NumberFormat = number_format
                    descriptor.NumberStyle = (
                        37 if normalized_scheme == "chinese-formal" and index == 1 else 0
                    )
                    descriptor.NumberPosition = (index - 1) * 18
                    descriptor.TextPosition = index * 18
                    descriptor.ResetOnHigher = 0 if index == 1 else index - 1
                    descriptor.StartAt = 1
                # Mirror the macOS add-in: link the built-in heading styles
                # to the template so numbering CONTINUES across headings.
                # Applying the template per range on WPS restarts the list
                # at 1 for every heading.
                for template_level in range(1, 5):
                    self._doc.Styles(-1 - template_level).LinkToListTemplate(
                        list_template, template_level
                    )
                cache[normalized_scheme] = list_template
            level_idx = int(level)
            heading_range = self._doc.Range(
                heading_start, self._native_position()
            )
            heading_range.Style = self._doc.Styles(-1 - level_idx)
            heading_range.ListFormat.ListLevelNumber = level_idx
            if not str(heading_range.ListFormat.ListString or "").strip():
                raise NativeWriterObjectError(
                    "EXECUTION_ABORTED", "heading numbering did not apply"
                )
        except NativeWriterObjectError:
            raise
        except Exception:
            raise NativeWriterObjectError(
                "EXECUTION_ABORTED", "heading numbering failed"
            ) from None

    def compact_terminal_paragraph(self):
        """Shrink only a final empty paragraph during the bounded M5 relayout."""
        try:
            paragraphs = self._doc.Paragraphs
            last = paragraphs(paragraphs.Count)
            text = str(getattr(last.Range, "Text", ""))
            if text.replace("\r", "").replace("\n", "").replace("\x07", "").strip():
                return
            last.Range.Font.Size = 1
            paragraph = last.Range.ParagraphFormat
            paragraph.SpaceBefore = 0
            paragraph.SpaceAfter = 0
            paragraph.LineSpacingRule = 4
            paragraph.LineSpacing = 1
            paragraph.KeepTogether = 0
            paragraph.KeepWithNext = 0
        except Exception:
            raise NativeWriterObjectError(
                "EXECUTION_ABORTED", "terminal paragraph compaction failed"
            ) from None

    # ================================================================
    # Long-form M3 native figure/table/caption/reference primitives
    # ================================================================

    def _native_fields(self):
        fields = getattr(self, "_wpsc_native_fields", None)
        if fields is None:
            fields = []
            self._wpsc_native_fields = fields
        return fields

    def _native_position(self):
        rng = self.selection.Range
        return int(getattr(rng, "End", getattr(self.selection, "End", 0)))

    def _native_set_position(self, position):
        self.selection.SetRange(int(position), int(position))

    def _native_insert_field(
        self, code, owner_node_id, kind, category, *, failure_code=None,
    ):
        rng = self.selection.Range
        try:
            rng.Collapse(0)
        except Exception:
            pass
        try:
            field = self._doc.Fields.Add(rng, -1, code, True)
        except Exception:
            if failure_code is not None:
                raise NativeWriterObjectError(failure_code) from None
            raise
        result = getattr(field, "Result", None)
        if result is not None:
            self.selection.SetRange(int(result.End), int(result.End))
            # WPS absorbs text typed at a field's result end into the field
            # result; a later Update() deletes it. Step right past the field
            # boundary before any further typing.
            try:
                self.selection.MoveRight(1, 1)
            except Exception:
                pass
        self._native_fields().append((owner_node_id or "doc:native", kind, field, category))
        return field

    def _native_document_end(self):
        content = getattr(getattr(self, "_doc", None), "Content", None)
        if content is not None:
            return int(content.End)
        return self._native_position()

    def _native_rollback(self, start, end):
        try:
            stop = int(end)
            if stop > int(start):
                self._doc.Range(int(start), stop).Delete()
            self._native_set_position(start)
        except Exception:
            raise NativeWriterObjectError(
                "LOCAL_MUTATION_ROLLBACK_FAILED", "native rollback failed"
            ) from None

    def _localized_styleref_code(self):
        """STYLEREF code naming the localized built-in Heading 1 style.

        WPS treats numeric STYLEREF arguments as literal style names and
        resolves built-in styles by localized UI name. Resolve the
        locale-independent built-in id (-2) at runtime, mirroring the
        macOS add-in.
        """
        try:
            name = str(self._doc.Styles(-2).NameLocal)
        except Exception:
            name = "Heading 1"
        if not name or '"' in name:
            name = "Heading 1"
        return f'STYLEREF "{name}" \\s'

    def _add_native_number_shell(self, numbering, bookmark_name, owner_node_id):
        style_code, sequence_code = _caption_field_codes(numbering)
        prefix = numbering["prefix"]
        suffix = numbering["suffix"]
        self.selection.TypeText(prefix)
        number_start = self._native_position()
        if style_code is None and numbering.get("mode") == "chapter":
            style_code = self._localized_styleref_code()
        if style_code is not None:
            self._native_insert_field(style_code, owner_node_id, "STYLEREF", "numbering")
            self.selection.TypeText("-")
        kind = {
            "WPSC_FIG": "SEQ_FIG",
            "WPSC_TAB": "SEQ_TAB",
            "WPSC_EQ": "SEQ_EQ",
        }[numbering["sequenceId"]]
        self._native_insert_field(sequence_code, owner_node_id, kind, "numbering")
        number_end = self._native_position()
        if bookmark_name is not None:
            if not _NATIVE_BOOKMARK_RE.fullmatch(bookmark_name):
                raise ValueError("invalid native bookmark")
            self._doc.Bookmarks.Add(
                bookmark_name, self._doc.Range(number_start, number_end)
            )
        if suffix:
            self.selection.TypeText(suffix)
        return number_start, number_end

    def _add_native_caption(
        self, caption, numbering, bookmark_name, owner_node_id,
        *, keep_with_next=False,
    ):
        start = self._native_position()
        self._add_native_number_shell(numbering, bookmark_name, owner_node_id)
        if caption:
            self.selection.TypeText(" " + str(caption))
        paragraph_range = self._doc.Range(start, self._native_position())
        paragraph_range.ParagraphFormat.Alignment = 1
        paragraph_range.ParagraphFormat.KeepTogether = -1
        paragraph_range.ParagraphFormat.KeepWithNext = -1 if keep_with_next else 0
        self.selection.TypeParagraph()
        return paragraph_range

    def _native_insert_figure_child(
        self, child, locator, owner_node_id, rollback_scope=None,
    ):
        start = self._native_position()
        try:
            shape = self.add_image(
                locator,
                width=float(child["displayWidthPt"]),
                height=float(child["displayHeightPt"]),
                inline=True,
                preserve_aspect=True,
                alt=owner_node_id,
            )
        except Exception:
            if rollback_scope is None:
                end = self._native_document_end()
            else:
                end = max(start, int(rollback_scope.End) - 1)
            self._native_rollback(start, end)
            raise NativeWriterObjectError("IMAGE_INSERT_FAILED") from None
        try:
            shape.Range.ParagraphFormat.Alignment = 1
            shape.Range.ParagraphFormat.KeepTogether = -1
            shape.Range.ParagraphFormat.KeepWithNext = -1
            self.selection.TypeParagraph()
            return shape
        except Exception:
            if rollback_scope is None:
                end = self._native_document_end()
            else:
                end = max(start, int(rollback_scope.End) - 1)
            self._native_rollback(start, end)
            raise

    def _native_columns_container(self, children):
        try:
            table = self._doc.Tables.Add(self.selection.Range, 1, 3)
            widths = (
                float(children[0]["displayWidthPt"]),
                12.0,
                float(children[1]["displayWidthPt"]),
            )
            for index, width in enumerate(widths, start=1):
                try:
                    table.Columns(index).SetWidth(width, 0)
                except Exception:
                    table.Columns(index).Width = width
            for border_id in range(-6, 0):
                table.Borders(border_id).LineStyle = 0
            table.Range.ParagraphFormat.KeepTogether = -1
            table.Range.ParagraphFormat.KeepWithNext = -1
            return table
        except Exception:
            raise NativeWriterObjectError("IMAGE_INSERT_FAILED") from None

    def _render_native_figure_stack(
        self, children, resource_locators, owner_node_id, *, retry_once,
    ):
        degraded = False
        for child in children:
            degradation = child.get("plannedDegradation")
            if degradation is not None:
                self.add_degradation_notice(
                    degradation["code"], degradation["message"],
                    degradation["fallback"], degradation.get("placement", "block"),
                )
                continue
            locator = resource_locators.get(child["resourceId"])
            if locator is None:
                raise NativeWriterObjectError("RESOURCE_HASH_MISMATCH")
            try:
                self._native_insert_figure_child(
                    child, locator, owner_node_id, rollback_scope=None
                )
            except NativeWriterObjectError as error:
                if error.code != "IMAGE_INSERT_FAILED" or not retry_once:
                    raise
                try:
                    self._native_insert_figure_child(
                        child, locator, owner_node_id, rollback_scope=None
                    )
                except NativeWriterObjectError as retry_error:
                    if retry_error.code != "IMAGE_INSERT_FAILED":
                        raise
                    self.add_degradation_notice(
                        "IMAGE_INSERT_FAILED", "Figure image could not be inserted",
                        "[IMAGE_INSERT_FAILED]", "block",
                    )
                degraded = True
        return degraded

    def add_captioned_figure_native(
        self, *, caption, numbering, indexable, referenceable, widthMode,
        orientation, kind, children, layout, keepWithCaption,
        owner_node_id=None, resource_locators=None, bookmarkName=None,
        explicitWidthPt=None, columns=None, controller_owned=False,
    ):
        """Insert normalized inline images followed by one native caption."""
        del indexable, referenceable, widthMode, kind, explicitWidthPt, columns
        if keepWithCaption is not True:
            raise ValueError("figure cohesion is required")
        resource_locators = dict(resource_locators or {})
        issues = []
        landscape = orientation == "landscape"
        owns_landscape_section = landscape and not self._current_section_is_landscape()
        if owns_landscape_section:
            self.add_section(landscape=True)
        issues = []
        try:
            for child in children:
                planned = child.get("plannedDegradation")
                if planned is not None:
                    issues.append({
                        "code": planned["code"],
                        "message": planned["message"],
                        "placement": planned.get("placement", "block"),
                    })
            can_use_columns = (
                layout == "columns"
                and all("resourceId" in child for child in children)
            )
            degraded = False
            if can_use_columns:
                operation_start = self._native_position()
                container = None
                try:
                    container = self._native_columns_container(children)
                    for index, child in enumerate(children):
                        locator = resource_locators.get(child["resourceId"])
                        if locator is None:
                            raise NativeWriterObjectError("RESOURCE_HASH_MISMATCH")
                        column = 1 if index == 0 else 3
                        cell_range = container.Cell(1, column).Range
                        self.selection.SetRange(cell_range.Start, cell_range.Start)
                        self._native_insert_figure_child(
                            child, locator, owner_node_id, rollback_scope=cell_range,
                        )
                    self.selection.SetRange(container.Range.End, container.Range.End)
                    self.selection.TypeParagraph()
                except NativeWriterObjectError as error:
                    if error.code != "IMAGE_INSERT_FAILED" or controller_owned:
                        raise
                    container_end = int(container.Range.End) if container is not None else self._native_document_end()
                    self._native_rollback(operation_start, container_end)
                    self._render_native_figure_stack(
                        children, resource_locators, owner_node_id, retry_once=False
                    )
                    degraded = True
            else:
                degraded = self._render_native_figure_stack(
                    children, resource_locators, owner_node_id,
                    retry_once=not controller_owned,
                )
            if degraded:
                issues.append({
                    "code": "IMAGE_INSERT_FAILED",
                    "message": "Figure used deterministic stack recovery",
                    "placement": "block",
                })
            if caption:
                self._add_native_caption(
                    caption, numbering, bookmarkName, owner_node_id,
                    keep_with_next=False,
                )
        finally:
            if owns_landscape_section:
                self.add_section(landscape=False)
        return {"issues": issues}

    def add_captioned_figure_fallback(
        self, *, children, caption="", numbering=None, orientation="portrait",
        bookmarkName=None, owner_node_id=None, resource_locators=None,
        failure_code="IMAGE_INSERT_FAILED", **kwargs,
    ):
        """Run the controller-owned bounded child-stack figure fallback."""
        del kwargs
        resource_locators = dict(resource_locators or {})
        issues = []
        landscape = orientation == "landscape"
        owns_landscape_section = landscape and not self._current_section_is_landscape()
        if owns_landscape_section:
            self.add_section(landscape=True)
        try:
            for child in children:
                planned = child.get("plannedDegradation")
                if planned is not None:
                    self.add_degradation_notice(
                        planned["code"], planned["message"], planned["fallback"],
                        planned.get("placement", "block"),
                    )
                    issues.append({
                        "code": planned["code"],
                        "message": planned["message"],
                        "placement": planned.get("placement", "block"),
                    })
                    continue
                locator = resource_locators.get(child.get("resourceId"))
                if locator is None:
                    raise NativeWriterObjectError("RESOURCE_HASH_MISMATCH")
                try:
                    self._native_insert_figure_child(
                        child, locator, owner_node_id, rollback_scope=None
                    )
                except NativeWriterObjectError as error:
                    if error.code != "IMAGE_INSERT_FAILED":
                        raise
                    self.add_degradation_notice(
                        failure_code, "Figure image could not be inserted",
                        "[IMAGE_INSERT_FAILED]", "block",
                    )
            if caption:
                self._add_native_caption(
                    caption, numbering or {}, bookmarkName, owner_node_id,
                    keep_with_next=False,
                )
        finally:
            if owns_landscape_section:
                self.add_section(landscape=False)
        return {"issues": issues}

    @staticmethod
    def _native_border_width(points):
        return {0.75: 6, 1.5: 12}.get(float(points), 2)

    def _apply_native_table_borders(self, table, border_spec):
        mapping = {
            "top": -1, "left": -2, "bottom": -3, "right": -4,
            "insideHorizontal": -5, "insideVertical": -6,
        }
        for key, border_id in mapping.items():
            points = float(border_spec[key])
            border = table.Borders(border_id)
            border.LineStyle = 0 if points == 0.0 else 1
            if points:
                border.LineWidth = self._native_border_width(points)
        header_points = float(border_spec["headerBottom"])
        header_border = table.Rows(1).Borders(-3)
        header_border.LineStyle = 0 if header_points == 0.0 else 1
        if header_points:
            header_border.LineWidth = self._native_border_width(header_points)

    def _create_native_table(
        self, headers, rows, alignments, border_spec, repeat_header,
        allow_row_split, cell_indent_pt, merges,
    ):
        data = [list(headers), *[list(row) for row in rows]]
        try:
            table = self._doc.Tables.Add(
                self.selection.Range, len(data), len(headers)
            )
        except Exception:
            raise NativeWriterObjectError("TABLE_INSERT_FAILED") from None
        alignment_codes = {"left": 0, "center": 1, "right": 2}
        try:
            for row_index, row in enumerate(data, start=1):
                for col_index, text in enumerate(row, start=1):
                    cell = table.Cell(row_index, col_index)
                    cell.Range.Text = str(text)
                    paragraph = cell.Range.ParagraphFormat
                    paragraph.FirstLineIndent = float(cell_indent_pt)
                    paragraph.LeftIndent = 0.0
                    paragraph.RightIndent = 0.0
                    paragraph.Alignment = alignment_codes[alignments[col_index - 1]]
            table.Rows.AllowBreakAcrossPages = -1 if allow_row_split else 0
            if repeat_header:
                table.Rows(1).HeadingFormat = -1
            self._fit_native_table_to_body(table, data)
            self._apply_native_table_borders(table, border_spec)
        except Exception:
            raise NativeWriterObjectError("TABLE_STYLE_APPLY_FAILED") from None
        try:
            for merge in merges:
                table.Cell(merge["top"], merge["left"]).Merge(
                    table.Cell(merge["bottom"], merge["right"])
                )
        except Exception:
            raise NativeWriterObjectError("TABLE_MERGE_APPLY_FAILED") from None
        try:
            self.selection.SetRange(table.Range.End, table.Range.End)
            self.selection.TypeParagraph()
        except Exception:
            raise NativeWriterObjectError("TABLE_INSERT_FAILED") from None
        return table

    def _table_overflow_group(self, table, merges):
        for merge in merges:
            if merge["top"] < 2 or merge["bottom"] <= merge["top"]:
                continue
            first = table.Cell(merge["top"], merge["left"]).Range
            last = table.Cell(merge["bottom"], merge["right"]).Range
            if int(first.Information(3)) != int(last.Information(3)):
                return merge["top"], merge["bottom"]
        return None

    def _add_native_table_text_fallback(self, headers, rows):
        self.selection.TypeText(" | ".join(str(value) for value in headers))
        self.selection.TypeParagraph()
        for row in rows:
            self.selection.TypeText(" | ".join(str(value) for value in row))
            self.selection.TypeParagraph()

    def _add_native_table_notice(self, code, message, fallback_text=""):
        start = self._native_position()
        self.add_degradation_notice(code, message, fallback_text, "block")
        notice = self._doc.Range(start, self._native_position())
        notice.ParagraphFormat.KeepTogether = -1
        notice.ParagraphFormat.KeepWithNext = -1

    def _apply_native_table_cell_degradations(self, table, degradations):
        for descriptor in degradations or ():
            try:
                cell_range = table.Cell(
                    int(descriptor["row"]), int(descriptor["column"])
                ).Range
                fallback_text = str(descriptor["fallbackText"])
                existing = str(getattr(cell_range, "Text", "")).rstrip(
                    "\r\n\x07"
                )
                if fallback_text not in existing:
                    cell_range.Text = (
                        f"{existing} {fallback_text}" if existing else fallback_text
                    )
                cell_range.Shading.BackgroundPatternColor = hex_to_rgb_long(
                    "#FCE8E6"
                )
            except Exception:
                raise NativeWriterObjectError(
                    "DEGRADATION_INSERT_FAILED",
                    "cell degradation styling failed",
                ) from None

    def add_semantic_table_native(
        self, *, caption, numbering, indexable, referenceable, headers, rows,
        alignments, style, orientation, borderSpec, merges, repeatHeader,
        allowRowSplit, cellIndentPt, plannedDegradation,
        keepCaptionWithFirstRow, owner_node_id=None, bookmarkName=None,
        cellDegradations=(), controller_owned=False, m5Relayout=None,
        includePreviousHeading=False, continuousExit=False,
    ):
        """Insert a caption, planned notices, and a resolved native table."""
        del indexable, referenceable, style, m5Relayout
        if keepCaptionWithFirstRow is not True:
            raise ValueError("table cohesion is required")
        landscape = orientation == "landscape"
        owns_landscape_section = landscape and not self._current_section_is_landscape()
        if owns_landscape_section:
            if includePreviousHeading:
                self.add_landscape_section_before_pending_heading()
            else:
                self.add_section(landscape=True)
        issues = []
        try:
            if caption:
                self._add_native_caption(
                    caption, numbering, bookmarkName, owner_node_id,
                    keep_with_next=True,
                )
            for degradation in plannedDegradation:
                self._add_native_table_notice(
                    degradation["code"], degradation["message"], "",
                )
                issues.append(dict(degradation))
            table_start = self._native_position()
            try:
                table = self._create_native_table(
                    headers, rows, alignments, borderSpec, repeatHeader,
                    allowRowSplit, cellIndentPt, merges,
                )
            except NativeWriterObjectError as error:
                if controller_owned or error.code not in {
                    "TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED",
                    "TABLE_INSERT_FAILED",
                }:
                    raise
                self._native_rollback(table_start, self._native_document_end())
                self._add_native_table_notice(
                    error.code, "Table used the deterministic grid fallback"
                )
                grid_start = self._native_position()
                try:
                    table = self._create_native_table(
                        headers, rows, alignments, {
                            "top": .75, "bottom": .75, "headerBottom": .75,
                            "left": .75, "right": .75,
                            "insideHorizontal": .75, "insideVertical": .75,
                        }, repeatHeader, True, cellIndentPt, (),
                    )
                    issues.append({"code": error.code, "message": "Table used deterministic grid fallback", "placement": "block"})
                except NativeWriterObjectError as grid_error:
                    if grid_error.code not in {
                        "TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED",
                        "TABLE_INSERT_FAILED",
                    }:
                        raise
                    self._native_rollback(grid_start, self._native_document_end())
                    self._add_native_table_text_fallback(headers, rows)
                    issues.append({"code": "TABLE_INSERT_FAILED", "message": "Table used deterministic text fallback", "placement": "block"})
                    return {"issues": issues}
            if self._table_overflow_group(table, merges) is not None:
                if controller_owned:
                    raise NativeWriterObjectError("TABLE_ROW_FORCED_SPLIT") from None
                self._native_rollback(table_start, self._native_document_end())
                self._add_native_table_notice(
                    "TABLE_ROW_FORCED_SPLIT",
                    "A vertical merge group exceeded the available page height", "",
                )
                issues.append({"code": "TABLE_ROW_FORCED_SPLIT", "message": "Vertical merge group rendered as splittable grid", "placement": "block"})
                grid_start = self._native_position()
                try:
                    table = self._create_native_table(
                        headers, rows, alignments, {
                            "top": .75, "bottom": .75, "headerBottom": .75,
                            "left": .75, "right": .75,
                            "insideHorizontal": .75, "insideVertical": .75,
                        }, repeatHeader, True, cellIndentPt, (),
                    )
                except NativeWriterObjectError as grid_error:
                    if grid_error.code not in {
                        "TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED",
                        "TABLE_INSERT_FAILED",
                    }:
                        raise
                    self._native_rollback(grid_start, self._native_document_end())
                    self._add_native_table_text_fallback(headers, rows)
                    issues.append({"code": "TABLE_INSERT_FAILED", "message": "Overflow grid used deterministic text fallback", "placement": "block"})
            self._apply_native_table_cell_degradations(
                table, cellDegradations
            )
        finally:
            if owns_landscape_section:
                self.add_section(
                    landscape=False, continuous=bool(continuousExit)
                )
        return {"issues": issues}

    def add_semantic_table_fallback(
        self, *, headers, rows, alignments, repeatHeader=True,
        cellIndentPt=0.0, orientation="portrait", failure_code="TABLE_INSERT_FAILED",
        owner_node_id=None, caption="", numbering=None, bookmarkName=None,
        plannedDegradation=(), cellDegradations=(), **kwargs,
    ):
        """Run the controller-owned grid-then-text table fallback once."""
        include_previous_heading = bool(kwargs.get("includePreviousHeading"))
        continuous_exit = bool(kwargs.get("continuousExit"))
        landscape = orientation == "landscape"
        owns_landscape_section = landscape and not self._current_section_is_landscape()
        if owns_landscape_section:
            if include_previous_heading:
                self.add_landscape_section_before_pending_heading()
            else:
                self.add_section(landscape=True)
        issues = []
        try:
            if caption:
                self._add_native_caption(
                    caption, numbering or {}, bookmarkName, owner_node_id,
                    keep_with_next=True,
                )
            for planned in plannedDegradation or ():
                self._add_native_table_notice(
                    planned["code"], planned.get("message", ""), ""
                )
                issues.append(dict(planned))
            self._add_native_table_notice(
                failure_code, "Table used the deterministic grid fallback"
            )
            grid_start = self._native_position()
            grid = {
                "top": .75, "bottom": .75, "headerBottom": .75,
                "left": .75, "right": .75,
                "insideHorizontal": .75, "insideVertical": .75,
            }
            try:
                table = self._create_native_table(
                    headers, rows, alignments, grid, repeatHeader,
                    True, cellIndentPt, (),
                )
                self._apply_native_table_cell_degradations(
                    table, cellDegradations
                )
            except NativeWriterObjectError as error:
                if error.code not in {
                    "TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED",
                    "TABLE_INSERT_FAILED",
                }:
                    raise
                self._native_rollback(grid_start, self._native_document_end())
                self._add_native_table_text_fallback(headers, rows)
        finally:
            if owns_landscape_section:
                self.add_section(landscape=False, continuous=continuous_exit)
        return {"issues": issues}

    def add_equation_number_native(
        self, *, source, numbering, bookmarkName, fallbackText,
        owner_node_id=None,
    ):
        """Insert readable M2 formula text plus the M3 native number shell."""
        start = self._native_position()
        self.selection.TypeText(str(source or fallbackText))
        self.selection.TypeText("\t")
        self._add_native_number_shell(numbering, bookmarkName, owner_node_id)
        paragraph = self._doc.Range(start, self._native_position())
        paragraph.ParagraphFormat.Alignment = 2
        paragraph.ParagraphFormat.KeepTogether = -1
        self.selection.TypeParagraph()
        return {"issues": []}

    def _native_formula_container(self):
        """Create the fixed borderless layout shell for one numbered formula."""
        table = self._doc.Tables.Add(self.selection.Range, 1, 3)
        for border_id in range(-6, 0):
            table.Borders(border_id).LineStyle = 0
        # Stay within the narrowest standard M2 body width while keeping equal
        # side columns so the middle formula is visually page-centered.
        for index, width in enumerate((36.0, 320.0, 36.0), start=1):
            column = table.Columns(index)
            try:
                column.SetWidth(width, 0)
            except Exception:
                column.Width = width
        table.Rows.AllowBreakAcrossPages = 0
        table.Range.ParagraphFormat.KeepTogether = -1
        table.Range.ParagraphFormat.KeepWithNext = 0
        center = table.Cell(1, 2).Range
        number = table.Cell(1, 3).Range
        center.ParagraphFormat.Alignment = 1
        center.ParagraphFormat.KeepTogether = -1
        number.ParagraphFormat.Alignment = 2
        number.ParagraphFormat.KeepTogether = -1
        return table

    def _native_formula_number_shell(
        self, table, numbering, bookmark_name, owner_node_id,
    ):
        number_range = table.Cell(1, 3).Range
        self.selection.SetRange(int(number_range.Start), int(number_range.Start))
        self._add_native_number_shell(
            numbering, bookmark_name, owner_node_id
        )
        number_range.ParagraphFormat.Alignment = 2
        number_range.ParagraphFormat.KeepTogether = -1
        self.selection.SetRange(int(table.Range.End), int(table.Range.End))
        self.selection.TypeParagraph()

    def _add_trusted_native_math(self, table, descriptor):
        if descriptor.get("syntax") != "wps-linear-v1":
            raise ValueError("untrusted native math syntax")
        center = table.Cell(1, 2).Range
        start = int(center.Start)
        end = int(center.End) - 1
        if end <= start:
            raise NativeWriterObjectError("EQUATION_INSERT_FAILED") from None
        math_range = self._doc.Range(start, end)
        math_range.Text = descriptor["linearText"]
        maths = self._doc.OMaths
        before = int(maths.Count)
        added_range = maths.Add(math_range)
        after = int(maths.Count)
        if added_range is None or after != before + 1:
            raise NativeWriterObjectError("EQUATION_INSERT_FAILED") from None
        added_omaths = getattr(added_range, "OMaths", None)
        if added_omaths is None or int(added_omaths.Count) != 1:
            raise NativeWriterObjectError("EQUATION_INSERT_FAILED") from None
        native = self._native_collection_item(added_omaths, 1)
        document_native = self._native_collection_item(maths, after)

        added_start = int(added_range.Start)
        added_end = int(added_range.End)
        native_range = native.Range
        document_range = document_native.Range
        native_start = int(native_range.Start)
        native_end = int(native_range.End)
        document_start = int(document_range.Start)
        document_end = int(document_range.End)
        if (
            added_end <= added_start
            or added_start < start
            or added_end > end
            or native_end <= native_start
            or native_start < added_start
            or native_end > added_end
            or (document_start, document_end) != (native_start, native_end)
        ):
            raise NativeWriterObjectError("EQUATION_INSERT_FAILED") from None

        native.BuildUp()
        if int(maths.Count) != after:
            raise NativeWriterObjectError("EQUATION_INSERT_FAILED") from None
        built_range = native.Range
        built_document_range = document_native.Range
        built_start = int(built_range.Start)
        built_end = int(built_range.End)
        if (
            built_end <= built_start
            or built_start < added_start
            or built_end > added_end
            or (
                int(built_document_range.Start),
                int(built_document_range.End),
            ) != (built_start, built_end)
        ):
            raise NativeWriterObjectError("EQUATION_INSERT_FAILED") from None
        center.ParagraphFormat.Alignment = 1
        center.ParagraphFormat.KeepTogether = -1
        return native

    def add_equation_native(
        self, *, renderMode, content, numbering, bookmarkName, fallbackText,
        owner_node_id=None, controller_owned=False,
        fallback_resource_locator=None,
    ):
        """Insert one M4 editable OMath; recovery remains controller-owned."""
        if renderMode != "native-m4":
            raise ValueError("invalid native M4 equation mode")
        del controller_owned
        planned = content.get("plannedDegradation")
        if planned is not None:
            self.add_equation_native_fallback(
                numbering=numbering,
                bookmarkName=bookmarkName,
                fallbackText=fallbackText,
                fallback_resource_locator=fallback_resource_locator,
                owner_node_id=owner_node_id,
                failure_code=planned["code"],
            )
            return {"issues": [{
                "code": planned["code"],
                "message": planned["reason"],
                "placement": planned["placement"],
                "fallback": planned["fallbackKind"],
            }]}
        table = self._native_formula_container()
        self._add_trusted_native_math(table, content["nativeMath"])
        self._native_formula_number_shell(
            table, numbering, bookmarkName, owner_node_id
        )
        return {"issues": []}

    def add_equation_native_fallback(
        self, *, numbering, bookmarkName, fallbackText,
        fallback_resource_locator=None, owner_node_id=None,
        failure_code="EQUATION_INSERT_FAILED",
    ):
        """Run the controller's one bounded image-or-source formula fallback."""
        table = self._native_formula_container()
        center = table.Cell(1, 2).Range
        self.selection.SetRange(int(center.Start), int(center.Start))
        image_inserted = False
        if fallback_resource_locator is not None:
            image_start = self._native_position()
            try:
                shape = self.add_image(
                    fallback_resource_locator,
                    max_width=290.0,
                    max_height=180.0,
                    inline=True,
                    preserve_aspect=True,
                    alt=owner_node_id,
                )
                if shape is None:
                    raise ValueError("formula fallback image was not created")
                shape_range = shape.Range
                shape_range.ParagraphFormat.Alignment = 1
                shape_range.ParagraphFormat.KeepTogether = -1
                image_inserted = True
            except Exception:
                self._native_rollback(
                    image_start,
                    max(image_start, int(center.End) - 1),
                )
                self.selection.SetRange(int(center.Start), int(center.Start))
        if image_inserted:
            self.add_inline_degradation(
                failure_code,
                "Formula used its validated image fallback",
                "formula image fallback",
            )
        else:
            self.add_inline_degradation(
                failure_code,
                "Formula native math could not be inserted",
                str(fallbackText),
            )
        center.ParagraphFormat.Alignment = 1
        center.ParagraphFormat.KeepTogether = -1
        self._native_formula_number_shell(
            table, numbering, bookmarkName, owner_node_id
        )
        return {"issues": []}

    def add_cross_reference_paragraph(
        self, *, runs, owner_node_id=None, listFormatting=None,
        controller_owned=False,
    ):
        """Insert ordered literal and native REF runs in one paragraph."""
        return self._add_native_run_paragraph(
            runs=runs,
            owner_node_id=owner_node_id,
            listFormatting=listFormatting,
            controller_owned=controller_owned,
        )

    def add_citation_paragraph(
        self, *, runs, owner_node_id=None, listFormatting=None,
        controller_owned=False,
    ):
        """Insert static numeric citations through the shared run primitive."""
        return self._add_native_run_paragraph(
            runs=runs,
            owner_node_id=owner_node_id,
            listFormatting=listFormatting,
            controller_owned=controller_owned,
        )

    def _add_native_run_paragraph(
        self, *, runs, owner_node_id=None, listFormatting=None,
        controller_owned=False,
    ):
        if listFormatting is not None:
            indent = float(listFormatting["indentPt"])
            self._reset_selection_to_normal()
            try:
                self.selection.Style = self._doc.Styles("List Paragraph")
            except Exception:
                pass
            paragraph_format = self.selection.ParagraphFormat
            paragraph_format.LeftIndent = indent
            paragraph_format.FirstLineIndent = -indent
            try:
                paragraph_format.TabStops.Add(indent)
            except Exception:
                pass
            self._set_line_spacing(paragraph_format, rule="one_and_half")
            paragraph_format.SpaceBefore = 0
            paragraph_format.SpaceAfter = 3
        degraded = False
        planned_issues = []
        for run in runs:
            run_type = run["type"]
            if run_type == "text":
                self.selection.TypeText(run["text"])
                continue
            if run_type == "citation":
                self.selection.TypeText(run["fallbackText"])
                continue
            if run_type == "degradation":
                self._add_inline_degradation_literal(run["fallbackText"])
                planned_issues.append({
                    "code": run["code"],
                    "message": "Reference target is unresolved",
                    "placement": "inline",
                    "fallback": "inline",
                    "nodeId": run["nodeId"],
                })
                continue
            self.selection.TypeText(run["prefix"])
            start = self._native_position()
            try:
                self._native_insert_field(
                    _reference_field_code(run["bookmarkName"]),
                    owner_node_id, "REF", "reference",
                    failure_code="CROSS_REFERENCE_FAILED",
                )
            except NativeWriterObjectError as error:
                if error.code != "CROSS_REFERENCE_FAILED" or controller_owned:
                    raise
                self._native_rollback(start, self._native_document_end())
                self.selection.TypeText(run["fallbackText"])
                degraded = True
            if run["suffix"]:
                self.selection.TypeText(run["suffix"])
        self.selection.TypeParagraph()
        if listFormatting is not None:
            self._reset_selection_to_normal()
        issues = list(planned_issues)
        if degraded:
            issues.append({
                "code": "CROSS_REFERENCE_FAILED",
                "message": "Cross-reference used its inline fallback",
                "placement": "inline",
            })
        return {"issues": issues}

    def _add_inline_degradation_literal(self, fallback_text):
        """Style an already-coded visible fallback without wrapping it again."""
        selection = self.selection
        start = int(selection.End)
        try:
            selection.TypeText(redact_private_text(str(fallback_text)))
            inserted = self._doc.Range(start, int(selection.End))
            self._style_degradation_range(inserted)
            return inserted
        except Exception:
            raise NativeWriterObjectError(
                "DEGRADATION_INSERT_FAILED",
                "inline degradation insertion failed",
            ) from None

    def add_bibliography_native(
        self, *, schemaVersion=1, entries, style, hangingIndentPt,
        leftIndentPt, spaceAfterPt, owner_node_id=None,
        controller_owned=False,
    ):
        """Insert already-ordered M4 numeric bibliography paragraphs."""
        del owner_node_id
        if schemaVersion != 1 or style != "numeric":
            raise ValueError("invalid structured bibliography contract")
        del controller_owned
        for entry in entries:
            start = self._native_position()
            try:
                self.selection.TypeText(
                    f"[{int(entry['number'])}] {entry['text']}"
                )
            except Exception:
                raise NativeWriterObjectError(
                    "BIBLIOGRAPHY_INSERT_FAILED"
                ) from None
            paragraph = self._doc.Range(start, self._native_position())
            paragraph.ParagraphFormat.Alignment = 0
            paragraph.ParagraphFormat.LeftIndent = float(leftIndentPt)
            paragraph.ParagraphFormat.FirstLineIndent = -float(hangingIndentPt)
            paragraph.ParagraphFormat.SpaceBefore = 0
            paragraph.ParagraphFormat.SpaceAfter = float(spaceAfterPt)
            paragraph.ParagraphFormat.KeepTogether = -1
            try:
                self.selection.TypeParagraph()
            except Exception:
                raise NativeWriterObjectError(
                    "BIBLIOGRAPHY_INSERT_FAILED"
                ) from None
        return {"issues": []}

    def add_bibliography_legacy(
        self, *, entries, style="numbered", owner_node_id=None,
    ):
        """Execute the unchanged legacy string bibliography shape."""
        del style, owner_node_id
        for entry in entries:
            self.selection.TypeText(str(entry))
            self.selection.TypeParagraph()
        return {"issues": []}

    def add_cross_reference_fallback(
        self, *, runs, owner_node_id=None, listFormatting=None,
        failure_code="CROSS_REFERENCE_FAILED",
    ):
        """Insert one controller-owned static inline reference fallback."""
        del owner_node_id, failure_code
        if listFormatting is not None:
            indent = float(listFormatting["indentPt"])
            self._reset_selection_to_normal()
            try:
                self.selection.Style = self._doc.Styles("List Paragraph")
            except Exception:
                pass
            paragraph_format = self.selection.ParagraphFormat
            paragraph_format.LeftIndent = indent
            paragraph_format.FirstLineIndent = -indent
            try:
                paragraph_format.TabStops.Add(indent)
            except Exception:
                pass
            self._set_line_spacing(paragraph_format, rule="one_and_half")
            paragraph_format.SpaceBefore = 0
            paragraph_format.SpaceAfter = 3
        for run in runs:
            if run["type"] == "text":
                self.selection.TypeText(run["text"])
            else:
                self.selection.TypeText(str(run.get("prefix", "")))
                start = self._native_position()
                self.selection.TypeText(str(run.get("fallbackText", "")))
                inserted = self._doc.Range(start, self._native_position())
                self._style_degradation_range(inserted)
                self.selection.TypeText(str(run.get("suffix", "")))
        self.selection.TypeParagraph()
        if listFormatting is not None:
            self._reset_selection_to_normal()

    def insert_caption_index_native(
        self, *, title, sequence_id, title_style_id, owner_node_id=None,
    ):
        if sequence_id not in {"WPSC_FIG", "WPSC_TAB"}:
            raise ValueError("invalid native caption index sequence")
        if title:
            try:
                self.selection.Style = self._doc.Styles(title_style_id)
            except Exception:
                # Built-in Body Text is explicitly non-outline and therefore
                # safe for the TOC when the named internal style is absent.
                self.selection.Style = self._doc.Styles("Body Text")
            self.selection.TypeText(str(title))
            self.selection.TypeParagraph()
        index = self._doc.TablesOfFigures.Add(self.selection.Range, sequence_id)
        self._native_fields().append((
            owner_node_id or "doc:index",
            "TOF_FIG" if sequence_id == "WPSC_FIG" else "TOF_TAB",
            index,
            "index",
        ))
        self.selection.EndKey(6)
        self.selection.TypeParagraph()
        return index

    def _update_tracked_native_fields(self, kinds):
        for _owner, kind, native, _category in self._native_fields():
            if kind in kinds:
                native.Update()

    @staticmethod
    def _native_collection_item(collection, index):
        item = getattr(collection, "Item", None)
        if callable(item):
            return item(index)
        return collection(index)

    @staticmethod
    def _native_field_kind(field):
        code = str(getattr(getattr(field, "Code", None), "Text", ""))
        token = code.strip().upper().split(None, 1)
        return token[0] if token and token[0] in {"PAGE", "NUMPAGES"} else None

    def _iter_section_page_fields(self):
        sections = self._doc.Sections
        for section_index in range(1, int(sections.Count) + 1):
            section = self._native_collection_item(sections, section_index)
            for story_name in ("Headers", "Footers"):
                stories = getattr(section, story_name)
                for story_index in range(1, int(stories.Count) + 1):
                    story = self._native_collection_item(stories, story_index)
                    exists = getattr(story, "Exists", True)
                    if exists is False or exists == 0:
                        continue
                    fields = story.Range.Fields
                    for field_index in range(1, int(fields.Count) + 1):
                        native = self._native_collection_item(fields, field_index)
                        kind = self._native_field_kind(native)
                        if kind is not None:
                            owner = (
                                f"section:{section_index}/"
                                f"{story_name.lower()}:{story_index}"
                            )
                            yield owner, kind, native

    @staticmethod
    def _native_range_page_span(native):
        native_range = native.Range
        start_position = int(native_range.Start)
        end_position = int(native_range.End)
        if end_position < start_position:
            raise ValueError("native index range is reversed")
        if end_position == start_position:
            # A truly empty native index has no content page to inspect but
            # still occupies its insertion page.
            return 1
        start = native_range.Duplicate
        end = native_range.Duplicate
        start.SetRange(start_position, start_position)
        last_content_position = end_position - 1
        end.SetRange(last_content_position, last_content_position)
        first_page = int(start.Information(3))
        last_page = int(end.Information(3))
        return max(1, last_page - first_page + 1)

    def repaginate_and_update_numbering(self):
        self._doc.Repaginate()
        self._update_tracked_native_fields({"STYLEREF", "SEQ_FIG", "SEQ_TAB", "SEQ_EQ"})

    def refresh_bookmarks_and_references(self):
        # Accessing Count is a required bookmark API health check.  Do not
        # swallow a missing/throwing native collection.
        int(self._doc.Bookmarks.Count)
        self._update_tracked_native_fields({"REF"})

    def refresh_indexes(self):
        for index in range(1, int(self._doc.TablesOfContents.Count) + 1):
            self._doc.TablesOfContents.Item(index).Update()
        for index in range(1, int(self._doc.TablesOfFigures.Count) + 1):
            self._doc.TablesOfFigures.Item(index).Update()

    def repaginate_and_update_page_fields(self):
        self._doc.Repaginate()
        self._update_tracked_native_fields({"PAGE", "NUMPAGES"})
        tracked_ids = {
            id(native)
            for _owner, kind, native, _category in self._native_fields()
            if kind in {"PAGE", "NUMPAGES"}
        }
        for _owner, _kind, native in self._iter_section_page_fields():
            if id(native) not in tracked_ids:
                native.Update()

    def snapshot_fields(self):
        from .longform.field_contract import snapshot_visible_field

        total_pages = int(self._doc.ComputeStatistics(2))
        tracked = tuple(self._native_fields())
        toc_pages = sum(
            self._native_range_page_span(native)
            for _owner, kind, native, _category in tracked
            if kind == "TOC"
        )
        figure_index_pages = sum(
            self._native_range_page_span(native)
            for _owner, kind, native, _category in tracked
            if kind == "TOF_FIG"
        )
        table_index_pages = sum(
            self._native_range_page_span(native)
            for _owner, kind, native, _category in tracked
            if kind == "TOF_TAB"
        )
        snapshots = []
        ordinals = {}
        for owner, kind, native, category in tracked:
            key = (owner, kind)
            ordinal = ordinals.get(key, 0)
            ordinals[key] = ordinal + 1
            result = getattr(native, "Result", None)
            if result is not None:
                visible = getattr(result, "Text", "")
            else:
                visible = getattr(getattr(native, "Range", None), "Text", "")
            snapshots.append(snapshot_visible_field(
                owner_node_id=owner,
                field_kind=kind,
                ordinal_within_node=ordinal,
                visible_result=visible,
                field_category=category,
                toc_page_count=toc_pages,
                figure_index_page_count=figure_index_pages,
                table_index_page_count=table_index_pages,
                total_pages=total_pages,
            ))
        story_ordinals = {}
        for owner, kind, native in self._iter_section_page_fields():
            key = (owner, kind)
            ordinal = story_ordinals.get(key, 0)
            story_ordinals[key] = ordinal + 1
            visible = getattr(getattr(native, "Result", None), "Text", "")
            snapshots.append(snapshot_visible_field(
                owner_node_id=owner,
                field_kind=kind,
                ordinal_within_node=ordinal,
                visible_result=visible,
                field_category="page",
                toc_page_count=toc_pages,
                figure_index_page_count=figure_index_pages,
                table_index_page_count=table_index_pages,
                total_pages=total_pages,
            ))
        return tuple(snapshots)

    def finalize_fields(self, *, max_rounds=3):
        """Update all fields (convergence helper)."""
        self.update_fields()

    def refresh_fields(self, round_index):
        """Return a deterministic field snapshot for the convergence loop."""
        from .longform.executor import FieldSnapshot

        # This is the single legacy mutation point used by the shared adapter.
        # Required field APIs intentionally propagate failures to the executor.
        for index in range(1, self._doc.TablesOfContents.Count + 1):
            self._doc.TablesOfContents.Item(index).Update()
        self._doc.Fields.Update()
        total_pages = int(self._doc.ComputeStatistics(2))
        toc_pages = int(self._doc.TablesOfContents.Count)
        return (
            FieldSnapshot(
                stable_key=("doc:finalize", "PAGE", 0),
                field_category="page",
                result_hash=str(total_pages) + "-" + str(toc_pages),
                toc_page_count=toc_pages,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=total_pages,
            ),
        )

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


    # ---- long-form M2 fallback / placeholder primitives ----
    def degradation_checkpoint(self):
        """Return a local end-of-document checkpoint for one plan node."""
        try:
            return self._native_document_end()
        except Exception:
            raise NativeWriterObjectError(
                "LOCAL_MUTATION_CHECKPOINT_FAILED", "checkpoint failed"
            ) from None

    def rollback_degradation_checkpoint(self, checkpoint):
        """Roll back one failed plan-node mutation to its saved checkpoint."""
        try:
            self._native_rollback(int(checkpoint), self._native_document_end())
        except Exception:
            raise NativeWriterObjectError(
                "LOCAL_MUTATION_ROLLBACK_FAILED", "rollback failed"
            ) from None

    @staticmethod
    def _degradation_display(code, fallback_text, *, inline=False):
        safe_code = (
            code
            if isinstance(code, str) and _PUBLIC_ISSUE_CODE_RE.fullmatch(code)
            else "DEGRADATION"
        )
        safe_text = redact_private_text(str(fallback_text or ""))
        prefix = f"[{safe_code}"
        if safe_text.startswith(prefix) and safe_text[len(prefix):len(prefix) + 1] in {
            "]",
            " ",
            ":",
        }:
            return safe_text
        if not safe_text or safe_text == safe_code:
            return f"[{safe_code}]"
        if inline:
            return f"[{safe_code}: {safe_text}]"
        return f"[{safe_code}] {safe_text}"

    @staticmethod
    def _style_degradation_range(native_range):
        """Apply the shared restrained fallback style to one native range."""
        native_range.Font.Italic = True
        native_range.Font.Color = hex_to_rgb_long("#9C0006")
        native_range.Shading.BackgroundPatternColor = hex_to_rgb_long("#FCE8E6")

    def _insert_degradation_box(self, display, target_range=None):
        """Insert one non-outline, single-cell degradation box."""
        target = None
        try:
            target = target_range or self._doc.Range(
                int(self.selection.End), int(self.selection.End)
            )
            table = self._doc.Tables.Add(target, 1, 1)
            cell_range = table.Cell(1, 1).Range
            cell_range.Text = display
            self._style_degradation_range(cell_range)
            cell_range.ParagraphFormat.SpaceBefore = 0
            cell_range.ParagraphFormat.SpaceAfter = 3
            cell_range.ParagraphFormat.KeepTogether = True
            cell_range.ParagraphFormat.OutlineLevel = 10
            table.Rows.AllowBreakAcrossPages = False
            return table
        except NativeWriterObjectError:
            raise
        except Exception:
            try:
                if target is None:
                    raise AttributeError("degradation anchor unavailable")
                start = int(target.Start)
                insert_after = getattr(target, "InsertAfter", None)
                if callable(insert_after):
                    insert_after(display)
                else:
                    target.Text = display
                inserted = self._doc.Range(start, start + len(display))
                self._style_degradation_range(inserted)
                inserted.ParagraphFormat.SpaceBefore = 0
                inserted.ParagraphFormat.SpaceAfter = 3
                inserted.ParagraphFormat.KeepTogether = True
                inserted.ParagraphFormat.OutlineLevel = 10
                return SimpleNamespace(Range=inserted)
            except Exception:
                raise NativeWriterObjectError(
                    "DEGRADATION_INSERT_FAILED", "degradation box insertion failed"
                ) from None

    def add_degradation_notice(self, code, message, fallback_text, placement="block"):
        """Insert a deterministic styled degradation at its semantic anchor."""
        if placement == "inline":
            return self.add_inline_degradation(code, message, fallback_text)
        display = self._degradation_display(code, fallback_text)
        return self._insert_degradation_box(display)

    def add_inline_degradation(self, code, message, fallback_text):
        """Insert and style a fallback without creating a new paragraph."""
        display = self._degradation_display(code, fallback_text, inline=True)
        selection = self.selection
        start = int(selection.End)
        try:
            selection.TypeText(display)
            inserted = self._doc.Range(start, int(selection.End))
            self._style_degradation_range(inserted)
            return inserted
        except Exception:
            raise NativeWriterObjectError(
                "DEGRADATION_INSERT_FAILED", "inline degradation insertion failed"
            ) from None

    def reserve_document_quality_anchor(self, title="生成质量提示", notices=()):
        """Reserve the fixed, body-style quality anchor even when it is empty."""
        if not hasattr(self, "_quality_notice_seen"):
            self._quality_notice_seen = set()
        if not hasattr(self, "_quality_notice_anchor_position"):
            safe_title = redact_private_text(str(title or "生成质量提示"))
            try:
                position = int(self.selection.End)
                anchor = self._doc.Range(
                    position,
                    position,
                )
                bookmarks = getattr(self._doc, "Bookmarks", None)
                if bookmarks is None or not callable(getattr(bookmarks, "Add", None)):
                    raise AttributeError("bookmark API unavailable")
                bookmarks.Add("wpsc_document_quality_anchor", anchor)
                self._quality_notice_anchor_position = position
                self._quality_notice_title = safe_title
            except Exception:
                raise NativeWriterObjectError(
                    "DEGRADATION_INSERT_FAILED", "quality anchor insertion failed"
                ) from None
        for notice in notices or ():
            self._upsert_quality_notice_mapping(notice)

    def _upsert_quality_notice_mapping(self, notice):
        raw_code = (
            notice.get("code")
            if isinstance(notice, dict)
            else getattr(notice, "code", None)
        )
        code = (
            raw_code
            if isinstance(raw_code, str) and _PUBLIC_ISSUE_CODE_RE.fullmatch(raw_code)
            else "QUALITY_NOTICE"
        )
        placement = (
            notice.get("placement", "document")
            if isinstance(notice, dict)
            else getattr(notice, "placement", "document")
        )
        node_id = (
            notice.get("nodeId")
            if isinstance(notice, dict)
            else getattr(notice, "node_id", None)
        )
        identity = (code, placement, redact_private_text(str(node_id or "")))
        if identity in self._quality_notice_seen:
            return
        fallback_text = (
            notice.get("fallbackText") or notice.get("message", "")
            if isinstance(notice, dict)
            else getattr(notice, "message", "")
        )
        try:
            position = int(self._quality_notice_anchor_position)
            target = self._doc.Range(position, position)
            display = self._degradation_display(code, fallback_text)
            if not self._quality_notice_seen:
                display = f"{self._quality_notice_title}\r{display}"
            table = self._insert_degradation_box(
                display, target
            )
            self._quality_notice_anchor_position = int(
                getattr(getattr(table, "Range", None), "End", self.selection.End)
            )
        except Exception:
            raise NativeWriterObjectError(
                "DEGRADATION_INSERT_FAILED", "quality notice upsert failed"
            ) from None
        self._quality_notice_seen.add(identity)

    def upsert_document_quality_notice(self, issue):
        """Upsert one runtime document issue at the reserved quality anchor."""
        if not hasattr(self, "_quality_notice_anchor_position"):
            raise NativeWriterObjectError(
                "DEGRADATION_INSERT_FAILED", "quality anchor is unavailable"
            )
        self._upsert_quality_notice_mapping(issue)

    def add_document_quality_notice(self, notices):
        """Backward-compatible document notice API using the fixed anchor."""
        if not hasattr(self, "_quality_notice_anchor_position"):
            self.reserve_document_quality_anchor(notices=notices)
            return
        for notice in notices or ():
            self._upsert_quality_notice_mapping(notice)

    def add_quality_notice_at_bookmark(
        self, *, code, message, fallback, node_id, page, bookmark_name=None
    ):
        """Insert one M5 notice beside a persisted native semantic bookmark."""
        try:
            name = bookmark_name or "wpsc_document_quality_anchor"
            bookmarks = self._doc.Bookmarks
            exists = getattr(bookmarks, "Exists", None)
            if callable(exists) and not exists(name):
                raise KeyError(name)
            try:
                bookmark = bookmarks(name)
            except Exception:
                bookmark = bookmarks.Item(name)
            anchor = bookmark.Range
            if bookmark_name:
                paragraph = anchor.Paragraphs(1).Range
                position = int(paragraph.End)
            else:
                position = int(anchor.Start)
            target = self._doc.Range(position, position)
            display = self._degradation_display(
                code,
                f"{redact_private_text(str(message))} "
                f"(page {int(page)}; {redact_private_text(str(fallback))})",
            )
            return self._insert_degradation_box(display, target)
        except Exception:
            raise NativeWriterObjectError(
                "DEGRADATION_INSERT_FAILED", "quality notice insertion failed"
            ) from None

    def pagination_fragment_for_bookmark(self, node_id, bookmark_name):
        """Return privacy-safe M5 point geometry for one persisted bookmark."""
        try:
            bookmarks = self._doc.Bookmarks
            try:
                bookmark = bookmarks(bookmark_name)
            except Exception:
                bookmark = bookmarks.Item(bookmark_name)
            rng = bookmark.Range.Paragraphs(1).Range
            start, end = int(rng.Start), int(rng.End)
            page = int(rng.Information(3))
            x = float(rng.Information(5))
            y = float(rng.Information(6))
            fragment = {"page": page}
            if all(math.isfinite(value) and value >= 0 for value in (x, y)):
                fragment["bounds"] = [x, y, x + 1.0, y + 12.0]
            return {
                "nodeId": redact_private_text(str(node_id)),
                "story": "main",
                "sections": ["body"],
                "pageStart": page,
                "pageEnd": page,
                "range": f"{start}:{end}",
                "fragments": [fragment],
            }
        except Exception:
            raise NativeWriterObjectError(
                "PAGINATION_SNAPSHOT_FAILED", "pagination snapshot failed"
            ) from None

    def pagination_map_for_ranges(self, tracked_ranges):
        """Build the M5 native range/page snapshot after final repagination."""
        try:
            repaginate = getattr(self._doc, "Repaginate", None)
            if not callable(repaginate):
                raise AttributeError("repagination API unavailable")
            repaginate()
            setup = self._doc.PageSetup
            page_width = float(getattr(setup, "PageWidth", 595.28))
            page_height = float(getattr(setup, "PageHeight", 841.89))
            left_margin = max(0.0, float(getattr(setup, "LeftMargin", 72.0)))
            right_margin = max(0.0, float(getattr(setup, "RightMargin", 72.0)))
            top_margin = max(0.0, float(getattr(setup, "TopMargin", 72.0)))
            bottom_margin = max(0.0, float(getattr(setup, "BottomMargin", 72.0)))
            visual_ops = {
                "writer.add_heading",
                "writer.add_captioned_figure",
                "writer.add_semantic_table",
                "writer.add_equation",
                "writer.add_degradation_notice",
                "writer.add_document_quality_notice",
            }
            nodes = []
            seen = set()
            content_end = max(0, int(self._doc.Content.End) - 1)
            for tracked in tracked_ranges:
                node_id = str(tracked.get("nodeId") or "")
                if not node_id or node_id in seen:
                    continue
                rng = tracked["range"]
                start, end = int(rng.Start), int(rng.End)
                if start < 0 or end < start:
                    raise ValueError("invalid native range")
                first_range = self._doc.Range(min(start, content_end), min(start, content_end))
                last_position = end - 1 if end > start else end
                last_position = min(max(0, last_position), content_end)
                last_range = self._doc.Range(last_position, last_position)
                first_page = int(first_range.Information(3))
                last_page = int(last_range.Information(3))
                if first_page < 1 or last_page < first_page:
                    raise ValueError("invalid native page span")
                visual = tracked.get("op") in visual_ops
                first_x = float(first_range.Information(5))
                first_y = float(first_range.Information(6))
                last_y = float(last_range.Information(6))
                usable_page = (
                    100.0 <= page_width <= 2000.0
                    and 100.0 <= page_height <= 2000.0
                    and left_margin + right_margin < page_width
                    and top_margin + bottom_margin < page_height
                )
                usable_points = (
                    usable_page
                    and all(
                        math.isfinite(value)
                        for value in (first_x, first_y, last_y)
                    )
                    and 0.0 <= first_x <= page_width
                    and 0.0 <= first_y <= page_height
                    and 0.0 <= last_y <= page_height
                )
                fragments = []
                for page in range(first_page, last_page + 1):
                    fragment = {"page": page}
                    if visual and usable_points:
                        x0 = max(0.0, first_x) if page == first_page and math.isfinite(first_x) else left_margin
                        y0 = max(0.0, first_y) if page == first_page and math.isfinite(first_y) else top_margin
                        x1 = max(x0 + 1.0, page_width - right_margin)
                        y1 = (
                            max(y0 + 1.0, min(page_height, last_y + 12.0))
                            if page == last_page and math.isfinite(last_y)
                            else max(y0 + 1.0, page_height - bottom_margin)
                        )
                        if not all(math.isfinite(value) for value in (x0, y0, x1, y1)):
                            raise ValueError("invalid native bounds")
                        fragment["bounds"] = [x0, y0, x1, y1]
                    fragments.append(fragment)
                seen.add(node_id)
                nodes.append({
                    "nodeId": redact_private_text(node_id),
                    "story": "main",
                    "sections": [str(tracked.get("role") or "body")],
                    "pageStart": first_page,
                    "pageEnd": last_page,
                    "range": f"{start}:{end}",
                    "fragments": fragments,
                })
            return {"version": "M5-v1", "nodes": nodes}
        except Exception:
            raise NativeWriterObjectError(
                "PAGINATION_SNAPSHOT_FAILED", "pagination snapshot failed"
            ) from None

    def insert_figure_index(self, title=None):
        """Insert a figure index placeholder (content population is M3)."""
        if title:
            self.add_heading(str(title), size=14, bold=True)
        self.add_paragraph("[Figure index placeholder]", style="Body Text")

    def insert_table_index(self, title=None):
        """Insert a table index placeholder (content population is M3)."""
        if title:
            self.add_heading(str(title), size=14, bold=True)
        self.add_paragraph("[Table index placeholder]", style="Body Text")

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

        # Stable paragraph ids read from the saved .docx on
        # disk. Only present for saved docs; unsaved/never-saved docs fall
        # back to purely positional ids. The list is 0-based, aligned to
        # Paragraphs(1..N); a count mismatch falls back to positional too.
        paraids = self._read_paraid_map(paragraph_count)

        paragraphs = []
        for index in range(1, min(paragraph_count, int(limit) if limit != float("inf") else paragraph_count) + 1):
            para = doc.Paragraphs(index)
            paraid = paraids[index - 1] if index <= len(paraids) else None
            # Prefer the stable id when available so agents naturally use it;
            # keep positional via the "index" field as a fallback.
            element_id = (
                "paragraph:@paraId=%s" % paraid if paraid
                else "paragraph:%d" % index
            )
            snap = self._range_snapshot(para.Range, element_id)
            snap["index"] = index
            if paraid:
                snap["para_id"] = paraid
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

        match = re.fullmatch(r"paragraph:@paraId=([0-9A-Fa-f]+)", target)
        if match:
            # Resolve a stable paraId back to a positional
            # Paragraphs(index) by re-reading the docx paraid list. Raises
            # ValueError (-> invalid_target) when the id is not present.
            index = self._paragraph_index_for_paraid(match.group(1))
            if index is None:
                raise ValueError(f"Unsupported Writer target: {target}")
            rng = self._doc.Paragraphs(index).Range
            return self._patch_range(rng, text, font, paragraph, style)

        match = re.fullmatch(r"paragraph:(\d+)", target)
        if match:
            index = int(match.group(1))
            count = int(safe_get(self._doc.Paragraphs, "Count", 0) or 0)
            if not 1 <= index <= count:
                raise ValueError(f"Unsupported Writer target: {target}")
            rng = self._doc.Paragraphs(index).Range
            return self._patch_range(rng, text, font, paragraph, style)

        match = re.fullmatch(r"range:(\d+)-(\d+)", target)
        if match:
            rng = self._doc.Range(int(match.group(1)), int(match.group(2)))
            return self._patch_range(rng, text, font, paragraph, style)

        match = re.fullmatch(r"table:(\d+)/cell:(\d+),(\d+)", target)
        if match:
            tindex = int(match.group(1))
            if not 1 <= tindex <= int(safe_get(self._doc.Tables, "Count", 0) or 0):
                raise ValueError(f"Unsupported Writer target: {target}")
            table = self._doc.Tables(tindex)
            try:
                cell = table.Cell(int(match.group(2)), int(match.group(3)))
            except Exception:
                raise ValueError(f"Unsupported Writer target: {target}")
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
            index = int(match.group(1))
            if not 1 <= index <= int(safe_get(self._doc.Shapes, "Count", 0) or 0):
                raise ValueError(f"Unsupported Writer target: {target}")
            shape = self._doc.Shapes(index)
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

    # ------------------------------------------------------------------
    # Stable paragraph-id support (w14:paraId). The alignment assumption:
    # the main-story walk of <w:p> in word/document.xml (text-box story
    # pruned, one None per table row-end mark) matches doc.Paragraphs(1..N).
    # ------------------------------------------------------------------
    def _read_paraid_map(self, expected_count):
        """Return the paraId list for the current document, or ``[]`` if it
        cannot be read (unsaved doc, non-docx source, read failure).

        Memoized per (document path, paragraph count): the docx on disk is
        stable while the composer holds the document open, so repeated
        ``@paraId`` resolutions in one ``edit()`` batch do not re-read the zip.
        """
        path = safe_get(self._doc, "FullName")
        cache = getattr(self, "_paraid_cache", None)
        if cache is not None and cache[0] == path and cache[1] == expected_count:
            return cache[2]
        if not path:
            return []
        paraids = read_paraids_from_docx(path)
        if expected_count is not None and len(paraids) != int(expected_count):
            # Mismatch means our document-order assumption is wrong for this
            # doc -- fall back to positional rather than emit wrong ids.
            return []
        self._paraid_cache = (path, expected_count, paraids)
        return paraids

    def _paragraph_index_for_paraid(self, paraid):
        """Return the 1-based Paragraphs index for *paraid*, or ``None``."""
        count = int(safe_get(self._doc.Paragraphs, "Count", 0) or 0)
        paraids = self._read_paraid_map(count)
        if not paraids:
            return None
        normalized = paraid.lower()
        for offset, value in enumerate(paraids):
            if value and value.lower() == normalized:
                return offset + 1
        return None

    # ------------------------------------------------------------------
    # Structural verbs (insert / remove / move / clone).
    # Uses standard Word/WPS primitives
    # (Range.InsertAfter / .Delete / .Cut / .Copy / .Paste). Returned
    # "path" is a best-effort POSITIONAL id (re-inspect for a stable
    # @paraId after a save). move/clone use clipboard (Cut/Copy + Paste).
    # ------------------------------------------------------------------
    def apply_structural_op(self, op):
        verb = op.get("op")
        if verb == "insert":
            return self._insert_element(
                op.get("parent", "body"), op.get("type"),
                op.get("props") or {}, op.get("position", "end"),
            )
        if verb == "remove":
            return self._remove_element(op.get("target"))
        if verb == "move":
            return self._move_or_clone(op.get("target"), op.get("to", "end"), cut=True)
        if verb == "clone":
            return self._move_or_clone(op.get("target"), op.get("to", "end"), cut=False)
        raise ValueError(f"Unsupported Writer structural op: {verb!r}")

    def _paragraph_range_for_target(self, target):
        """Resolve a paragraph target (positional or @paraId) to a COM Range,
        or raise ValueError."""
        match = re.fullmatch(r"paragraph:(\d+)", target or "")
        if match:
            index = int(match.group(1))
            if not 1 <= index <= int(safe_get(self._doc.Paragraphs, "Count", 0) or 0):
                raise ValueError(f"Unsupported Writer target: {target}")
            return self._doc.Paragraphs(index).Range
        match = re.fullmatch(r"paragraph:@paraId=([0-9A-Fa-f]+)", target or "")
        if match:
            index = self._paragraph_index_for_paraid(match.group(1))
            if index is None:
                raise ValueError(f"Unsupported Writer target: {target}")
            return self._doc.Paragraphs(index).Range
        raise ValueError(f"Unsupported Writer target: {target}")

    def _resolve_insert_range(self, position):
        """Return a COM Range at which to insert (collapsed appropriately).
        Supports 'end', 'start', {'after': ptarget}, {'before': ptarget},
        {'index': N}."""
        doc = self._doc
        if position in (None, "end"):
            rng = doc.Content
            rng.Collapse(0)  # wdCollapseEnd
            return rng
        if position == "start":
            rng = doc.Content
            rng.Collapse(1)  # wdCollapseStart
            return rng
        if isinstance(position, dict):
            if "after" in position:
                rng = self._paragraph_range_for_target(position["after"])
                rng.Collapse(0)
                return rng
            if "before" in position:
                rng = self._paragraph_range_for_target(position["before"])
                rng.Collapse(1)  # wdCollapseStart
                return rng
            if "index" in position:
                index = int(position["index"])
                if not 1 <= index <= int(safe_get(doc.Paragraphs, "Count", 0) or 0):
                    raise ValueError(f"Unsupported Writer insert position: {position!r}")
                return doc.Paragraphs(index).Range
        raise ValueError(f"Unsupported Writer insert position: {position!r}")

    def _insert_element(self, parent, etype, props, position):
        doc = self._doc
        if etype == "page_break":
            rng = self._resolve_insert_range(position)
            rng.InsertBreak(7)  # wdPageBreak
            return {"type": "page_break"}
        if etype in ("paragraph", "heading"):
            text = str(props.get("text", ""))
            rng = self._resolve_insert_range(position)
            rng.InsertAfter(text + "\r")
            if etype == "heading":
                # InsertAfter expanded rng over the new text, so the style
                # applies to the inserted paragraph. Use the locale-independent
                # built-in style id (wdStyleHeading1..9 = -2..-10): the English
                # name "Heading N" does not exist in localized WPS builds.
                level = min(max(int(props.get("level", 1)), 1), 9)
                try:
                    rng.Style = self._doc.Styles(-1 - level)
                except Exception:
                    pass
            try:
                new_index = int(safe_get(doc.Paragraphs, "Count", 0) or 0)
                if position in (None, "end") and new_index:
                    return {"type": etype, "path": f"paragraph:{new_index}"}
            except Exception:
                pass
            return {"type": etype}
        if etype == "table":
            rng = self._resolve_insert_range(position)
            rows = int(props.get("rows", 2))
            cols = int(props.get("cols", 2))
            table = doc.Tables.Add(rng, rows, cols)
            data = props.get("data") or []
            for r, row in enumerate(data[:rows], start=1):
                for c, value in enumerate(row[:cols], start=1):
                    try:
                        table.Cell(r, c).Range.Text = str(value)
                    except Exception:
                        pass
            table_index = int(safe_get(doc.Tables, "Count", 0) or 0)
            return {"type": "table", "path": f"table:{table_index}"}
        if etype == "image":
            path = str(props.get("path", ""))
            rng = self._resolve_insert_range(position)
            doc.InlineShapes.AddPicture(path, False, True, rng)
            shape_index = int(safe_get(doc.InlineShapes, "Count", 0) or 0)
            return {"type": "image", "path": f"inline_shape:{shape_index}"}
        if etype == "textbox":
            left = float(props.get("left", 100))
            top = float(props.get("top", 100))
            width = float(props.get("width", 200))
            height = float(props.get("height", 50))
            shape = doc.Shapes.AddTextbox(1, left, top, width, height)
            text = props.get("text")
            if text is not None:
                tr = safe_get(safe_get(shape, "TextFrame"), "TextRange")
                if tr is not None:
                    safe_set(tr, "Text", str(text))
            shape_index = int(safe_get(doc.Shapes, "Count", 0) or 0)
            return {"type": "textbox", "path": f"shape:{shape_index}"}
        raise ValueError(f"Unsupported Writer insert type: {etype!r}")

    def _structural_target(self, target):
        """Resolve a structural target to (kind, object). kind is 'paragraph',
        'shape', 'inline_shape', or 'table'; object is the COM Range/Shape.
        Raises ValueError on unrecognised targets."""
        target = target or ""
        m = re.fullmatch(r"paragraph:(\d+)", target)
        if m:
            index = int(m.group(1))
            if not 1 <= index <= int(safe_get(self._doc.Paragraphs, "Count", 0) or 0):
                raise ValueError(f"Unsupported Writer target: {target}")
            return "paragraph", self._doc.Paragraphs(index).Range
        m = re.fullmatch(r"paragraph:@paraId=([0-9A-Fa-f]+)", target)
        if m:
            idx = self._paragraph_index_for_paraid(m.group(1))
            if idx is None:
                raise ValueError(f"Unsupported Writer target: {target}")
            return "paragraph", self._doc.Paragraphs(idx).Range
        m = re.fullmatch(r"shape:(\d+)", target)
        if m:
            index = int(m.group(1))
            if not 1 <= index <= int(safe_get(self._doc.Shapes, "Count", 0) or 0):
                raise ValueError(f"Unsupported Writer target: {target}")
            return "shape", self._doc.Shapes(index)
        m = re.fullmatch(r"inline_shape:(\d+)", target)
        if m:
            index = int(m.group(1))
            if not 1 <= index <= int(safe_get(self._doc.InlineShapes, "Count", 0) or 0):
                raise ValueError(f"Unsupported Writer target: {target}")
            return "inline_shape", self._doc.InlineShapes(index).Range
        m = re.fullmatch(r"table:(\d+)", target)
        if m:
            index = int(m.group(1))
            if not 1 <= index <= int(safe_get(self._doc.Tables, "Count", 0) or 0):
                raise ValueError(f"Unsupported Writer target: {target}")
            return "table", self._doc.Tables(index)
        raise ValueError(f"Unsupported Writer target: {target}")

    def _remove_element(self, target):
        kind, obj = self._structural_target(target)
        obj.Delete()
        return {"removed": target, "kind": kind}

    def _move_or_clone(self, target, to, *, cut):
        """Move (cut=True) or clone (cut=False) a paragraph/table/shape to *to*
        via the clipboard.

        WPS quirks handled here (verified on Windows):
        * ``Table.Range.Cut()`` is a silent no-op (paste then duplicates) —
          move is Copy + explicit ``Table.Delete()``.
        * ``Shape.Cut()``/``Copy()`` raise ``<unknown>.Cut`` — go through
          ``Select()`` + ``Selection.Cut()/Copy()`` instead.
        """
        kind, obj = self._structural_target(target)
        if kind == "shape":
            obj.Select()
            selection = self._app.Selection
            if cut:
                selection.Cut()
            else:
                selection.Copy()
        elif kind == "table":
            obj.Range.Copy()
            if cut:
                obj.Delete()
        else:
            if cut:
                obj.Cut()
            else:
                obj.Copy()
        dest = self._resolve_insert_range(to)
        dest.Paste()
        return {"type": kind, "moved": cut, "from": target}

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
