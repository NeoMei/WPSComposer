"""SheetComposer - WPS Spreadsheets (Ket.Application) to xlsx.

Styled spreadsheets with formulas, charts, conditional formatting,
freeze panes, and merged cells.
"""

from __future__ import annotations

import re

from ._dispatch import (
    _abs, FMT_XLSX, FMT_PDF_FROM_XLS, FMT_XLS, FMT_XLSM, FMT_XLSB,
    FMT_XLTX, FMT_XLTM, FMT_CSV, FMT_TSV, FMT_ODS,
)
from ._colors import hex_to_rgb_long
from ._base import BaseComposer
from .formatting import (
    apply_fill,
    apply_font,
    apply_geometry,
    apply_line,
    color_hex,
    fill_snapshot,
    font_snapshot,
    geometry_snapshot,
    line_snapshot,
    merge_results,
    page_setup_snapshot,
    safe_get,
    safe_set,
)


# ===========================================================================

class SheetComposer(BaseComposer):
    _progids = ("Ket.Application", "Excel.Application")
    _doc_type = "calc"
    _native_fmt = FMT_XLSX
    _pdf_fmt = FMT_PDF_FROM_XLS
    _formats_by_extension = {
        ".xlsx": FMT_XLSX, ".xls": FMT_XLS, ".xlsm": FMT_XLSM,
        ".xlsb": FMT_XLSB, ".xltx": FMT_XLTX, ".xltm": FMT_XLTM,
        ".csv": FMT_CSV, ".tsv": FMT_TSV, ".ods": FMT_ODS,
    }

    @staticmethod
    def _create_doc(app):
        return app.Workbooks.Add()

    @staticmethod
    def _open_document(app, path, read_only=False):
        return app.Workbooks.Open(path, 0, bool(read_only))

    @staticmethod
    def _active_document(app):
        return app.ActiveWorkbook

    @property
    def ws(self):
        return self._ws


    def select_sheet(self, index):
        self._ws = self._doc.Worksheets(index)
        return self._ws

    def rename_sheet(self, index, name):
        ws = self._doc.Worksheets(index)
        ws.Name = name
        return ws

    def add_sheet(self, name=None):
        count = self._doc.Worksheets.Count
        ws = self._doc.Worksheets.Add(None, self._doc.Worksheets(count))
        if name:
            ws.Name = name
        self._ws = ws
        return ws

    def write_cell(self, row, col, value):
        self.ws.Cells(row, col).Value = value

    def set_formula(self, row, col, formula):
        self.ws.Cells(row, col).Formula = formula

    @property
    def ws(self):
        if not hasattr(self, "_ws") or self._ws is None:
            self._ws = self._doc.Worksheets(1)
        return self._ws

    @ws.setter
    def ws(self, value):
        self._ws = value

    def write_table(self, start_row, start_col, data, header_bold=True,
                    header_shade="#4472C4", header_font_color="#FFFFFF",
                    font_size=11):
        for r, rowvals in enumerate(data):
            for c, v in enumerate(rowvals):
                cell = self.ws.Cells(start_row + r, start_col + c)
                cell.Value = v
                cell.Font.Size = font_size
                if r == 0:
                    if header_bold:
                        cell.Font.Bold = True
                    if header_font_color:
                        cell.Font.Color = hex_to_rgb_long(header_font_color)
                    if header_shade:
                        cell.Interior.Color = hex_to_rgb_long(header_shade)
        if header_shade or header_bold:
            for c in range(len(data[0])):
                cell = self.ws.Cells(start_row, start_col + c)
                if header_bold:
                    cell.Font.Bold = True
                if header_shade:
                    cell.Interior.Color = hex_to_rgb_long(header_shade)

    def set_cell_style(self, row, col, bold=None, italic=None, font_size=None,
                       font_color=None, fill_color=None, align=None,
                       number_format=None):
        cell = self.ws.Cells(row, col)
        if bold is not None:
            cell.Font.Bold = bold
        if italic is not None:
            cell.Font.Italic = italic
        if font_size is not None:
            cell.Font.Size = font_size
        if font_color is not None:
            cell.Font.Color = hex_to_rgb_long(font_color)
        if fill_color is not None:
            cell.Interior.Color = hex_to_rgb_long(fill_color)
        if align is not None:
            cell.HorizontalAlignment = align  # -4108 center, -4131 left, -4152 right
        if number_format is not None:
            cell.NumberFormat = number_format

    def set_range_style(self, range_str, bold=None, italic=None, font_size=None,
                        font_color=None, fill_color=None, align=None,
                        number_format=None):
        rng = self.ws.Range(range_str)
        if bold is not None:
            rng.Font.Bold = bold
        if italic is not None:
            rng.Font.Italic = italic
        if font_size is not None:
            rng.Font.Size = font_size
        if font_color is not None:
            rng.Font.Color = hex_to_rgb_long(font_color)
        if fill_color is not None:
            rng.Interior.Color = hex_to_rgb_long(fill_color)
        if align is not None:
            rng.HorizontalAlignment = align
        if number_format is not None:
            rng.NumberFormat = number_format

    def set_borders(self, range_str, style=1, weight=2, color="#000000"):
        """style: 1=continuous, weight: 2=thin."""
        rng = self.ws.Range(range_str)
        for edge in (7, 8, 9, 10, 11, 12):  # left,right,top,bottom,insideH,insideV
            try:
                b = rng.Borders(edge)
                b.LineStyle = style
                b.Weight = weight
                b.Color = hex_to_rgb_long(color)
            except Exception:
                pass

    def merge_cells(self, range_str):
        self.ws.Range(range_str).Merge()

    def freeze_panes(self, cell):
        self._app.ActiveWindow.SplitColumn = 0
        self.ws.Range(cell).Select()
        self._app.ActiveWindow.FreezePanes = True

    def set_column_width(self, col, width):
        self.ws.Columns(col).ColumnWidth = width

    def set_row_height(self, row, height):
        self.ws.Rows(row).RowHeight = height

    def autofit(self):
        try:
            self.ws.UsedRange.Columns.AutoFit()
        except Exception:
            pass

    def add_chart(self, chart_type=4, left=100, top=100, width=400, height=300,
                  source_range=None, title=None):
        """chart_type: xlColumnClustered=51, xlLine=4, xlPie=5, xlBarClustered=57.
        Returns chart object."""
        if source_range:
            rng = self.ws.Range(source_range) if isinstance(source_range, str) else source_range
            charts = self.ws.ChartObjects()
            co = charts.Add(left, top, width, height)
            ch = co.Chart
            ch.SetSourceData(rng)
            try:
                ch.ChartType = chart_type
            except Exception:
                pass
            if title:
                try:
                    ch.HasTitle = True
                    ch.ChartTitle.Text = title
                except Exception:
                    pass
            return co
        else:
            return self.ws.Shapes.AddChart(chart_type, left, top, width, height)

    def add_title_row(self, range_str, text, fill_color="#4472C4",
                      font_color="#FFFFFF", size=14):
        rng = self.ws.Range(range_str)
        rng.Merge()
        rng.Value = text
        rng.Font.Bold = True
        rng.Font.Size = size
        rng.Font.Color = hex_to_rgb_long(font_color)
        rng.Interior.Color = hex_to_rgb_long(fill_color)
        rng.HorizontalAlignment = -4108  # center

    def conditional_format(self, range_str, rule_type="cellvalue",
                           operator=1, formula="0", fill_color="#FFC7CE"):
        """Simple highlight rule. operator: 1=between-like; use formula for value."""
        rng = self.ws.Range(range_str)
        try:
            rng.FormatConditions.Delete()
            fmt = rng.FormatConditions.Add(rule_type, operator, formula)
            fmt.Interior.Color = hex_to_rgb_long(fill_color)
        except Exception:
            pass

    def set_header_footer(self, left=None, center=None, right=None):
        ps = self.ws.PageSetup
        if left:
            ps.LeftHeader = left
        if center:
            ps.CenterHeader = center
        if right:
            ps.RightHeader = right

    # ================================================================
    # Existing-workbook inspection and element-level editing
    # ================================================================

    def inspect_selection(self):
        """Return the active selected range and its effective formatting."""
        selection = safe_get(self._app, "Selection")
        return self._range_snapshot(selection, "selection", include_value=True)

    def inspect_document(self, include_cells=True, max_cells=1000,
                         include_empty=False):
        """Return workbook structure, values, formulas, and effective formats.

        ``max_cells`` prevents an accidentally formatted entire column from
        flooding the agent context.  Increase it explicitly when required.
        """
        workbook = self._doc
        sheet_count = int(safe_get(workbook.Worksheets, "Count", 0) or 0)
        sheets = []
        remaining = max_cells
        truncated = False

        for sheet_index in range(1, sheet_count + 1):
            ws = workbook.Worksheets(sheet_index)
            used = ws.UsedRange
            row_count = int(safe_get(used.Rows, "Count", 0) or 0)
            col_count = int(safe_get(used.Columns, "Count", 0) or 0)
            first_row = int(safe_get(used, "Row", 1) or 1)
            first_col = int(safe_get(used, "Column", 1) or 1)
            cells = []
            if include_cells:
                for r_offset in range(row_count):
                    for c_offset in range(col_count):
                        if remaining <= 0:
                            truncated = True
                            break
                        cell = ws.Cells(first_row + r_offset, first_col + c_offset)
                        value = safe_get(cell, "Value")
                        formula = safe_get(cell, "Formula")
                        if not include_empty and value in (None, "") and formula in (None, ""):
                            continue
                        cells.append(self._range_snapshot(
                            cell,
                            f"sheet:{sheet_index}/cell:{safe_get(cell, 'Address', '')}",
                            include_value=True,
                        ))
                        remaining -= 1
                    if truncated:
                        break

            shapes = []
            shape_count = int(safe_get(ws.Shapes, "Count", 0) or 0)
            for shape_index in range(1, shape_count + 1):
                shape = ws.Shapes(shape_index)
                shapes.append({
                    "id": f"sheet:{sheet_index}/shape:{shape_index}",
                    "index": shape_index,
                    "name": safe_get(shape, "Name"),
                    "type": safe_get(shape, "Type"),
                    "geometry": geometry_snapshot(shape),
                    "fill": fill_snapshot(safe_get(shape, "Fill")),
                    "line": line_snapshot(safe_get(shape, "Line")),
                })

            chart_count = int(safe_get(ws.ChartObjects(), "Count", 0) or 0)
            charts = []
            for chart_index in range(1, chart_count + 1):
                chart_object = ws.ChartObjects(chart_index)
                chart = chart_object.Chart
                charts.append({
                    "id": f"sheet:{sheet_index}/chart:{chart_index}",
                    "index": chart_index,
                    "name": safe_get(chart_object, "Name"),
                    "chart_type": safe_get(chart, "ChartType"),
                    "has_title": safe_get(chart, "HasTitle"),
                    "title": safe_get(safe_get(chart, "ChartTitle"), "Text"),
                    "geometry": geometry_snapshot(chart_object),
                })

            sheets.append({
                "id": f"sheet:{sheet_index}",
                "index": sheet_index,
                "name": safe_get(ws, "Name"),
                "visible": safe_get(ws, "Visible"),
                "used_range": safe_get(used, "Address"),
                "used_rows": row_count,
                "used_columns": col_count,
                "page_setup": page_setup_snapshot(ws.PageSetup),
                "freeze_panes": safe_get(safe_get(self._app, "ActiveWindow"), "FreezePanes")
                    if safe_get(workbook, "ActiveSheet") == ws else None,
                "cells": cells,
                "shapes": shapes,
                "charts": charts,
            })

        return {
            "kind": "sheet",
            "name": safe_get(workbook, "Name"),
            "path": safe_get(workbook, "FullName"),
            "saved": safe_get(workbook, "Saved"),
            "sheet_count": sheet_count,
            "cells_truncated": truncated,
            "sheets": sheets,
        }

    def apply_format_patch(self, target, *, value=None, formula=None, font=None,
                           fill=None, line=None, geometry=None,
                           number_format=None, horizontal_alignment=None,
                           vertical_alignment=None, wrap_text=None, indent=None,
                           row_height=None, column_width=None, borders=None,
                           page_setup=None, name=None, chart_type=None,
                           chart_title=None):
        """Apply a partial patch to a range, shape, chart, or worksheet.

        Targets include ``selection``, ``sheet:N/cell:A1``,
        ``sheet:N/range:A1:C8``, ``sheet:N/shape:N``, ``sheet:N/chart:N``, and
        ``sheet:N``.  A leading ``$`` in cell addresses is accepted.
        """
        font, fill, line, geometry = font or {}, fill or {}, line or {}, geometry or {}
        if target == "selection":
            return self._patch_range(
                self._app.Selection, value, formula, font, fill, number_format,
                horizontal_alignment, vertical_alignment, wrap_text, indent,
                row_height, column_width, borders,
            )

        match = re.fullmatch(r"sheet:(\d+)/(cell|range):(.+)", target)
        if match:
            ws = self._doc.Worksheets(int(match.group(1)))
            rng = ws.Range(match.group(3))
            return self._patch_range(
                rng, value, formula, font, fill, number_format,
                horizontal_alignment, vertical_alignment, wrap_text, indent,
                row_height, column_width, borders,
            )

        match = re.fullmatch(r"sheet:(\d+)/shape:(\d+)", target)
        if match:
            shape = self._doc.Worksheets(int(match.group(1))).Shapes(int(match.group(2)))
            results = [
                self._prefix_result(apply_geometry(shape, geometry), "geometry"),
                self._prefix_result(apply_fill(shape.Fill, fill), "fill"),
                self._prefix_result(apply_line(shape.Line, line), "line"),
            ]
            if name is not None:
                ok = safe_set(shape, "Name", name)
                results.append({"accepted": ["name"] if ok else [],
                                "rejected": [] if ok else ["name"]})
            return merge_results(*results)

        match = re.fullmatch(r"sheet:(\d+)/chart:(\d+)", target)
        if match:
            co = self._doc.Worksheets(int(match.group(1))).ChartObjects(int(match.group(2)))
            chart = co.Chart
            results = [self._prefix_result(apply_geometry(co, geometry), "geometry")]
            accepted, rejected = [], []
            if chart_type is not None:
                (accepted if safe_set(chart, "ChartType", chart_type) else rejected).append("chart_type")
            if chart_title is not None:
                try:
                    chart.HasTitle = True
                    chart.ChartTitle.Text = chart_title
                    ok = True
                except Exception:
                    ok = False
                (accepted if ok else rejected).append("chart_title")
            results.append({"accepted": accepted, "rejected": rejected})
            return merge_results(*results)

        match = re.fullmatch(r"sheet:(\d+)", target)
        if match:
            ws = self._doc.Worksheets(int(match.group(1)))
            accepted, rejected = [], []
            if name is not None:
                (accepted if safe_set(ws, "Name", name) else rejected).append("name")
            mapping = {
                "orientation": "Orientation", "top_margin": "TopMargin",
                "bottom_margin": "BottomMargin", "left_margin": "LeftMargin",
                "right_margin": "RightMargin", "header_margin": "HeaderMargin",
                "footer_margin": "FooterMargin", "paper_size": "PaperSize",
                "zoom": "Zoom", "fit_to_pages_wide": "FitToPagesWide",
                "fit_to_pages_tall": "FitToPagesTall", "print_area": "PrintArea",
                "print_title_rows": "PrintTitleRows", "print_title_columns": "PrintTitleColumns",
            }
            for key, item in (page_setup or {}).items():
                ok = key in mapping and safe_set(ws.PageSetup, mapping[key], item)
                (accepted if ok else rejected).append(f"page_setup.{key}")
            return {"accepted": accepted, "rejected": rejected}

        raise ValueError(f"Unsupported Sheet target: {target}")

    def _range_snapshot(self, rng, element_id, include_value=False):
        if rng is None:
            return {"id": element_id}
        interior = safe_get(rng, "Interior")
        merge_area = safe_get(rng, "MergeArea") if safe_get(rng, "MergeCells") else None
        result = {
            "id": element_id,
            "address": safe_get(rng, "Address"),
            "font": font_snapshot(safe_get(rng, "Font")),
            "fill": {"color": color_hex(safe_get(interior, "Color"))},
            "number_format": safe_get(rng, "NumberFormat"),
            "horizontal_alignment": safe_get(rng, "HorizontalAlignment"),
            "vertical_alignment": safe_get(rng, "VerticalAlignment"),
            "wrap_text": safe_get(rng, "WrapText"),
            "indent": safe_get(rng, "IndentLevel"),
            "row_height": safe_get(rng, "RowHeight"),
            "column_width": safe_get(rng, "ColumnWidth"),
            "merged": safe_get(rng, "MergeCells"),
            "merge_area": safe_get(merge_area, "Address"),
        }
        if include_value:
            result["value"] = safe_get(rng, "Value")
            result["formula"] = safe_get(rng, "Formula")
            result["display_text"] = safe_get(rng, "Text")
        return result

    def _patch_range(self, rng, value, formula, font, fill, number_format,
                     horizontal_alignment, vertical_alignment, wrap_text,
                     indent, row_height, column_width, borders):
        results = [self._prefix_result(apply_font(rng.Font, font), "font")]
        accepted, rejected = [], []
        if value is not None:
            (accepted if safe_set(rng, "Value", value) else rejected).append("value")
        if formula is not None:
            (accepted if safe_set(rng, "Formula", formula) else rejected).append("formula")
        if "color" in fill:
            ok = safe_set(rng.Interior, "Color", hex_to_rgb_long(fill["color"]))
            (accepted if ok else rejected).append("fill.color")
        scalar = {
            "number_format": ("NumberFormat", number_format),
            "horizontal_alignment": ("HorizontalAlignment", horizontal_alignment),
            "vertical_alignment": ("VerticalAlignment", vertical_alignment),
            "wrap_text": ("WrapText", wrap_text),
            "indent": ("IndentLevel", indent),
            "row_height": ("RowHeight", row_height),
            "column_width": ("ColumnWidth", column_width),
        }
        for key, (prop, item) in scalar.items():
            if item is not None:
                (accepted if safe_set(rng, prop, item) else rejected).append(key)
        if borders:
            for edge, spec in borders.items():
                try:
                    border = rng.Borders(int(edge))
                    if "style" in spec:
                        border.LineStyle = spec["style"]
                    if "weight" in spec:
                        border.Weight = spec["weight"]
                    if "color" in spec:
                        border.Color = hex_to_rgb_long(spec["color"])
                    accepted.append(f"borders.{edge}")
                except Exception:
                    rejected.append(f"borders.{edge}")
        results.append({"accepted": accepted, "rejected": rejected})
        return merge_results(*results)

    @staticmethod
    def _prefix_result(result, prefix):
        return {
            "accepted": [f"{prefix}.{key}" for key in result["accepted"]],
            "rejected": [f"{prefix}.{key}" for key in result["rejected"]],
        }

    def save_xlsx(self, path):
        p = _abs(path)
        self._doc.SaveAs(p, FMT_XLSX)
        return p

    def export_pdf(self, path):
        p = _abs(path)
        ws1 = self._doc.Worksheets(1) if hasattr(self._doc, "Worksheets") else self._doc.Sheets(1)
        ws1.ExportAsFixedFormat(FMT_PDF_FROM_XLS, p)
        return p


# ===========================================================================
#  SLIDE  (pptx)
# ===========================================================================
