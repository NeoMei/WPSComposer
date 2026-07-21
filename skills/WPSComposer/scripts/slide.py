"""SlideComposer - WPS Presentation (KWpp.Application) to pptx.

Presentation slides with title/text layouts, shapes, images,
tables, and design preset integration.
"""

from __future__ import annotations

import re

from ._dispatch import (
    _abs, FMT_PPTX, FMT_PDF_FROM_PPT, FMT_PPT, FMT_PPSX, FMT_PPTM,
    FMT_POTX, FMT_POTM, FMT_PPSM, FMT_ODP,
)
from ._colors import hex_to_rgb_long
from ._base import BaseComposer
from .design_presets import DesignPreset
from .layout_templates import LayoutTemplate
from .formatting import (
    apply_fill,
    apply_font,
    apply_geometry,
    apply_line,
    apply_paragraph,
    fill_snapshot,
    font_snapshot,
    geometry_snapshot,
    line_snapshot,
    merge_results,
    page_setup_snapshot,
    paragraph_snapshot,
    safe_get,
    safe_set,
)


# ===========================================================================

class SlideComposer(BaseComposer):
    _progids = ("KWpp.Application", "Wpp.Application", "PowerPoint.Application")
    _doc_type = "impress"
    _native_fmt = FMT_PPTX
    _pdf_fmt = FMT_PDF_FROM_PPT
    _formats_by_extension = {
        ".pptx": FMT_PPTX, ".ppt": FMT_PPT, ".ppsx": FMT_PPSX,
        ".pptm": FMT_PPTM, ".potx": FMT_POTX, ".potm": FMT_POTM,
        ".ppsm": FMT_PPSM, ".odp": FMT_ODP,
    }

    MSO_TEXTBOX = 17
    MSO_RECTANGLE = 1
    MSO_ROUNDED_RECTANGLE = 5
    MSO_OVAL = 9
    MSO_RIGHT_ARROW = 13
    MSO_LEFT_RIGHT_ARROW = 37

    @staticmethod
    def _create_doc(app):
        return app.Presentations.Add()

    @staticmethod
    def _open_document(app, path, read_only=False):
        # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
        return app.Presentations.Open(path, bool(read_only), False, True)

    @staticmethod
    def _active_document(app):
        return app.ActivePresentation

    LAYOUT_TITLE = 1
    LAYOUT_TITLE_CONTENT = 1
    LAYOUT_BLANK = 12
    LAYOUT_SECTION = 2
    LAYOUT_TWO_CONTENT = 3

    # shape types
    MSO_TEXTBOX = 17
    MSO_RECTANGLE = 1
    MSO_ROUNDED_RECTANGLE = 5
    MSO_OVAL = 9
    MSO_RIGHT_ARROW = 13
    MSO_LEFT_RIGHT_ARROW = 37

    @property
    def pres(self):
        return self._doc

    @property
    def slide_count(self):
        """Return the number of slides owned by this presentation."""
        return int(self._doc.Slides.Count)

    def set_slide_size(self, width_pt=960, height_pt=540):
        try:
            ps = self._doc.PageSetup
            ps.SlideWidth = width_pt
            ps.SlideHeight = height_pt
        except Exception:
            pass

    def _new_slide(self, layout=12):
        idx = self._doc.Slides.Count + 1
        return self._doc.Slides.Add(idx, layout), idx

    def add_title_slide(self, title, subtitle="", title_size=40, sub_size=20,
                        title_color=None):
        slide, idx = self._new_slide(self.LAYOUT_TITLE)
        try:
            tr = slide.Shapes.Title.TextFrame.TextRange
            tr.Text = title
            tr.Font.Size = title_size
            if title_color:
                tr.Font.Color.RGB = hex_to_rgb_long(title_color)
        except Exception:
            pass
        if subtitle:
            try:
                tr2 = slide.Shapes.Placeholders(2).TextFrame.TextRange
                tr2.Text = subtitle
                tr2.Font.Size = sub_size
            except Exception:
                pass
        return idx

    def add_section_slide(self, title):
        slide, idx = self._new_slide(self.LAYOUT_SECTION)
        try:
            slide.Shapes.Title.TextFrame.TextRange.Text = title
        except Exception:
            pass
        return idx

    def add_text_slide(self, title, body, title_size=32, body_size=18,
                       title_color=None, body_color=None,
                       bullets=True):
        slide, idx = self._new_slide(self.LAYOUT_TITLE_CONTENT)
        try:
            tr = slide.Shapes.Title.TextFrame.TextRange
            tr.Text = title
            tr.Font.Size = title_size
            if title_color:
                tr.Font.Color.RGB = hex_to_rgb_long(title_color)
        except Exception:
            pass
        try:
            tr2 = slide.Shapes.Placeholders(2).TextFrame.TextRange
            if bullets and isinstance(body, list):
                tr2.Text = "\r".join(body)
            else:
                tr2.Text = body
            tr2.Font.Size = body_size
            if body_color:
                tr2.Font.Color.RGB = hex_to_rgb_long(body_color)
        except Exception:
            pass
        return idx

    def add_bullets_slide(self, title, items, title_size=32, body_size=18):
        return self.add_text_slide(title, items, title_size, body_size, bullets=True)

    def add_blank_slide(self):
        slide, idx = self._new_slide(self.LAYOUT_BLANK)
        return slide, idx

    def set_background_color(self, slide_index, color):
        slide = self._doc.Slides(slide_index)
        try:
            slide.FollowMasterBackground = False
            fill = slide.Background.Fill
            fill.ForeColor.RGB = hex_to_rgb_long(color)
            fill.Visible = 1
            fill.Solid()
        except Exception:
            pass

    def add_textbox(self, slide_index, text, left, top, width, height,
                    size=18, bold=False, color=None, align=1,
                    fill_color=None, shape_type=None):
        """align: 1=left,2=center,3=right. shape_type None=plain textbox."""
        slide = self._doc.Slides(slide_index)
        st = shape_type if shape_type else self.MSO_TEXTBOX
        shape = slide.Shapes.AddShape(st, left, top, width, height) if shape_type else \
                slide.Shapes.AddTextbox(1, left, top, width, height)
        tr = shape.TextFrame.TextRange
        tr.Text = text
        tr.Font.Size = size
        tr.Font.Bold = bold
        if color:
            tr.Font.Color.RGB = hex_to_rgb_long(color)
        try:
            tr.ParagraphFormat.Alignment = align
        except Exception:
            pass
        if fill_color:
            try:
                shape.Fill.ForeColor.RGB = hex_to_rgb_long(fill_color)
                shape.Fill.Visible = 1
            except Exception:
                pass
        else:
            try:
                shape.Fill.Visible = False
                shape.Line.Visible = False
            except Exception:
                pass
        return shape

    def add_shape(self, slide_index, shape_type, left, top, width, height,
                  fill_color=None, line_color=None, text="", size=14):
        slide = self._doc.Slides(slide_index)
        shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
        if fill_color:
            try:
                shape.Fill.ForeColor.RGB = hex_to_rgb_long(fill_color)
                shape.Fill.Visible = 1
            except Exception:
                pass
        else:
            try:
                shape.Fill.Visible = False
            except Exception:
                pass
        if line_color:
            try:
                shape.Line.ForeColor.RGB = hex_to_rgb_long(line_color)
                shape.Line.Visible = 1
            except Exception:
                pass
        else:
            try:
                shape.Line.Visible = False
            except Exception:
                pass
        if text:
            tr = shape.TextFrame.TextRange
            tr.Text = text
            tr.Font.Size = size
        return shape

    def add_image(self, slide_index, path, left, top, width=None, height=None):
        slide = self._doc.Slides(slide_index)
        p = _abs(path)
        if width and height:
            return slide.Shapes.AddPicture(p, False, True, left, top, width, height)
        return slide.Shapes.AddPicture(p, False, True, left, top)

    def add_table(self, slide_index, rows, cols, left, top, width, height,
                  data, header_shade="#4472C4", header_font="#FFFFFF",
                  font_size=11):
        slide = self._doc.Slides(slide_index)
        shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
        tbl = shape.Table
        for r in range(rows):
            for c in range(cols):
                if r < len(data) and c < len(data[r]):
                    cell = tbl.Cell(r + 1, c + 1)
                    cell.Shape.TextFrame.TextRange.Text = str(data[r][c])
                    cell.Shape.TextFrame.TextRange.Font.Size = font_size
                    if r == 0:
                        cell.Shape.TextFrame.TextRange.Font.Bold = True
                        cell.Shape.TextFrame.TextRange.Font.Color.RGB = hex_to_rgb_long(header_font)
                        try:
                            cell.Shape.Fill.ForeColor.RGB = hex_to_rgb_long(header_shade)
                        except Exception:
                            pass
        return shape

    def set_notes(self, slide_index, text):
        slide = self._doc.Slides(slide_index)
        try:
            slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text
        except Exception:
            pass

    def save_pptx(self, path):
        p = _abs(path)
        self._doc.SaveAs(p, FMT_PPTX)
        return p

    def export_pdf(self, path):
        p = _abs(path)
        self._doc.SaveAs(p, FMT_PDF_FROM_PPT)
        return p

    def apply_design_preset(self, preset):
        """Apply a DesignPreset to the entire presentation.

        Sets the slide master background colour and updates the
        default text style to match the preset.

        Args:
            preset: design_presets.DesignPreset instance.
        """
        from .design_presets import DesignPreset
        from ._colors import resolve_color
        if not isinstance(preset, DesignPreset):
            raise TypeError("preset must be a DesignPreset instance")

        # Slide master background
        try:
            master = self._doc.SlideMaster
            master.Background.Fill.ForeColor.RGB = (
                hex_to_rgb_long(resolve_color("bg", preset))
            )
        except Exception:
            pass

        # Default text style on first slide (best-effort)
        try:
            slide = self._doc.Slides(1)
            for shape in slide.Shapes:
                try:
                    if shape.HasTextFrame:
                        tf = shape.TextFrame.TextRange
                        if tf.Font.Size > 30:
                            title_font = preset.get_font("title")
                            tf.Font.Name = title_font[0]
                            tf.Font.Color.RGB = hex_to_rgb_long(title_font[2])
                except Exception:
                    pass
        except Exception:
            pass

    def apply_layout_template(self, layout, preset=None):
        """Apply a LayoutTemplate to the current slide.

        Renders bg, line, box, and text elements from the template
        descriptor onto the active slide.  Uses *preset* to resolve
        semantic colour names when provided.

        Args:
            layout: layout_templates.LayoutTemplate instance.
            preset: Optional DesignPreset for colour-name resolution.
        """
        from .layout_templates import LayoutTemplate, resolve_layout_colors
        if not isinstance(layout, LayoutTemplate):
            raise TypeError("layout must be a LayoutTemplate instance")

        # Resolve colours if preset given
        if preset is not None:
            try:
                layout = resolve_layout_colors(layout, preset)
            except Exception:
                pass

        # Get current slide index
        try:
            slide_idx = (
                self._app.ActiveWindow.Selection.SlideRange.SlideIndex
            )
        except Exception:
            slide_idx = 1
        slide = self._doc.Slides(slide_idx)

        for elem in layout.elements:
            etype = elem.get("type")
            if etype == "bg":
                try:
                    color = elem.get("color", "#FFFFFF")
                    slide.Background.Fill.ForeColor.RGB = (
                        hex_to_rgb_long(color)
                    )
                except Exception:
                    pass

            elif etype == "line":
                try:
                    left = elem.get("x", 0)
                    top = elem.get("y", 0)
                    w = elem.get("w", 100)
                    h = elem.get("h", 3)
                    color = elem.get("color", "#000000")
                    shape = slide.Shapes.AddLine(
                        left, top, left + w, top + h
                    )
                    shape.Line.ForeColor.RGB = hex_to_rgb_long(color)
                    shape.Line.Weight = max(h, 1)
                except Exception:
                    pass

            elif etype == "box":
                try:
                    left = elem.get("x", 0)
                    top = elem.get("y", 0)
                    w = elem.get("w", 100)
                    h = elem.get("h", 100)
                    color = elem.get("color", "#EEEEEE")
                    text = elem.get("text", "")
                    shape = slide.Shapes.AddShape(
                        1, left, top, w, h  # msoShapeRectangle
                    )
                    shape.Fill.ForeColor.RGB = hex_to_rgb_long(color)
                    shape.Line.Visible = 0
                    if text:
                        shape.TextFrame.TextRange.Text = text
                        shape.TextFrame.TextRange.Font.Size = 12
                except Exception:
                    pass

            elif etype == "text":
                try:
                    text = elem.get("text", "")
                    if not text:
                        continue
                    fs = elem.get("fs", 18)
                    color = elem.get("color", "#000000")
                    bold = elem.get("bold", False)
                    left = elem.get("x", 0)
                    top = elem.get("y", 0)
                    w = elem.get("w", 300)
                    h = elem.get("h", 50)
                    align = elem.get("align", 0)
                    shape = slide.Shapes.AddTextbox(
                        1, left, top, w, h
                    )
                    shape.TextFrame.TextRange.Text = text
                    shape.TextFrame.TextRange.Font.Size = fs
                    shape.TextFrame.TextRange.Font.Bold = bold
                    if align:
                        try:
                            pf = shape.TextFrame.TextRange.ParagraphFormat
                            pf.Alignment = align
                        except Exception:
                            pass
                    try:
                        shape.TextFrame.TextRange.Font.Color.RGB = (
                            hex_to_rgb_long(color)
                        )
                    except Exception:
                        pass
                except Exception:
                    pass

    # ================================================================
    # Existing-presentation inspection and element-level editing
    # ================================================================

    def inspect_selection(self):
        """Describe the shapes currently selected in the WPS presentation."""
        selection = safe_get(safe_get(self._app, "ActiveWindow"), "Selection")
        text_range = safe_get(selection, "TextRange")
        if text_range is not None:
            return {
                "kind": "slide_text_selection",
                "text": safe_get(text_range, "Text", ""),
                "font": font_snapshot(safe_get(text_range, "Font")),
                "paragraph": paragraph_snapshot(safe_get(text_range, "ParagraphFormat")),
            }
        shape_range = safe_get(selection, "ShapeRange")
        count = int(safe_get(shape_range, "Count", 0) or 0)
        shapes = []
        for index in range(1, count + 1):
            shape = shape_range(index)
            shapes.append(self._shape_snapshot(shape, None, index, True, "selection"))
        return {"kind": "slide_selection", "count": count, "shapes": shapes}

    def inspect_document(self, include_text=True, max_shapes=None):
        """Return slides, shapes, text paragraphs/runs, tables, and formats."""
        pres = self._doc
        slide_count = int(safe_get(pres.Slides, "Count", 0) or 0)
        slides = []
        remaining = max_shapes if max_shapes is not None else float("inf")
        truncated = False

        for slide_index in range(1, slide_count + 1):
            slide = pres.Slides(slide_index)
            shape_count = int(safe_get(slide.Shapes, "Count", 0) or 0)
            shapes = []
            for shape_index in range(1, shape_count + 1):
                if remaining <= 0:
                    truncated = True
                    break
                shapes.append(self._shape_snapshot(
                    slide.Shapes(shape_index), slide_index, shape_index,
                    include_text, f"slide:{slide_index}",
                ))
                remaining -= 1

            notes = None
            try:
                notes = slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
            except Exception:
                pass
            background = safe_get(safe_get(slide, "Background"), "Fill")
            slides.append({
                "id": f"slide:{slide_index}",
                "index": slide_index,
                "name": safe_get(slide, "Name"),
                "layout": safe_get(slide, "Layout"),
                "follow_master_background": safe_get(slide, "FollowMasterBackground"),
                "background": fill_snapshot(background),
                "shape_count": shape_count,
                "shapes": shapes,
                "notes": notes if include_text else None,
            })

        return {
            "kind": "slide",
            "name": safe_get(pres, "Name"),
            "path": safe_get(pres, "FullName"),
            "saved": safe_get(pres, "Saved"),
            "page_setup": page_setup_snapshot(pres.PageSetup),
            "slide_count": slide_count,
            "shapes_truncated": truncated,
            "slides": slides,
        }

    def apply_format_patch(self, target, *, text=None, font=None,
                           paragraph=None, geometry=None, fill=None, line=None,
                           text_frame=None, name=None, shape_type=None,
                           background=None, follow_master_background=None,
                           page_setup=None, vertical_alignment=None):
        """Apply a partial patch to a slide, shape, paragraph, run, or cell.

        Targets include ``presentation``, ``slide:N``, ``slide:N/shape:N``,
        ``slide:N/shape:N/paragraph:N``, ``.../run:N``, and
        ``slide:N/shape:N/table/cell:R,C``.
        """
        font, paragraph = font or {}, paragraph or {}
        geometry, fill, line = geometry or {}, fill or {}, line or {}
        text_frame, background = text_frame or {}, background or {}

        if target == "selection":
            selection = safe_get(safe_get(self._app, "ActiveWindow"), "Selection")
            selected_text = safe_get(selection, "TextRange")
            if selected_text is not None:
                return self._patch_text_range(selected_text, text, font, paragraph)
            shape_range = safe_get(selection, "ShapeRange")
            count = int(safe_get(shape_range, "Count", 0) or 0)
            if count == 0:
                raise ValueError("The active presentation selection has no text or shapes")
            results = []
            for index in range(1, count + 1):
                results.append(self._patch_text_shape(
                    shape_range(index), text, font, paragraph, geometry, fill,
                    line, dict(text_frame), vertical_alignment,
                ))
            return merge_results(*results)

        if target == "presentation":
            accepted, rejected = [], []
            mapping = {"slide_width": "SlideWidth", "slide_height": "SlideHeight"}
            for key, value in (page_setup or {}).items():
                ok = key in mapping and safe_set(self._doc.PageSetup, mapping[key], value)
                (accepted if ok else rejected).append(f"page_setup.{key}")
            return {"accepted": accepted, "rejected": rejected}

        match = re.fullmatch(r"slide:(\d+)", target)
        if match:
            slide = self._doc.Slides(int(match.group(1)))
            accepted, rejected = [], []
            if name is not None:
                (accepted if safe_set(slide, "Name", name) else rejected).append("name")
            if follow_master_background is not None:
                ok = safe_set(slide, "FollowMasterBackground", follow_master_background)
                (accepted if ok else rejected).append("follow_master_background")
            if background:
                result = self._prefix_result(apply_fill(slide.Background.Fill, background), "background")
                accepted.extend(result["accepted"])
                rejected.extend(result["rejected"])
            return {"accepted": accepted, "rejected": rejected}

        match = re.fullmatch(r"slide:(\d+)/shape:(\d+)/table/cell:(\d+),(\d+)", target)
        if match:
            shape = self._doc.Slides(int(match.group(1))).Shapes(int(match.group(2)))
            cell_shape = shape.Table.Cell(int(match.group(3)), int(match.group(4))).Shape
            return self._patch_text_shape(
                cell_shape, text, font, paragraph, {}, fill, line, text_frame,
                vertical_alignment,
            )

        match = re.fullmatch(
            r"slide:(\d+)/shape:(\d+)/paragraph:(\d+)(?:/run:(\d+))?", target
        )
        if match:
            shape = self._doc.Slides(int(match.group(1))).Shapes(int(match.group(2)))
            text_range = shape.TextFrame.TextRange
            para = text_range.Paragraphs(int(match.group(3)), 1)
            rng = para.Runs(int(match.group(4)), 1) if match.group(4) else para
            return self._patch_text_range(rng, text, font, paragraph)

        match = re.fullmatch(r"slide:(\d+)/shape:(\d+)", target)
        if match:
            shape = self._doc.Slides(int(match.group(1))).Shapes(int(match.group(2)))
            result = self._patch_text_shape(
                shape, text, font, paragraph, geometry, fill, line, text_frame,
                vertical_alignment,
            )
            if name is not None:
                ok = safe_set(shape, "Name", name)
                result["accepted" if ok else "rejected"].append("name")
            return result

        raise ValueError(f"Unsupported Slide target: {target}")

    def _shape_snapshot(self, shape, slide_index, shape_index, include_text, prefix):
        element_id = f"{prefix}/shape:{shape_index}"
        result = {
            "id": element_id,
            "index": shape_index,
            "name": safe_get(shape, "Name"),
            "type": safe_get(shape, "Type"),
            "geometry": geometry_snapshot(shape),
            "fill": fill_snapshot(safe_get(shape, "Fill")),
            "line": line_snapshot(safe_get(shape, "Line")),
            "placeholder_type": safe_get(safe_get(shape, "PlaceholderFormat"), "Type"),
        }
        text_frame = safe_get(shape, "TextFrame")
        text_range = safe_get(text_frame, "TextRange")
        if text_range is not None:
            result["text_frame"] = {
                "margin_left": safe_get(text_frame, "MarginLeft"),
                "margin_right": safe_get(text_frame, "MarginRight"),
                "margin_top": safe_get(text_frame, "MarginTop"),
                "margin_bottom": safe_get(text_frame, "MarginBottom"),
                "word_wrap": safe_get(text_frame, "WordWrap"),
                "auto_size": safe_get(text_frame, "AutoSize"),
                "vertical_anchor": safe_get(text_frame, "VerticalAnchor"),
            }
            if include_text:
                result["text"] = safe_get(text_range, "Text", "")
            result["font"] = font_snapshot(safe_get(text_range, "Font"))
            result["paragraph"] = paragraph_snapshot(safe_get(text_range, "ParagraphFormat"))
            result["paragraphs"] = self._paragraphs_snapshot(text_range, element_id, include_text)

        table = safe_get(shape, "Table") if safe_get(shape, "HasTable") else None
        if table is not None:
            cells = []
            rows = int(safe_get(table.Rows, "Count", 0) or 0)
            cols = int(safe_get(table.Columns, "Count", 0) or 0)
            for row in range(1, rows + 1):
                for col in range(1, cols + 1):
                    cell_shape = table.Cell(row, col).Shape
                    cell_text = safe_get(safe_get(cell_shape, "TextFrame"), "TextRange")
                    cell_snap = {
                        "id": f"{element_id}/table/cell:{row},{col}",
                        "row": row,
                        "column": col,
                        "fill": fill_snapshot(safe_get(cell_shape, "Fill")),
                    }
                    if cell_text is not None:
                        if include_text:
                            cell_snap["text"] = safe_get(cell_text, "Text", "")
                        cell_snap["font"] = font_snapshot(safe_get(cell_text, "Font"))
                        cell_snap["paragraph"] = paragraph_snapshot(
                            safe_get(cell_text, "ParagraphFormat")
                        )
                    cells.append(cell_snap)
            result["table"] = {"rows": rows, "columns": cols, "cells": cells}
        return result

    def _paragraphs_snapshot(self, text_range, element_id, include_text):
        try:
            paragraph_range = text_range.Paragraphs()
            count = int(safe_get(paragraph_range, "Count", 0) or 0)
        except Exception:
            count = 0
        paragraphs = []
        for index in range(1, count + 1):
            para = text_range.Paragraphs(index, 1)
            snap = {
                "id": f"{element_id}/paragraph:{index}",
                "index": index,
                "font": font_snapshot(safe_get(para, "Font")),
                "paragraph": paragraph_snapshot(safe_get(para, "ParagraphFormat")),
                "runs": [],
            }
            if include_text:
                snap["text"] = safe_get(para, "Text", "")
            try:
                run_range = para.Runs()
                run_count = int(safe_get(run_range, "Count", 0) or 0)
            except Exception:
                run_count = 0
            for run_index in range(1, run_count + 1):
                run = para.Runs(run_index, 1)
                run_snap = {
                    "id": f"{element_id}/paragraph:{index}/run:{run_index}",
                    "index": run_index,
                    "font": font_snapshot(safe_get(run, "Font")),
                }
                if include_text:
                    run_snap["text"] = safe_get(run, "Text", "")
                snap["runs"].append(run_snap)
            paragraphs.append(snap)
        return paragraphs

    def _patch_text_shape(self, shape, text, font, paragraph, geometry, fill,
                          line, text_frame_patch, vertical_alignment):
        results = [
            self._prefix_result(apply_geometry(shape, geometry), "geometry"),
            self._prefix_result(apply_fill(safe_get(shape, "Fill"), fill), "fill"),
            self._prefix_result(apply_line(safe_get(shape, "Line"), line), "line"),
        ]
        text_frame = safe_get(shape, "TextFrame")
        text_range = safe_get(text_frame, "TextRange")
        if text_range is not None:
            results.append(self._patch_text_range(text_range, text, font, paragraph))
            accepted, rejected = [], []
            mapping = {
                "margin_left": "MarginLeft", "margin_right": "MarginRight",
                "margin_top": "MarginTop", "margin_bottom": "MarginBottom",
                "word_wrap": "WordWrap", "auto_size": "AutoSize",
                "vertical_anchor": "VerticalAnchor",
            }
            if vertical_alignment is not None:
                text_frame_patch["vertical_anchor"] = vertical_alignment
            for key, value in text_frame_patch.items():
                ok = key in mapping and safe_set(text_frame, mapping[key], value)
                (accepted if ok else rejected).append(f"text_frame.{key}")
            results.append({"accepted": accepted, "rejected": rejected})
        elif text is not None or font or paragraph or text_frame_patch:
            results.append({"accepted": [], "rejected": ["text_frame"]})
        return merge_results(*results)

    def _patch_text_range(self, text_range, text, font, paragraph):
        results = [
            self._prefix_result(apply_font(text_range.Font, font), "font"),
            self._prefix_result(apply_paragraph(text_range.ParagraphFormat, paragraph), "paragraph"),
        ]
        if text is not None:
            ok = safe_set(text_range, "Text", text)
            results.append({"accepted": ["text"] if ok else [],
                            "rejected": [] if ok else ["text"]})
        return merge_results(*results)

    @staticmethod
    def _prefix_result(result, prefix):
        return {
            "accepted": [f"{prefix}.{key}" for key in result["accepted"]],
            "rejected": [f"{prefix}.{key}" for key in result["rejected"]],
        }

# ===========================================================================
#  PDF  (edit existing PDFs — pure Python, no WPS dependency)
# ===========================================================================
# WPS exposes KPDF.Application but its Dispatch blocks headless (GUI RPC),
# so PDF *editing* uses pypdf/pdfplumber instead. PDF *generation* still goes
# through the three WPS composers above (export_pdf).
