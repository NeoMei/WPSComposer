"""Renderer-facing composers that record closed generation plans."""

from __future__ import annotations

from pathlib import Path

from .generation_plan import (
    GenerationOperation,
    GenerationPlan,
    GenerationResource,
    OperationPlanError,
    RecordedGeneration,
    validate_generation_plan,
)


_STYLE_KEYS = {
    "based_on": "basedOn",
    "font_name": "fontName",
    "font_name_ascii": "fontNameAscii",
    "font_size": "fontSize",
    "indent_first": "indentFirst",
    "left_indent": "leftIndent",
    "right_indent": "rightIndent",
    "line_spacing": "lineSpacing",
    "line_spacing_rule": "lineSpacingRule",
    "space_before": "spaceBefore",
    "space_after": "spaceAfter",
    "left_border": "leftBorder",
    "border_color": "borderColor",
    "keep_with_next": "keepWithNext",
    "outline_level": "outlineLevel",
}


def _without_none(**values):
    return {key: value for key, value in values.items() if value is not None}


def _normalize_style(key, props, *, outline_level=None):
    normalized = {"name": props.get("name", str(key))}
    for name, value in props.items():
        if name == "name" or value is None:
            continue
        normalized[_STYLE_KEYS.get(name, name)] = value
    if outline_level is not None:
        normalized["outlineLevel"] = int(outline_level)
    return normalized


class RecordingWriterComposer:
    """Record Writer renderer calls without opening WPS or touching files."""

    def __init__(self):
        self._operations = []
        self._resources = []
        self._resource_ids = {}

    def __enter__(self):
        self._record("writer.reset")
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False

    def _record(self, op, **args):
        self._operations.append(GenerationOperation(op, args))

    def set_margins(self, top, bottom, left, right):
        self._record(
            "writer.configure_page",
            marginTop=top,
            marginBottom=bottom,
            marginLeft=left,
            marginRight=right,
        )

    def ensure_styles(self, styles_dict):
        self._record(
            "writer.ensure_styles",
            styles=[
                _normalize_style(key, props)
                for key, props in styles_dict.items()
            ],
        )

    def ensure_heading_styles(self, styles_by_level):
        self._record(
            "writer.ensure_styles",
            styles=[
                _normalize_style(
                    f"Heading {int(level)}",
                    props,
                    outline_level=level,
                )
                for level, props in styles_by_level.items()
            ],
        )

    def apply_heading_text_color(self, color):
        self._record(
            "writer.ensure_styles",
            styles=[
                {"name": f"Heading {level}", "color": color, "outlineLevel": level}
                for level in range(1, 7)
            ],
        )

    def add_styled_paragraph(self, text, style_name):
        self._record("writer.add_paragraph", text=str(text), style=style_name)

    def add_rich_paragraph(self, spans, style_name):
        normalized_spans = [
            {
                "text": span.text,
                "bold": bool(span.bold),
                "italic": bool(span.italic),
                "code": bool(span.code),
                "strikethrough": bool(getattr(span, "strikethrough", False)),
                "link": span.link,
                "linkTitle": getattr(span, "link_title", None),
            }
            for span in spans
        ]
        self._record(
            "writer.add_paragraph",
            text="".join(span["text"] for span in normalized_spans),
            style=style_name,
            spans=normalized_spans,
        )

    def add_code_lines(self, lines):
        for line in lines:
            self.add_styled_paragraph(line if line else " ", "Source Code")
        self.add_paragraph("", size=4)

    def add_paragraph(
        self,
        text,
        size=None,
        bold=None,
        italic=None,
        color=None,
        align=None,
        indent_first=None,
        line_spacing=None,
        space_after=None,
        font_name=None,
        line_spacing_rule=None,
        font_name_ascii=None,
        space_before=None,
        style=None,
    ):
        self._record(
            "writer.add_paragraph",
            **_without_none(
                text=str(text),
                style=style,
                size=size,
                bold=bold,
                italic=italic,
                color=color,
                align=align,
                indentFirst=indent_first,
                lineSpacing=line_spacing,
                lineSpacingRule=line_spacing_rule,
                spaceBefore=space_before,
                spaceAfter=space_after,
                fontName=font_name,
                fontNameAscii=font_name_ascii,
            ),
        )

    def add_heading_level(
        self,
        text,
        level=1,
        size=None,
        color=None,
        line_spacing=None,
        space_after=None,
        line_spacing_rule=None,
    ):
        self._record(
            "writer.add_heading",
            **_without_none(
                text=str(text),
                level=min(max(int(level), 1), 6),
                size=size,
                color=color,
                lineSpacing=line_spacing,
                lineSpacingRule=line_spacing_rule,
                spaceAfter=space_after,
            ),
        )

    def add_bullet_list(self, items, glyph="\u2022", indent=24):
        self._record(
            "writer.add_list",
            items=[str(item) for item in items],
            ordered=False,
            glyph=glyph,
            indent=indent,
        )

    def add_numbered_list(self, items, indent=24):
        self._record(
            "writer.add_list",
            items=[str(item) for item in items],
            ordered=True,
            indent=indent,
        )

    def add_table(
        self,
        rows,
        cols,
        data,
        shade_header="#4472C4",
        header_color="#FFFFFF",
        font_size=10,
        col_widths=None,
        alignments=None,
        banded_rows=True,
        auto_fit=True,
        repeat_header=True,
        border_color="#D0D0D0",
    ):
        self._record(
            "writer.add_table",
            **_without_none(
                rows=int(rows),
                cols=int(cols),
                data=[list(row) for row in data],
                shadeHeader=shade_header,
                headerColor=header_color,
                fontSize=font_size,
                columnWidths=list(col_widths) if col_widths is not None else None,
                alignments=list(alignments) if alignments is not None else None,
                bandedRows=bool(banded_rows),
                autoFit=bool(auto_fit),
                repeatHeader=bool(repeat_header),
                borderColor=border_color,
            ),
        )

    def _image_id(self, path):
        source_path = Path(path).expanduser().resolve()
        if source_path in self._resource_ids:
            return self._resource_ids[source_path]
        media_types = {
            ".bmp": "image/bmp",
            ".gif": "image/gif",
            ".jpeg": "image/jpeg",
            ".jpg": "image/jpeg",
            ".png": "image/png",
            ".tif": "image/tiff",
            ".tiff": "image/tiff",
        }
        try:
            media_type = media_types[source_path.suffix.lower()]
        except KeyError as error:
            raise OperationPlanError(
                f"unsupported image resource: {source_path.suffix or '<none>'}"
            ) from error
        resource_id = f"image-{len(self._resources) + 1}"
        self._resources.append(
            GenerationResource(resource_id, source_path, media_type)
        )
        self._resource_ids[source_path] = resource_id
        return resource_id

    def add_image(
        self,
        path,
        width=None,
        height=None,
        wrap=0,
        *,
        max_width=None,
        max_height=None,
        inline=True,
        preserve_aspect=True,
        alt=None,
    ):
        self._record(
            "writer.add_image",
            **_without_none(
                imageId=self._image_id(path),
                width=width,
                height=height,
                maxWidth=max_width,
                maxHeight=max_height,
                wrap=wrap,
                inline=bool(inline),
                preserveAspect=bool(preserve_aspect),
                alt=alt,
            ),
        )

    def add_image_block(self, *args, **kwargs):
        self.add_image(*args, **kwargs)

    def add_page_break(self):
        self._record("writer.add_page_break")

    def add_section(self):
        self._record("writer.add_section")

    def add_horizontal_line(self):
        self._record("writer.add_horizontal_line")

    def add_paragraph_horizontal_line(self):
        self._record("writer.add_horizontal_line")

    def insert_toc(self, title="Table of Contents"):
        self._record("writer.insert_toc", title=title)

    def set_page_number_in_footer(self):
        self._record("writer.set_page_number")

    def update_fields(self):
        self._record("writer.update_fields")

    def save_docx(self, path):
        plan = GenerationPlan("writer", tuple(self._operations))
        validated = validate_generation_plan(plan.to_dict(), "writer")
        return RecordedGeneration(validated, tuple(self._resources))


class RecordingSheetComposer:
    """Record Sheet renderer calls without opening WPS or touching files."""

    def __init__(self):
        self._operations = []

    def __enter__(self):
        self._record("sheet.reset")
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False

    def _record(self, op, **args):
        self._operations.append(GenerationOperation(op, args))

    def rename_sheet(self, index, name):
        self._record("sheet.rename", index=int(index), name=str(name))

    def add_sheet(self, name=None):
        self._record("sheet.add", name=str(name))

    def select_sheet(self, index):
        self._record("sheet.select", index=int(index))

    def write_table(
        self,
        start_row,
        start_col,
        data,
        header_bold=True,
        header_shade="#4472C4",
        header_font_color="#FFFFFF",
        font_size=11,
    ):
        if header_shade is True:
            header_shade = "#4472C4"
        self._record(
            "sheet.write_table",
            startRow=int(start_row),
            startCol=int(start_col),
            values=[list(row) for row in data],
            headerBold=bool(header_bold),
            headerShade=header_shade,
            headerFontColor=header_font_color,
            fontSize=font_size,
        )

    def set_column_width(self, column, width):
        self._record("sheet.set_column_width", column=str(column), width=width)

    def autofit(self):
        self._record("sheet.autofit")

    def save_xlsx(self, path):
        plan = GenerationPlan("spreadsheet", tuple(self._operations))
        validated = validate_generation_plan(plan.to_dict(), "spreadsheet")
        return RecordedGeneration(validated, ())
