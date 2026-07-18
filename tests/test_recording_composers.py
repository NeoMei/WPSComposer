import importlib.util

from pathlib import Path

from skills.WPSComposer.scripts import recording_composers
from skills.WPSComposer.scripts.document_model import Span
from skills.WPSComposer.scripts.generation_plan import (
    ALLOWED_OPERATIONS,
    validate_generation_plan,
)
from skills.WPSComposer.scripts.document_model import (
    Section,
    StructuredDocument,
    TableBlock,
)
from skills.WPSComposer.scripts.renderers import sheet_renderer


def test_recording_composers_module_exists():
    assert importlib.util.find_spec(
        "skills.WPSComposer.scripts.recording_composers"
    ) is not None


def test_recording_writer_composer_is_publicly_defined():
    assert hasattr(recording_composers, "RecordingWriterComposer")


def test_recording_writer_composer_records_every_closed_writer_operation(tmp_path):
    composer = recording_composers.RecordingWriterComposer()
    image_path = tmp_path / "chart.png"

    with composer as writer:
        writer.set_margins(72, 73, 90, 91)
        writer.ensure_styles(
            {
                "BodyText": {
                    "name": "Body Text",
                    "type": "paragraph",
                    "font_name": "FangSong",
                    "font_size": 12,
                    "indent_first": 24,
                }
            }
        )
        writer.add_paragraph("Summary", style="Body Text")
        writer.add_heading_level("Results", level=2)
        writer.add_bullet_list(["First", "Second"], glyph="-", indent=20)
        writer.add_table(2, 2, [["Item", "Amount"], ["A", 10]])
        writer.add_image(image_path, width=240, inline=True, alt="Chart")
        writer.add_page_break()
        writer.add_section()
        writer.add_horizontal_line()
        writer.insert_toc("Contents")
        writer.set_page_number_in_footer()
        writer.update_fields()
        recorded = writer.save_docx("ignored.docx")

    names = [operation.op for operation in recorded.plan.operations]
    assert names == [
        "writer.reset",
        "writer.configure_page",
        "writer.ensure_styles",
        "writer.add_paragraph",
        "writer.add_heading",
        "writer.add_list",
        "writer.add_table",
        "writer.add_image",
        "writer.add_page_break",
        "writer.add_section",
        "writer.add_horizontal_line",
        "writer.insert_toc",
        "writer.set_page_number",
        "writer.update_fields",
    ]
    assert set(names) == ALLOWED_OPERATIONS["writer"]
    assert validate_generation_plan(recorded.plan.to_dict(), "writer") == recorded.plan
    assert recorded.resources[0].id == "image-1"
    assert recorded.resources[0].source_path == image_path.resolve()
    assert recorded.resources[0].media_type == "image/png"
    assert "ignored.docx" not in repr(recorded)
    assert str(image_path.resolve()) not in repr(recorded.plan.to_dict())


def test_recording_writer_normalizes_styles_spans_and_reuses_logical_image_ids(
    tmp_path,
):
    composer = recording_composers.RecordingWriterComposer()
    image_path = tmp_path / "figure.jpeg"

    with composer as writer:
        writer.ensure_heading_styles(
            {1: {"font_name": "SimHei", "font_size": 16, "bold": True}}
        )
        writer.apply_heading_text_color("#000000")
        writer.add_rich_paragraph(
            [
                Span("Bold", bold=True),
                Span(" code", code=True, link_title="sample"),
            ],
            "First Paragraph",
        )
        writer.add_code_lines(["print('ok')", " "])
        writer.add_image(image_path)
        writer.add_image(Path(image_path))
        recorded = writer.save_docx("unused.docx")

    operations = [operation.to_dict() for operation in recorded.plan.operations]
    heading_styles = operations[1]["args"]["styles"]
    assert heading_styles == [
        {
            "name": "Heading 1",
            "fontName": "SimHei",
            "fontSize": 16,
            "bold": True,
            "outlineLevel": 1,
        }
    ]
    rich = next(
        operation
        for operation in operations
        if operation["op"] == "writer.add_paragraph"
        and operation["args"].get("spans")
    )
    assert rich["args"] == {
        "text": "Bold code",
        "style": "First Paragraph",
        "spans": [
            {
                "text": "Bold",
                "bold": True,
                "italic": False,
                "code": False,
                "strikethrough": False,
                "link": None,
                "linkTitle": None,
            },
            {
                "text": " code",
                "bold": False,
                "italic": False,
                "code": True,
                "strikethrough": False,
                "link": None,
                "linkTitle": "sample",
            },
        ],
    }
    code_lines = [
        operation["args"]
        for operation in operations
        if operation["op"] == "writer.add_paragraph"
        and operation["args"].get("style") == "Source Code"
    ]
    assert code_lines == [
        {"text": "print('ok')", "style": "Source Code"},
        {"text": " ", "style": "Source Code"},
    ]
    image_ids = [
        operation["args"]["imageId"]
        for operation in operations
        if operation["op"] == "writer.add_image"
    ]
    assert image_ids == ["image-1", "image-1"]
    assert len(recorded.resources) == 1
    assert validate_generation_plan(recorded.plan.to_dict(), "writer") == recorded.plan


def test_recording_sheet_renderer_records_multi_table_plan_in_windows_call_order():
    document = StructuredDocument(
        sections=[
            Section(
                level=2,
                heading="Summary",
                elements=[TableBlock(["Item", "Amount"], [["A", 10]])],
            ),
            Section(
                level=2,
                heading="Details",
                elements=[TableBlock(["Code", "Owner"], [["X", "Ada"]])],
            ),
        ]
    )
    composer = recording_composers.RecordingSheetComposer()

    recorded = sheet_renderer.render(
        document,
        "ignored.xlsx",
        composer_factory=lambda: composer,
    )

    operations = [operation.to_dict() for operation in recorded.plan.operations]
    assert operations == [
        {"op": "sheet.reset", "args": {}},
        {"op": "sheet.rename", "args": {"index": 1, "name": "Summary"}},
        {
            "op": "sheet.write_table",
            "args": {
                "startRow": 1,
                "startCol": 1,
                "values": [["Item", "Amount"], ["A", 10]],
                "headerBold": True,
                "headerShade": "#4472C4",
                "headerFontColor": "#FFFFFF",
                "fontSize": 11,
            },
        },
        {"op": "sheet.set_column_width", "args": {"column": "A", "width": 15}},
        {"op": "sheet.set_column_width", "args": {"column": "B", "width": 12}},
        {"op": "sheet.add", "args": {"name": "Details"}},
        {
            "op": "sheet.write_table",
            "args": {
                "startRow": 1,
                "startCol": 1,
                "values": [["Code", "Owner"], ["X", "Ada"]],
                "headerBold": True,
                "headerShade": "#4472C4",
                "headerFontColor": "#FFFFFF",
                "fontSize": 11,
            },
        },
        {"op": "sheet.set_column_width", "args": {"column": "A", "width": 15}},
        {"op": "sheet.set_column_width", "args": {"column": "B", "width": 12}},
        {"op": "sheet.select", "args": {"index": 1}},
        {"op": "sheet.autofit", "args": {}},
    ]
    assert {operation["op"] for operation in operations} == ALLOWED_OPERATIONS[
        "spreadsheet"
    ]
    assert validate_generation_plan(recorded.plan.to_dict(), "spreadsheet") == recorded.plan
