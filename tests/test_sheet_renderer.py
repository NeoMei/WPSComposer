from pathlib import Path

from skills.WPSComposer.scripts.document_model import (
    Section,
    StructuredDocument,
    TableBlock,
)
from skills.WPSComposer.scripts.recording_composers import RecordingSheetComposer
from skills.WPSComposer.scripts.renderers import sheet_renderer
from skills.WPSComposer.scripts.sheet import SheetComposer


def test_sheet_composer_rename_sheet_preserves_windows_object_model_calls():
    events = []

    class Worksheet:
        Name = "Sheet1"

    worksheet = Worksheet()

    class Worksheets:
        def __call__(self, index):
            events.append(("worksheet", index))
            return worksheet

    composer = object.__new__(SheetComposer)
    composer._doc = type("Workbook", (), {"Worksheets": Worksheets()})()

    result = composer.rename_sheet(1, "Summary")

    assert events == [("worksheet", 1)]
    assert worksheet.Name == "Summary"
    assert result is worksheet


def test_sheet_renderer_preserves_windows_table_calls_with_injected_composer():
    document = StructuredDocument(
        sections=[
            Section(
                level=2,
                heading="Summary",
                elements=[TableBlock(["Item"], [["A"]])],
            )
        ]
    )
    events = []

    class Composer:
        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc_value, traceback):
            return False

        def rename_sheet(self, index, name):
            events.append(("rename", index, name))

        def write_table(self, **kwargs):
            events.append(("write", kwargs))

        def set_column_width(self, column, width):
            events.append(("width", column, width))

        def autofit(self):
            events.append(("autofit",))

        def save_xlsx(self, output):
            events.append(("save", output))
            return output

    preset = type("Preset", (), {"get_color": lambda self, key: "#123456"})()

    result = sheet_renderer.render(
        document,
        "ignored.xlsx",
        preset=preset,
        composer_factory=Composer,
    )

    assert result == "ignored.xlsx"
    assert events == [
        ("rename", 1, "Summary"),
        (
            "write",
            {
                "start_row": 1,
                "start_col": 1,
                "data": [["Item"], ["A"]],
                "header_shade": True,
                "header_font_color": "#FFFFFF",
            },
        ),
        ("width", "A", 15),
        ("autofit",),
        ("save", "ignored.xlsx"),
    ]


def test_sheet_renderer_has_no_private_composer_access():
    source = Path(sheet_renderer.__file__).read_text(encoding="utf-8")
    assert "s._doc" not in source
