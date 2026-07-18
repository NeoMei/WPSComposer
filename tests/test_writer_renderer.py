from pathlib import Path

from skills.WPSComposer.scripts.document_model import (
    CodeBlock,
    ImageBlock,
    ListBlock,
    Paragraph,
    Section,
    Span,
    StructuredDocument,
    TableBlock,
)
from skills.WPSComposer.scripts.design_presets import PRESETS
from skills.WPSComposer.scripts.generation_plan import validate_generation_plan
from skills.WPSComposer.scripts.recording_composers import RecordingWriterComposer
from skills.WPSComposer.scripts.renderers import writer_renderer
from skills.WPSComposer.scripts.writer import WriterComposer


class _InjectedComposer:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False

    def set_margins(self, *args):
        pass

    def ensure_styles(self, styles):
        pass

    def ensure_heading_styles(self, styles):
        pass

    def insert_toc(self, title):
        pass

    def set_page_number_in_footer(self):
        pass

    def update_fields(self):
        pass

    def add_paragraph(self, *args, **kwargs):
        pass

    def add_horizontal_line(self):
        pass

    def add_page_break(self):
        pass

    def save_docx(self, output_path):
        return output_path


def test_writer_renderer_accepts_an_injected_composer_factory():
    document = StructuredDocument()
    composer = _InjectedComposer()

    result = writer_renderer.render(
        document,
        "ignored.docx",
        composer_factory=lambda: composer,
    )

    assert result == "ignored.docx"


def test_writer_composer_exposes_renderer_facing_methods():
    assert hasattr(WriterComposer, "apply_heading_text_color")
    assert hasattr(WriterComposer, "add_rich_paragraph")
    assert hasattr(WriterComposer, "add_code_lines")


def test_writer_styled_paragraph_keeps_existing_windows_call_semantics():
    events = []

    class Selection:
        def TypeText(self, text):
            events.append(("text", text))

        def TypeParagraph(self):
            events.append(("paragraph",))

    class Styles:
        def __call__(self, name):
            events.append(("style", name))
            return f"style:{name}"

    composer = object.__new__(WriterComposer)
    composer._app = type("App", (), {"Selection": Selection()})()
    composer._doc = type("Document", (), {"Styles": Styles()})()
    composer._reset_selection_to_normal = lambda: events.append(("reset",))

    composer.add_styled_paragraph("Summary", "Body Text")

    assert events == [
        ("reset",),
        ("style", "Body Text"),
        ("text", "Summary"),
        ("paragraph",),
        ("reset",),
    ]


def test_writer_renderer_records_semantic_plan_and_resources(tmp_path):
    image_path = tmp_path / "chart.png"
    document = StructuredDocument(
        title="Quarterly report",
        metadata={"author": "Ada", "date": "2026-07-19"},
        sections=[
            Section(
                level=1,
                heading="Overview",
                elements=[
                    Paragraph([Span("Strong", bold=True), Span(" code", code=True)]),
                    ListBlock([[Span("First")], [Span("Second")]], ordered=False),
                    TableBlock(["Item", "Amount"], [["A", "10"]], ["left", "right"]),
                    CodeBlock("print('ok')\n"),
                    ImageBlock(str(image_path), "Chart", 320, 180),
                ],
            ),
            Section(level=1, heading="Next", elements=[]),
        ],
    )
    recorder = RecordingWriterComposer()

    recorded = writer_renderer.render(
        document,
        "ignored.docx",
        preset=PRESETS["business"],
        composer_factory=lambda: recorder,
    )

    names = [operation.op for operation in recorded.plan.operations]
    assert names[0] == "writer.reset"
    assert "writer.configure_page" in names
    assert "writer.ensure_styles" in names
    assert "writer.add_heading" in names
    assert "writer.add_list" in names
    assert "writer.add_table" in names
    assert "writer.add_image" in names
    assert "writer.add_section" in names
    assert names[-2:] == ["writer.set_page_number", "writer.update_fields"]
    assert recorded.resources[0].source_path == image_path.resolve()
    assert validate_generation_plan(recorded.plan.to_dict(), "writer") == recorded.plan


def test_writer_renderer_has_no_com_or_private_composer_access():
    text = Path(writer_renderer.__file__).read_text(encoding="utf-8")
    assert "w.selection" not in text
    assert "w.doc" not in text
    assert "._set_font_family" not in text
    assert "._set_line_spacing" not in text
    assert "._reset_selection_to_normal" not in text
