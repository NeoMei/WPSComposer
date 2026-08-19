from pathlib import Path
import struct

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
from skills.WPSComposer.scripts.reference_styles import STYLES
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


def test_render_body_skips_title_and_leaves_pre_chapter_headings_unnumbered():
    # The first H1 equals doc.title (rendered on the cover page): it must
    # not appear in the body, in the TOC, or take part in numbering.
    document = StructuredDocument(
        title="知识库应用价值与材料清单（建议稿）",
        sections=[
            Section(level=1, heading="知识库应用价值与材料清单（建议稿）"),
            Section(level=2, heading="核心组织机制：部门文档柜"),
            Section(level=3, heading="制度法规一键查询"),
        ],
    )
    recorder = RecordingWriterComposer()
    writer_renderer._render_body(recorder, document, None)

    heads = [
        op.args["text"]
        for op in recorder._operations
        if op.op == "writer.add_heading"
    ]
    assert "第一章 知识库应用价值与材料清单（建议稿）" not in heads
    # A preface before the first chapter is outside the numbered hierarchy.
    assert heads == [
        "核心组织机制：部门文档柜",
        "制度法规一键查询",
    ]


def test_render_body_preserves_key_method_prefix_for_native_numbering_pass():
    document = StructuredDocument(
        title="技术标",
        sections=[
            Section(level=1, heading="技术标"),
            Section(level=1, heading="第三章 施工方案"),
            Section(level=2, heading="3.4 工程专项施工方法"),
            Section(level=3, heading="3.4.1 安全隐患消除工程"),
            Section(level=4, heading="关键工法01：削坡与回填筑坡"),
        ],
    )
    recorder = RecordingWriterComposer()

    writer_renderer._render_body(recorder, document, None)

    heads = [
        op.args["text"]
        for op in recorder._operations
        if op.op == "writer.add_heading"
    ]
    assert heads[-1] == "关键工法01：削坡与回填筑坡"


def test_writer_renderer_wraps_wide_png_in_landscape_sections(tmp_path):
    image = tmp_path / "wide.png"
    image.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + b"\x00\x00\x00\rIHDR"
        + struct.pack(">II", 3040, 2240)
    )
    document = StructuredDocument(
        title="图示报告",
        sections=[
            Section(level=1, heading="图示报告"),
            Section(
                level=1,
                heading="第一章 进度计划",
                elements=[ImageBlock(path=str(image), alt="图08 总进度横道图")],
            ),
        ],
    )
    recorder = RecordingWriterComposer()

    writer_renderer._render_body(recorder, document, None)

    operations = [operation.to_dict() for operation in recorder._operations]
    image_index = next(
        index
        for index, operation in enumerate(operations)
        if operation["op"] == "writer.add_image"
    )
    assert operations[image_index - 1] == {
        "op": "writer.add_section",
        "args": {"landscape": True},
    }
    assert operations[image_index + 2] == {
        "op": "writer.add_section",
        "args": {"landscape": False},
    }


def test_writer_body_styles_keep_paragraphs_together_before_chapter_breaks():
    assert STYLES["BodyText"]["keep_together"] is True
    assert STYLES["FirstParagraph"]["keep_together"] is True


def test_render_body_keeps_elements_from_document_title_section():
    document = StructuredDocument(
        title="Report",
        sections=[
            Section(
                level=1,
                heading="Report",
                elements=[Paragraph([Span("Intro must survive.")])],
            ),
            Section(level=2, heading="Details", elements=[]),
        ],
    )
    recorder = RecordingWriterComposer()

    writer_renderer._render_body(recorder, document, None)

    operations = [operation.to_dict() for operation in recorder._operations]
    headings = [op["args"]["text"] for op in operations if op["op"] == "writer.add_heading"]
    paragraphs = [op["args"]["text"] for op in operations if op["op"] == "writer.add_paragraph"]
    assert "Report" not in headings
    assert "Intro must survive." in paragraphs


def test_render_body_keeps_multiple_h1_without_cover_title():
    # Only the title-section is skipped; additional H1s still render.
    # Manual 第一章 / 第二章 prefixes are detected and kept as-is.
    document = StructuredDocument(
        title="报告",
        sections=[
            Section(level=1, heading="报告"),
            Section(level=1, heading="第一章 引言"),
            Section(level=1, heading="第二章 结论"),
        ],
    )
    recorder = RecordingWriterComposer()
    writer_renderer._render_body(recorder, document, None)
    heads = [
        op.args["text"]
        for op in recorder._operations
        if op.op == "writer.add_heading"
    ]
    assert heads == ["第一章 引言", "第二章 结论"]
