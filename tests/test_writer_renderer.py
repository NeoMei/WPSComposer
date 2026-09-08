from pathlib import Path
import struct

import pytest

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


def _operation_dicts(recorder):
    return [operation.to_dict() for operation in recorder._operations]


def _section_args(operations):
    return [
        operation["args"]
        for operation in operations
        if operation["op"] == "writer.add_section"
    ]


def _assert_no_adjacent_section_operations(operations):
    assert all(
        first["op"] != "writer.add_section"
        or second["op"] != "writer.add_section"
        for first, second in zip(operations, operations[1:])
    )


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


def test_configure_style_accepts_camelcase_longform_keys():
    """Regression: long-form plan ops emit camelCase style keys; the Windows
    COM path must translate them instead of silently dropping properties."""
    applied = {}

    class Font:
        Name = NameFarEast = NameAscii = NameOther = NameBi = None
        Size = Bold = None

    class ParagraphFormat:
        Alignment = SpaceBefore = SpaceAfter = KeepWithNext = None

    class Style:
        def __init__(self):
            self.Font = Font()
            self.ParagraphFormat = ParagraphFormat()

    composer = object.__new__(WriterComposer)
    composer._doc = None

    style = Style()
    composer._configure_style(
        style,
        {
            "name": "Heading 1",
            "fontName": "黑体",
            "fontSize": 16,
            "bold": True,
            "align": 1,
            "spaceBefore": 16,
            "spaceAfter": 5,
            "keepWithNext": True,
        },
    )

    assert style.Font.Name == "黑体"
    assert style.Font.Size == 16
    assert style.Font.Bold is True
    assert style.ParagraphFormat.Alignment == 1
    assert style.ParagraphFormat.SpaceBefore == 16
    assert style.ParagraphFormat.SpaceAfter == 5
    assert style.ParagraphFormat.KeepWithNext is True


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


def test_writer_renderer_leaves_a_document_end_wide_png_in_landscape(tmp_path):
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

    operations = _operation_dicts(recorder)
    image_index = next(
        index
        for index, operation in enumerate(operations)
        if operation["op"] == "writer.add_image"
    )
    assert operations[image_index - 1] == {
        "op": "writer.add_section",
        "args": {"landscape": True},
    }
    assert _section_args(operations) == [{}, {"landscape": True}]
    _assert_no_adjacent_section_operations(operations)


def test_wide_png_followed_by_chapter_reuses_portrait_section_break(tmp_path):
    image = tmp_path / "wide.png"
    image.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + b"\x00\x00\x00\rIHDR"
        + struct.pack(">II", 1600, 800)
    )
    document = StructuredDocument(
        title="图示报告",
        sections=[
            Section(level=1, heading="图示报告"),
            Section(
                level=1,
                heading="第一章 进度计划",
                elements=[ImageBlock(path=str(image), alt="进度图")],
            ),
            Section(level=1, heading="第二章 实施方案"),
        ],
    )
    recorder = RecordingWriterComposer()

    writer_renderer._render_body(recorder, document, None)

    operations = _operation_dicts(recorder)
    assert _section_args(operations) == [
        {},
        {"landscape": True},
        {"landscape": False},
    ]
    _assert_no_adjacent_section_operations(operations)
    second_chapter = next(
        index
        for index, operation in enumerate(operations)
        if operation["op"] == "writer.add_heading"
        and operation["args"]["text"] == "第二章 实施方案"
    )
    assert operations[second_chapter - 1] == {
        "op": "writer.add_section",
        "args": {"landscape": False},
    }


def test_consecutive_wide_pngs_share_one_landscape_section(tmp_path):
    images = []
    for name in ("first.png", "second.png"):
        image = tmp_path / name
        image.write_bytes(
            b"\x89PNG\r\n\x1a\n"
            + b"\x00\x00\x00\rIHDR"
            + struct.pack(">II", 1600, 800)
        )
        images.append(image)
    document = StructuredDocument(
        title="图示报告",
        sections=[
            Section(level=1, heading="图示报告"),
            Section(
                level=1,
                heading="第一章 进度计划",
                elements=[
                    ImageBlock(path=str(images[0]), alt="图一"),
                    ImageBlock(path=str(images[1]), alt="图二"),
                ],
            ),
        ],
    )
    recorder = RecordingWriterComposer()

    writer_renderer._render_body(recorder, document, None)

    operations = _operation_dicts(recorder)
    assert _section_args(operations) == [{}, {"landscape": True}]
    assert sum(op["op"] == "writer.add_image" for op in operations) == 2
    _assert_no_adjacent_section_operations(operations)


def test_local_portrait_png_ihdr_overrides_wide_display_dimensions(tmp_path):
    image = tmp_path / "portrait.png"
    image.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + b"\x00\x00\x00\rIHDR"
        + struct.pack(">II", 800, 1200)
    )
    block = ImageBlock(
        path=str(image), alt="portrait", width=400, height=200
    )

    assert writer_renderer._image_is_landscape(block) is False


def test_local_wide_png_ihdr_overrides_square_display_dimensions(tmp_path):
    image = tmp_path / "wide.png"
    image.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + b"\x00\x00\x00\rIHDR"
        + struct.pack(">II", 1600, 800)
    )
    block = ImageBlock(
        path=str(image), alt="wide", width=300, height=300
    )

    assert writer_renderer._image_is_landscape(block) is True


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


@pytest.mark.parametrize('transparent, style_name, outline', [
    (True, 'WPSC Sequence Transparent Heading 1', 10),
    (False, 'WPSC Unnumbered Heading 1', 1),
])
def test_native_unnumbered_heading_keeps_appearance_without_inheriting_later_numbering(transparent, style_name, outline):
    from copy import deepcopy
    from types import SimpleNamespace

    class Format(SimpleNamespace):
        @property
        def Duplicate(self):
            return deepcopy(self)

    source = SimpleNamespace(
        Font=Format(NameFarEast='黑体', Size=16, Bold=True, Color=0),
        ParagraphFormat=Format(OutlineLevel=1, Alignment=1, SpaceBefore=16, SpaceAfter=5, KeepWithNext=True),
    )
    normal = SimpleNamespace(NameLocal='Normal')
    named = {}

    class Styles:
        def __call__(self, key):
            if key == -2:
                return source
            if key == -1:
                return normal
            return named[key]

        def Add(self, name, style_type):
            assert style_type == 1
            target = SimpleNamespace(Font=Format(), ParagraphFormat=Format())
            named[name] = target
            return target

    class ListFormat:
        ListString = '2'

        def RemoveNumbers(self, number_type):
            assert number_type == 1
            self.ListString = ''

    native_range = SimpleNamespace(
        Style=source, Font=source.Font,
        ParagraphFormat=Format(OutlineLevel=1), ListFormat=ListFormat(),
    )
    writer = object.__new__(WriterComposer)
    position = [0]
    writer._native_position = lambda: position[0]
    writer.add_heading_level = lambda text, level: position.__setitem__(0, len(text) + 1)
    writer._doc = SimpleNamespace(Styles=Styles(), Range=lambda start, end: native_range)
    writer.add_heading_level_native('Unnumbered boundary', 1, numbering=False, sequence_transparent=transparent)
    assert native_range.Style is named[style_name]
    assert native_range.Style.BaseStyle is normal
    assert native_range.Style.Font.Size == 16
    assert native_range.Style.Font.NameFarEast == '黑体'
    assert native_range.Style.Font.Bold is True
    assert native_range.Style.ParagraphFormat.Alignment == 1
    assert native_range.Style.ParagraphFormat.SpaceBefore == 16
    assert native_range.Style.ParagraphFormat.SpaceAfter == 5
    assert native_range.Style.ParagraphFormat.KeepWithNext is True
    assert native_range.Style.ParagraphFormat.OutlineLevel == outline
    assert native_range.ParagraphFormat.OutlineLevel == outline
    assert native_range.ListFormat.ListString == ''
    assert source.ParagraphFormat.OutlineLevel == 1
    # A later numbered chapter links a list to the same built-in heading.
    # Effective inheritance of the already emitted paragraph must remain empty.
    source.LinkToListTemplate = lambda template, level: setattr(source, 'ListTemplate', template)
    template = object()
    source.LinkToListTemplate(template, 1)
    effective = native_range.Style
    while effective is not None and not getattr(effective, 'ListTemplate', None):
        effective = getattr(effective, 'BaseStyle', None)
    assert effective is None
    assert source.ListTemplate is template


def test_native_sequence_transparent_heading_rejects_numbering_before_writing():
    import pytest
    writer = object.__new__(WriterComposer)
    writer._native_position = lambda: pytest.fail('native document inspected before conflicting flags rejected')
    with pytest.raises(ValueError, match='number'):
        writer.add_heading_level_native('Conflict', 1, numbering=True, sequence_transparent=True)


def test_native_heading_none_numbering_keeps_legacy_path():
    from types import SimpleNamespace
    writer = object.__new__(WriterComposer)
    writer._native_position = lambda: 0
    state = SimpleNamespace(text='')
    writer.add_heading_level = lambda text, level: setattr(state, 'text', text)
    # Legacy None does not require detached-style host APIs.
    writer.add_heading_level_native('Legacy heading', 1)
    assert state.text == 'Legacy heading'
