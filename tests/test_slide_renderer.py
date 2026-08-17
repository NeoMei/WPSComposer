import pytest

from skills.WPSComposer.scripts.document_model import (
    ExcalidrawBlock,
    ImageBlock,
    Paragraph,
    Section,
    Span,
    StructuredDocument,
    TableBlock,
)
from skills.WPSComposer.scripts.renderers import slide_renderer
from skills.WPSComposer.scripts.slide import SlideComposer
from skills.WPSComposer.scripts.recording_composers import RecordingSlideComposer
from skills.WPSComposer.scripts.md_parser import parse


def test_slide_composer_exposes_public_slide_count():
    composer = object.__new__(SlideComposer)
    composer._doc = type(
        "Presentation",
        (),
        {"Slides": type("Slides", (), {"Count": 7})()},
    )()

    assert composer.slide_count == 7


def test_slide_renderer_uses_public_slide_count_for_table_and_image_slides():
    class SeamComposer:
        def __init__(self):
            self.slide_count = 0
            self.calls = []

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc_value, traceback):
            return False

        def set_slide_size(self, width, height):
            self.calls.append(("size", width, height))

        def add_bullets_slide(self, *args, **kwargs):
            self.slide_count += 1

        def add_blank_slide(self):
            self.slide_count += 1
            return None, self.slide_count

        def add_table(self, slide_index, *args, **kwargs):
            self.calls.append(("table", slide_index))

        def add_image(self, slide_index, *args, **kwargs):
            self.calls.append(("image", slide_index))

        def save_pptx(self, path):
            return self.calls

    composer = SeamComposer()
    document = StructuredDocument(
        sections=[
            Section(
                level=2,
                heading="Content",
                elements=[
                    TableBlock(["A"], [["B"]]),
                    ImageBlock(path="figure.png"),
                ],
            )
        ]
    )

    calls = slide_renderer.render(
        document,
        "ignored.pptx",
        composer_factory=lambda: composer,
    )

    assert ("table", 1) in calls
    assert ("image", 2) in calls


def test_slide_renderer_does_not_duplicate_title_h1_and_keeps_its_body():
    from skills.WPSComposer.scripts.document_model import Paragraph, Span

    document = StructuredDocument(
        title="Title",
        sections=[
            Section(1, "Title", [Paragraph([Span("Intro must survive.")])]),
        ],
    )

    recorded = slide_renderer.render(
        document,
        "ignored.pptx",
        composer_factory=RecordingSlideComposer,
    )
    operations = [operation.to_dict() for operation in recorded.plan.operations]

    assert [op["op"] for op in operations].count("slide.add_title") == 1
    assert not any(op["op"] == "slide.add_section" for op in operations)
    bullet_ops = [op for op in operations if op["op"] == "slide.add_bullets"]
    assert [op["args"]["items"] for op in bullet_ops] == [["Intro must survive."]]


def test_slide_renderer_paginates_all_table_rows():
    rows = [[str(index)] for index in range(1, 13)]
    document = StructuredDocument(
        sections=[Section(2, "Metrics", [TableBlock(["n"], rows)])]
    )

    recorded = slide_renderer.render(
        document,
        "ignored.pptx",
        composer_factory=RecordingSlideComposer,
    )
    tables = [
        op.to_dict()["args"]["data"]
        for op in recorded.plan.operations
        if op.op == "slide.add_table"
    ]

    assert len(tables) == 2
    assert all(data[0] == ["n"] for data in tables)
    assert [row for data in tables for row in data[1:]] == rows


def test_slide_renderer_uses_preset_table_colors():
    from skills.WPSComposer.scripts.design_presets import PRESETS

    document = StructuredDocument(
        sections=[Section(2, "Metrics", [TableBlock(["n"], [["1"]])])]
    )

    recorded = slide_renderer.render(
        document,
        "ignored.pptx",
        preset=PRESETS["tech"],
        composer_factory=RecordingSlideComposer,
    )
    table = next(
        op.to_dict()["args"]
        for op in recorded.plan.operations
        if op.op == "slide.add_table"
    )

    assert table["headerShade"] == PRESETS["tech"].get_color("primary")


def test_slide_renderer_preserves_wikilink_image_width_and_height(tmp_path):
    document = parse(
        "![[photo.png|320x200]]", base_dir=str(tmp_path)
    )

    recorded = slide_renderer.render(
        document, "ignored.pptx", composer_factory=RecordingSlideComposer
    )
    image = next(
        op.to_dict()["args"]
        for op in recorded.plan.operations
        if op.op == "slide.add_image"
    )

    assert image["width"] == 320
    assert image["height"] == 200


def test_slide_renderer_preserves_wikilink_image_width_only(tmp_path):
    document = parse("![[photo.png|320]]", base_dir=str(tmp_path))

    recorded = slide_renderer.render(
        document, "ignored.pptx", composer_factory=RecordingSlideComposer
    )
    image = next(
        op.to_dict()["args"]
        for op in recorded.plan.operations
        if op.op == "slide.add_image"
    )

    assert image["width"] == 320
    assert "height" not in image


def test_slide_renderer_preserves_interleaved_element_order():
    document = StructuredDocument(
        sections=[Section(2, "Flow", [
            Paragraph([Span("before")]),
            ImageBlock(path="figure.png"),
            Paragraph([Span("after")]),
            TableBlock(["h"], [["cell"]]),
        ])]
    )

    recorded = slide_renderer.render(
        document, "ignored.pptx", composer_factory=RecordingSlideComposer
    )
    operations = [operation.op for operation in recorded.plan.operations]

    assert operations == [
        "slide.reset",
        "slide.set_size",
        "slide.add_bullets",
        "slide.add_blank",
        "slide.add_image",
        "slide.add_bullets",
        "slide.add_bullets",
        "slide.add_table",
    ]
    bullet_items = [
        operation.to_dict()["args"]["items"]
        for operation in recorded.plan.operations
        if operation.op == "slide.add_bullets"
    ]
    assert bullet_items[:2] == [["before"], ["after"]]


def test_slide_renderer_requires_excalidraw_preprocessing():
    document = StructuredDocument(
        sections=[Section(2, "Diagram", [
            ExcalidrawBlock(path="architecture.excalidraw.md")
        ])]
    )

    with pytest.raises(RuntimeError, match="Excalidraw.*excalidraw"):
        slide_renderer.render(
            document, "ignored.pptx", composer_factory=RecordingSlideComposer
        )


@pytest.mark.parametrize("failure", ["preset", "table", "image"])
def test_slide_renderer_propagates_content_and_preset_failures(failure):
    class FailingComposer:
        def __init__(self):
            self.slide_count = 0
            self.saved = False

        def __enter__(self):
            return self

        def __exit__(self, *_):
            return False

        def set_slide_size(self, *_):
            pass

        def apply_design_preset(self, _preset):
            if failure == "preset":
                raise RuntimeError("preset failed")

        def add_bullets_slide(self, *_args, **_kwargs):
            self.slide_count += 1

        def add_blank_slide(self):
            self.slide_count += 1

        def add_table(self, *_args, **_kwargs):
            if failure == "table":
                raise RuntimeError("table failed")

        def add_image(self, *_args, **_kwargs):
            if failure == "image":
                raise RuntimeError("image failed")

        def save_pptx(self, _path):
            self.saved = True

    composer = FailingComposer()
    elements = {
        "preset": [Paragraph([Span("text")])],
        "table": [TableBlock(["h"], [["cell"]])],
        "image": [ImageBlock(path="broken.png")],
    }[failure]
    document = StructuredDocument(sections=[Section(2, "Content", elements)])

    with pytest.raises(RuntimeError, match=failure):
        slide_renderer.render(
            document,
            "ignored.pptx",
            preset=object(),
            composer_factory=lambda: composer,
        )

    assert composer.saved is False
