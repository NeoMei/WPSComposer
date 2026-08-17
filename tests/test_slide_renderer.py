from skills.WPSComposer.scripts.document_model import (
    ImageBlock,
    Section,
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
