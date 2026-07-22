from skills.WPSComposer.scripts.document_model import (
    ImageBlock,
    Section,
    StructuredDocument,
    TableBlock,
)
from skills.WPSComposer.scripts.renderers import slide_renderer
from skills.WPSComposer.scripts.slide import SlideComposer


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
