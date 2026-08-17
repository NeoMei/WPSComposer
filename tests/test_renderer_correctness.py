from __future__ import annotations

from types import SimpleNamespace

from skills.WPSComposer.scripts._colors import resolve_color
from skills.WPSComposer.scripts.design_presets import PRESETS
from skills.WPSComposer.scripts.quality_checks import validate_wps_slide
from skills.WPSComposer.scripts.pdf import _validate_page_indices
from skills.WPSComposer.scripts.slide import SlideComposer
from skills.WPSComposer.scripts.writer import WriterComposer


def test_white_is_literal_white_for_dark_preset():
    assert resolve_color("white", PRESETS["tech"]) == "#FFFFFF"
    assert resolve_color("white", PRESETS["tech"]) != PRESETS["tech"].get_color("bg")


def test_quality_inspection_failure_fails_closed():
    result = validate_wps_slide(object(), PRESETS["business"])

    assert result["pass"] is False
    assert result["score"] == 0
    assert result["status"] == "error"


def test_pdf_page_indices_reject_zero_negative_bool_and_out_of_range():
    for invalid in (0, -1, True, 3):
        try:
            _validate_page_indices([invalid], 2)
        except ValueError as exc:
            assert "page index" in str(exc)
        else:
            raise AssertionError(f"accepted invalid page index: {invalid!r}")


def _slide_with_picture_shape(width=640, height=480):
    calls = []
    shape = SimpleNamespace(Width=width, Height=height, LockAspectRatio=False)

    class Shapes:
        def AddPicture(self, *args):
            calls.append(args)
            return shape

    slide = SimpleNamespace(Shapes=Shapes())
    composer = object.__new__(SlideComposer)
    composer._doc = SimpleNamespace(Slides=lambda index: slide)
    return composer, shape, calls


def test_slide_add_image_honors_width_only_and_preserves_aspect_ratio():
    composer, shape, calls = _slide_with_picture_shape()

    returned = composer.add_image(1, "photo.png", 10, 20, width=400)

    assert returned is shape
    assert len(calls[0]) == 5
    assert shape.LockAspectRatio is True
    assert shape.Width == 400
    assert shape.Height == 480


def test_slide_add_image_honors_height_only_and_preserves_aspect_ratio():
    composer, shape, calls = _slide_with_picture_shape()

    returned = composer.add_image(1, "photo.png", 10, 20, height=300)

    assert returned is shape
    assert len(calls[0]) == 5
    assert shape.LockAspectRatio is True
    assert shape.Width == 640
    assert shape.Height == 300


def test_writer_add_heading_forwards_explicit_false_bold():
    composer = object.__new__(WriterComposer)
    observed = []
    composer.add_heading_level = lambda *args, **kwargs: observed.append((args, kwargs))

    composer.add_heading("Appendix", bold=False)

    assert observed == [(("Appendix",), {"level": 1, "size": None, "bold": False, "color": None})]


def test_windows_slide_preset_applies_to_slides_created_after_preset():
    title_font = SimpleNamespace(Size=None, Name=None, Color=SimpleNamespace(RGB=None))
    subtitle_font = SimpleNamespace(Size=None, Name=None, Color=SimpleNamespace(RGB=None))
    title_range = SimpleNamespace(Text="", Font=title_font)
    subtitle_range = SimpleNamespace(Text="", Font=subtitle_font)
    slide = SimpleNamespace(
        Shapes=SimpleNamespace(
            Title=SimpleNamespace(TextFrame=SimpleNamespace(TextRange=title_range)),
            Placeholders=lambda index: SimpleNamespace(
                TextFrame=SimpleNamespace(TextRange=subtitle_range)
            ),
        )
    )

    class Slides:
        Count = 0

        def Add(self, index, layout):
            self.Count += 1
            return slide

        def __call__(self, index):
            if self.Count == 0:
                raise IndexError(index)
            return slide

    master = SimpleNamespace(
        Background=SimpleNamespace(
            Fill=SimpleNamespace(ForeColor=SimpleNamespace(RGB=None))
        )
    )
    composer = object.__new__(SlideComposer)
    composer._doc = SimpleNamespace(Slides=Slides(), SlideMaster=master)

    composer.apply_design_preset(PRESETS["academic"])
    composer.add_title_slide("Title", "Subtitle")

    assert title_font.Name == PRESETS["academic"].get_font("title")[0]
    assert title_font.Size == PRESETS["academic"].get_font("title")[1]
    assert subtitle_font.Name == PRESETS["academic"].get_font("subtitle")[0]
    assert subtitle_font.Size == PRESETS["academic"].get_font("subtitle")[1]
