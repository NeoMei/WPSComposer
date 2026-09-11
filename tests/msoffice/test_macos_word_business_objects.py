"""Direct macOS Word table/image/text-box business API parity."""
from __future__ import annotations

import inspect
import importlib.util
from pathlib import Path
import re
import subprocess
import sys

from PIL import Image
import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import (
    MacWordSession,
    NativeWordCapabilityError,
)
from skills.WPSComposer.scripts.writer import WriterComposer


METHODS = ["add_table", "add_merged_table", "add_image", "add_image_block", "add_floating_textbox"]


@pytest.fixture
def session(tmp_path, monkeypatch):
    value = MacWordSession()
    value.staging_root = tmp_path
    calls = []

    def execute(lines, **kwargs):
        calls.append(lines)
        script = "\n".join(lines)
        if '"created", "table"' in script:
            rows = int(re.search(r"number of rows:(\d+)", script).group(1))
            cols = int(re.search(r"number of columns:(\d+)", script).group(1))
            return [["created", "table", 2, 1, rows, cols, False]]
        if '"created", "inline_shape"' in script:
            width, height = 320.0, 180.0
            direct_width = re.search(r"^set width of insertedImage to ([0-9.]+)$", script, re.MULTILINE)
            direct_height = re.search(r"^set height of insertedImage to ([0-9.]+)$", script, re.MULTILINE)
            boxed_width = re.search(r"^set scaleFactor to ([0-9.]+) / naturalWidth$", script, re.MULTILINE)
            boxed_height = re.search(r"^if \(([0-9.]+) / naturalHeight\)", script, re.MULTILINE)
            if boxed_width and boxed_height:
                scale = min(float(boxed_width.group(1)) / width, float(boxed_height.group(1)) / height)
                width, height = width * scale, height * scale
            elif direct_width:
                requested = float(direct_width.group(1))
                if "lock aspect ratio of insertedImage to true" in script:
                    height *= requested / width
                width = requested
            elif direct_height:
                requested = float(direct_height.group(1))
                if "lock aspect ratio of insertedImage to true" in script:
                    width *= requested / height
                height = requested
            max_width = re.search(r"^set maxWidth to ([0-9.]+)$", script, re.MULTILINE)
            max_height = re.search(r"^set maxHeight to ([0-9.]+)$", script, re.MULTILINE)
            scale = min(1.0, float(max_width.group(1)) / width if max_width else 1.0,
                        float(max_height.group(1)) / height if max_height else 1.0)
            width, height = width * scale, height * scale
            alt_line = re.search(r'^set alternative text of insertedImage to "(.*)"$', script, re.MULTILINE)
            alt = alt_line.group(1) if alt_line else "missing value"
            return [["created", "inline_shape", 3, 2, 320.0, 180.0, width, height, alt]]
        if '"created", "shape"' in script and "make new text box" in script:
            geometry = re.search(r"left position:([0-9.]+), top:([0-9.]+), width:([0-9.]+), height:([0-9.]+)", script)
            text = re.search(r'^set content of text range of text frame of insertedShape to "(.*)"$', script, re.MULTILINE).group(1)
            size = float(re.search(r"font size of font object of text range of text frame of insertedShape to ([0-9.]+)", script).group(1))
            bold = "bold of font object of text range of text frame of insertedShape to true" in script
            fill = "visible of fill format of insertedShape to true" in script
            wrap = 4 if "wrap type of wrap format of insertedShape to wrap top bottom" in script else 0
            return [["created", "shape", 2, 1, True, text + "\r",
                     *(float(value) for value in geometry.groups()), wrap, size, bold,
                     fill, [65535, 61680, 46260]]]
        if '"created", "shape"' in script:
            width = float(re.search(r"^set width of insertedImage to ([0-9.]+)$", script, re.MULTILINE).group(1))
            height = float(re.search(r"^set height of insertedImage to ([0-9.]+)$", script, re.MULTILINE).group(1))
            return [["created", "shape", 1, 0, 320.0, 180.0, width, height, 0, True]]
        return [["ok"]]

    monkeypatch.setattr(value, "_execute", execute)
    return value, calls


def _png(path: Path, size=(320, 180)) -> Path:
    Image.new("RGB", size, "#4472C4").save(path)
    return path


@pytest.mark.parametrize("name", METHODS)
def test_exact_object_method_signatures(name):
    assert inspect.signature(getattr(MacWordSession, name)) == inspect.signature(getattr(WriterComposer, name))


def test_table_compiles_one_complete_native_batch_and_returns_readback_target(session):
    value, calls = session
    target = value.add_table(
        3,
        3,
        [["Header 中文😀", "H2", "H3"], ["L", "C", "R"], ["B", "wrap", "T"]],
        col_widths=[90, 150, 105],
        alignments=["left", "center", "right"],
    )
    assert target == "table:2"
    assert len(calls) == 1
    script = "\n".join(calls[0])
    for token in (
        "number of rows:3", "number of columns:3", "Header 中文😀",
        "first line indent", "character unit first line indent", "paragraph format left indent",
        "paragraph format right indent", "space before", "space after", "line space single",
        "cell align vertical center", "align paragraph left", "align paragraph center",
        "align paragraph right", "heading format of row 1", "allow break across pages",
        "background pattern color", "line style single", "width of ownCell to 150",
        '"created", "table"',
    ):
        assert token in script
    assert value._structural_changed


def test_default_table_widths_are_content_aware_fixed_ratios(session):
    value, calls = session
    value.add_table(2, 2, [["ID", "Long narrative content"], ["1", "detail"]])
    script = "\n".join(calls[0])
    assert "set availableWidth to" in script
    assert "set allow auto fit of insertedTable to false" in script
    assert "availableWidth *" in script
    assert "set previousTableCount to count tables of boundDoc" in script
    assert "tableIndex is not previousTableCount + 1" in script


def test_structural_object_batches_activate_only_the_exact_bound_window(session, tmp_path):
    value, calls = session
    source = _png(tmp_path / "source.png")
    value.add_table(1, 1, [["cell"]])
    value.add_image(source, width=72)
    value.add_floating_textbox("text", 72, 360, 220, 54)
    for lines in calls:
        script = "\n".join(lines)
        assert "activate object boundWindow" in script
        assert script.index("activate object boundWindow") < min(
            index for token in ("make new table", "make new inline picture", "make new text box")
            if (index := script.find(token)) >= 0
        )


def test_merged_table_preflights_all_merges_and_sorts_bottom_right_first(session):
    value, calls = session
    target = value.add_merged_table(
        [["A", "", "C"], ["D", "E", "F"], ["G", "H", ""]],
        merges=[(1, 1, 1, 2), (2, 3, 3, 3)],
    )
    assert target == "table:2"
    script = "\n".join(calls[0])
    assert script.index("set mergeStart to get cell from table insertedTable row 2 column 3") < script.index("set mergeStart to get cell from table insertedTable row 1 column 1")
    assert script.count("merge cell mergeStart with mergeEnd") == 2


@pytest.mark.parametrize("merges", [[(0, 1, 1, 1)], [(1, 1, 4, 1)], [(1, 1, 2, 2), (2, 2, 3, 3)]])
def test_invalid_late_merge_prevents_whole_native_call(session, merges):
    value, calls = session
    with pytest.raises(ValueError):
        value.add_merged_table([["a", "b", "c"], ["d", "e", "f"], ["g", "h", "i"]], merges=merges)
    assert calls == []


@pytest.mark.parametrize(
    "args,kwargs",
    [
        ((0, 2, []), {}),
        ((2, 2, [["a", "b", "overflow"]]), {}),
        ((2, 2, [["a"]]), {"shade_header": "bad"}),
        ((2, 2, [["a"]]), {"font_size": float("nan")}),
        ((2, 2, [["a"]]), {"col_widths": [10]}),
        ((2, 2, [["a"]]), {"alignments": ["left", "bad"]}),
        ((2, 2, [["a"]]), {"banded_rows": 1}),
    ],
)
def test_invalid_table_arguments_reject_before_execute(session, args, kwargs):
    value, calls = session
    with pytest.raises((TypeError, ValueError)):
        value.add_table(*args, **kwargs)
    assert calls == []


def test_empty_merged_table_is_transport_free_even_read_only(session):
    value, calls = session
    value._read_only = True
    assert value.add_merged_table([]) is None
    assert calls == []


def test_inline_image_stages_private_copy_and_supports_complete_size_math(session, tmp_path):
    value, calls = session
    source = _png(tmp_path / "source.png")
    target = value.add_image(source, width=100, height=100, max_width=90, max_height=60, alt="alt")
    assert target == "inline_shape:3"
    script = "\n".join(calls[0])
    assert str(source) not in script
    assert "image-" in script and "link to file:false" in script and "save with document:true" in script
    assert "naturalWidth" in script and "scaleFactor" in script
    assert "alternative text of insertedImage to \"alt\"" in script
    assert '"created", "inline_shape"' in script


@pytest.mark.parametrize(
    "kwargs,required",
    [
        ({}, "naturalWidth"),
        ({"height": 50}, "height of insertedImage to 50"),
        ({"width": 80, "preserve_aspect": False}, "width of insertedImage to 80"),
        ({"max_width": 70}, "maxWidth"),
        ({"max_height": 40}, "maxHeight"),
    ],
)
def test_inline_image_compiles_baseline_size_branches(session, tmp_path, kwargs, required):
    value, calls = session
    source = _png(tmp_path / "source.png")
    value.add_image(source, **kwargs)
    assert required in "\n".join(calls[0])


def test_floating_image_uses_anchor_geometry_wrap_and_rejects_unverified_alt(session, tmp_path):
    value, calls = session
    source = _png(tmp_path / "source.png")
    assert value.add_image(source, width=96, height=54, inline=False, wrap=0) == "shape:1"
    script = "\n".join(calls[0])
    assert "make new picture" in script and "anchor:imageRange" in script
    assert "wrap square" in script and '"created", "shape"' in script
    calls.clear()
    with pytest.raises(NativeWordCapabilityError, match="floating.*alt"):
        value.add_image(source, inline=False, alt="blocked")
    assert calls == []


@pytest.mark.parametrize(
    "method,kwargs",
    [
        ("add_image", {"wrap": 1}),
        ("add_image_block", {"inline": False}),
    ],
)
def test_unproved_image_layout_branches_reject_before_staging(session, tmp_path, method, kwargs):
    value, calls = session
    source = _png(tmp_path / "source.png")
    before = set(value.staging_root.iterdir())
    with pytest.raises(NativeWordCapabilityError):
        getattr(value, method)(source, **kwargs)
    assert calls == []
    assert set(value.staging_root.iterdir()) == before


def test_floating_two_dimension_branch_matches_writer_stretch_semantics(session, tmp_path):
    value, calls = session
    source = _png(tmp_path / "source.png")
    value.add_image(source, width=100, height=100, inline=False, preserve_aspect=True)
    script = "\n".join(calls[0])
    assert "set lock aspect ratio of insertedImage to false" in script
    assert "set width of insertedImage to 100" in script
    assert "set height of insertedImage to 100" in script


@pytest.mark.parametrize(
    "values,actual,expected",
    [
        ({"width": 100, "height": None, "max_width": 70, "max_height": None}, (70, 39.375), True),
        ({"width": None, "height": 80, "max_width": None, "max_height": 40}, (71.111111, 40), True),
        ({"width": 100, "height": 100, "max_width": None, "max_height": None}, (1, 1), False),
        ({"width": 100, "height": 100, "max_width": None, "max_height": None}, (100, 56.25), True),
    ],
)
def test_image_ack_requires_exact_size_policy(values, actual, expected):
    row = [["created", "inline_shape", 1, 0, 320, 180, actual[0], actual[1], "missing value"]]
    assert MacWordSession._image_ack(row, "inline_shape", values, True, None, 0) is expected


@pytest.mark.parametrize("natural", [(0, 180), (float("inf"), 180)])
def test_image_ack_rejects_malformed_natural_dimensions(natural):
    values = {"width": 72, "height": None, "max_width": None, "max_height": None}
    row = [["created", "inline_shape", 1, 0, natural[0], natural[1], 72, 40.5, "missing value"]]
    assert MacWordSession._image_ack(row, "inline_shape", values, True, None, 0) is False


def test_invalid_alt_control_character_rejects_before_staging(session, tmp_path):
    value, calls = session
    source = _png(tmp_path / "source.png")
    before = set(value.staging_root.iterdir())
    with pytest.raises(ValueError):
        value.add_image(source, alt="bad\u0000alt")
    assert calls == [] and set(value.staging_root.iterdir()) == before


def test_svg_is_rejected_before_execute_and_partial_stage_is_removed(session, tmp_path):
    value, calls = session
    source = tmp_path / "source.svg"
    source.write_text('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"/>')
    before = set(value.staging_root.iterdir())
    with pytest.raises(NativeWordCapabilityError):
        value.add_image(source)
    assert calls == []
    assert set(value.staging_root.iterdir()) == before


def test_multiframe_resource_uses_shared_first_frame_png_normalization(session, tmp_path):
    value, calls = session
    source = tmp_path / "animated.gif"
    frames = [Image.new("RGB", (12, 8), color) for color in ("red", "blue")]
    frames[0].save(source, save_all=True, append_images=frames[1:], format="GIF")
    value.add_image(source)
    script = "\n".join(calls[0])
    staged = next(value.staging_root.glob("image-*.png"))
    with Image.open(staged) as normalized:
        assert normalized.format == "PNG"
        assert getattr(normalized, "n_frames", 1) == 1
    from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
    assert apple_string(str(staged)) in script


def test_image_block_is_one_native_batch_and_keeps_following_boundary(session, tmp_path):
    value, calls = session
    source = _png(tmp_path / "source.png")
    assert value.add_image_block(source, width=84) == "inline_shape:3"
    assert len(calls) == 1
    script = "\n".join(calls[0])
    assert "align paragraph center" in script
    assert "keep with next" in script
    assert "space after" in script
    assert "set content of followingBoundary to return" in script


def test_textbox_uses_page_relative_geometry_and_exact_readback_target(session):
    value, calls = session
    assert value.add_floating_textbox("Text 中文😀", 72, 360, 220, 54, wrap=0,
                                      fill_color="#FFF0B4", font_size=13, bold=True) == "shape:2"
    script = "\n".join(calls[0])
    assert "make new text box" in script
    assert script.index("relative horizontal position") < script.index("left position of insertedShape to 72")
    assert "relative vertical position page" in script
    assert "wrap square" in script and "font size" in script and "fill format" in script
    assert '"created", "shape"' in script


@pytest.mark.parametrize("rows", [[], [[]], [["wrong", "kind"]], [["created", "shape", 0]], [["created", "shape", -1]]])
def test_malformed_created_object_acknowledgement_is_controlled_and_retained(session, tmp_path, monkeypatch, rows):
    value, calls = session
    source = _png(tmp_path / "source.png")
    monkeypatch.setattr(value, "_execute", lambda lines, **kwargs: rows)
    with pytest.raises(Exception) as captured:
        value.add_image(source)
    assert type(captured.value).__name__ == "NativeWordError"
    assert value._retain_evidence is True
    assert value._structural_changed is True


@pytest.mark.parametrize(
    "method,args,bad_rows",
    [
        ("add_table", (1, 1, [["x"]]), [["created", "table", 1, 0, 2, 1, False]]),
        ("add_image", ("source.png",), [["created", "inline_shape", 1, 0, 320, 180, 0, 0, "missing value"]]),
        ("add_floating_textbox", ("x", 0, 0, 10, 10), [["created", "shape", 1, 0, False, "x", 0, 0, 10, 10, 0, 11, False, False, [0, 0, 0]]]),
    ],
)
def test_wrong_native_object_postcondition_is_rejected_and_retained(
    session, tmp_path, monkeypatch, method, args, bad_rows
):
    value, calls = session
    if method == "add_image":
        args = (_png(tmp_path / "source.png"),)
    monkeypatch.setattr(value, "_execute", lambda lines, **kwargs: bad_rows)
    with pytest.raises(Exception) as captured:
        getattr(value, method)(*args)
    assert type(captured.value).__name__ == "NativeWordError"
    assert value._structural_changed is True
    assert value._retain_evidence is True


def test_native_object_failure_invalidates_structure_and_retains_evidence(
    session, tmp_path, monkeypatch
):
    value, calls = session
    source = _png(tmp_path / "source.png")

    def fail(lines, **kwargs):
        from skills.WPSComposer.scripts.msoffice.errors import NativeWordError
        raise NativeWordError("NATIVE_WORD_EXECUTION_FAILED")

    monkeypatch.setattr(value, "_execute", fail)
    with pytest.raises(Exception):
        value.add_image(source)
    assert value._structural_changed is True
    assert value._retain_evidence is True


@pytest.mark.parametrize(
    "method,args,kwargs",
    [
        ("add_image", ("missing.png",), {"width": -1}),
        ("add_image", ("missing.png",), {"max_height": float("inf")}),
        ("add_image", ("missing.png",), {"wrap": 99}),
        ("add_floating_textbox", ("x", 0, 0, -1, 10), {}),
        ("add_floating_textbox", (object(), 0, 0, 10, 10), {}),
        ("add_floating_textbox", ("x", 0, 0, 10, 10), {"fill_color": "bad"}),
        ("add_floating_textbox", ("x", 0, 0, 10, 10), {"wrap": 2}),
    ],
)
def test_invalid_object_arguments_reject_before_execute_or_staging(session, method, args, kwargs):
    value, calls = session
    before = set(value.staging_root.iterdir())
    with pytest.raises((TypeError, ValueError, NativeWordCapabilityError)):
        getattr(value, method)(*args, **kwargs)
    assert calls == []
    assert set(value.staging_root.iterdir()) == before


@pytest.mark.parametrize("method,args", [("add_table", (1, 1, [["x"]])), ("add_image", ("missing.png",)), ("add_floating_textbox", ("x", 0, 0, 10, 10))])
def test_read_only_rejects_before_transport_and_image_staging(session, method, args):
    value, calls = session
    value._read_only = True
    before = set(value.staging_root.iterdir())
    with pytest.raises(ValueError, match="read-only"):
        getattr(value, method)(*args)
    assert calls == [] and set(value.staging_root.iterdir()) == before


@pytest.mark.skipif(sys.platform != "darwin" or not Path("/Applications/Microsoft Word.app").is_dir(), reason="Installed Word dictionary required")
def test_all_object_compiler_branches_are_syntax_valid(session, tmp_path):
    value, calls = session
    source = _png(tmp_path / "source.png")
    value.add_table(2, 2, [["a", "b"], ["c", "d"]])
    value.add_merged_table([["a", ""], ["c", "d"]], [(1, 1, 1, 2)])
    value.add_image(source, width=72, alt="inline")
    value.add_image(source, width=96, height=54, inline=False, wrap=0)
    value.add_image_block(source, max_width=84)
    value.add_floating_textbox("text", 72, 360, 220, 54, wrap=4)
    for index, lines in enumerate(calls):
        script = tmp_path / f"object-{index}.applescript"
        script.write_text('tell application "/Applications/Microsoft Word.app"\nset boundDoc to document 1\n' + "\n".join(lines) + "\nend tell\n")
        result = subprocess.run(["/usr/bin/osacompile", "-o", str(tmp_path / f"object-{index}.scpt"), str(script)], capture_output=True, text=True, timeout=20)
        assert result.returncode == 0, result.stderr


def test_native_object_fixture_import_has_no_office_side_effects(monkeypatch):
    def forbidden(*args, **kwargs):
        raise AssertionError("native fixture ran during import")
    monkeypatch.setattr(MacWordSession, "_prepare", forbidden)
    path = Path(__file__).resolve().parents[2] / "fixtures/microsoft_parity/macos_word_business_objects.py"
    spec = importlib.util.spec_from_file_location("native_word_business_objects", path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    assert callable(module.run) and callable(module.main)
