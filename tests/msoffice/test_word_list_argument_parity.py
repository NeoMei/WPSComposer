"""Pure parity tests for the closed ``writer.add_list`` operation."""

from __future__ import annotations

from dataclasses import replace

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    GenerationOperation,
    OperationPlanError,
)
from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan


_MISSING = object()


def _compile_list(
    tmp_path,
    *,
    items=("one", "two"),
    ordered=False,
    glyph=_MISSING,
    indent=_MISSING,
):
    build = build_longform_generation("# Report\n\n- seed")
    list_args = {"items": list(items), "ordered": ordered}
    if glyph is not _MISSING:
        list_args["glyph"] = glyph
    if indent is not _MISSING:
        list_args["indent"] = indent
    plan = replace(
        build.plan,
        operations=tuple(
            replace(operation, args=list_args)
            if operation.op == "writer.add_list"
            else operation
            for operation in build.plan.operations
        ),
    )
    return compile_plan(plan, {}, tmp_path / "owned.docx", timeout=30)


def test_mac_custom_glyph_uses_literal_prefix_and_requested_hanging_indent(tmp_path):
    source = _compile_list(tmp_path, glyph="→", indent=28).source

    assert 'my appendText(ownedDoc, "→" & tab & "one" & return)' in source
    assert 'my appendText(ownedDoc, "→" & tab & "two" & return)' in source
    assert source.count("set style of r to style list paragraph") == 2
    assert source.count(
        "set paragraph format left indent of paragraph format of r to 28"
    ) == 2
    assert source.count("set first line indent of paragraph format of r to -28") == 2
    assert source.count(
        "make new tab stop at paragraph 1 of r with properties {tab stop position:28}"
    ) == 2
    assert "apply bullet default" not in source
    assert "apply number default" not in source


def test_mac_list_clears_inherited_direct_format_and_normalizes_its_tail(tmp_path):
    source = _compile_list(tmp_path, items=("item",), glyph="→", indent=28).source
    list_block = source.split('log "WPSC_OP:writer.add_list"', 1)[1].split(
        'log "WPSC_OP:writer.finalize_fields"', 1
    )[0]

    style = list_block.index("set style of r to style list paragraph")
    reset_font = list_block.index("reset font object of r")
    reset_paragraph = list_block.index("reset paragraph format of r")
    left_indent = list_block.index(
        "set paragraph format left indent of paragraph format of r to 28"
    )
    assert style < reset_font < reset_paragraph < left_indent
    assert "set style of trailingRange to style normal" in list_block
    assert "reset font object of trailingRange" in list_block
    assert "reset paragraph format of trailingRange" in list_block


def test_mac_default_bullet_and_empty_glyph_remain_literal_caller_data(tmp_path):
    default_source = _compile_list(tmp_path, items=("default",)).source
    empty_source = _compile_list(tmp_path, items=("empty",), glyph="").source

    assert 'my appendText(ownedDoc, "•" & tab & "default" & return)' in default_source
    assert 'my appendText(ownedDoc, "" & tab & "empty" & return)' in empty_source


def test_mac_ordered_list_ignores_even_an_unsafe_glyph(tmp_path):
    source = _compile_list(
        tmp_path,
        items=("first", "second"),
        ordered=True,
        glyph="ignored\x00glyph",
        indent=31.5,
    ).source

    assert 'my appendText(ownedDoc, "1." & tab & "first" & return)' in source
    assert 'my appendText(ownedDoc, "2." & tab & "second" & return)' in source
    assert "ignored" not in source
    assert "tab stop position:31.5" in source


@pytest.mark.parametrize("indent", [0, 31.5, -4.25])
def test_mac_accepts_every_finite_plan_indent_without_normalizing(tmp_path, indent):
    source = _compile_list(tmp_path, items=("item",), indent=indent).source

    assert (
        f"set paragraph format left indent of paragraph format of r to {indent}"
        in source
    )
    assert f"set first line indent of paragraph format of r to {-indent}" in source
    assert f"tab stop position:{indent}" in source


@pytest.mark.parametrize("indent", [float("nan"), float("inf"), float("-inf"), True])
def test_mac_rejects_nonfinite_or_boolean_indent_before_script_creation(
    tmp_path, indent
):
    with pytest.raises(OperationPlanError, match="JSON-compatible|finite number"):
        _compile_list(tmp_path, items=("item",), indent=indent)


def test_mac_empty_items_do_not_materialize_a_blank_native_list(tmp_path):
    source = _compile_list(tmp_path, items=(), glyph="■", indent=19).source

    assert "WPSC_OP:writer.add_list" in source
    assert "style list paragraph" not in source
    assert "tab stop position:19" not in source
    assert "apply bullet default" not in source


def test_mac_quotes_script_like_glyph_and_rejects_control_data(tmp_path):
    source = _compile_list(
        tmp_path,
        items=("item",),
        glyph='"; error "PWN',
    ).source

    assert (
        'my appendText(ownedDoc, "\\"; error \\"PWN" & tab & "item" & return)'
        in source
    )
    assert 'error "PWN"' not in source
    with pytest.raises(ValueError, match="unsupported control characters"):
        _compile_list(tmp_path, items=("item",), glyph="bad\x00glyph")


class _ListRecorder:
    def __init__(self):
        self.calls = []

    def add_bullet_list(self, items, glyph="•", indent=24):
        self.calls.append(("bullet", list(items), glyph, indent))

    def add_numbered_list(self, items, indent=24):
        self.calls.append(("ordered", list(items), indent))


def test_windows_forwards_bullet_glyph_and_indent_and_uses_default_glyph():
    composer = _ListRecorder()
    executor = WindowsLongformExecutor()

    executor._dispatch_one(
        composer,
        GenerationOperation(
            op="writer.add_list",
            args={"items": ["custom"], "ordered": False, "glyph": "→", "indent": 28},
        ),
    )
    executor._dispatch_one(
        composer,
        GenerationOperation(
            op="writer.add_list",
            args={"items": ["default"], "ordered": False, "indent": 19.5},
        ),
    )

    assert composer.calls == [
        ("bullet", ["custom"], "→", 28),
        ("bullet", ["default"], "•", 19.5),
    ]


def test_windows_forwards_ordered_indent_and_ignores_glyph():
    composer = _ListRecorder()
    executor = WindowsLongformExecutor()

    executor._dispatch_one(
        composer,
        GenerationOperation(
            op="writer.add_list",
            args={
                "items": ["first", "second"],
                "ordered": True,
                "glyph": "ignored\x00glyph",
                "indent": 31.5,
            },
        ),
    )

    assert composer.calls == [("ordered", ["first", "second"], 31.5)]
