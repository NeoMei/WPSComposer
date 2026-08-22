from __future__ import annotations

from dataclasses import FrozenInstanceError

import pytest

from skills.WPSComposer.scripts.design_presets import PRESETS
from skills.WPSComposer.scripts.document_model import SemanticTableBlock, TableMerge
from skills.WPSComposer.scripts.longform.policy import resolve_table_policy
from skills.WPSComposer.scripts.longform.table_policy import (
    TABLE_ROW_FORCED_SPLIT,
    TablePolicy,
    row_forced_split_degradation,
)


def _table(*, style: str = "", repeat_header: bool = True) -> SemanticTableBlock:
    return SemanticTableBlock(
        caption="Data",
        headers=["Name", "Value", "Status"],
        rows=[["alpha", "1", "ok"], ["beta", "2", "ok"]],
        alignments=["left", "right", "center"],
        style=style,
        repeat_header=repeat_header,
    )


@pytest.mark.parametrize("preset", ["academic", PRESETS["academic"]])
def test_academic_preset_defaults_to_three_line(preset: object) -> None:
    policy, issues = resolve_table_policy(_table(), preset)

    assert policy.style == "three-line"
    assert issues == ()


@pytest.mark.parametrize("preset", ["business", "consultant", "tech", "proposal"])
def test_nonacademic_presets_default_to_grid(preset: str) -> None:
    policy, issues = resolve_table_policy(_table(), preset)

    assert policy.style == "grid"
    assert issues == ()


@pytest.mark.parametrize(
    ("preset", "explicit_style"),
    [("academic", "grid"), ("business", "three-line"), ("tech", " THREE-LINE ")],
)
def test_explicit_style_overrides_preset(preset: str, explicit_style: str) -> None:
    policy, _ = resolve_table_policy(_table(style=explicit_style), preset)

    assert policy.style == explicit_style.strip().lower()


def test_three_line_policy_declares_only_the_three_visible_rules() -> None:
    policy, _ = resolve_table_policy(_table(style="three-line"), "business")

    assert policy.borders == {
        "top": 1.5,
        "bottom": 1.5,
        "header_bottom": 0.75,
        "left": 0.0,
        "right": 0.0,
        "inside_horizontal": 0.0,
        "inside_vertical": 0.0,
    }


def test_grid_policy_declares_all_grid_borders() -> None:
    policy, _ = resolve_table_policy(_table(style="grid"), "academic")

    assert policy.borders == {
        "top": 0.75,
        "bottom": 0.75,
        "header_bottom": 0.75,
        "left": 0.75,
        "right": 0.75,
        "inside_horizontal": 0.75,
        "inside_vertical": 0.75,
    }


def test_table_policy_repeats_header_disables_row_split_and_uses_zero_indent() -> None:
    policy, _ = resolve_table_policy(_table(repeat_header=False), "business")

    assert policy.repeat_header is False
    assert policy.allow_row_split is False
    assert policy.cell_indent_pt == 0.0


def test_declared_cell_alignments_remain_direct_and_unchanged() -> None:
    table = _table()
    original_alignments = list(table.alignments)

    resolve_table_policy(table, "academic")

    assert table.alignments == ["left", "right", "center"]
    assert table.alignments == original_alignments


def test_policy_and_merge_coordinates_are_immutable() -> None:
    merge = TableMerge(2, 1, 3, 1)
    policy = TablePolicy("grid", {}, (merge,), True, False, 0.0)

    with pytest.raises(FrozenInstanceError):
        merge.top = 1  # type: ignore[misc]
    with pytest.raises(FrozenInstanceError):
        policy.style = "three-line"  # type: ignore[misc]


def test_vertical_body_merges_mark_the_transitive_row_group_indivisible() -> None:
    table = SemanticTableBlock(
        headers=["A", "B", "C"],
        rows=[
            ["one", "x", ""],
            ["", "two", "y"],
            ["three", "", "z"],
        ],
        merge_spec="A2:A3;B3:B4",
    )

    policy, issues = resolve_table_policy(table, "business")

    assert issues == ()
    assert policy.indivisible_row_groups == ((2, 4),)
    assert policy.allow_row_split is False


def test_normal_rows_have_no_indivisible_group() -> None:
    policy, _ = resolve_table_policy(_table(), "business")

    assert policy.indivisible_row_groups == ()
    assert policy.allow_row_split is False


def test_forced_split_metadata_is_runtime_only_and_caption_anchored() -> None:
    metadata = row_forced_split_degradation(2, 4)

    assert metadata == {
        "code": TABLE_ROW_FORCED_SPLIT,
        "message": (
            "Table row group 2:4 exceeded the available page height; "
            "remove its vertical merges and allow row splitting."
        ),
        "placement": "block",
        "insert_after": "caption",
        "trigger": "indivisible-row-group-exceeds-available-page",
        "recovery": "unmerge-group-and-allow-row-split",
        "row_group": {"top": 2, "bottom": 4},
    }


def test_forced_split_metadata_also_supports_one_oversized_normal_row() -> None:
    metadata = row_forced_split_degradation(3, 3)

    assert metadata["code"] == TABLE_ROW_FORCED_SPLIT
    assert metadata["row_group"] == {"top": 3, "bottom": 3}
    assert metadata["recovery"] == "unmerge-group-and-allow-row-split"


def test_forced_split_metadata_does_not_measure_pages() -> None:
    with pytest.raises(ValueError, match="body row group"):
        row_forced_split_degradation(1, 2)
    with pytest.raises(ValueError, match="ordered"):
        row_forced_split_degradation(4, 3)
