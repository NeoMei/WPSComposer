from __future__ import annotations

from copy import deepcopy

import pytest

from skills.WPSComposer.scripts.document_model import SemanticTableBlock, TableMerge
from skills.WPSComposer.scripts.longform.table_policy import (
    TABLE_MERGE_INVALID,
    parse_a1_merge_ranges,
    resolve_table_policy,
)
from skills.WPSComposer.scripts.md_parser import parse_markdown


def _table(*, merges: str = "") -> SemanticTableBlock:
    return SemanticTableBlock(
        identifier="tab:data",
        caption="Data",
        headers=["Group", "Metric", "Value"],
        rows=[
            ["A", "Count", "1"],
            ["", "Rate", "2"],
            ["B", "Count", "3"],
        ],
        alignments=["left", "center", "right"],
        merge_spec=merges,
    )


def test_a1_ranges_are_one_based_case_normalized_and_ordered() -> None:
    assert parse_a1_merge_ranges(" a2:a3 ; b2:c2 ") == (
        TableMerge(top=2, left=1, bottom=3, right=1),
        TableMerge(top=2, left=2, bottom=2, right=3),
    )


@pytest.mark.parametrize(
    "merge_spec",
    [
        "B2:A2",  # reversed columns
        "A3:A2",  # reversed rows
        "A2:A2",  # degenerate single cell
        "A2",  # no rectangle
        "A2-B2",  # wrong separator
        "A2:B2:C2",  # more than two corners
        "A0:B1",  # zero is not a one-based row
        "A01:B01",  # non-canonical row spelling
        "1A:1B",  # column must precede row
        "A2,B2",  # cell list is not a rectangle
        "A2:B2;;A3:B3",  # empty declaration
    ],
)
def test_reversed_degenerate_and_nonrectangular_syntax_is_rejected(
    merge_spec: str,
) -> None:
    with pytest.raises(ValueError):
        parse_a1_merge_ranges(merge_spec)


def test_multi_letter_columns_use_exact_one_based_coordinates() -> None:
    assert parse_a1_merge_ranges("Z2:AA3") == (
        TableMerge(top=2, left=26, bottom=3, right=27),
    )


@pytest.mark.parametrize(
    "coordinates",
    [
        (0, 1, 1, 2),
        (1, -1, 1, 2),
        (2, 1, 1, 2),
        (1, 2, 1, 1),
        (1, 1, 1, 1),
    ],
)
def test_direct_table_merge_construction_enforces_one_based_rectangle(
    coordinates: tuple[int, int, int, int],
) -> None:
    with pytest.raises(ValueError):
        TableMerge(*coordinates)


def test_header_horizontal_merge_is_accepted_when_covered_cells_are_empty() -> None:
    table = _table(merges="A1:B1")
    table.headers = ["Grouped heading", "", "Value"]

    policy, issues = resolve_table_policy(table, "academic")

    assert issues == ()
    assert policy.merges == (TableMerge(1, 1, 1, 2),)


def test_body_vertical_merge_is_accepted_when_covered_cells_are_empty() -> None:
    policy, issues = resolve_table_policy(_table(merges="A2:A3"), "business")

    assert issues == ()
    assert policy.merges == (TableMerge(2, 1, 3, 1),)


@pytest.mark.parametrize(
    "merge_spec",
    [
        "A1:A2",  # header/body crossing and header vertical merge
        "A1:B2",  # rectangular but crosses repeated-header boundary
        "A2:A5",  # row out of bounds
        "A2:D2",  # column out of bounds
        "A2:A3;A3:A4",  # overlap
    ],
)
def test_semantically_invalid_ranges_discard_all_merges(merge_spec: str) -> None:
    policy, issues = resolve_table_policy(_table(merges=merge_spec), "business")

    assert policy.merges == ()
    assert len(issues) == 1
    assert issues[0].code == TABLE_MERGE_INVALID
    assert issues[0].placement == "block"


@pytest.mark.parametrize(
    ("merge_spec", "covered_value"),
    [
        ("A2:A3", "not empty"),
        ("A1:B1", "not empty"),
        ("A2:B3", "not empty"),
    ],
)
def test_every_covered_cell_except_top_left_must_be_empty(
    merge_spec: str,
    covered_value: str,
) -> None:
    table = _table(merges=merge_spec)
    if merge_spec == "A1:B1":
        table.headers = ["anchor", covered_value, "Value"]
    elif merge_spec == "A2:A3":
        table.rows[1][0] = covered_value
    else:
        table.rows[0] = ["anchor", covered_value, "1"]
        table.rows[1] = ["", "", "2"]

    policy, issues = resolve_table_policy(table, "business")

    assert policy.merges == ()
    assert [issue.code for issue in issues] == [TABLE_MERGE_INVALID]


def test_whitespace_only_covered_cells_are_empty() -> None:
    table = _table(merges="A2:A3")
    table.rows[1][0] = " \t "

    policy, issues = resolve_table_policy(table, "business")

    assert issues == ()
    assert policy.merges == (TableMerge(2, 1, 3, 1),)


def test_one_invalid_range_discards_every_requested_merge_without_mutating_grid() -> None:
    table = _table(merges="A2:A3;B2:C2;A3:A4")
    table.rows[0][2] = ""
    original = deepcopy(table)

    policy, issues = resolve_table_policy(table, "academic")

    assert policy.merges == ()
    assert len(issues) == 1
    assert issues[0].code == TABLE_MERGE_INVALID
    assert issues[0].placement == "block"
    assert issues[0].insert_after == "caption"
    assert issues[0].to_dict() == {
        "code": TABLE_MERGE_INVALID,
        "message": (
            "Table merge declaration is invalid; all merges were discarded and "
            "the complete unmerged grid was preserved."
        ),
        "placement": "block",
        "insertAfter": "caption",
        "trigger": "invalid-merge-declaration",
        "recoveryScope": "complete-table",
        "actions": ["discard-all-merges", "preserve-complete-grid"],
    }
    assert "complete unmerged grid" in issues[0].message
    assert table == original


def test_many_invalid_ranges_still_produce_one_stable_issue() -> None:
    table = _table(merges="A1:A2;D2:E9;B2:B2")

    first = resolve_table_policy(table, "academic")
    second = resolve_table_policy(table, "academic")

    assert first == second
    assert len(first[1]) == 1
    assert first[1][0].code == TABLE_MERGE_INVALID


def test_empty_merge_declaration_is_an_unmerged_table_without_an_issue() -> None:
    policy, issues = resolve_table_policy(_table(merges="  "), "academic")

    assert policy.merges == ()
    assert issues == ()


def test_parser_preserves_empty_covered_cells_and_direct_alignments() -> None:
    document = parse_markdown(
        """:::table {#tab:grouped caption="Grouped" merges="a1:b1;A2:A3"}
| Group |  | Value |
|:------|:-:|------:|
| A     | x | 1     |
|       | y | 2     |
:::
""",
        longform=True,
    )
    table = next(
        element
        for section in document.sections
        for element in section.elements
        if isinstance(element, SemanticTableBlock)
    )

    policy, issues = resolve_table_policy(table, "academic")

    assert table.headers == ["Group", "", "Value"]
    assert table.rows == [["A", "x", "1"], ["", "y", "2"]]
    assert table.alignments == ["left", "center", "right"]
    assert policy.merges == (TableMerge(1, 1, 1, 2), TableMerge(2, 1, 3, 1))
    assert issues == ()
