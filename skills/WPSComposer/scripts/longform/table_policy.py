"""Pure table-style and constrained-merge policy for long-form documents."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Sequence

from ..design_presets import PRESETS
from ..document_model import DocumentIssue, TableMerge


TABLE_MERGE_INVALID = "TABLE_MERGE_INVALID"
TABLE_ROW_FORCED_SPLIT = "TABLE_ROW_FORCED_SPLIT"

_TABLE_MERGE_INVALID_MESSAGE = (
    "Table merge declaration is invalid; all merges were discarded and the "
    "complete unmerged grid was preserved."
)
_A1_RANGE_RE = re.compile(
    r"(?P<left>[A-Za-z]+)(?P<top>[1-9][0-9]*):"
    r"(?P<right>[A-Za-z]+)(?P<bottom>[1-9][0-9]*)"
)

_THREE_LINE_BORDERS = {
    "top": 1.5,
    "bottom": 1.5,
    "header_bottom": 0.75,
    "left": 0.0,
    "right": 0.0,
    "inside_horizontal": 0.0,
    "inside_vertical": 0.0,
}
_GRID_BORDERS = {
    "top": 0.75,
    "bottom": 0.75,
    "header_bottom": 0.75,
    "left": 0.75,
    "right": 0.75,
    "inside_horizontal": 0.75,
    "inside_vertical": 0.75,
}


@dataclass(frozen=True)
class TablePolicy:
    """Resolved platform-independent table policy."""

    style: str
    borders: dict[str, float]
    merges: tuple[TableMerge, ...]
    repeat_header: bool
    allow_row_split: bool
    cell_indent_pt: float

    @property
    def indivisible_row_groups(self) -> tuple[tuple[int, int], ...]:
        """Body-row intervals that must stay together to retain vertical merges."""
        intervals = sorted(
            (merge.top, merge.bottom)
            for merge in self.merges
            if merge.top >= 2 and merge.bottom > merge.top
        )
        groups: list[tuple[int, int]] = []
        for top, bottom in intervals:
            if groups and top <= groups[-1][1]:
                groups[-1] = (groups[-1][0], max(groups[-1][1], bottom))
            else:
                groups.append((top, bottom))
        return tuple(groups)


def parse_a1_merge_ranges(merge_spec: str) -> tuple[TableMerge, ...]:
    """Parse semicolon-separated A1 rectangles without applying table semantics."""
    text = str(merge_spec or "").strip()
    if not text:
        return ()

    merges = []
    for declaration in text.split(";"):
        declaration = declaration.strip()
        match = _A1_RANGE_RE.fullmatch(declaration)
        if match is None:
            raise ValueError("invalid A1 merge range")
        left = _column_number(match.group("left"))
        right = _column_number(match.group("right"))
        top = int(match.group("top"))
        bottom = int(match.group("bottom"))
        if right < left or bottom < top:
            raise ValueError("A1 merge range is reversed")
        if right == left and bottom == top:
            raise ValueError("A1 merge range is degenerate")
        merges.append(TableMerge(top, left, bottom, right))
    return tuple(merges)


def resolve_table_policy(
    table: Any,
    preset: Any,
) -> tuple[TablePolicy, tuple[DocumentIssue, ...]]:
    """Resolve style and validate every requested merge as one atomic declaration."""
    style = _resolve_style(getattr(table, "style", ""), preset)
    merges: tuple[TableMerge, ...] = ()
    issues: tuple[DocumentIssue, ...] = ()
    merge_spec = str(getattr(table, "merge_spec", "") or "")
    if merge_spec.strip():
        try:
            requested_merges = parse_a1_merge_ranges(merge_spec)
            _validate_merges(
                requested_merges,
                getattr(table, "headers", ()),
                getattr(table, "rows", ()),
            )
            merges = requested_merges
        except (IndexError, TypeError, ValueError):
            issues = (
                DocumentIssue(
                    code=TABLE_MERGE_INVALID,
                    message=_TABLE_MERGE_INVALID_MESSAGE,
                    placement="block",
                ),
            )

    policy = TablePolicy(
        style=style,
        borders=dict(_THREE_LINE_BORDERS if style == "three-line" else _GRID_BORDERS),
        merges=merges,
        repeat_header=bool(getattr(table, "repeat_header", True)),
        allow_row_split=False,
        cell_indent_pt=0.0,
    )
    return policy, issues


def row_forced_split_degradation(top: int, bottom: int) -> dict[str, Any]:
    """Describe runtime recovery for an over-page body row or merged group."""
    if top < 2:
        raise ValueError("body row group must start at row 2 or later")
    if bottom < top:
        raise ValueError("body row group coordinates must be ordered")
    return {
        "code": TABLE_ROW_FORCED_SPLIT,
        "message": (
            f"Table row group {top}:{bottom} exceeded the available page height; "
            "remove its vertical merges and allow row splitting."
        ),
        "placement": "block",
        "insert_after": "caption",
        "trigger": "indivisible-row-group-exceeds-available-page",
        "recovery": "unmerge-group-and-allow-row-split",
        "row_group": {"top": top, "bottom": bottom},
    }


def _column_number(label: str) -> int:
    value = 0
    for character in label.upper():
        value = value * 26 + ord(character) - ord("A") + 1
    return value


def _preset_key(preset: Any) -> str:
    if isinstance(preset, str):
        return preset.strip().lower()
    for key, candidate in PRESETS.items():
        if preset is candidate:
            return key
    return str(getattr(preset, "name", "") or "").strip().lower()


def _resolve_style(explicit_style: Any, preset: Any) -> str:
    style = str(explicit_style or "").strip().lower()
    if style in {"three-line", "grid"}:
        return style
    return "three-line" if _preset_key(preset) == "academic" else "grid"


def _validate_merges(
    merges: Sequence[TableMerge],
    headers: Sequence[Any],
    rows: Sequence[Sequence[Any]],
) -> None:
    column_count = len(headers)
    row_count = 1 + len(rows)
    if column_count <= 0:
        raise ValueError("table has no header columns")
    if any(len(row) != column_count for row in rows):
        raise ValueError("table grid is not rectangular")

    occupied: set[tuple[int, int]] = set()
    for merge in merges:
        if merge.bottom > row_count or merge.right > column_count:
            raise ValueError("merge is outside the table grid")
        if merge.top == 1 and merge.bottom != 1:
            raise ValueError("merge crosses the repeated-header boundary")

        coordinates = {
            (row, column)
            for row in range(merge.top, merge.bottom + 1)
            for column in range(merge.left, merge.right + 1)
        }
        if occupied.intersection(coordinates):
            raise ValueError("merge ranges overlap")
        occupied.update(coordinates)

        for row, column in sorted(coordinates):
            if row == merge.top and column == merge.left:
                continue
            if str(_cell_value(headers, rows, row, column)).strip():
                raise ValueError("covered merge cell is not empty")


def _cell_value(
    headers: Sequence[Any],
    rows: Sequence[Sequence[Any]],
    row: int,
    column: int,
) -> Any:
    if row == 1:
        return headers[column - 1]
    return rows[row - 2][column - 1]


__all__ = [
    "TABLE_MERGE_INVALID",
    "TABLE_ROW_FORCED_SPLIT",
    "TableMerge",
    "TablePolicy",
    "parse_a1_merge_ranges",
    "resolve_table_policy",
    "row_forced_split_degradation",
]
