"""Pure table-style and constrained-merge policy for long-form documents."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, ClassVar, Iterator, Mapping, Optional, Sequence

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

@dataclass(frozen=True, eq=False)
class TableBorders(Mapping[str, float]):
    """Fixed, deeply immutable border values with mapping-style lookup."""

    top: float
    bottom: float
    header_bottom: float
    left: float
    right: float
    inside_horizontal: float
    inside_vertical: float

    _KEYS: ClassVar[tuple[str, ...]] = (
        "top",
        "bottom",
        "header_bottom",
        "left",
        "right",
        "inside_horizontal",
        "inside_vertical",
    )

    @classmethod
    def from_mapping(cls, values: Mapping[str, float]) -> "TableBorders":
        if isinstance(values, cls):
            return values
        return cls(*(float(values.get(key, 0.0)) for key in cls._KEYS))

    def __getitem__(self, key: str) -> float:
        if key not in self._KEYS:
            raise KeyError(key)
        return float(getattr(self, key))

    def __iter__(self) -> Iterator[str]:
        return iter(self._KEYS)

    def __len__(self) -> int:
        return len(self._KEYS)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, Mapping):
            return False
        return self.to_dict() == dict(other.items())

    def __hash__(self) -> int:
        return hash(tuple(self.items()))

    def to_dict(self) -> dict[str, float]:
        return {key: self[key] for key in self._KEYS}


_THREE_LINE_BORDERS = TableBorders(1.5, 1.5, 0.75, 0.0, 0.0, 0.0, 0.0)
_GRID_BORDERS = TableBorders(0.75, 0.75, 0.75, 0.75, 0.75, 0.75, 0.75)


@dataclass(frozen=True)
class TableDocumentIssue(DocumentIssue):
    """A DocumentIssue with immutable table-local recovery metadata."""

    trigger: str = ""
    recovery_scope: str = ""
    actions: tuple[str, ...] = ()
    row_group: Optional[tuple[int, int]] = None
    placement: str = "block"
    insert_after: str = "caption"

    def __post_init__(self) -> None:
        object.__setattr__(self, "actions", tuple(self.actions))
        if self.row_group is not None:
            group = tuple(self.row_group)
            if (
                len(group) != 2
                or any(type(value) is not int for value in group)
                or group[0] < 2
                or group[1] < group[0]
            ):
                raise ValueError("table degradation row group is invalid")
            object.__setattr__(self, "row_group", group)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "code": self.code,
            "message": self.message,
            "placement": self.placement,
            "insertAfter": self.insert_after,
            "trigger": self.trigger,
            "recoveryScope": self.recovery_scope,
            "actions": list(self.actions),
        }
        if self.row_group is not None:
            result["rowGroup"] = {
                "top": self.row_group[0],
                "bottom": self.row_group[1],
            }
        return result


# Backward-compatible name from the first Task 3 review; this is now a
# DocumentIssue subtype rather than an independent metadata object.
TableDegradationMetadata = TableDocumentIssue


@dataclass(frozen=True)
class TablePolicy:
    """Resolved platform-independent table policy."""

    style: str
    borders: TableBorders
    merges: tuple[TableMerge, ...]
    repeat_header: bool
    allow_row_split: bool
    cell_indent_pt: float

    def __post_init__(self) -> None:
        object.__setattr__(self, "borders", TableBorders.from_mapping(self.borders))
        object.__setattr__(self, "merges", tuple(self.merges))

    def to_dict(self) -> dict[str, Any]:
        """Return a detached, deterministic JSON-ready policy mapping."""
        return {
            "style": self.style,
            "borders": dict(sorted(self.borders.to_dict().items())),
            "merges": [
                {
                    "top": merge.top,
                    "left": merge.left,
                    "bottom": merge.bottom,
                    "right": merge.right,
                }
                for merge in self.merges
            ],
            "repeatHeader": self.repeat_header,
            "allowRowSplit": self.allow_row_split,
            "cellIndentPt": self.cell_indent_pt,
        }

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
                TableDocumentIssue(
                    code=TABLE_MERGE_INVALID,
                    message=_TABLE_MERGE_INVALID_MESSAGE,
                    trigger="invalid-merge-declaration",
                    recovery_scope="complete-table",
                    actions=("discard-all-merges", "preserve-complete-grid"),
                ),
            )

    policy = TablePolicy(
        style=style,
        borders=_THREE_LINE_BORDERS if style == "three-line" else _GRID_BORDERS,
        merges=merges,
        repeat_header=bool(getattr(table, "repeat_header", True)),
        allow_row_split=False,
        cell_indent_pt=0.0,
    )
    return policy, issues


def row_forced_split_degradation(
    top: int,
    bottom: int,
) -> TableDocumentIssue:
    """Describe runtime recovery for an over-page body row or merged group."""
    if top < 2:
        raise ValueError("body row group must start at row 2 or later")
    if bottom < top:
        raise ValueError("body row group coordinates must be ordered")
    if top == bottom:
        return TableDocumentIssue(
            code=TABLE_ROW_FORCED_SPLIT,
            message=(
                f"Table row {top} exceeded the available page height; allow that "
                "row to split without removing table merges."
            ),
            trigger="row-exceeds-available-page",
            recovery_scope="row",
            actions=("allow-row-split",),
            row_group=(top, bottom),
        )
    return TableDocumentIssue(
        code=TABLE_ROW_FORCED_SPLIT,
        message=(
            f"Table row group {top}:{bottom} exceeded the available page height; "
            "render the complete table as an unmerged splittable grid."
        ),
        trigger="vertical-merge-group-exceeds-available-page",
        recovery_scope="complete-table",
        actions=("discard-all-merges", "apply-grid-style", "allow-row-split"),
        row_group=(top, bottom),
    )


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
    "TableBorders",
    "TableDegradationMetadata",
    "TableDocumentIssue",
    "TableMerge",
    "TablePolicy",
    "parse_a1_merge_ranges",
    "resolve_table_policy",
    "row_forced_split_degradation",
]
