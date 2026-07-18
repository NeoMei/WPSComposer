"""Closed, JSON-compatible operation plans for macOS WPS generation."""

from __future__ import annotations

import json
import math
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType
from typing import Any, Mapping


MAX_PLAN_BYTES = 2_000_000
MAX_OPERATIONS = 10_000
MAX_STRING_CHARS = 100_000
MAX_TABLE_CELLS = 10_000

ALLOWED_MEDIA_TYPES = frozenset(
    {
        "image/bmp",
        "image/gif",
        "image/jpeg",
        "image/png",
        "image/tiff",
    }
)

ALLOWED_OPERATIONS = {
    "writer": {
        "writer.reset",
        "writer.configure_page",
        "writer.ensure_styles",
        "writer.add_paragraph",
        "writer.add_heading",
        "writer.add_list",
        "writer.add_table",
        "writer.add_image",
        "writer.add_page_break",
        "writer.add_section",
        "writer.add_horizontal_line",
        "writer.insert_toc",
        "writer.set_page_number",
        "writer.update_fields",
    },
    "spreadsheet": {
        "sheet.reset",
        "sheet.rename",
        "sheet.add",
        "sheet.select",
        "sheet.write_table",
        "sheet.set_column_width",
        "sheet.autofit",
    },
    "presentation": {
        "slide.reset",
        "slide.set_size",
        "slide.apply_preset",
        "slide.add_title",
        "slide.add_section",
        "slide.add_bullets",
        "slide.add_blank",
        "slide.add_image",
        "slide.add_table",
    },
}
ALLOWED_OPERATIONS = MappingProxyType(
    {
        component: frozenset(operations)
        for component, operations in ALLOWED_OPERATIONS.items()
    }
)


class OperationPlanError(ValueError):
    """Raised when a generation operation plan violates the closed protocol."""


def _freeze(value: Any) -> Any:
    if isinstance(value, Mapping):
        return MappingProxyType({key: _freeze(item) for key, item in value.items()})
    if isinstance(value, (list, tuple)):
        return tuple(_freeze(item) for item in value)
    return value


def _thaw(value: Any) -> Any:
    if isinstance(value, Mapping):
        return {key: _thaw(item) for key, item in value.items()}
    if isinstance(value, tuple):
        return [_thaw(item) for item in value]
    return value


@dataclass(frozen=True)
class GenerationOperation:
    op: str
    args: Mapping[str, Any]

    def __post_init__(self) -> None:
        object.__setattr__(self, "args", _freeze(dict(self.args)))

    def to_dict(self) -> dict[str, Any]:
        return {"op": self.op, "args": _thaw(self.args)}


@dataclass(frozen=True)
class GenerationPlan:
    component: str
    operations: tuple[GenerationOperation, ...]

    def __post_init__(self) -> None:
        object.__setattr__(self, "operations", tuple(self.operations))

    def to_dict(self) -> dict[str, Any]:
        return {
            "component": self.component,
            "operations": [operation.to_dict() for operation in self.operations],
        }


@dataclass(frozen=True)
class GenerationResource:
    id: str
    source_path: Path
    media_type: str

    def __post_init__(self) -> None:
        if not isinstance(self.id, str) or not self.id:
            raise OperationPlanError("resource id is required")
        if "/" in self.id or "\\" in self.id:
            raise OperationPlanError("resource id must be a logical identifier")
        if self.media_type not in ALLOWED_MEDIA_TYPES:
            raise OperationPlanError(f"unsupported media type: {self.media_type}")
        try:
            resolved = Path(self.source_path).expanduser().resolve()
        except (TypeError, ValueError) as error:
            raise OperationPlanError(
                "resource source_path must be a host path"
            ) from error
        object.__setattr__(self, "source_path", resolved)


@dataclass(frozen=True)
class RecordedGeneration:
    plan: GenerationPlan
    resources: tuple[GenerationResource, ...]

    def __post_init__(self) -> None:
        object.__setattr__(self, "resources", tuple(self.resources))


def _validate_json_value(value: Any) -> None:
    pending = [value]
    seen_containers = set()
    while pending:
        item = pending.pop()
        if isinstance(item, str):
            if len(item) > MAX_STRING_CHARS:
                raise OperationPlanError("string exceeds 100,000 characters")
        elif item is None or isinstance(item, (bool, int)):
            continue
        elif isinstance(item, float):
            if not math.isfinite(item):
                raise OperationPlanError(
                    "operation arguments must be JSON-compatible"
                )
        elif isinstance(item, (list, dict)):
            identity = id(item)
            if identity in seen_containers:
                continue
            seen_containers.add(identity)
            if isinstance(item, list):
                pending.extend(item)
            else:
                for key, nested in item.items():
                    if not isinstance(key, str):
                        raise OperationPlanError(
                            "operation arguments must be JSON-compatible"
                        )
                    pending.extend((key, nested))
        else:
            raise OperationPlanError(
                "operation arguments must be JSON-compatible"
            )


def _table_cell_count(args: Mapping[str, Any]) -> int:
    counts = []
    for key in ("data", "values"):
        data = args.get(key)
        if isinstance(data, list):
            counts.append(
                sum(len(row) if isinstance(row, list) else 1 for row in data)
            )
    rows = args.get("rows")
    cols = args.get("cols")
    if (
        isinstance(rows, int)
        and not isinstance(rows, bool)
        and isinstance(cols, int)
        and not isinstance(cols, bool)
        and rows >= 0
        and cols >= 0
    ):
        counts.append(rows * cols)
    return max(counts, default=0)


def _validate_image_args(op: str, args: Mapping[str, Any]) -> None:
    if not op.endswith(".add_image"):
        return
    image_id = args.get("imageId")
    if not isinstance(image_id, str) or not image_id:
        raise OperationPlanError(f"{op} requires imageId")
    if "/" in image_id or "\\" in image_id:
        raise OperationPlanError("imageId must be a logical identifier")
    if _contains_path_key(args):
        raise OperationPlanError("image paths are not allowed; use imageId")


def _contains_path_key(value: Any) -> bool:
    if isinstance(value, Mapping):
        for key, item in value.items():
            normalized = key.replace("_", "").replace("-", "").lower()
            if normalized.endswith("path") or _contains_path_key(item):
                return True
    elif isinstance(value, (list, tuple)):
        return any(_contains_path_key(item) for item in value)
    return False


def _serialized_plan(raw: Mapping[str, Any]) -> bytes:
    try:
        return json.dumps(
            raw,
            allow_nan=False,
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
    except (RecursionError, TypeError, ValueError) as error:
        raise OperationPlanError(
            "operation arguments must be JSON-compatible"
        ) from error


def validate_generation_plan(
    raw: Mapping[str, Any], component: str
) -> GenerationPlan:
    """Parse and defensively copy a plan that satisfies the closed protocol."""

    if not isinstance(raw, Mapping):
        raise OperationPlanError("generation plan must be an object")
    if set(raw) != {"component", "operations"}:
        raise OperationPlanError(
            "generation plan must contain exactly component and operations"
        )
    if component not in ALLOWED_OPERATIONS:
        raise OperationPlanError(f"unsupported component: {component}")
    if raw.get("component") != component:
        raise OperationPlanError(
            f"component mismatch: expected {component}, got {raw.get('component')}"
        )
    operations = raw.get("operations")
    if not isinstance(operations, list):
        raise OperationPlanError("operations must be a list")
    if not operations:
        raise OperationPlanError("generation plan requires at least one operation")
    if len(operations) > MAX_OPERATIONS:
        raise OperationPlanError("generation plan has too many operations")

    _validate_json_value(dict(raw))
    serialized = _serialized_plan(raw)
    if len(serialized) > MAX_PLAN_BYTES:
        raise OperationPlanError("generation plan exceeds 2,000,000 bytes")

    normalized = json.loads(serialized.decode("utf-8"))

    parsed = []
    allowed = ALLOWED_OPERATIONS[component]
    for operation in normalized["operations"]:
        if not isinstance(operation, dict):
            raise OperationPlanError("operation must be an object")
        if set(operation) != {"op", "args"}:
            raise OperationPlanError("operation must contain exactly op and args")
        op = operation.get("op")
        if not isinstance(op, str):
            raise OperationPlanError("operation op must be a string")
        if op not in allowed:
            raise OperationPlanError(f"unsupported operation: {op}")
        args = operation.get("args")
        if not isinstance(args, dict):
            raise OperationPlanError("operation args must be an object")
        if op.endswith(".add_table") or op == "sheet.write_table":
            if _table_cell_count(args) > MAX_TABLE_CELLS:
                raise OperationPlanError("table exceeds 10,000 cells")
        _validate_image_args(op, args)
        parsed.append(GenerationOperation(op, dict(args)))
    return GenerationPlan(component, tuple(parsed))
