"""Closed, JSON-compatible operation plans for macOS WPS generation."""

from __future__ import annotations

import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType
from typing import Any, Callable, Mapping, Optional


MAX_PLAN_BYTES = 2_000_000
MAX_OPERATIONS = 10_000
MAX_STRING_CHARS = 100_000
MAX_TABLE_CELLS = 10_000
MAX_NESTING_DEPTH = 64
MAX_SAFE_NUMBER = 2**53 - 1

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
    pending = [(value, 0)]
    seen_containers = set()
    while pending:
        item, depth = pending.pop()
        if depth > MAX_NESTING_DEPTH:
            raise OperationPlanError("JSON nesting exceeds 64 levels")
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
                pending.extend((nested, depth + 1) for nested in item)
            else:
                for key, nested in item.items():
                    if not isinstance(key, str):
                        raise OperationPlanError(
                            "operation arguments must be JSON-compatible"
                        )
                    pending.extend(((key, depth + 1), (nested, depth + 1)))
        else:
            raise OperationPlanError(
                "operation arguments must be JSON-compatible"
            )


ArgumentValidator = Callable[[Any, str], None]


@dataclass(frozen=True)
class _ObjectSchema:
    required: frozenset[str]
    fields: Mapping[str, ArgumentValidator]


def _schema(
    required: tuple[str, ...] = (),
    **fields: ArgumentValidator,
) -> _ObjectSchema:
    return _ObjectSchema(frozenset(required), MappingProxyType(dict(fields)))


def _invalid(path: str, expected: str) -> None:
    raise OperationPlanError(f"invalid argument {path}: expected {expected}")


def _string(value: Any, path: str) -> None:
    if not isinstance(value, str):
        _invalid(path, "string")


_COLOR_RE = re.compile(r"^#[0-9A-Fa-f]{6}$")


def _color(value: Any, path: str) -> None:
    if not isinstance(value, str) or not _COLOR_RE.match(value):
        _invalid(path, "#RRGGBB color string")


def _color_or_false(value: Any, path: str) -> None:
    if value is False:
        return
    _color(value, path)


def _nullable_string(value: Any, path: str) -> None:
    if value is not None:
        _string(value, path)


def _boolean(value: Any, path: str) -> None:
    if not isinstance(value, bool):
        _invalid(path, "boolean")


def _integer(value: Any, path: str) -> None:
    if not isinstance(value, int) or isinstance(value, bool):
        _invalid(path, "integer")
    if not -MAX_SAFE_NUMBER <= value <= MAX_SAFE_NUMBER:
        _invalid(path, "integer in the interoperable safe range")


def _number(value: Any, path: str) -> None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        _invalid(path, "finite number in the interoperable safe range")
    if isinstance(value, int):
        if not -MAX_SAFE_NUMBER <= value <= MAX_SAFE_NUMBER:
            _invalid(path, "number in the interoperable safe range")
        return
    if not math.isfinite(value) or abs(value) > MAX_SAFE_NUMBER:
        _invalid(path, "finite number in the interoperable safe range")


def _bounded_integer(minimum: int, maximum: Optional[int] = None) -> ArgumentValidator:
    def validate(value: Any, path: str) -> None:
        _integer(value, path)
        if value < minimum or (maximum is not None and value > maximum):
            expected = f"integer >= {minimum}"
            if maximum is not None:
                expected = f"integer in [{minimum}, {maximum}]"
            _invalid(path, expected)

    return validate


def _bounded_number(minimum: float, exclusive: bool = False) -> ArgumentValidator:
    def validate(value: Any, path: str) -> None:
        _number(value, path)
        if (exclusive and value <= minimum) or (not exclusive and value < minimum):
            op = ">" if exclusive else ">="
            _invalid(path, f"number {op} {minimum}")

    return validate


_POSITIVE_INT = _bounded_integer(1)
_POSITIVE_NUMBER = _bounded_number(0, exclusive=True)
_NONNEGATIVE_NUMBER = _bounded_number(0)


def _cell(value: Any, path: str) -> None:
    if value is None or isinstance(value, (str, bool)):
        return
    if isinstance(value, int):
        if -MAX_SAFE_NUMBER <= value <= MAX_SAFE_NUMBER:
            return
        _invalid(path, "integer in the interoperable safe range")
    if (
        isinstance(value, float)
        and math.isfinite(value)
        and abs(value) <= MAX_SAFE_NUMBER
    ):
        return
    _invalid(
        path,
        "string, safe-range finite number, boolean, or null",
    )


def _validate_object(value: Any, path: str, schema: _ObjectSchema) -> None:
    if not isinstance(value, dict):
        _invalid(path, "object")
    unknown = set(value) - set(schema.fields)
    if unknown:
        names = ", ".join(sorted(unknown))
        raise OperationPlanError(f"unknown argument in {path}: {names}")
    missing = schema.required - set(value)
    if missing:
        names = ", ".join(sorted(missing))
        raise OperationPlanError(f"missing required argument in {path}: {names}")
    for name, item in value.items():
        schema.fields[name](item, f"{path}.{name}")


def _list_of(validator: ArgumentValidator) -> ArgumentValidator:
    def validate(value: Any, path: str) -> None:
        if not isinstance(value, list):
            _invalid(path, "array")
        for index, item in enumerate(value):
            validator(item, f"{path}[{index}]")

    return validate


_STRING_LIST = _list_of(_string)
_NUMBER_LIST = _list_of(_number)

_SPAN_SCHEMA = _schema(
    ("text",),
    text=_string,
    bold=_boolean,
    italic=_boolean,
    strikethrough=_boolean,
    code=_boolean,
    link=_nullable_string,
    linkTitle=_nullable_string,
)


def _span(value: Any, path: str) -> None:
    _validate_object(value, path, _SPAN_SCHEMA)


_STYLE_SCHEMA = _schema(
    ("name",),
    name=_string,
    type=_string,
    basedOn=_string,
    fontName=_string,
    fontNameAscii=_string,
    fontSize=_POSITIVE_NUMBER,
    bold=_boolean,
    italic=_boolean,
    underline=_boolean,
    strikethrough=_boolean,
    color=_color,
    align=_integer,
    indentFirst=_number,
    leftIndent=_number,
    rightIndent=_number,
    lineSpacing=_number,
    lineSpacingRule=_string,
    spaceBefore=_number,
    spaceAfter=_number,
    shading=_string,
    leftBorder=_boolean,
    borderColor=_color,
    keepWithNext=_boolean,
    outlineLevel=_integer,
)


def _style(value: Any, path: str) -> None:
    _validate_object(value, path, _STYLE_SCHEMA)


def _table(value: Any, path: str) -> None:
    if not isinstance(value, list):
        _invalid(path, "array of rows")
    width = None
    for row_index, row in enumerate(value):
        if not isinstance(row, list):
            _invalid(f"{path}[{row_index}]", "array of cells")
        if width is None:
            width = len(row)
        elif len(row) != width:
            _invalid(path, "rectangular table")
        for col_index, item in enumerate(row):
            _cell(item, f"{path}[{row_index}][{col_index}]")


_FONT_SCHEMA = _schema(
    ("family", "size", "color"),
    family=_string,
    size=_POSITIVE_NUMBER,
    color=_color,
)


def _font(value: Any, path: str) -> None:
    _validate_object(value, path, _FONT_SCHEMA)


_SPACING_SCHEMA = _schema(
    ("margin", "gap", "cardPadding", "lineHeight"),
    margin=_number,
    gap=_number,
    cardPadding=_number,
    lineHeight=_number,
)
def _role_map(required: frozenset, item_validator: ArgumentValidator) -> ArgumentValidator:
    """Preset role map: required roles must exist, extra custom roles pass."""

    def validate(value: Any, path: str) -> None:
        if not isinstance(value, dict):
            _invalid(path, "object")
        missing = required - set(value)
        if missing:
            names = ", ".join(sorted(missing))
            raise OperationPlanError(
                f"invalid argument {path}: missing roles: {names}"
            )
        for name, item in value.items():
            item_validator(item, f"{path}.{name}")

    return validate


_PRESET_SCHEMA = _schema(
    ("name", "colors", "fonts"),
    name=_string,
    colors=_role_map(frozenset({"primary", "dark", "background"}), _color),
    fonts=_role_map(frozenset({"title", "body"}), _font),
    spacing=lambda value, path: _validate_object(value, path, _SPACING_SCHEMA),
)


def _preset(value: Any, path: str) -> None:
    _validate_object(value, path, _PRESET_SCHEMA)


_OPERATION_ARG_SCHEMAS = MappingProxyType(
    {
        "writer.reset": _schema(),
        "writer.configure_page": _schema(
            ("marginTop", "marginBottom", "marginLeft", "marginRight"),
            marginTop=_NONNEGATIVE_NUMBER,
            marginBottom=_NONNEGATIVE_NUMBER,
            marginLeft=_NONNEGATIVE_NUMBER,
            marginRight=_NONNEGATIVE_NUMBER,
            pageWidth=_POSITIVE_NUMBER,
            pageHeight=_POSITIVE_NUMBER,
            landscape=_boolean,
            columns=_POSITIVE_INT,
            header=_string,
            footer=_string,
        ),
        "writer.ensure_styles": _schema(
            ("styles",), styles=_list_of(_style)
        ),
        "writer.add_paragraph": _schema(
            ("text",),
            text=_string,
            style=_string,
            spans=_list_of(_span),
            size=_POSITIVE_NUMBER,
            bold=_boolean,
            italic=_boolean,
            color=_color,
            align=_integer,
            indentFirst=_number,
            lineSpacing=_number,
            lineSpacingRule=_string,
            spaceBefore=_number,
            spaceAfter=_number,
            fontName=_string,
            fontNameAscii=_string,
        ),
        "writer.add_heading": _schema(
            ("text", "level"),
            text=_string,
            level=_bounded_integer(1, 6),
            size=_POSITIVE_NUMBER,
            bold=_boolean,
            color=_color,
            lineSpacing=_number,
            lineSpacingRule=_string,
            spaceAfter=_number,
        ),
        "writer.add_list": _schema(
            ("items", "ordered"),
            items=_STRING_LIST,
            ordered=_boolean,
            glyph=_string,
            indent=_number,
        ),
        "writer.add_table": _schema(
            ("rows", "cols", "data"),
            rows=_POSITIVE_INT,
            cols=_POSITIVE_INT,
            data=_table,
            shadeHeader=_color,
            headerColor=_color,
            fontSize=_POSITIVE_NUMBER,
            columnWidths=_NUMBER_LIST,
            alignments=_STRING_LIST,
            bandedRows=_boolean,
            autoFit=_boolean,
            repeatHeader=_boolean,
            borderColor=_color,
        ),
        "writer.add_image": _schema(
            ("imageId",),
            imageId=_string,
            width=_POSITIVE_NUMBER,
            height=_POSITIVE_NUMBER,
            maxWidth=_POSITIVE_NUMBER,
            maxHeight=_POSITIVE_NUMBER,
            wrap=_integer,
            inline=_boolean,
            preserveAspect=_boolean,
            alt=_nullable_string,
        ),
        "writer.add_page_break": _schema(),
        "writer.add_section": _schema(),
        "writer.add_horizontal_line": _schema(),
        "writer.insert_toc": _schema(("title",), title=_string),
        "writer.set_page_number": _schema(),
        "writer.update_fields": _schema(),
        "sheet.reset": _schema(),
        "sheet.rename": _schema(("index", "name"), index=_POSITIVE_INT, name=_string),
        "sheet.add": _schema(("name",), name=_string),
        "sheet.select": _schema(("index",), index=_POSITIVE_INT),
        "sheet.write_table": _schema(
            ("startRow", "startCol", "values"),
            startRow=_POSITIVE_INT,
            startCol=_POSITIVE_INT,
            values=_table,
            headerBold=_boolean,
            headerShade=_color_or_false,
            headerFontColor=_color,
            fontSize=_POSITIVE_NUMBER,
        ),
        "sheet.set_column_width": _schema(
            ("column", "width"), column=_string, width=_POSITIVE_NUMBER
        ),
        "sheet.autofit": _schema(),
        "slide.reset": _schema(),
        "slide.set_size": _schema(
            ("width", "height"), width=_POSITIVE_NUMBER, height=_POSITIVE_NUMBER
        ),
        "slide.apply_preset": _schema(("preset",), preset=_preset),
        "slide.add_title": _schema(
            ("title",),
            title=_string,
            subtitle=_nullable_string,
            titleSize=_POSITIVE_NUMBER,
            subtitleSize=_POSITIVE_NUMBER,
            titleColor=_color,
        ),
        "slide.add_section": _schema(("title",), title=_string),
        "slide.add_bullets": _schema(
            ("title", "items"),
            title=_string,
            items=_STRING_LIST,
            titleSize=_POSITIVE_NUMBER,
            bodySize=_POSITIVE_NUMBER,
        ),
        "slide.add_blank": _schema(),
        "slide.add_image": _schema(
            ("slide", "imageId", "left", "top"),
            slide=_POSITIVE_INT,
            imageId=_string,
            left=_NONNEGATIVE_NUMBER,
            top=_NONNEGATIVE_NUMBER,
            width=_POSITIVE_NUMBER,
            height=_POSITIVE_NUMBER,
        ),
        "slide.add_table": _schema(
            ("slide", "rows", "cols", "left", "top", "width", "height", "data"),
            slide=_POSITIVE_INT,
            rows=_POSITIVE_INT,
            cols=_POSITIVE_INT,
            left=_NONNEGATIVE_NUMBER,
            top=_NONNEGATIVE_NUMBER,
            width=_POSITIVE_NUMBER,
            height=_POSITIVE_NUMBER,
            data=_table,
            headerShade=_color,
            headerFont=_color,
            fontSize=_POSITIVE_NUMBER,
        ),
    }
)


def _validate_operation_args(op: str, args: Mapping[str, Any]) -> None:
    schema = _OPERATION_ARG_SCHEMAS[op]
    _validate_object(args, f"{op}.args", schema)
    _validate_table_shape(op, args)


def _validate_table_shape(op: str, args: Mapping[str, Any]) -> None:
    """Cross-check grid data against declared rows/cols (add-ins reject
    mismatches; fail here before booting WPS)."""
    data_key = "data" if "data" in args else "values" if "values" in args else None
    if data_key is None:
        return
    data = args[data_key]
    if not isinstance(data, list) or not data:
        if data_key == "values" or "rows" in args:
            raise OperationPlanError(
                f"invalid argument {op}.args.{data_key}: expected a non-empty array"
            )
        return
    rows = args.get("rows")
    cols = args.get("cols")
    if rows is None and cols is None:
        for index, row in enumerate(data):
            if not isinstance(row, list) or not row:
                raise OperationPlanError(
                    f"invalid argument {op}.args.{data_key}[{index}]: "
                    "expected a non-empty array"
                )
        return
    if rows is not None and len(data) != rows:
        raise OperationPlanError(
            f"invalid argument {op}.args.{data_key}: {len(data)} rows, expected {rows}"
        )
    if cols is not None:
        for index, row in enumerate(data):
            width = len(row) if isinstance(row, list) else 1
            if width != cols:
                raise OperationPlanError(
                    f"invalid argument {op}.args.{data_key}[{index}]: "
                    f"{width} cells, expected {cols}"
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
        _validate_operation_args(op, args)
        parsed.append(GenerationOperation(op, dict(args)))
    return GenerationPlan(component, tuple(parsed))
