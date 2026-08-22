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
        "image/svg+xml",
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

PROTOCOL_VERSION_V2 = 2
SEMANTIC_VERSION_LONGFORM = "longform-1"
RESOURCE_MANIFEST_VERSION = 1

_V2_ENVELOPE_KEYS = frozenset({
    "component",
    "operations",
    "protocolVersion",
    "semanticVersion",
    "resourceManifestVersion",
    "resourceManifestDigest",
})

_LONGFORM_WRITER_OPERATIONS = frozenset(
    {
        "writer.configure_front_matter",
        "writer.configure_section",
        "writer.configure_toc_styles",
        "writer.add_page_break",
        "writer.add_captioned_figure",
        "writer.add_semantic_table",
        "writer.add_equation",
        "writer.add_cross_reference",
        "writer.insert_figure_index",
        "writer.insert_table_index",
        "writer.add_bibliography",
        "writer.add_inline_degradation",
        "writer.add_degradation_notice",
        "writer.add_document_quality_notice",
        "writer.finalize_fields",
        "writer.set_page_role",
        "writer.set_page_numbering",
        "writer.set_header_footer",
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
    node_id: Optional[str] = None
    failure_policy: Optional[Mapping[str, Any]] = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "args", _freeze(dict(self.args)))
        if self.failure_policy is not None:
            object.__setattr__(
                self, "failure_policy", _freeze(dict(self.failure_policy))
            )

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {"op": self.op, "args": _thaw(self.args)}
        if self.node_id is not None:
            result["nodeId"] = self.node_id
        if self.failure_policy is not None:
            result["failurePolicy"] = _thaw(self.failure_policy)
        return result


@dataclass(frozen=True)
class GenerationPlan:
    component: str
    operations: tuple[GenerationOperation, ...]
    protocol_version: Optional[int] = None
    semantic_version: Optional[str] = None
    resource_manifest_version: Optional[int] = None
    resource_manifest_digest: Optional[str] = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "operations", tuple(self.operations))

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "component": self.component,
            "operations": [operation.to_dict() for operation in self.operations],
        }
        if self.protocol_version is not None:
            result["protocolVersion"] = self.protocol_version
        if self.semantic_version is not None:
            result["semanticVersion"] = self.semantic_version
        if self.resource_manifest_version is not None:
            result["resourceManifestVersion"] = self.resource_manifest_version
        if self.resource_manifest_digest is not None:
            result["resourceManifestDigest"] = self.resource_manifest_digest
        return result


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
    try:
        utf16_units = len(value.encode("utf-16-le")) // 2
    except UnicodeEncodeError as error:
        raise OperationPlanError(f"invalid argument {path}: invalid Unicode string") from error
    if utf16_units > MAX_STRING_CHARS:
        _invalid(path, "string no longer than 100,000 UTF-16 code units")


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

_PAGE_ROLES = frozenset({"cover", "front_matter", "body", "landscape"})
_PAGE_NUMBER_FORMATS = frozenset({"none", "roman", "arabic", "continue"})
_NUMBERING_SCHEMES = frozenset({"none", "chinese-formal", "decimal", "hybrid-bid"})
_TOC_DENSITY_LEVELS = frozenset({"toc1", "toc2", "toc3"})
_TOC_DENSITY_BOUNDS: dict[str, dict[str, float]] = {
    "minFontSizePt": {"toc1": 10.5, "toc2": 10.0, "toc3": 10.0},
    "minSpaceBeforePt": {"toc1": 0.0, "toc2": 0.0, "toc3": 0.0},
    "minSpaceAfterPt": {"toc1": 0.0, "toc2": 0.0, "toc3": 0.0},
}


def _enum(values: frozenset[str], name: str) -> ArgumentValidator:
    def validate(value: Any, path: str) -> None:
        _string(value, path)
        if value not in values:
            _invalid(path, name)
    return validate


def _toc_density_map(bounds: dict[str, float]) -> ArgumentValidator:
    def validate(value: Any, path: str) -> None:
        if not isinstance(value, dict):
            _invalid(path, "object")
        unknown = set(value) - _TOC_DENSITY_LEVELS
        if unknown:
            names = ", ".join(sorted(unknown))
            raise OperationPlanError(f"unknown argument in {path}: {names}")
        missing = _TOC_DENSITY_LEVELS - set(value)
        if missing:
            names = ", ".join(sorted(missing))
            raise OperationPlanError(f"missing required argument in {path}: {names}")
        for level, minimum in bounds.items():
            item_path = f"{path}.{level}"
            number = value[level]
            _number(number, item_path)
            if number < minimum:
                _invalid(item_path, f"number >= {minimum}")
    return validate


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
    keepTogether=_boolean,
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
            numbering=_boolean,
            numberingScheme=_enum(_NUMBERING_SCHEMES, "numbering scheme"),
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
        "writer.add_section": _schema(landscape=_boolean),
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



_FAILURE_POLICY_SCHEMA = _schema(
    ("mode",),
    mode=_string,
    recoverableCodes=_STRING_LIST,
    fallback=_string,
)


def _failure_policy(value: Any, path: str) -> None:
    _validate_object(value, path, _FAILURE_POLICY_SCHEMA)
    if value["mode"] not in ("fail", "degrade"):
        _invalid(f"{path}.mode", "fail or degrade")
    if value["mode"] == "degrade":
        if not value.get("recoverableCodes") or not value.get("fallback"):
            raise OperationPlanError(
                f"{path}: degrade requires recoverableCodes and fallback"
            )


_M3_FAILURE_POLICIES = {
    "writer.add_captioned_figure": (
        ("IMAGE_INSERT_FAILED",),
        "figure-child-stack-then-notice",
    ),
    "writer.add_semantic_table": (
        (
            "TABLE_STYLE_APPLY_FAILED",
            "TABLE_MERGE_APPLY_FAILED",
            "TABLE_ROW_FORCED_SPLIT",
            "TABLE_INSERT_FAILED",
        ),
        "grid-then-text",
    ),
    "writer.add_cross_reference": (
        ("CROSS_REFERENCE_FAILED",),
        "inline-fallback",
    ),
}


def _validate_m3_failure_policy(
    op: str,
    value: Optional[Mapping[str, Any]],
    args: Mapping[str, Any],
) -> None:
    if not _is_m3_shape(op, args):
        return
    expected = _M3_FAILURE_POLICIES.get(op)
    if expected is not None:
        if value is None:
            raise OperationPlanError(f"{op}.failurePolicy is required")
        codes, fallback = expected
        if (
            value.get("mode") != "degrade"
            or tuple(value.get("recoverableCodes", ())) != codes
            or value.get("fallback") != fallback
        ):
            _invalid(f"{op}.failurePolicy", "exact M3 recovery allowlist and fallback")
    elif op in {
        "writer.add_equation",
        "writer.insert_figure_index",
        "writer.insert_table_index",
        "writer.finalize_fields",
    } and value is not None and value.get("mode") != "fail":
        _invalid(f"{op}.failurePolicy", "fatal field/index policy")


_NOTICE_ITEM_SCHEMA = _schema(
    ("code", "message", "fallbackText", "placement"),
    code=_string,
    message=_string,
    fallbackText=_string,
    placement=_string,
)


def _notice_item(value: Any, path: str) -> None:
    _validate_object(value, path, _NOTICE_ITEM_SCHEMA)


_PLANNED_DEGRADATION_SCHEMA = _schema(
    ("code", "message", "fallback", "placement"),
    code=_string,
    message=_string,
    fallback=_string,
    placement=_string,
)


def _planned_degradation(value: Any, path: str) -> None:
    _validate_object(value, path, _PLANNED_DEGRADATION_SCHEMA)
    if value["code"] not in {
        "RESOURCE_NOT_FOUND",
        "RESOURCE_PATH_ESCAPES_BASE",
        "RESOURCE_ABSOLUTE_PATH_OUTSIDE",
        "RESOURCE_MEDIA_TYPE_UNSUPPORTED",
        "RESOURCE_READ_FAILED",
        "RESOURCE_DECODE_FAILED",
        "RESOURCE_TOO_LARGE",
        "RESOURCE_PIXEL_LIMIT_EXCEEDED",
        "RESOURCE_SIDE_LIMIT_EXCEEDED",
    }:
        _invalid(f"{path}.code", "controlled M3 resource degradation code")
    if value["placement"] != "block":
        _invalid(f"{path}.placement", "block")

_BOOKMARK_RE = re.compile(r"^wpsc_(fig|tab|eq|ref|head|para)_[0-9a-f]{24}$")


def _bookmark(value: Any, path: str) -> None:
    if not isinstance(value, str) or _BOOKMARK_RE.fullmatch(value) is None:
        _invalid(path, "generated WPSComposer bookmark name")


def _logical_id(value: Any, path: str) -> None:
    _string(value, path)
    if not value or "/" in value or "\\" in value:
        _invalid(path, "logical identifier")


_CAPTION_MODES = frozenset({"global", "chapter"})
_SEQUENCE_IDS = frozenset({"WPSC_FIG", "WPSC_TAB", "WPSC_EQ"})
_OBJECT_KINDS = frozenset({"figure", "table", "equation"})
_ORIENTATIONS = frozenset({"portrait", "landscape"})
_FIGURE_LAYOUTS = frozenset({"stack", "columns"})
_FIGURE_KINDS = frozenset({"auto", "photo", "scan", "screenshot", "diagram"})
_WIDTH_MODES = frozenset({"auto", "column", "full", "explicit"})
_TABLE_STYLES = frozenset({"three-line", "grid"})
_ALIGNMENTS = frozenset({"left", "center", "right"})
_NORMALIZER_IDS = frozenset({
    "none-v1", "exif-transpose-png-v1", "gif-first-frame-png-v1",
    "tiff-first-page-png-v1", "svg-static-v1",
})


_NUMBERING_SCHEMA = _schema(
    ("mode", "sequenceId", "chapterStyleLevel", "resetLevel", "prefix", "suffix"),
    mode=_enum(_CAPTION_MODES, "caption numbering mode"),
    sequenceId=_enum(_SEQUENCE_IDS, "controlled sequence identifier"),
    chapterStyleLevel=lambda value, path: _integer(value, path) if value is not None else None,
    resetLevel=lambda value, path: _integer(value, path) if value is not None else None,
    prefix=_string,
    suffix=_string,
)


def _numbering(value: Any, path: str) -> None:
    _validate_object(value, path, _NUMBERING_SCHEMA)
    if value["mode"] == "global":
        if value["chapterStyleLevel"] is not None or value["resetLevel"] is not None:
            _invalid(path, "global numbering without chapter/reset levels")
    elif value["chapterStyleLevel"] != 1 or value["resetLevel"] != 1:
        _invalid(path, "chapter numbering at heading/reset level 1")


_FIGURE_CHILD_SCHEMA = _schema(
    ("nodeId",),
    nodeId=_string,
    resourceId=_logical_id,
    displayWidthPt=_POSITIVE_NUMBER,
    displayHeightPt=_POSITIVE_NUMBER,
    effectiveDpi=_NONNEGATIVE_NUMBER,
    mediaType=_enum(ALLOWED_MEDIA_TYPES, "allowed M3 media type"),
    normalizerId=_enum(_NORMALIZER_IDS, "controlled normalizer identifier"),
    plannedDegradation=_planned_degradation,
)


def _figure_child(value: Any, path: str) -> None:
    _validate_object(value, path, _FIGURE_CHILD_SCHEMA)
    resource_fields = {
        "resourceId", "displayWidthPt", "displayHeightPt", "effectiveDpi",
        "mediaType", "normalizerId",
    }
    has_resource = bool(resource_fields & set(value))
    has_degradation = "plannedDegradation" in value
    if has_resource == has_degradation:
        _invalid(path, "exactly one complete resource or planned degradation")
    if has_resource and not resource_fields <= set(value):
        _invalid(path, "complete resolved image metadata")


def _figure_children(value: Any, path: str) -> None:
    _list_of(_figure_child)(value, path)
    if not 1 <= len(value) <= 2:
        _invalid(path, "one or two figure children")


_BORDER_SCHEMA = _schema(
    ("top", "bottom", "headerBottom", "left", "right", "insideHorizontal", "insideVertical"),
    top=_NONNEGATIVE_NUMBER,
    bottom=_NONNEGATIVE_NUMBER,
    headerBottom=_NONNEGATIVE_NUMBER,
    left=_NONNEGATIVE_NUMBER,
    right=_NONNEGATIVE_NUMBER,
    insideHorizontal=_NONNEGATIVE_NUMBER,
    insideVertical=_NONNEGATIVE_NUMBER,
)


def _border_spec(value: Any, path: str) -> None:
    _validate_object(value, path, _BORDER_SCHEMA)


_MERGE_SCHEMA = _schema(
    ("top", "left", "bottom", "right"),
    top=_POSITIVE_INT,
    left=_POSITIVE_INT,
    bottom=_POSITIVE_INT,
    right=_POSITIVE_INT,
)


def _merge(value: Any, path: str) -> None:
    _validate_object(value, path, _MERGE_SCHEMA)
    if value["bottom"] < value["top"] or value["right"] < value["left"]:
        _invalid(path, "ordered rectangular merge coordinates")
    if value["bottom"] == value["top"] and value["right"] == value["left"]:
        _invalid(path, "non-degenerate merge rectangle")


_TABLE_DEGRADATION_SCHEMA = _schema(
    ("code", "message", "placement", "insertAfter", "trigger", "recoveryScope", "actions"),
    code=_string,
    message=_string,
    placement=_enum(frozenset({"block"}), "block placement"),
    insertAfter=_enum(frozenset({"caption"}), "caption adjacency"),
    trigger=_string,
    recoveryScope=_string,
    actions=_STRING_LIST,
    rowGroup=lambda value, path: _validate_object(
        value,
        path,
        _schema(("top", "bottom"), top=_POSITIVE_INT, bottom=_POSITIVE_INT),
    ),
)


def _table_degradation(value: Any, path: str) -> None:
    _validate_object(value, path, _TABLE_DEGRADATION_SCHEMA)
    if value["code"] not in {"TABLE_MERGE_INVALID", "TABLE_ROW_FORCED_SPLIT"}:
        _invalid(f"{path}.code", "controlled M3 table degradation code")
    if value["recoveryScope"] not in {"row", "complete-table"}:
        _invalid(f"{path}.recoveryScope", "row or complete-table")
    allowed_actions = {
        "discard-all-merges", "preserve-complete-grid", "allow-row-split",
        "apply-grid-style",
    }
    if not value["actions"] or any(item not in allowed_actions for item in value["actions"]):
        _invalid(f"{path}.actions", "controlled non-empty table recovery actions")
    if "rowGroup" in value and value["rowGroup"]["bottom"] < value["rowGroup"]["top"]:
        _invalid(f"{path}.rowGroup", "ordered row interval")


_TEXT_RUN_SCHEMA = _schema(("type", "text"), type=_enum(frozenset({"text"}), "text run"), text=_string)
_REFERENCE_RUN_SCHEMA = _schema(
    ("type", "targetNodeId", "targetKind", "bookmarkName", "prefix", "suffix", "fallbackText"),
    type=_enum(frozenset({"reference"}), "reference run"),
    targetNodeId=_string,
    targetKind=_enum(_OBJECT_KINDS, "reference target kind"),
    bookmarkName=_bookmark,
    prefix=_string,
    suffix=_string,
    fallbackText=_string,
)


def _reference_run(value: Any, path: str) -> None:
    if not isinstance(value, dict):
        _invalid(path, "text or reference run")
    if value.get("type") == "text":
        _validate_object(value, path, _TEXT_RUN_SCHEMA)
    elif value.get("type") == "reference":
        _validate_object(value, path, _REFERENCE_RUN_SCHEMA)
    else:
        _invalid(f"{path}.type", "text or reference")


def _reference_runs(value: Any, path: str) -> None:
    _list_of(_reference_run)(value, path)
    if not value or not any(item.get("type") == "reference" for item in value):
        _invalid(path, "non-empty runs containing a resolved reference")

_LONGFORM_OPERATION_ARG_SCHEMAS: dict[str, _ObjectSchema] = {
    "writer.configure_front_matter": _schema(
        (),
        title=_nullable_string,
        shortTitle=_nullable_string,
        author=_nullable_string,
        date=_nullable_string,
        header=_nullable_string,
        titlePage=_boolean,
    ),
   "writer.configure_section": _schema(
       (),
       landscape=_boolean,
        role=_enum(_PAGE_ROLES, "page role"),
       pageSize=_string,
       margins=_schema(
           ("top", "bottom", "left", "right"),
            top=_number,
            bottom=_number,
            left=_number,
            right=_number,
        ),
        restartPageNumbering=_boolean,
        pageNumberFormat=_enum(_PAGE_NUMBER_FORMATS, "page number format"),
        startPageNumber=_integer,
        headerText=_string,
        footerText=_string,
        linkToPreviousHeader=_boolean,
        linkToPreviousFooter=_boolean,
    ),
    "writer.configure_toc_styles": _schema(
        (),
        tocTitle=_nullable_string,
        levels=_integer,
        includeFigureIndex=_boolean,
        includeTableIndex=_boolean,
        figureIndexTitle=_nullable_string,
        tableIndexTitle=_nullable_string,
        minFontSizePt=_toc_density_map(_TOC_DENSITY_BOUNDS["minFontSizePt"]),
        minSpaceBeforePt=_toc_density_map(_TOC_DENSITY_BOUNDS["minSpaceBeforePt"]),
        minSpaceAfterPt=_toc_density_map(_TOC_DENSITY_BOUNDS["minSpaceAfterPt"]),
    ),
    "writer.set_page_role": _schema(
        ("role",),
        role=_enum(_PAGE_ROLES, "page role"),
    ),
    "writer.set_page_numbering": _schema(
        ("format",),
        format=_enum(_PAGE_NUMBER_FORMATS, "page number format"),
        start=_integer,
        restart=_boolean,
    ),
    "writer.set_header_footer": _schema(
        (),
        headerText=_string,
        footerText=_string,
        linkToPreviousHeader=_boolean,
        linkToPreviousFooter=_boolean,
    ),
    "writer.add_captioned_figure": _schema(
        (
            "caption", "numbering", "widthMode", "orientation", "kind",
            "indexable", "referenceable", "children", "layout", "keepWithCaption",
        ),
        caption=_string,
        numbering=_numbering,
        bookmarkName=_bookmark,
        indexable=_boolean,
        referenceable=_boolean,
        widthMode=_enum(_WIDTH_MODES, "figure width mode"),
        explicitWidthPt=_POSITIVE_NUMBER,
        orientation=_enum(_ORIENTATIONS, "page orientation"),
        kind=_enum(_FIGURE_KINDS, "figure kind"),
        children=_figure_children,
        layout=_enum(_FIGURE_LAYOUTS, "figure layout"),
        columns=_bounded_integer(2, 2),
        keepWithCaption=_boolean,
    ),
    "writer.add_semantic_table": _schema(
        (
            "caption", "numbering", "indexable", "referenceable", "headers", "rows",
            "alignments", "style", "orientation", "borderSpec", "merges",
            "repeatHeader", "allowRowSplit", "cellIndentPt",
            "plannedDegradation", "keepCaptionWithFirstRow",
        ),
        caption=_string,
        numbering=_numbering,
        bookmarkName=_bookmark,
        indexable=_boolean,
        referenceable=_boolean,
        headers=_STRING_LIST,
        rows=_list_of(_STRING_LIST),
        alignments=_list_of(_enum(_ALIGNMENTS, "cell alignment")),
        style=_enum(_TABLE_STYLES, "table style"),
        orientation=_enum(_ORIENTATIONS, "page orientation"),
        borderSpec=_border_spec,
        merges=_list_of(_merge),
        repeatHeader=_boolean,
        allowRowSplit=_boolean,
        cellIndentPt=_NONNEGATIVE_NUMBER,
        plannedDegradation=_list_of(_table_degradation),
        keepCaptionWithFirstRow=_boolean,
    ),
    "writer.add_equation": _schema(
        ("source", "numbering", "bookmarkName", "fallbackText"),
        source=_string,
        numbering=_numbering,
        bookmarkName=_bookmark,
        fallbackText=_string,
    ),
    "writer.add_cross_reference": _schema(
        ("runs",),
        runs=_reference_runs,
    ),
    "writer.insert_figure_index": _schema(
        ("title", "sequenceId", "titleStyleId"),
        title=_nullable_string,
        sequenceId=_enum(frozenset({"WPSC_FIG"}), "figure sequence identifier"),
        titleStyleId=_enum(frozenset({"WPSC_INDEX_TITLE"}), "index title style"),
    ),
    "writer.insert_table_index": _schema(
        ("title", "sequenceId", "titleStyleId"),
        title=_nullable_string,
        sequenceId=_enum(frozenset({"WPSC_TAB"}), "table sequence identifier"),
        titleStyleId=_enum(frozenset({"WPSC_INDEX_TITLE"}), "index title style"),
    ),
    "writer.add_bibliography": _schema(
        ("entries",),
        entries=_STRING_LIST,
        style=_string,
    ),
    "writer.add_inline_degradation": _schema(
        ("code", "message", "fallbackText"),
        code=_string,
        message=_string,
        fallbackText=_string,
    ),
    "writer.add_degradation_notice": _schema(
        ("code", "message", "fallbackText", "placement"),
        code=_string,
        message=_string,
        fallbackText=_string,
        placement=_string,
    ),
    "writer.add_document_quality_notice": _schema(
        ("notices",),
        notices=_list_of(_notice_item),
    ),
    "writer.finalize_fields": _schema(
        (),
        maxRounds=_integer,
    ),
}


_M2_COMPAT_OPERATION_ARG_SCHEMAS: dict[str, _ObjectSchema] = {
    "writer.add_captioned_figure": _schema(
        ("caption", "children", "layout"),
        caption=_string,
        children=_list_of(
            lambda value, path: _validate_object(
                value,
                path,
                _schema(
                    ("nodeId",),
                    nodeId=_string,
                    resourceId=_nullable_string,
                    plannedDegradation=_planned_degradation,
                ),
            )
        ),
        layout=_string,
        columns=_integer,
    ),
    "writer.add_semantic_table": _schema(
        ("caption", "headers", "rows"),
        caption=_string,
        headers=_STRING_LIST,
        rows=_list_of(_STRING_LIST),
        alignments=_STRING_LIST,
        style=_string,
        orientation=_string,
    ),
    "writer.add_equation": _schema(
        ("source",),
        source=_string,
        number=_nullable_string,
        fallbackText=_string,
    ),
    "writer.add_cross_reference": _schema(
        ("targetId", "kind", "fallbackText"),
        targetId=_string,
        kind=_string,
        fallbackText=_string,
    ),
    "writer.insert_figure_index": _schema((), title=_nullable_string),
    "writer.insert_table_index": _schema((), title=_nullable_string),
}


def _is_m3_shape(op: str, args: Mapping[str, Any]) -> bool:
    markers = {
        "writer.add_captioned_figure": "numbering",
        "writer.add_semantic_table": "numbering",
        "writer.add_equation": "numbering",
        "writer.add_cross_reference": "runs",
        "writer.insert_figure_index": "sequenceId",
        "writer.insert_table_index": "sequenceId",
    }
    marker = markers.get(op)
    return marker is None or marker in args


def _validate_operation_args(op: str, args: Mapping[str, Any]) -> None:
    schema = _OPERATION_ARG_SCHEMAS.get(op)
    if schema is None and op in _M2_COMPAT_OPERATION_ARG_SCHEMAS and not _is_m3_shape(op, args):
        schema = _M2_COMPAT_OPERATION_ARG_SCHEMAS[op]
    if schema is None:
        schema = _LONGFORM_OPERATION_ARG_SCHEMAS.get(op)
    if schema is None:
        raise OperationPlanError(f"unknown operation schema: {op}")
    _validate_object(args, f"{op}.args", schema)
    _validate_table_shape(op, args)
    if _is_m3_shape(op, args):
        _validate_m3_operation_contract(op, args)


def _validate_m3_operation_contract(op: str, args: Mapping[str, Any]) -> None:
    """Validate invariants spanning nested fields of one M3 operation."""
    expected = {
        "writer.add_captioned_figure": ("WPSC_FIG", "图 ", "", "fig"),
        "writer.add_semantic_table": ("WPSC_TAB", "表 ", "", "tab"),
        "writer.add_equation": ("WPSC_EQ", "(", ")", "eq"),
    }
    if op in expected:
        sequence_id, prefix, suffix, bookmark_kind = expected[op]
        numbering = args["numbering"]
        if (
            numbering["sequenceId"] != sequence_id
            or numbering["prefix"] != prefix
            or numbering["suffix"] != suffix
        ):
            _invalid(f"{op}.args.numbering", "controlled type-specific numbering")
        bookmark = args.get("bookmarkName")
        if bookmark is not None and not bookmark.startswith(f"wpsc_{bookmark_kind}_"):
            _invalid(f"{op}.args.bookmarkName", f"{bookmark_kind} bookmark")
        if op in {"writer.add_captioned_figure", "writer.add_semantic_table"}:
            if args["referenceable"] and bookmark is None:
                _invalid(f"{op}.args.bookmarkName", "bookmark for referenceable object")
            if not args["referenceable"] and bookmark is not None:
                _invalid(f"{op}.args.bookmarkName", "omitted for non-referenceable object")
            if not args["caption"] and (
                args["indexable"] or args["referenceable"] or bookmark is not None
            ):
                _invalid(f"{op}.args.caption", "empty caption omitted from indexes and targets")

    if op == "writer.add_captioned_figure":
        children = args["children"]
        if args["layout"] == "columns":
            if len(children) != 2 or args.get("columns") != 2:
                _invalid(f"{op}.args.columns", "2 with exactly two children")
        elif "columns" in args:
            _invalid(f"{op}.args.columns", "omitted for stack layout")
        if args["widthMode"] == "explicit" and "explicitWidthPt" not in args:
            _invalid(f"{op}.args.explicitWidthPt", "positive width for explicit mode")
        if args["widthMode"] != "explicit" and "explicitWidthPt" in args:
            _invalid(f"{op}.args.explicitWidthPt", "omitted unless widthMode is explicit")
        for index, child in enumerate(children):
            if "resourceId" in child and args["layout"] == "columns":
                if child["displayWidthPt"] <= 0 or child["displayHeightPt"] <= 0:
                    _invalid(f"{op}.args.children[{index}]", "positive dimensions")

    if op == "writer.add_semantic_table":
        width = len(args["headers"])
        if width == 0 or len(args["alignments"]) != width:
            _invalid(f"{op}.args.alignments", "one alignment per table column")
        if any(len(row) != width for row in args["rows"]):
            _invalid(f"{op}.args.rows", "rectangular rows matching headers")
        row_count = 1 + len(args["rows"])
        occupied: set[tuple[int, int]] = set()
        for index, merge in enumerate(args["merges"]):
            if merge["bottom"] > row_count or merge["right"] > width:
                _invalid(f"{op}.args.merges[{index}]", "coordinates inside table grid")
            if merge["top"] == 1 and merge["bottom"] != 1:
                _invalid(f"{op}.args.merges[{index}]", "merge not crossing header boundary")
            coordinates = {
                (row, col)
                for row in range(merge["top"], merge["bottom"] + 1)
                for col in range(merge["left"], merge["right"] + 1)
            }
            if occupied & coordinates:
                _invalid(f"{op}.args.merges[{index}]", "non-overlapping merge")
            occupied.update(coordinates)

    if op == "writer.add_cross_reference":
        for index, run in enumerate(args["runs"]):
            if run["type"] != "reference":
                continue
            expected_affixes = {
                "figure": ("fig", "图 ", ""),
                "table": ("tab", "表 ", ""),
                "equation": ("eq", "(", ")"),
            }[run["targetKind"]]
            kind, prefix, suffix = expected_affixes
            if (
                not run["bookmarkName"].startswith(f"wpsc_{kind}_")
                or run["prefix"] != prefix
                or run["suffix"] != suffix
            ):
                _invalid(f"{op}.args.runs[{index}]", "controlled resolved reference")


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
    if component not in ALLOWED_OPERATIONS:
        raise OperationPlanError(f"unsupported component: {component}")
    if raw.get("component") != component:
        raise OperationPlanError(
            f"component mismatch: expected {component}, got {raw.get('component')}"
        )

    protocol_version = raw.get("protocolVersion")
    is_v2 = protocol_version == PROTOCOL_VERSION_V2

    if is_v2:
        if set(raw) != _V2_ENVELOPE_KEYS:
            missing = _V2_ENVELOPE_KEYS - set(raw)
            extra = set(raw) - _V2_ENVELOPE_KEYS
            details = []
            if missing:
                details.append(f"missing: {', '.join(sorted(missing))}")
            if extra:
                details.append(f"extra: {', '.join(sorted(extra))}")
            raise OperationPlanError(
                "generation plan must match protocol v2 envelope: " + "; ".join(details)
            )
        if raw.get("semanticVersion") != SEMANTIC_VERSION_LONGFORM:
            raise OperationPlanError(
                f"semanticVersion must be {SEMANTIC_VERSION_LONGFORM!r}"
            )
        if raw.get("resourceManifestVersion") != RESOURCE_MANIFEST_VERSION:
            raise OperationPlanError(
                f"resourceManifestVersion must be {RESOURCE_MANIFEST_VERSION}"
            )
        if not isinstance(raw.get("resourceManifestDigest"), str):
            raise OperationPlanError("resourceManifestDigest must be a string")
    else:
        if set(raw) != {"component", "operations"}:
            raise OperationPlanError(
                "generation plan must contain exactly component and operations"
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
    allowed_v1 = ALLOWED_OPERATIONS[component]
    allowed_v2 = _LONGFORM_WRITER_OPERATIONS if component == "writer" else frozenset()
    for operation in normalized["operations"]:
        if not isinstance(operation, dict):
            raise OperationPlanError("operation must be an object")
        op_name = operation.get("op")
        if not isinstance(op_name, str):
            raise OperationPlanError("operation op must be a string")
        if is_v2:
            allowed = allowed_v1 | allowed_v2
        else:
            allowed = allowed_v1
            if set(operation) != {"op", "args"}:
                raise OperationPlanError(
                    "operation must contain exactly op and args"
                )
        required_keys = {"op", "args"}
        unknown = set(operation) - required_keys - {"nodeId", "failurePolicy"}
        if unknown:
            raise OperationPlanError(
                f"operation contains unknown keys: {', '.join(sorted(unknown))}"
            )
        if not required_keys <= set(operation):
            raise OperationPlanError("operation must contain op and args")
        if op_name not in allowed:
            raise OperationPlanError(f"unsupported operation: {op_name}")

        node_id = operation.get("nodeId")
        if node_id is not None and not isinstance(node_id, str):
            raise OperationPlanError("operation nodeId must be a string")
        if (
            is_v2
            and op_name in _LONGFORM_OPERATION_ARG_SCHEMAS
            and not node_id
            and op_name
            not in {
                "writer.finalize_fields",
                "writer.insert_figure_index",
                "writer.insert_table_index",
                "writer.add_document_quality_notice",
                "writer.configure_front_matter",
                "writer.configure_section",
                "writer.configure_toc_styles",
            }
        ):
            raise OperationPlanError(f"{op_name} requires a nodeId")

        failure_policy = operation.get("failurePolicy")
        if failure_policy is not None:
            _failure_policy(failure_policy, f"{op_name}.failurePolicy")
        if not is_v2 and failure_policy is not None:
            raise OperationPlanError("failurePolicy is only allowed in protocol v2")

        args = operation.get("args")
        if not isinstance(args, dict):
            raise OperationPlanError("operation args must be an object")
        if op_name.endswith(".add_table") or op_name == "sheet.write_table":
            if _table_cell_count(args) > MAX_TABLE_CELLS:
                raise OperationPlanError("table exceeds 10,000 cells")
        _validate_image_args(op_name, args)
        _validate_operation_args(op_name, args)
        if is_v2:
            _validate_m3_failure_policy(op_name, failure_policy, args)
        parsed.append(
            GenerationOperation(
                op_name,
                dict(args),
                node_id=node_id,
                failure_policy=failure_policy,
            )
        )

    if is_v2:
        _validate_m3_plan_state(normalized["operations"])
        return GenerationPlan(
            component,
            tuple(parsed),
            protocol_version=normalized["protocolVersion"],
            semantic_version=normalized["semanticVersion"],
            resource_manifest_version=normalized["resourceManifestVersion"],
            resource_manifest_digest=normalized["resourceManifestDigest"],
        )
    return GenerationPlan(component, tuple(parsed))


def _validate_m3_plan_state(operations: list[dict[str, Any]]) -> None:
    """Enforce deterministic ownership and terminal ordering for complete M3 plans."""
    if not any(_is_m3_shape(item["op"], item["args"]) for item in operations):
        return

    finalizers = [index for index, item in enumerate(operations) if item["op"] == "writer.finalize_fields"]
    if finalizers and (len(finalizers) != 1 or finalizers[0] != len(operations) - 1):
        raise OperationPlanError("writer.finalize_fields must occur exactly once and last")

    body_ops = {
        "writer.add_captioned_figure",
        "writer.add_semantic_table",
        "writer.add_equation",
        "writer.add_cross_reference",
    }
    owned: set[str] = set()
    object_positions: list[int] = []
    referenceable_targets: set[str] = set()
    has_native_targets = False
    for index, item in enumerate(operations):
        if item["op"] not in body_ops or not _is_m3_shape(item["op"], item["args"]):
            continue
        object_positions.append(index)
        node_id = item.get("nodeId")
        if node_id in owned:
            raise OperationPlanError(f"semantic node is owned more than once: {node_id}")
        owned.add(node_id)
        if item["op"] == "writer.add_captioned_figure":
            for child in item["args"]["children"]:
                child_node_id = child["nodeId"]
                if child_node_id in owned:
                    raise OperationPlanError(
                        f"semantic node is owned more than once: {child_node_id}"
                    )
                owned.add(child_node_id)
        if item["op"] == "writer.add_equation" or (
            item["op"] in {"writer.add_captioned_figure", "writer.add_semantic_table"}
            and item["args"].get("referenceable", True)
        ):
            referenceable_targets.add(node_id)
        if item["op"] in {
            "writer.add_captioned_figure", "writer.add_semantic_table", "writer.add_equation"
        }:
            has_native_targets = True

    is_sectioned_plan = any(
        item["op"] == "writer.configure_section"
        and item["args"].get("role") in {"front_matter", "body", "landscape"}
        for item in operations
    )
    if object_positions and is_sectioned_plan:
        first_body = min(object_positions)
        for index, item in enumerate(operations):
            if item["op"] in {"writer.insert_figure_index", "writer.insert_table_index"} and index > first_body:
                raise OperationPlanError("native figure/table indexes must precede body objects")

    if has_native_targets:
        for item in operations:
            if item["op"] != "writer.add_cross_reference" or not _is_m3_shape(item["op"], item["args"]):
                continue
            for run in item["args"]["runs"]:
                if run["type"] == "reference" and run["targetNodeId"] not in referenceable_targets:
                    raise OperationPlanError("reference run target is not a referenceable native object")
