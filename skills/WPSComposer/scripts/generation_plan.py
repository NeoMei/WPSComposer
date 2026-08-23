"""Closed, JSON-compatible operation plans for macOS WPS generation."""

from __future__ import annotations

import json
import math
import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from types import MappingProxyType
from typing import Any, Callable, Mapping, Optional

from .longform.citation_ids import table_cell_citation_node_id
from .longform.native_math import (
    contains_forbidden_formula_command,
    validate_wps_linear_text,
)
from .longform.privacy import contains_private_plan_text


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
        "writer.reserve_document_quality_anchor",
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

_PAGE_ROLES = frozenset({"cover", "front_matter", "body", "landscape", "bibliography"})
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
_M3_EXISTING_LOCAL_FAILURE_POLICIES = {
    "writer.add_bibliography": (("BIBLIOGRAPHY_INSERT_FAILED",), "notice"),
    "writer.add_inline_degradation": (("DEGRADATION_INSERT_FAILED",), "inline"),
    "writer.add_degradation_notice": (("DEGRADATION_INSERT_FAILED",), "notice"),
}


def _validate_m3_failure_policy(
    op: str,
    value: Optional[Mapping[str, Any]],
    native_mode: str,
    args: Optional[Mapping[str, Any]] = None,
) -> None:
    if native_mode != "m3":
        return
    if value is not None and value.get("mode") == "fail" and set(value) != {"mode"}:
        _invalid(f"{op}.failurePolicy", "fatal policy containing only mode=fail")
    expected = (
        (("EQUATION_INSERT_FAILED",), "explicit-image-then-source-notice")
        if op == "writer.add_equation"
        and args is not None
        and args.get("renderMode") == "native-m4"
        else _M3_FAILURE_POLICIES.get(op)
    )
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
        return
    if value is None or value.get("mode") == "fail":
        return
    existing = _M3_EXISTING_LOCAL_FAILURE_POLICIES.get(op)
    if existing is not None:
        codes, fallback = existing
        if (
            tuple(value.get("recoverableCodes", ())) == codes
            and value.get("fallback") == fallback
        ):
            return
    _invalid(f"{op}.failurePolicy", "fatal policy or exact named local recovery")


_M2_FAILURE_POLICIES = {
    "writer.add_captioned_figure": (("IMAGE_INSERT_FAILED",), "notice"),
    "writer.add_semantic_table": (("TABLE_INSERT_FAILED",), "notice"),
    "writer.add_equation": (("EQUATION_INSERT_FAILED",), "inline"),
    "writer.add_cross_reference": (("CROSS_REFERENCE_FAILED",), "inline"),
}


def _validate_m2_failure_policy(
    op: str,
    value: Optional[Mapping[str, Any]],
    native_mode: str,
) -> None:
    if native_mode != "legacy" or value is None or op not in _M2_FAILURE_POLICIES:
        return
    if value.get("mode") == "fail" and set(value) == {"mode"}:
        return
    codes, fallback = _M2_FAILURE_POLICIES[op]
    if (
        value.get("mode") != "degrade"
        or tuple(value.get("recoverableCodes", ())) != codes
        or value.get("fallback") != fallback
    ):
        _invalid(f"{op}.failurePolicy", "exact legacy recovery allowlist and fallback")


def _validate_v2_failure_policy(
    op: str,
    value: Optional[Mapping[str, Any]],
    native_mode: str,
    args: Optional[Mapping[str, Any]] = None,
) -> None:
    """Reject recovery policies outside the closed longform-v2 allowlist."""
    if value is None:
        return
    if value.get("mode") == "fail":
        if set(value) != {"mode"}:
            _invalid(f"{op}.failurePolicy", "fatal policy containing only mode=fail")
        return

    expected = _M3_EXISTING_LOCAL_FAILURE_POLICIES.get(op)
    if expected is None:
        if (
            op == "writer.add_equation"
            and args is not None
            and args.get("renderMode") == "native-m4"
        ):
            expected = (
                ("EQUATION_INSERT_FAILED",),
                "explicit-image-then-source-notice",
            )
        else:
            expected = (
                _M3_FAILURE_POLICIES.get(op)
                if native_mode == "m3"
                else _M2_FAILURE_POLICIES.get(op)
            )
    if expected is None:
        _invalid(f"{op}.failurePolicy", "fatal policy or exact named recovery")
    codes, fallback = expected
    if (
        tuple(value.get("recoverableCodes", ())) != codes
        or value.get("fallback") != fallback
    ):
        _invalid(f"{op}.failurePolicy", "exact recovery allowlist and fallback")


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
_RESOURCE_MANIFEST_DIGEST_RE = re.compile(r"^sha256:[0-9a-f]{64}$")


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
    if has_resource:
        media_type = value["mediaType"]
        normalizer = value["normalizerId"]
        if media_type == "image/svg+xml":
            if normalizer != "svg-static-v1":
                _invalid(path, "SVG media paired with svg-static-v1 normalizer")
        elif normalizer == "svg-static-v1":
            _invalid(path, "raster media paired with a raster normalizer")
        elif normalizer in {
            "exif-transpose-png-v1",
            "gif-first-frame-png-v1",
            "tiff-first-page-png-v1",
        } and media_type != "image/png":
            _invalid(path, "normalized PNG media for PNG-producing normalizer")


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
    code = value["code"]
    if code not in {"TABLE_MERGE_INVALID", "TABLE_ROW_FORCED_SPLIT"}:
        _invalid(f"{path}.code", "controlled M3 table degradation code")
    if code == "TABLE_MERGE_INVALID":
        if (
            value["trigger"] != "invalid-merge-declaration"
            or value["recoveryScope"] != "complete-table"
            or tuple(value["actions"])
            != ("discard-all-merges", "preserve-complete-grid")
            or "rowGroup" in value
        ):
            _invalid(path, "exact TABLE_MERGE_INVALID degradation recovery")
        return
    row_group = value.get("rowGroup")
    if row_group is None or row_group["top"] < 2 or row_group["bottom"] < row_group["top"]:
        _invalid(f"{path}.rowGroup", "ordered body-row interval")
    if row_group["top"] == row_group["bottom"]:
        expected = (
            "row-exceeds-available-page",
            "row",
            ("allow-row-split",),
        )
    else:
        expected = (
            "vertical-merge-group-exceeds-available-page",
            "complete-table",
            ("discard-all-merges", "apply-grid-style", "allow-row-split"),
        )
    if (
        value["trigger"], value["recoveryScope"], tuple(value["actions"])
    ) != expected:
        _invalid(path, "exact TABLE_ROW_FORCED_SPLIT degradation recovery")


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

_M4_FORMULA_CONTENT_CODES = frozenset({
    "FORMULA_FORBIDDEN_PRIMITIVE",
    "FORMULA_MALFORMED",
    "FORMULA_NATIVE_EQUIVALENCE_UNSUPPORTED",
    "FORMULA_NESTING_TOO_DEEP",
    "FORMULA_TOO_COMPLEX",
    "FORMULA_TOO_LONG",
    "FORMULA_UNKNOWN_COMMAND",
})
_M4_FORMULA_RESOURCE_CODES = frozenset({"FORMULA_FALLBACK_IMAGE_UNAVAILABLE"})
_M4_INLINE_DEGRADATION_CODES = frozenset({"REFERENCE_UNRESOLVED"})
_M4_SOURCE_HASH_RE = re.compile(r"^[0-9a-f]{64}$")


def _m4_text(value: Any, path: str, *, maximum: int = 10_000) -> None:
    _string(value, path)
    if not value or len(value) > maximum:
        _invalid(path, f"non-empty string no longer than {maximum} Unicode code points")
    if any(
        unicodedata.category(char) in {"Cf", "Cs"}
        or (unicodedata.category(char) == "Cc" and char not in "\t\n\r")
        for char in value
    ):
        _invalid(path, "Unicode text without control characters")
    if re.search(r"data:[^,;]{0,80};base64,", value, re.IGNORECASE):
        _invalid(path, "privacy-safe text without embedded base64 payloads")


def _privacy_safe_text(value: Any, path: str, *, maximum: int = 10_000) -> None:
    _m4_text(value, path, maximum=maximum)
    if contains_private_plan_text(value):
        _invalid(path, "privacy-safe text without paths, hashes, exception reprs, or base64")


def _m4_id(value: Any, path: str) -> None:
    _m4_text(value, path, maximum=128)
    if (
        "/" in value
        or "\\" in value
        or any(unicodedata.category(char).startswith("C") for char in value)
    ):
        _invalid(path, "bounded logical identifier")


def _m4_node_id(value: Any, path: str) -> None:
    _m4_text(value, path, maximum=256)
    if (
        value.startswith("/")
        or re.match(r"^[A-Za-z][A-Za-z0-9+.-]*://", value)
        or contains_private_plan_text(value)
        or "\\" in value
        or any(
            unicodedata.category(char).startswith("C") for char in value
        )
    ):
        _invalid(path, "bounded semantic node identifier")


_NATIVE_MATH_SCHEMA = _schema(
    ("syntax", "linearText", "sourceHash"),
    syntax=_enum(frozenset({"wps-linear-v1"}), "trusted native-math syntax"),
    linearText=lambda value, path: _m4_text(value, path, maximum=10_000),
    sourceHash=lambda value, path: (
        None
        if isinstance(value, str) and _M4_SOURCE_HASH_RE.fullmatch(value)
        else _invalid(path, "64 lowercase hexadecimal characters")
    ),
)


def _native_math(value: Any, path: str) -> None:
    _validate_object(value, path, _NATIVE_MATH_SCHEMA)
    try:
        validate_wps_linear_text(value["linearText"])
    except (TypeError, ValueError):
        _invalid(f"{path}.linearText", "closed trusted WPS linear math")


_M4_DEGRADATION_SCHEMA = _schema(
    ("code", "placement", "objectLabel", "reason", "fallbackText", "fallbackKind"),
    code=lambda value, path: _m4_text(value, path, maximum=64),
    placement=_enum(frozenset({"block"}), "block placement"),
    objectLabel=lambda value, path: _m4_text(value, path, maximum=64),
    reason=lambda value, path: _privacy_safe_text(value, path, maximum=500),
    fallbackText=lambda value, path: _privacy_safe_text(value, path, maximum=10_000),
    fallbackKind=_enum(frozenset({"source", "none"}), "controlled formula fallback kind"),
)


def _m4_formula_degradation(value: Any, path: str) -> None:
    _validate_object(value, path, _M4_DEGRADATION_SCHEMA)
    code = value["code"]
    if code in _M4_FORMULA_CONTENT_CODES:
        expected = ("block", "formula", "source")
    elif code in _M4_FORMULA_RESOURCE_CODES:
        expected = ("block", "formula", "none")
    else:
        _invalid(f"{path}.code", "controlled M4 formula degradation code")
    if (value["placement"], value["objectLabel"], value["fallbackKind"]) != expected:
        _invalid(path, "exact controlled formula degradation placement and fallback")
    if contains_forbidden_formula_command(value["fallbackText"]):
        _invalid(f"{path}.fallbackText", "formula fallback without forbidden commands")


def _m4_formula_content(value: Any, path: str) -> None:
    if not isinstance(value, dict):
        _invalid(path, "object")
    if set(value) == {"nativeMath"}:
        _native_math(value["nativeMath"], f"{path}.nativeMath")
    elif set(value) == {"plannedDegradation"}:
        _m4_formula_degradation(
            value["plannedDegradation"], f"{path}.plannedDegradation"
        )
        if value["plannedDegradation"]["code"] not in _M4_FORMULA_CONTENT_CODES:
            _invalid(f"{path}.plannedDegradation.code", "formula content degradation code")
    else:
        _invalid(path, "exactly one nativeMath or plannedDegradation member")


def _m4_formula_fallback_resource(value: Any, path: str) -> None:
    if not isinstance(value, dict):
        _invalid(path, "object")
    if set(value) == {"fallbackResourceId"}:
        _m4_id(value["fallbackResourceId"], f"{path}.fallbackResourceId")
    elif set(value) == {"fallbackResourcePlannedDegradation"}:
        descriptor = value["fallbackResourcePlannedDegradation"]
        _m4_formula_degradation(
            descriptor, f"{path}.fallbackResourcePlannedDegradation"
        )
        if descriptor["code"] not in _M4_FORMULA_RESOURCE_CODES:
            _invalid(
                f"{path}.fallbackResourcePlannedDegradation.code",
                "formula fallback-resource degradation code",
            )
    else:
        _invalid(
            path,
            "exactly one fallbackResourceId or fallbackResourcePlannedDegradation member",
        )


_M3_EQUATION_SCHEMA = _schema(
    ("source", "numbering", "bookmarkName", "fallbackText"),
    source=_string,
    numbering=_numbering,
    bookmarkName=_bookmark,
    fallbackText=_string,
)


def _m4_formula_fallback_text(value: Any, path: str) -> None:
    _privacy_safe_text(value, path, maximum=10_000)
    if contains_forbidden_formula_command(value):
        _invalid(path, "formula fallback without forbidden commands")


_M4_EQUATION_SCHEMA = _schema(
    ("renderMode", "content", "numbering", "bookmarkName", "fallbackText"),
    renderMode=_enum(frozenset({"native-m4"}), "native-m4 render mode"),
    content=_m4_formula_content,
    fallbackResource=_m4_formula_fallback_resource,
    numbering=_numbering,
    bookmarkName=_bookmark,
    fallbackText=_m4_formula_fallback_text,
)


def _equation_args(value: Any, path: str) -> None:
    if not isinstance(value, dict):
        _invalid(path, "object")
    schema = _M4_EQUATION_SCHEMA if "renderMode" in value else _M3_EQUATION_SCHEMA
    _validate_object(value, path, schema)


_CITATION_RUN_SCHEMA = _schema(
    ("type", "nodeId", "targetId", "targetNodeId", "number", "fallbackText"),
    type=_enum(frozenset({"citation"}), "citation run"),
    nodeId=_m4_node_id,
    targetId=_m4_id,
    targetNodeId=_m4_node_id,
    number=_bounded_integer(1, 10_000),
    fallbackText=lambda value, path: _m4_text(value, path, maximum=128),
)
_INLINE_DEGRADATION_RUN_SCHEMA = _schema(
    ("type", "nodeId", "code", "fallbackText"),
    type=_enum(frozenset({"degradation"}), "degradation run"),
    nodeId=_m4_node_id,
    code=_enum(_M4_INLINE_DEGRADATION_CODES, "controlled inline degradation code"),
    fallbackText=lambda value, path: _m4_text(value, path, maximum=500),
)


def _reference_run(value: Any, path: str) -> None:
    if not isinstance(value, dict):
        _invalid(path, "text or reference run")
    if value.get("type") == "text":
        _validate_object(value, path, _TEXT_RUN_SCHEMA)
    elif value.get("type") == "reference":
        _validate_object(value, path, _REFERENCE_RUN_SCHEMA)
    elif value.get("type") == "citation":
        _validate_object(value, path, _CITATION_RUN_SCHEMA)
        if value["fallbackText"] != f"[{value['number']}]":
            _invalid(f"{path}.fallbackText", "numeric citation fallback matching number")
    elif value.get("type") == "degradation":
        _validate_object(value, path, _INLINE_DEGRADATION_RUN_SCHEMA)
        if value["fallbackText"] != "[REFERENCE_UNRESOLVED 引用目标未解析]":
            _invalid(f"{path}.fallbackText", "controlled unresolved-reference fallback")
    else:
        _invalid(f"{path}.type", "text, reference, citation, or degradation")


def _reference_runs(value: Any, path: str) -> None:
    _list_of(_reference_run)(value, path)
    if not 1 <= len(value) <= 10_000:
        _invalid(path, "one to 10,000 closed inline runs")


_CELL_DEGRADATION_SCHEMA = _schema(
    ("row", "column", "code", "fallbackText"),
    row=_bounded_integer(1, 10_001),
    column=_bounded_integer(1, 10_000),
    code=_enum(_M4_INLINE_DEGRADATION_CODES, "controlled cell degradation code"),
    fallbackText=lambda value, path: _m4_text(value, path, maximum=500),
)


def _cell_degradation(value: Any, path: str) -> None:
    _validate_object(value, path, _CELL_DEGRADATION_SCHEMA)
    if value["fallbackText"] != "[REFERENCE_UNRESOLVED 引用目标未解析]":
        _invalid(f"{path}.fallbackText", "controlled unresolved-reference fallback")


_CELL_CITATION_SCHEMA = _schema(
    ("row", "column", "nodeId", "targetId", "targetNodeId", "number", "fallbackText"),
    row=_bounded_integer(1, 10_001),
    column=_bounded_integer(1, 10_000),
    nodeId=_m4_node_id,
    targetId=_m4_id,
    targetNodeId=_m4_node_id,
    number=_bounded_integer(1, 10_000),
    fallbackText=lambda value, path: _m4_text(value, path, maximum=128),
)


def _cell_citation(value: Any, path: str) -> None:
    _validate_object(value, path, _CELL_CITATION_SCHEMA)
    if value["fallbackText"] != f"[{value['number']}]":
        _invalid(f"{path}.fallbackText", "numeric citation fallback matching number")


def _cell_citations(value: Any, path: str) -> None:
    _list_of(_cell_citation)(value, path)
    if len(value) > 10_000:
        _invalid(path, "at most 10,000 explicit table-cell citations")


def _bibliography_text(value: Any, path: str) -> None:
    _m4_text(value, path, maximum=10_000)
    if value != value.strip() or any(char in value for char in "\r\n\t\v\f\u2028\u2029"):
        _invalid(path, "one trimmed bibliography paragraph")


_BIBLIOGRAPHY_ENTRY_SCHEMA = _schema(
    ("id", "nodeId", "number", "text", "cited"),
    id=_m4_id,
    nodeId=_m4_node_id,
    number=_bounded_integer(1, 10_000),
    text=_bibliography_text,
    cited=_boolean,
)


def _bibliography_entry(value: Any, path: str) -> None:
    _validate_object(value, path, _BIBLIOGRAPHY_ENTRY_SCHEMA)


_LEGACY_BIBLIOGRAPHY_SCHEMA = _schema(
    ("entries",), entries=_STRING_LIST, style=_string
)
_STRUCTURED_BIBLIOGRAPHY_SCHEMA = _schema(
    (
        "schemaVersion", "entries", "style", "hangingIndentPt",
        "leftIndentPt", "spaceAfterPt",
    ),
    schemaVersion=_bounded_integer(1, 1),
    entries=_list_of(_bibliography_entry),
    style=_enum(frozenset({"numeric"}), "numeric bibliography style"),
    hangingIndentPt=_NONNEGATIVE_NUMBER,
    leftIndentPt=_NONNEGATIVE_NUMBER,
    spaceAfterPt=_NONNEGATIVE_NUMBER,
)


def _bibliography_args(value: Any, path: str) -> None:
    if not isinstance(value, dict):
        _invalid(path, "object")
    if "schemaVersion" not in value:
        _validate_object(value, path, _LEGACY_BIBLIOGRAPHY_SCHEMA)
        return
    _validate_object(value, path, _STRUCTURED_BIBLIOGRAPHY_SCHEMA)
    entries = value["entries"]
    if not entries or len(entries) > 10_000:
        _invalid(f"{path}.entries", "one to 10,000 structured entries")
    numbers = [entry["number"] for entry in entries]
    if numbers != sorted(set(numbers)):
        _invalid(f"{path}.entries", "strictly increasing unique numeric order")
    if len({entry["id"] for entry in entries}) != len(entries):
        _invalid(f"{path}.entries", "unique bibliography identifiers")
    if len({entry["nodeId"] for entry in entries}) != len(entries):
        _invalid(f"{path}.entries", "unique bibliography node owners")
    seen_uncited = False
    for entry in entries:
        if not entry["cited"]:
            seen_uncited = True
        elif seen_uncited:
            _invalid(f"{path}.entries", "cited entries before uncited entries")
    if (
        value["hangingIndentPt"] != 18.0
        or value["leftIndentPt"] != 18.0
        or value["spaceAfterPt"] != 6.0
    ):
        _invalid(path, "fixed numeric bibliography geometry")


_M4_NOTICE_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]{0,63}$")


def _m4_notice_code(value: Any, path: str) -> None:
    if (
        not isinstance(value, str)
        or not _M4_NOTICE_CODE_RE.fullmatch(value)
        or re.fullmatch(r"(?i)[0-9a-f]{64}", value)
    ):
        _invalid(path, "bounded stable non-hash issue code")


_M4_NOTICE_SCHEMA = _schema(
    ("code", "message", "fallbackText", "placement"),
    code=_m4_notice_code,
    message=lambda value, path: _privacy_safe_text(value, path, maximum=500),
    fallbackText=lambda value, path: _privacy_safe_text(value, path, maximum=500),
    placement=_enum(frozenset({"document"}), "document placement"),
)


def _m4_notice(value: Any, path: str) -> None:
    _validate_object(value, path, _M4_NOTICE_SCHEMA)


_QUALITY_ANCHOR_SCHEMA = _schema(
    ("title", "notices"),
    title=_enum(frozenset({"生成质量提示"}), "fixed quality-anchor title"),
    notices=_list_of(_m4_notice),
)


def _quality_anchor_args(value: Any, path: str) -> None:
    _validate_object(value, path, _QUALITY_ANCHOR_SCHEMA)
    keys = [(notice["code"], notice["message"]) for notice in value["notices"]]
    if len(keys) != len(set(keys)):
        _invalid(f"{path}.notices", "unique stable code and message keys")


_LIST_FORMATTING_SCHEMA = _schema(
    ("kind", "indentPt"),
    kind=_enum(frozenset({"bullet", "ordered"}), "controlled list kind"),
    indentPt=_number,
)


def _list_formatting(value: Any, path: str) -> None:
    _validate_object(value, path, _LIST_FORMATTING_SCHEMA)
    if type(value["indentPt"]) not in {int, float} or value["indentPt"] != 24.0:
        _invalid(f"{path}.indentPt", "24.0")

_LONGFORM_OPERATION_ARG_SCHEMAS: dict[str, Any] = {
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
        cellDegradations=_list_of(_cell_degradation),
        cellCitations=_cell_citations,
        keepCaptionWithFirstRow=_boolean,
    ),
    "writer.add_equation": _equation_args,
    "writer.add_cross_reference": _schema(
        ("runs",),
        runs=_reference_runs,
        listFormatting=_list_formatting,
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
    "writer.add_bibliography": _bibliography_args,
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
    "writer.reserve_document_quality_anchor": _quality_anchor_args,
    "writer.finalize_fields": _schema(
        (),
        maxRounds=_bounded_integer(1, 3),
    ),
}


_M2_FIGURE_CHILD_SCHEMA = _schema(
    ("nodeId",),
    nodeId=_string,
    resourceId=_logical_id,
    plannedDegradation=_planned_degradation,
)


def _m2_figure_child(value: Any, path: str) -> None:
    _validate_object(value, path, _M2_FIGURE_CHILD_SCHEMA)
    if ("resourceId" in value) == ("plannedDegradation" in value):
        _invalid(path, "exactly one logical resource or planned degradation")


_M2_COMPAT_OPERATION_ARG_SCHEMAS: dict[str, _ObjectSchema] = {
    "writer.add_captioned_figure": _schema(
        ("caption", "children", "layout"),
        caption=_string,
        children=_list_of(_m2_figure_child),
        layout=_enum(frozenset({"stack", "side-by-side", "columns"}), "legacy figure layout"),
        columns=_bounded_integer(1, 2),
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
        kind=_enum(_OBJECT_KINDS, "legacy reference kind"),
        fallbackText=_string,
    ),
    "writer.insert_figure_index": _schema((), title=_nullable_string),
    "writer.insert_table_index": _schema((), title=_nullable_string),
}


_NATIVE_OBJECT_OPERATIONS = frozenset(_M2_COMPAT_OPERATION_ARG_SCHEMAS)


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
    return marker is not None and marker in args


def _native_plan_mode(operations: list[dict[str, Any]]) -> str:
    native = [
        item
        for item in operations
        if isinstance(item, dict) and item.get("op") in _NATIVE_OBJECT_OPERATIONS
    ]
    if not native:
        return "legacy"
    shapes = [
        _is_m3_shape(
            item["op"], item.get("args") if isinstance(item.get("args"), dict) else {}
        )
        for item in native
    ]
    if any(shapes) and not all(shapes):
        raise OperationPlanError("M2 legacy and M3 native object shapes cannot mix in one plan")
    return "m3" if all(shapes) else "legacy"


def _validate_operation_args(
    op: str,
    args: Mapping[str, Any],
    native_mode: str,
) -> None:
    schema = _OPERATION_ARG_SCHEMAS.get(op)
    if schema is None and op in _M2_COMPAT_OPERATION_ARG_SCHEMAS and native_mode == "legacy":
        schema = _M2_COMPAT_OPERATION_ARG_SCHEMAS[op]
    if schema is None:
        schema = _LONGFORM_OPERATION_ARG_SCHEMAS.get(op)
    if schema is None:
        raise OperationPlanError(f"unknown operation schema: {op}")
    if isinstance(schema, _ObjectSchema):
        _validate_object(args, f"{op}.args", schema)
    else:
        schema(args, f"{op}.args")
    _validate_table_shape(op, args)
    if native_mode == "m3" and op in _NATIVE_OBJECT_OPERATIONS:
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
        if args["keepWithCaption"] is not True:
            _invalid(f"{op}.args.keepWithCaption", "true")
        for index, child in enumerate(children):
            if "resourceId" in child and args["layout"] == "columns":
                if child["displayWidthPt"] <= 0 or child["displayHeightPt"] <= 0:
                    _invalid(f"{op}.args.children[{index}]", "positive dimensions")

    if op == "writer.add_semantic_table":
        width = len(args["headers"])
        if width * (1 + len(args["rows"])) > MAX_TABLE_CELLS:
            _invalid(f"{op}.args.rows", "table no larger than 10,000 cells")
        if width == 0 or len(args["alignments"]) != width:
            _invalid(f"{op}.args.alignments", "one alignment per table column")
        if any(len(row) != width for row in args["rows"]):
            _invalid(f"{op}.args.rows", "rectangular rows matching headers")
        if args["keepCaptionWithFirstRow"] is not True:
            _invalid(f"{op}.args.keepCaptionWithFirstRow", "true")
        if args["allowRowSplit"] is not False:
            _invalid(f"{op}.args.allowRowSplit", "false in the initial descriptor")
        if args["cellIndentPt"] != 0.0:
            _invalid(f"{op}.args.cellIndentPt", "0.0 in the initial descriptor")
        expected_borders = (
            {
                "top": 1.5,
                "bottom": 1.5,
                "headerBottom": 0.75,
                "left": 0.0,
                "right": 0.0,
                "insideHorizontal": 0.0,
                "insideVertical": 0.0,
            }
            if args["style"] == "three-line"
            else {
                "top": 0.75,
                "bottom": 0.75,
                "headerBottom": 0.75,
                "left": 0.75,
                "right": 0.75,
                "insideHorizontal": 0.75,
                "insideVertical": 0.75,
            }
        )
        if args["borderSpec"] != expected_borders:
            _invalid(
                f"{op}.args.borderSpec",
                f"exact {args['style']} border policy",
            )
        grid = [args["headers"], *args["rows"]]
        row_count = 1 + len(args["rows"])
        occupied: set[tuple[int, int]] = set()
        vertical_intervals: list[tuple[int, int]] = []
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
            covered = coordinates - {(merge["top"], merge["left"])}
            if any(str(grid[row - 1][col - 1]).strip() for row, col in covered):
                _invalid(f"{op}.args.merges[{index}]", "empty covered merge cells")
            occupied.update(coordinates)
            if merge["top"] >= 2 and merge["bottom"] > merge["top"]:
                vertical_intervals.append((merge["top"], merge["bottom"]))
        vertical_groups: list[tuple[int, int]] = []
        for top, bottom in sorted(vertical_intervals):
            if vertical_groups and top <= vertical_groups[-1][1]:
                vertical_groups[-1] = (
                    vertical_groups[-1][0],
                    max(vertical_groups[-1][1], bottom),
                )
            else:
                vertical_groups.append((top, bottom))
        for index, degradation in enumerate(args["plannedDegradation"]):
            row_group = degradation.get("rowGroup")
            if row_group is not None and row_group["bottom"] > row_count:
                _invalid(
                    f"{op}.args.plannedDegradation[{index}].rowGroup",
                    "body-row interval inside table grid",
                )
            if degradation["code"] == "TABLE_MERGE_INVALID" and args["merges"]:
                _invalid(
                    f"{op}.args.plannedDegradation[{index}]",
                    "merge-invalid degradation with an empty resolved merge set",
                )
            if (
                degradation["code"] == "TABLE_ROW_FORCED_SPLIT"
                and row_group is not None
                and row_group["bottom"] > row_group["top"]
                and (row_group["top"], row_group["bottom"]) not in vertical_groups
            ):
                _invalid(
                    f"{op}.args.plannedDegradation[{index}].rowGroup",
                    "an indivisible vertical merge group",
                )
        occupied_degradations: set[tuple[int, int]] = set()
        for index, degradation in enumerate(args.get("cellDegradations", [])):
            coordinate = (degradation["row"], degradation["column"])
            if coordinate[0] > row_count or coordinate[1] > width:
                _invalid(
                    f"{op}.args.cellDegradations[{index}]",
                    "coordinates inside table grid",
                )
            if coordinate in occupied_degradations:
                _invalid(
                    f"{op}.args.cellDegradations[{index}]",
                    "one degradation per table cell",
                )
            occupied_degradations.add(coordinate)
            if degradation["fallbackText"] not in grid[coordinate[0] - 1][coordinate[1] - 1]:
                _invalid(
                    f"{op}.args.cellDegradations[{index}].fallbackText",
                    "fallback text visible in the target table cell",
                )
        occupied_citations: set[tuple[int, int, str, int]] = set()
        for index, citation in enumerate(args.get("cellCitations", [])):
            coordinate = (citation["row"], citation["column"])
            if coordinate[0] > row_count or coordinate[1] > width:
                _invalid(
                    f"{op}.args.cellCitations[{index}]",
                    "coordinates inside table grid",
                )
            identity = (*coordinate, citation["targetId"], citation["number"])
            if identity in occupied_citations:
                _invalid(
                    f"{op}.args.cellCitations[{index}]",
                    "unique citation identity per table cell",
                )
            occupied_citations.add(identity)
            if citation["fallbackText"] not in grid[coordinate[0] - 1][coordinate[1] - 1]:
                _invalid(
                    f"{op}.args.cellCitations[{index}].fallbackText",
                    "fallback text visible in the target table cell",
                )

    if op == "writer.add_cross_reference":
        has_reference = any(
            run["type"] in {"reference", "citation", "degradation"}
            for run in args["runs"]
        )
        list_formatting = args.get("listFormatting")
        if not has_reference and list_formatting is None:
            _invalid(f"{op}.args.runs", "a resolved reference outside list paragraphs")
        if list_formatting is not None:
            first = args["runs"][0]
            marker = first.get("text", "") if first.get("type") == "text" else ""
            if list_formatting["kind"] == "bullet":
                valid_marker = marker.startswith("•\t")
            else:
                valid_marker = re.match(r"[1-9][0-9]*\.\t", marker) is not None
            if not valid_marker:
                _invalid(f"{op}.args.runs[0]", "controlled list marker and tab")
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
        digest = raw.get("resourceManifestDigest")
        if not isinstance(digest, str) or _RESOURCE_MANIFEST_DIGEST_RE.fullmatch(digest) is None:
            raise OperationPlanError(
                "resourceManifestDigest must match sha256:<64 lowercase hex characters>"
            )
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
    native_mode = _native_plan_mode(normalized["operations"]) if is_v2 else "legacy"

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
        if node_id is not None:
            _string(node_id, "operation.nodeId")
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
        _validate_operation_args(op_name, args, native_mode)
        if is_v2:
            _validate_m3_failure_policy(
                op_name, failure_policy, native_mode, args
            )
            _validate_m2_failure_policy(op_name, failure_policy, native_mode)
            _validate_v2_failure_policy(
                op_name, failure_policy, native_mode, args
            )
        parsed.append(
            GenerationOperation(
                op_name,
                dict(args),
                node_id=node_id,
                failure_policy=failure_policy,
            )
        )

    if is_v2:
        _validate_v2_plan_lifecycle(normalized["operations"])
        if native_mode == "m3":
            _validate_m3_plan_state(normalized["operations"])
        else:
            _validate_m2_plan_state(normalized["operations"])
        _validate_m4_plan_state(normalized["operations"])
        return GenerationPlan(
            component,
            tuple(parsed),
            protocol_version=normalized["protocolVersion"],
            semantic_version=normalized["semanticVersion"],
            resource_manifest_version=normalized["resourceManifestVersion"],
            resource_manifest_digest=normalized["resourceManifestDigest"],
        )
    return GenerationPlan(component, tuple(parsed))


def _validate_v2_plan_lifecycle(operations: list[dict[str, Any]]) -> None:
    """Enforce the lifecycle shared by every complete longform-v2 plan."""
    finalizers = [index for index, item in enumerate(operations) if item["op"] == "writer.finalize_fields"]
    if len(finalizers) != 1 or finalizers[0] != len(operations) - 1:
        raise OperationPlanError("writer.finalize_fields must occur exactly once and last")


def _is_m4_plan(operations: list[dict[str, Any]]) -> bool:
    for item in operations:
        op = item["op"]
        args = item.get("args", {})
        if op == "writer.reserve_document_quality_anchor":
            return True
        if op == "writer.add_equation" and args.get("renderMode") == "native-m4":
            return True
        if op == "writer.add_bibliography" and args.get("schemaVersion") == 1:
            return True
        if op == "writer.add_semantic_table" and (
            "cellDegradations" in args or "cellCitations" in args
        ):
            return True
        if op == "writer.add_cross_reference" and any(
            run.get("type") in {"citation", "degradation"}
            for run in args.get("runs", [])
        ):
            return True
    return False


def _validate_m4_plan_state(operations: list[dict[str, Any]]) -> None:
    """Validate M4 ownership, anchor, citation, and bibliography invariants."""
    if not _is_m4_plan(operations):
        return

    anchors = [
        (index, item)
        for index, item in enumerate(operations)
        if item["op"] == "writer.reserve_document_quality_anchor"
    ]
    if len(anchors) != 1:
        raise OperationPlanError(
            "M4 plans require writer.reserve_document_quality_anchor exactly once"
        )
    anchor_index, anchor = anchors[0]
    if (
        anchor.get("nodeId") != "doc:quality"
        or anchor.get("failurePolicy") != {"mode": "fail"}
    ):
        raise OperationPlanError(
            "document quality anchor must own doc:quality and fail hard"
        )
    if len(anchor["args"]["notices"]) > 1_000:
        raise OperationPlanError("document quality anchor has too many initial notices")
    if any(
        item["op"] == "writer.add_document_quality_notice"
        for item in operations
    ):
        raise OperationPlanError(
            "M4 quality issues must upsert through the reserved anchor"
        )
    cover_positions = [
        index
        for index, item in enumerate(operations)
        if item["op"] == "writer.configure_section"
        and item["args"].get("role") == "cover"
    ]
    if cover_positions and anchor_index != max(cover_positions) + 1:
        raise OperationPlanError(
            "document quality anchor must immediately follow the cover transition"
        )
    non_cover_transitions = [
        index
        for index, item in enumerate(operations)
        if item["op"] == "writer.configure_section"
        and item["args"].get("role") != "cover"
    ]
    if non_cover_transitions and anchor_index >= min(non_cover_transitions):
        raise OperationPlanError(
            "document quality anchor must precede the first non-cover section transition"
        )
    content_ops = {
        "writer.insert_toc",
        "writer.insert_figure_index",
        "writer.insert_table_index",
        "writer.add_heading",
        "writer.add_paragraph",
        "writer.add_list",
        "writer.add_captioned_figure",
        "writer.add_semantic_table",
        "writer.add_equation",
        "writer.add_cross_reference",
        "writer.add_bibliography",
        "writer.add_inline_degradation",
        "writer.add_degradation_notice",
        "writer.add_page_break",
    }
    first_content = min(
        (index for index, item in enumerate(operations) if item["op"] in content_ops),
        default=len(operations),
    )
    if anchor_index >= first_content:
        raise OperationPlanError(
            "document quality anchor must precede TOC, indexes, and body content"
        )

    semantic_owner_ops = {
        "writer.add_heading",
        "writer.add_paragraph",
        "writer.add_list",
        "writer.add_captioned_figure",
        "writer.add_semantic_table",
        "writer.add_equation",
        "writer.add_cross_reference",
        "writer.add_bibliography",
        "writer.add_inline_degradation",
        "writer.add_degradation_notice",
        "writer.add_page_break",
        "writer.reserve_document_quality_anchor",
    }
    owned: set[str] = set()

    def own(node_id: Any) -> None:
        if not node_id:
            return
        _m4_node_id(node_id, "M4 semantic owner nodeId")
        if node_id in owned:
            raise OperationPlanError(f"semantic node is owned more than once: {node_id}")
        owned.add(node_id)

    current_role: Optional[str] = None
    bibliography_started = False
    bibliography: dict[str, tuple[str, int, bool]] = {}
    bibliography_numbers: dict[int, str] = {}
    bibliography_sequence: list[tuple[int, bool]] = []
    citations: list[dict[str, Any]] = []
    cell_citations: list[dict[str, Any]] = []
    for item in operations:
        op = item["op"]
        args = item["args"]
        if op == "writer.configure_section":
            if item.get("nodeId") is not None:
                _m4_node_id(item["nodeId"], "M4 section nodeId")
            current_role = args.get("role")
            if bibliography_started and current_role in {"body", "landscape"}:
                raise OperationPlanError(
                    "bibliography back matter must remain after all body content"
                )
        if op in semantic_owner_ops:
            own(item.get("nodeId"))
        if op == "writer.add_captioned_figure":
            for child in args["children"]:
                own(child["nodeId"])
        if op == "writer.add_cross_reference":
            for run in args["runs"]:
                if run["type"] in {"citation", "degradation"}:
                    own(run["nodeId"])
                if run["type"] == "citation":
                    citations.append(run)
        if op == "writer.add_semantic_table":
            for citation in args.get("cellCitations", []):
                expected_node_id = table_cell_citation_node_id(
                    item.get("nodeId"),
                    citation["row"],
                    citation["column"],
                    citation["targetId"],
                )
                if citation["nodeId"] != expected_node_id:
                    raise OperationPlanError(
                        "table-cell citation node must deterministically bind table, cell, and target"
                    )
                own(citation["nodeId"])
                cell_citations.append(citation)
        if op == "writer.add_bibliography" and args.get("schemaVersion") == 1:
            if current_role != "bibliography":
                raise OperationPlanError(
                    "structured bibliography must be emitted in a bibliography section"
                )
            bibliography_started = True
            for entry in args["entries"]:
                own(entry["nodeId"])
                if entry["id"] in bibliography:
                    raise OperationPlanError(
                        f"duplicate structured bibliography id: {entry['id']}"
                    )
                if entry["number"] in bibliography_numbers:
                    raise OperationPlanError(
                        "bibliography numbers must map to exactly one target"
                    )
                bibliography[entry["id"]] = (
                    entry["nodeId"], entry["number"], entry["cited"]
                )
                bibliography_numbers[entry["number"]] = entry["id"]
                bibliography_sequence.append((entry["number"], entry["cited"]))

    if bibliography_sequence:
        if [number for number, _ in bibliography_sequence] != list(
            range(1, len(bibliography_sequence) + 1)
        ):
            raise OperationPlanError(
                "structured bibliography must be emitted in unique gap-free numeric order"
            )
        seen_uncited = False
        for _, cited in bibliography_sequence:
            if not cited:
                seen_uncited = True
            elif seen_uncited:
                raise OperationPlanError(
                    "structured bibliography must emit cited entries before uncited entries"
                )

    citation_numbers: dict[int, str] = {}
    for run in [*citations, *cell_citations]:
        target = bibliography.get(run["targetId"])
        if (
            target is None
            or target[:2] != (run["targetNodeId"], run["number"])
            or target[2] is not True
        ):
            raise OperationPlanError(
                "citation target id, node, number, and cited flag must exactly match the final bibliography"
            )
        previous_target = citation_numbers.get(run["number"])
        if previous_target is not None and previous_target != run["targetId"]:
            raise OperationPlanError(
                "one citation number cannot identify multiple bibliography targets"
            )
        citation_numbers[run["number"]] = run["targetId"]

    cited_targets = {
        identifier for identifier, (_, _, cited) in bibliography.items() if cited
    }
    referenced_targets = {run["targetId"] for run in [*citations, *cell_citations]}
    if cited_targets != referenced_targets:
        raise OperationPlanError(
            "cited bibliography entries must have explicit inline or table-cell citation metadata"
        )


def _validate_m3_plan_state(operations: list[dict[str, Any]]) -> None:
    """Enforce deterministic ownership and terminal ordering for complete M3 plans."""

    body_ops = {
        "writer.add_captioned_figure",
        "writer.add_semantic_table",
        "writer.add_equation",
        "writer.add_cross_reference",
    }
    semantic_owner_ops = {
        "writer.add_heading",
        "writer.add_paragraph",
        "writer.add_list",
        "writer.add_captioned_figure",
        "writer.add_semantic_table",
        "writer.add_equation",
        "writer.add_cross_reference",
        "writer.add_bibliography",
        "writer.add_inline_degradation",
        "writer.add_degradation_notice",
        "writer.add_page_break",
    }
    owned: set[str] = set()
    object_positions: list[int] = []
    target_descriptors: dict[str, tuple[str, str]] = {}
    bookmark_owners: dict[str, str] = {}
    index_counts = {"figure": 0, "table": 0}
    index_targets = {"figure": 0, "table": 0}
    current_section_role: Optional[str] = None
    for index, item in enumerate(operations):
        if item["op"] == "writer.configure_section":
            current_section_role = item["args"].get("role")
        index_kind = {
            "writer.insert_figure_index": "figure",
            "writer.insert_table_index": "table",
        }.get(item["op"])
        if index_kind is not None:
            if current_section_role != "front_matter":
                raise OperationPlanError(
                    "native figure/table indexes must precede body objects and be in a front-matter section"
                )
            index_counts[index_kind] += 1
            if index_counts[index_kind] > 1:
                raise OperationPlanError(f"duplicate native {index_kind} index; at most one is allowed")
        node_id = item.get("nodeId")
        if item["op"] in semantic_owner_ops and node_id:
            if node_id in owned:
                raise OperationPlanError(f"semantic node is owned more than once: {node_id}")
            owned.add(node_id)
        if item["op"] not in body_ops:
            continue
        object_positions.append(index)
        if item["op"] == "writer.add_captioned_figure":
            for child in item["args"]["children"]:
                child_node_id = child["nodeId"]
                if child_node_id in owned:
                    raise OperationPlanError(
                        f"semantic node is owned more than once: {child_node_id}"
                    )
                owned.add(child_node_id)
        target_kind = {
            "writer.add_captioned_figure": "figure",
            "writer.add_semantic_table": "table",
            "writer.add_equation": "equation",
        }.get(item["op"])
        if target_kind in index_targets and item["args"]["indexable"] and item["args"]["caption"]:
            index_targets[target_kind] += 1
        bookmark = item["args"].get("bookmarkName") if target_kind is not None else None
        if bookmark is not None:
            owner = bookmark_owners.get(bookmark)
            if owner is not None and owner != node_id:
                raise OperationPlanError(
                    f"bookmarkName must be globally unique to one target owner: {bookmark}"
                )
            bookmark_owners[bookmark] = node_id
        if target_kind is not None and (
            item["op"] == "writer.add_equation"
            or item["args"].get("referenceable", False)
        ):
            target_descriptors[node_id] = (
                target_kind,
                item["args"]["bookmarkName"],
            )

    if object_positions:
        first_body = min(object_positions)
        for index, item in enumerate(operations):
            if item["op"] in {"writer.insert_figure_index", "writer.insert_table_index"} and index > first_body:
                raise OperationPlanError("native figure/table indexes must precede body objects")

    for kind, count in index_counts.items():
        if count and index_targets[kind] == 0:
            raise OperationPlanError(
                f"native {kind} index requires an indexable target with a non-empty caption"
            )

    for item in operations:
        if item["op"] != "writer.add_cross_reference":
            continue
        for run in item["args"]["runs"]:
            if run["type"] != "reference":
                continue
            target = target_descriptors.get(run["targetNodeId"])
            actual = (run["targetKind"], run["bookmarkName"])
            if target is None or target != actual:
                raise OperationPlanError(
                    "reference target node, kind, and bookmark must exactly match a referenceable native object"
                )


def _validate_m2_plan_state(operations: list[dict[str, Any]]) -> None:
    """Retain safe target semantics for a complete legacy native-object plan."""
    targets = {
        item.get("nodeId"): {
            "writer.add_captioned_figure": "figure",
            "writer.add_semantic_table": "table",
            "writer.add_equation": "equation",
        }[item["op"]]
        for item in operations
        if item["op"] in {
            "writer.add_captioned_figure",
            "writer.add_semantic_table",
            "writer.add_equation",
        }
    }
    for item in operations:
        if item["op"] != "writer.add_cross_reference":
            continue
        if targets.get(item["args"]["targetId"]) != item["args"]["kind"]:
            raise OperationPlanError("legacy reference target must exist and match its kind")
