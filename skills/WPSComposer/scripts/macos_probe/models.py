from __future__ import annotations

from dataclasses import dataclass
import math
from pathlib import Path
import re
from typing import Any, Mapping, Optional, Union
from uuid import uuid4

from ..longform.degradation import controlled_token
from ..longform.privacy import redact_private_text

COMPONENTS = ("writer", "presentation", "spreadsheet")
METHOD_COMPONENT = {
    "probe_capabilities": None,
    "smoke_docx": "writer",
    "smoke_pdf": "writer",
    "smoke_pptx": "presentation",
    "smoke_xlsx": "spreadsheet",
    "convert_writer_pdf": "writer",
    "convert_workbook_pdf": "spreadsheet",
    "convert_presentation_pdf": "presentation",
    "generate_writer_document": "writer",
    "generate_spreadsheet_workbook": "spreadsheet",
    "generate_presentation_deck": "presentation",
    "inspect_presentation": "presentation",
    "edit_presentation": "presentation",
    "inspect_document": "writer",
    "inspect_workbook": "spreadsheet",
    "probe_longform_m0": "writer",
    "generate_longform_document": "writer",
    "mutate_longform_document": "writer",
    "patch_longform_quality_notices": "writer",
}


class ProtocolError(ValueError):
    """Raised when a bridge message violates the probe protocol."""


_LONGFORM_RESULT_KEYS = frozenset(
    {
        "outputPath",
        "appliedOperations",
        "issueCodes",
        "fieldSnapshots",
        "childResults",
        "paginationMap",
    }
)

_LONGFORM_REQUEST_KEYS = frozenset({"plan", "outputPath", "resources"})
_NOTICE_PATCH_REQUEST_KEYS = frozenset({"sourcePath", "outputPath", "notices"})
_NOTICE_PATCH_RESULT_KEYS = frozenset({
    "outputPath", "appliedNotices", "issueCodes", "fieldSnapshots", "paginationMap",
})
_BOOKMARK_RE = re.compile(r"^wpsc_(?:fig|tab|eq|ref|head|para)_[0-9a-f]{24}$")


def validate_longform_generation_request(raw: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the closed private Writer long-form request envelope."""

    if not isinstance(raw, Mapping) or set(raw) != _LONGFORM_REQUEST_KEYS:
        raise ProtocolError("Long-form generation request is invalid")
    plan = raw.get("plan")
    output = raw.get("outputPath")
    resources = raw.get("resources")
    if (
        not isinstance(plan, Mapping)
        or plan.get("protocolVersion") != 2
        or plan.get("component") != "writer"
        or not isinstance(output, str)
        or not output
        or not isinstance(resources, Mapping)
    ):
        raise ProtocolError("Long-form generation request is invalid")
    if any(
        not isinstance(resource_id, str)
        or not resource_id
        or not isinstance(locator, str)
        or not locator
        for resource_id, locator in resources.items()
    ):
        raise ProtocolError("Long-form generation request is invalid")
    return {
        "plan": dict(plan),
        "outputPath": output,
        "resources": dict(resources),
    }


def validate_longform_generation_value(raw: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the privacy-safe result envelope returned by the Writer add-in.

    The private request may contain resource locators, but the response is a
    closed summary.  It cannot echo locators, resource identities/hashes,
    bookmark maps, or visible native field results.
    """

    if not isinstance(raw, Mapping) or set(raw) - _LONGFORM_RESULT_KEYS:
        raise ProtocolError("Long-form generation result is invalid")
    output = raw.get("outputPath")
    applied = raw.get("appliedOperations", 0)
    issues = raw.get("issueCodes", [])
    history = raw.get("fieldSnapshots", [])
    children = raw.get("childResults", [])
    pagination = raw.get("paginationMap", {})
    if not isinstance(output, str) or not output:
        raise ProtocolError("Long-form generation result is invalid")
    if not isinstance(applied, int) or isinstance(applied, bool) or applied < 0:
        raise ProtocolError("Long-form generation result is invalid")
    if not isinstance(issues, list) or not isinstance(history, list):
        raise ProtocolError("Long-form generation result is invalid")
    if not isinstance(children, list) or not isinstance(pagination, dict):
        raise ProtocolError("Long-form generation result is invalid")
    for item in issues:
        if (
            not isinstance(item, dict)
            or set(item) - {
                "code", "message", "placement", "nodeId",
                "stage", "fallback", "recoverable",
            }
            or not _is_closed_issue_code(item.get("code"))
            or not isinstance(item.get("message"), str)
            or redact_private_text(item.get("message")) != item.get("message")
            or item.get("placement") not in {"block", "inline", "document"}
            or (
                item.get("nodeId") is not None
                and (
                    not isinstance(item.get("nodeId"), str)
                    or redact_private_text(item.get("nodeId")) != item.get("nodeId")
                )
            )
            or (
                item.get("stage") is not None
                and controlled_token(item.get("stage")) != item.get("stage")
            )
            or (
                item.get("fallback") is not None
                and controlled_token(item.get("fallback")) != item.get("fallback")
            )
            or (
                "recoverable" in item
                and type(item.get("recoverable")) is not bool
            )
        ):
            raise ProtocolError("Long-form generation result is invalid")
    for item in children:
        issue_code = item.get("issueCode") if isinstance(item, dict) else None
        status = item.get("status") if isinstance(item, dict) else None
        if (
            not isinstance(item, dict)
            or set(item) - {"nodeId", "status", "issueCode"}
            or not isinstance(item.get("nodeId"), str)
            or redact_private_text(item.get("nodeId")) != item.get("nodeId")
            or status not in {"applied", "degraded", "skipped"}
            or (
                issue_code is not None
                and not _is_closed_issue_code(issue_code)
            )
            or (status == "applied" and issue_code is not None)
            or (status != "applied" and issue_code is None)
        ):
            raise ProtocolError("Long-form generation result is invalid")
    _validate_pagination_map(pagination, "Long-form generation result is invalid")
    # Reject sensitive response key names recursively.  Field snapshot hashes
    # are deliberately allowed; raw visible field results are not.
    def walk(value: Any) -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                lowered = str(key).lower()
                if lowered in {
                    "path", "resourceid", "resources", "resourcesha256",
                    "payloadsha256", "sourcesha256", "bookmarkmap",
                    "bookmarkmapping", "fieldresult", "visibletext",
                }:
                    raise ProtocolError("Long-form generation result is invalid")
                walk(child)
        elif isinstance(value, (list, tuple)):
            for child in value:
                walk(child)

    walk({key: value for key, value in raw.items() if key != "outputPath"})
    result = dict(raw)
    result.setdefault("appliedOperations", applied)
    result.setdefault("childResults", list(children))
    return result


def _positive_integer(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def _valid_bounds(value: Any) -> bool:
    if value is None:
        return True
    if not isinstance(value, list) or len(value) != 4:
        return False
    if any(
        not isinstance(item, (int, float))
        or isinstance(item, bool)
        or not math.isfinite(float(item))
        for item in value
    ):
        return False
    x0, y0, x1, y1 = (float(item) for item in value)
    return x0 >= 0 and y0 >= 0 and x1 > x0 and y1 > y0


def _validate_pagination_map(value: Any, message: str) -> None:
    if not isinstance(value, dict) or set(value) - {"version", "nodes"}:
        raise ProtocolError(message)
    version = value.get("version")
    nodes = value.get("nodes", [])
    if version not in {"M2-stub", "M5-v1"} or not isinstance(nodes, list):
        raise ProtocolError(message)
    for node in nodes:
        if not isinstance(node, dict) or not isinstance(node.get("nodeId"), str):
            raise ProtocolError(message)
        if redact_private_text(node["nodeId"]) != node["nodeId"]:
            raise ProtocolError(message)
        fragments = node.get("fragments")
        if not isinstance(fragments, list):
            raise ProtocolError(message)
        if version == "M2-stub":
            if set(node) - {"nodeId", "fragments"}:
                raise ProtocolError(message)
        else:
            if set(node) - {
                "nodeId", "story", "sections", "pageStart", "pageEnd", "range", "fragments",
            }:
                raise ProtocolError(message)
            if node.get("story") is not None and node.get("story") not in {
                "main", "header", "footer", "footnote",
            }:
                raise ProtocolError(message)
            if not isinstance(node.get("sections", []), list) or any(
                not isinstance(item, str) or controlled_token(item) != item
                for item in node.get("sections", [])
            ):
                raise ProtocolError(message)
            start = node.get("pageStart")
            end = node.get("pageEnd")
            if start is not None and not _positive_integer(start):
                raise ProtocolError(message)
            if end is not None and not _positive_integer(end):
                raise ProtocolError(message)
            if start is not None and end is not None and end < start:
                raise ProtocolError(message)
            range_value = node.get("range")
            if range_value is not None and not re.fullmatch(r"[0-9]+:[0-9]+", str(range_value)):
                raise ProtocolError(message)
        for fragment in fragments:
            if not isinstance(fragment, dict) or not _positive_integer(fragment.get("page")):
                raise ProtocolError(message)
            if version == "M2-stub":
                if set(fragment) != {"page"}:
                    raise ProtocolError(message)
            elif set(fragment) - {"page", "bounds"} or not _valid_bounds(fragment.get("bounds")):
                raise ProtocolError(message)


def validate_longform_notice_patch_request(raw: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the one-shot, privacy-bounded notice patch envelope."""

    if not isinstance(raw, Mapping) or set(raw) != _NOTICE_PATCH_REQUEST_KEYS:
        raise ProtocolError("Long-form notice patch request is invalid")
    source = raw.get("sourcePath")
    output = raw.get("outputPath")
    notices = raw.get("notices")
    if (
        not isinstance(source, str) or not source
        or not isinstance(output, str) or not output
        or not isinstance(notices, list) or not notices
    ):
        raise ProtocolError("Long-form notice patch request is invalid")
    for notice in notices:
        if (
            not isinstance(notice, dict)
            or set(notice) != {
                "code", "message", "fallback", "placement", "nodeId", "page", "bookmarkName",
            }
            or not _is_closed_issue_code(notice.get("code"))
            or not isinstance(notice.get("message"), str)
            or redact_private_text(notice["message"]) != notice["message"]
            or controlled_token(notice.get("fallback")) != notice.get("fallback")
            or notice.get("placement") not in {"block", "document"}
            or not isinstance(notice.get("nodeId"), str)
            or redact_private_text(notice["nodeId"]) != notice["nodeId"]
            or not _positive_integer(notice.get("page"))
            or (
                notice.get("bookmarkName") is not None
                and (
                    not isinstance(notice.get("bookmarkName"), str)
                    or not _BOOKMARK_RE.fullmatch(notice["bookmarkName"])
                )
            )
            or (notice.get("placement") == "block" and notice.get("bookmarkName") is None)
        ):
            raise ProtocolError("Long-form notice patch request is invalid")
    return {"sourcePath": source, "outputPath": output, "notices": [dict(item) for item in notices]}


def validate_longform_notice_patch_value(raw: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the closed result returned by a notice-only native patch."""

    if not isinstance(raw, Mapping) or set(raw) - _NOTICE_PATCH_RESULT_KEYS:
        raise ProtocolError("Long-form notice patch result is invalid")
    output = raw.get("outputPath")
    applied = raw.get("appliedNotices")
    issues = raw.get("issueCodes", [])
    history = raw.get("fieldSnapshots", [])
    pagination = raw.get("paginationMap", {})
    if (
        not isinstance(output, str) or not output
        or not _positive_integer(applied)
        or not isinstance(issues, list)
        or not isinstance(history, list)
    ):
        raise ProtocolError("Long-form notice patch result is invalid")
    try:
        validated = validate_longform_generation_value({
            "outputPath": output,
            "appliedOperations": applied,
            "issueCodes": issues,
            "fieldSnapshots": history,
            "childResults": [],
            "paginationMap": pagination,
        })
    except ProtocolError:
        raise ProtocolError("Long-form notice patch result is invalid") from None
    return {
        "outputPath": output,
        "appliedNotices": applied,
        "issueCodes": list(validated.get("issueCodes", [])),
        "fieldSnapshots": list(validated.get("fieldSnapshots", [])),
        "paginationMap": dict(validated.get("paginationMap", {})),
    }


def _is_closed_issue_code(value: Any) -> bool:
    return (
        isinstance(value, str)
        and 3 <= len(value) <= 64
        and value[0].isalpha()
        and all(
            character in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_"
            for character in value
        )
    )


@dataclass(frozen=True)
class ProbeCommand:
    id: str
    component: str
    method: str
    params: Mapping[str, Any]

    @classmethod
    def create(
        cls, component: str, method: str, params: Mapping[str, Any]
    ) -> "ProbeCommand":
        if component not in COMPONENTS:
            raise ProtocolError(f"Unsupported component: {component}")
        if method not in METHOD_COMPONENT:
            raise ProtocolError(f"Unsupported method: {method}")
        required = METHOD_COMPONENT[method]
        if required is not None and required != component:
            raise ProtocolError(f"{method} requires {required}, got {component}")
        return cls(uuid4().hex, component, method, dict(params))

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "component": self.component,
            "method": self.method,
            "params": dict(self.params),
        }


@dataclass(frozen=True)
class ProbeResult:
    id: str
    ok: bool
    value: Mapping[str, Any]
    error: Optional[Mapping[str, Any]]

    @classmethod
    def from_dict(cls, raw: Mapping[str, Any]) -> "ProbeResult":
        command_id = str(raw.get("id", ""))
        if not command_id:
            raise ProtocolError("Result id is required")
        ok = raw.get("ok")
        if not isinstance(ok, bool):
            raise ProtocolError("Result ok must be boolean")
        value = raw.get("value") or {}
        error = raw.get("error")
        if not isinstance(value, dict):
            raise ProtocolError("Result value must be an object")
        if error is not None and not isinstance(error, dict):
            raise ProtocolError("Result error must be an object or null")
        return cls(command_id, ok, value, error)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "ok": self.ok,
            "value": dict(self.value),
            "error": None if self.error is None else dict(self.error),
        }


class PathPolicy:
    def __init__(self, roots: tuple[Path, ...]):
        self._roots = tuple(root.resolve() for root in roots)

    def require_allowed(self, value: Union[str, Path]) -> Path:
        candidate = Path(value).expanduser().resolve()
        if not any(candidate == root or root in candidate.parents for root in self._roots):
            raise ProtocolError(f"Path is outside allowed roots: {candidate}")
        return candidate
