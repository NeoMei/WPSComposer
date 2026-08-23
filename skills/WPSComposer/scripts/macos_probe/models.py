from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
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
            or status not in {"applied", "degraded", "skipped"}
            or (
                issue_code is not None
                and not _is_closed_issue_code(issue_code)
            )
            or (status == "applied" and issue_code is not None)
            or (status != "applied" and issue_code is None)
        ):
            raise ProtocolError("Long-form generation result is invalid")
    if set(pagination) - {"version", "nodes"}:
        raise ProtocolError("Long-form generation result is invalid")
    if pagination.get("version") != "M2-stub" or not isinstance(
        pagination.get("nodes", []), list
    ):
        raise ProtocolError("Long-form generation result is invalid")
    for node in pagination.get("nodes", []):
        if (
            not isinstance(node, dict)
            or set(node) - {"nodeId", "fragments"}
            or not isinstance(node.get("nodeId"), str)
            or not isinstance(node.get("fragments"), list)
        ):
            raise ProtocolError("Long-form generation result is invalid")
        for fragment in node["fragments"]:
            if (
                not isinstance(fragment, dict)
                or set(fragment) != {"page"}
                or not isinstance(fragment.get("page"), int)
                or isinstance(fragment.get("page"), bool)
                or fragment["page"] < 1
            ):
                raise ProtocolError("Long-form generation result is invalid")
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
