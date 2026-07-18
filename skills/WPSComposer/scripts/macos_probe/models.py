from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Optional, Union
from uuid import uuid4

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
}


class ProtocolError(ValueError):
    """Raised when a bridge message violates the probe protocol."""


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
