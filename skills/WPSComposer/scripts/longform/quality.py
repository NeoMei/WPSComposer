"""Typed, platform-pure contracts for the M5 PDF quality gate."""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import math
import re
from typing import Any, Mapping, Optional, Tuple

from .privacy import redact_private_text


_CODE_RE = re.compile(r"^[A-Z][A-Z0-9_]{1,63}$")
_TOKEN_RE = re.compile(r"^[a-z][a-z0-9-]{0,63}$")
_EVIDENCE_KEYS = frozenset({
    "actualDpi",
    "availableHeightPoints",
    "availableWidthPoints",
    "blankCharacterCount",
    "captionPage",
    "fieldCount",
    "figureIndexPageCount",
    "finalPageUtilization",
    "headerDisplayUnits",
    "objectPage",
    "overflowPoints",
    "pageRole",
    "refreshRounds",
    "tableIndexPageCount",
    "tocPageCount",
    "totalPages",
    "utilization",
})
_EVIDENCE_STRING_VALUES = frozenset({
    "bibliography",
    "body",
    "chapter_start",
    "cover",
    "explicit_break",
    "front_matter",
    "landscape",
    "toc",
})


class QualitySeverity(str, Enum):
    FATAL = "fatal"
    DEGRADED = "degraded"
    WARNING = "warning"
    INFO = "info"


class QualityConfidence(str, Enum):
    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"


class PageRole(str, Enum):
    COVER = "cover"
    FRONT_MATTER = "front_matter"
    TOC = "toc"
    CHAPTER_START = "chapter_start"
    BODY = "body"
    LANDSCAPE = "landscape"
    BIBLIOGRAPHY = "bibliography"
    EXPLICIT_BREAK = "explicit_break"


def normalize_bounds(value: Any) -> Optional[Tuple[float, float, float, float]]:
    """Validate and normalize one top-left point-coordinate bounding box."""

    if value is None:
        return None
    if isinstance(value, Mapping):
        try:
            value = (value["x0"], value["y0"], value["x1"], value["y1"])
        except KeyError as exc:
            raise ValueError("bounds mapping must contain x0, y0, x1, and y1") from exc
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        raise TypeError("bounds must contain exactly four point values")
    result = tuple(float(item) for item in value)
    if not all(math.isfinite(item) for item in result):
        raise ValueError("bounds values must be finite")
    x0, y0, x1, y1 = result
    if x0 < 0 or y0 < 0 or x1 <= x0 or y1 <= y0:
        raise ValueError("bounds must be a positive non-inverted rectangle")
    return result  # type: ignore[return-value]


@dataclass(frozen=True)
class QualityEvidence:
    """Closed, privacy-safe evidence attached to a quality finding."""

    values: Tuple[Tuple[str, Any], ...] = ()

    @classmethod
    def from_mapping(cls, raw: Optional[Mapping[str, Any]]) -> "QualityEvidence":
        items = []
        for key, value in sorted((raw or {}).items()):
            if key not in _EVIDENCE_KEYS:
                raise ValueError(f"uncontrolled quality evidence key: {key}")
            if isinstance(value, bool):
                normalized = value
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                normalized = float(value)
                if not math.isfinite(normalized):
                    raise ValueError("quality evidence numbers must be finite")
            elif isinstance(value, str) and value in _EVIDENCE_STRING_VALUES:
                normalized = value
            else:
                raise ValueError("quality evidence contains a private or uncontrolled value")
            items.append((key, normalized))
        return cls(tuple(items))

    def to_dict(self) -> dict[str, Any]:
        return dict(self.values)


@dataclass(frozen=True)
class QualityFinding:
    """One deterministic quality observation mapped to a page/node when possible."""

    code: str
    severity: QualitySeverity
    confidence: QualityConfidence
    message: str
    node_id: Optional[str] = None
    page: Optional[int] = None
    bounds: Optional[Tuple[float, float, float, float]] = None
    evidence: QualityEvidence = field(default_factory=QualityEvidence)
    repair_key: Optional[str] = None

    def __post_init__(self) -> None:
        if not _CODE_RE.fullmatch(self.code):
            raise ValueError("quality code must be a controlled uppercase token")
        object.__setattr__(self, "severity", QualitySeverity(self.severity))
        object.__setattr__(self, "confidence", QualityConfidence(self.confidence))
        object.__setattr__(self, "message", redact_private_text(str(self.message)))
        if self.page is not None and (type(self.page) is not int or self.page <= 0):
            raise ValueError("quality page must be a positive integer")
        object.__setattr__(self, "bounds", normalize_bounds(self.bounds))
        if self.bounds is not None and self.page is None:
            raise ValueError("bounded quality findings require a page")
        if self.repair_key is not None and not _TOKEN_RE.fullmatch(self.repair_key):
            raise ValueError("repair key must be a controlled token")
        if not isinstance(self.evidence, QualityEvidence):
            object.__setattr__(
                self, "evidence", QualityEvidence.from_mapping(self.evidence)
            )

    def sort_key(self) -> tuple[Any, ...]:
        return (
            self.page or 0,
            self.node_id or "",
            self.code,
            self.severity.value,
            self.confidence.value,
        )

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "code": self.code,
            "severity": self.severity.value,
            "confidence": self.confidence.value,
            "message": self.message,
            "evidence": self.evidence.to_dict(),
        }
        if self.node_id is not None:
            result["nodeId"] = redact_private_text(self.node_id)
        if self.page is not None:
            result["page"] = self.page
        if self.bounds is not None:
            result["bounds"] = list(self.bounds)
        if self.repair_key is not None:
            result["repairKey"] = self.repair_key
        return result

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "QualityFinding":
        return cls(
            code=str(data["code"]),
            severity=QualitySeverity(data["severity"]),
            confidence=QualityConfidence(data["confidence"]),
            message=str(data["message"]),
            node_id=data.get("nodeId"),
            page=data.get("page"),
            bounds=data.get("bounds"),
            evidence=QualityEvidence.from_mapping(data.get("evidence")),
            repair_key=data.get("repairKey"),
        )


@dataclass(frozen=True)
class QualityReport:
    findings: Tuple[QualityFinding, ...] = ()
    page_count: int = 0

    def __post_init__(self) -> None:
        if type(self.page_count) is not int or self.page_count < 0:
            raise ValueError("page_count must be a non-negative integer")
        object.__setattr__(self, "findings", tuple(sorted(self.findings, key=lambda i: i.sort_key())))

    @property
    def degraded(self) -> bool:
        return any(
            item.severity in {QualitySeverity.FATAL, QualitySeverity.DEGRADED}
            for item in self.findings
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "pageCount": self.page_count,
            "findings": [item.to_dict() for item in self.findings],
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "QualityReport":
        return cls(
            findings=tuple(
                QualityFinding.from_dict(item) for item in data.get("findings", ())
            ),
            page_count=int(data.get("pageCount", 0)),
        )


@dataclass(frozen=True)
class GenerationOutcome:
    """Private public-route result; `generate()` still returns only `path`."""

    path: str
    issues: Tuple[QualityFinding, ...] = ()

    @property
    def degraded(self) -> bool:
        return any(
            item.severity in {QualitySeverity.FATAL, QualitySeverity.DEGRADED}
            for item in self.issues
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "degraded": self.degraded,
            "issues": [item.to_dict() for item in self.issues],
        }


__all__ = [
    "GenerationOutcome",
    "PageRole",
    "QualityConfidence",
    "QualityEvidence",
    "QualityFinding",
    "QualityReport",
    "QualitySeverity",
    "normalize_bounds",
]
