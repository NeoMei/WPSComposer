"""Closed, deterministic planner for the single M5 full-relayout pass."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional, Tuple

from .quality import (
    QualityConfidence,
    QualityFinding,
    QualityReport,
    QualitySeverity,
)


_RELAYOUT_MATRIX: dict[str, tuple[int, str, dict[str, Any], bool]] = {
    "compact-toc": (10, "compact-toc", {"minimumOnly": True}, False),
    "shorten-header": (20, "shorten-header", {"maxDisplayUnits": 32}, False),
    "fit-image": (30, "fit-image", {"fitToBodyWidth": True}, True),
    "compress-table": (40, "compress-table", {"attempts": 1}, True),
    "force-table-split": (
        50,
        "force-table-split",
        {"removeVerticalMerge": True},
        True,
    ),
    "keep-heading": (60, "keep-heading", {"minimumFollowingLines": 2}, True),
    "keep-caption": (70, "keep-caption", {"samePage": True}, True),
    "remove-unexpected-blank": (
        80,
        "remove-unexpected-blank",
        {"maximumPages": 1},
        False,
    ),
}
_NOTICE_ONLY = frozenset({"low-dpi-notice"})


@dataclass(frozen=True)
class RelayoutDirective:
    kind: str
    node_id: Optional[str]
    args: dict[str, Any]

    def to_dict(self) -> dict[str, Any]:
        result = {"kind": self.kind, "args": dict(self.args)}
        if self.node_id is not None:
            result["nodeId"] = self.node_id
        return result


@dataclass(frozen=True)
class RelayoutDecision:
    directives: Tuple[RelayoutDirective, ...] = ()
    notice_findings: Tuple[QualityFinding, ...] = ()

    @property
    def requires_relayout(self) -> bool:
        return bool(self.directives)


def _eligible(finding: QualityFinding) -> bool:
    return (
        finding.severity is QualitySeverity.DEGRADED
        and finding.confidence is QualityConfidence.HIGH
    )


def build_relayout_directives(
    report: QualityReport,
    *,
    relayout_already_used: bool = False,
) -> RelayoutDecision:
    """Map high-confidence findings through the closed repair matrix once."""

    eligible = tuple(item for item in report.findings if _eligible(item))
    if relayout_already_used:
        return RelayoutDecision(notice_findings=eligible)

    directives: list[tuple[int, RelayoutDirective]] = []
    notices: list[QualityFinding] = []
    seen: set[tuple[str, Optional[str]]] = set()
    for finding in eligible:
        repair_key = finding.repair_key
        if repair_key in _NOTICE_ONLY:
            notices.append(finding)
            continue
        matrix = _RELAYOUT_MATRIX.get(repair_key or "")
        if matrix is None:
            notices.append(finding)
            continue
        priority, kind, args, requires_node = matrix
        if requires_node and not finding.node_id:
            raise ValueError(f"{kind} relayout requires a mapped node_id")
        identity = (kind, finding.node_id)
        if identity in seen:
            continue
        seen.add(identity)
        directives.append(
            (
                priority,
                RelayoutDirective(
                    kind=kind,
                    node_id=finding.node_id,
                    args=dict(args),
                ),
            )
        )

    directives.sort(key=lambda item: (item[0], item[1].node_id or "", item[1].kind))
    notices.sort(key=lambda item: item.sort_key())
    return RelayoutDecision(
        directives=tuple(item[1] for item in directives),
        notice_findings=tuple(notices),
    )


__all__ = [
    "RelayoutDecision",
    "RelayoutDirective",
    "build_relayout_directives",
]
