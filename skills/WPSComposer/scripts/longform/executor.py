"""Shared long-form executor interface and field-refresh convergence loop.

This module is pure: it does not import WPS/COM/JSAPI/subprocess modules.
Concrete platform executors (Windows, macOS) live in sibling modules.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional, Protocol, Tuple, runtime_checkable

from ..generation_plan import GenerationPlan, GenerationResource


FIELD_REFRESH_UNSTABLE = "FIELD_REFRESH_UNSTABLE"


# ---------------------------------------------------------------------------
# Pagination map (M2 stub: typed and serializable)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class PaginationFragment:
    """A single physical page fragment for a pagination node."""

    page: int
    bounds: Optional[dict[str, Any]] = None

    def to_dict(self) -> dict[str, Any]:
        return {"page": self.page, "bounds": self.bounds}

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PaginationFragment:
        return cls(page=int(data["page"]), bounds=data.get("bounds"))


@dataclass(frozen=True)
class PaginationNode:
    """Pagination information for one semantic node."""

    node_id: str
    story: Optional[str] = None
    sections: Tuple[str, ...] = ()
    page_start: Optional[int] = None
    page_end: Optional[int] = None
    range: Optional[str] = None
    fragments: Tuple[PaginationFragment, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "nodeId": self.node_id,
            "story": self.story,
            "sections": list(self.sections),
            "pageStart": self.page_start,
            "pageEnd": self.page_end,
            "range": self.range,
            "fragments": [f.to_dict() for f in self.fragments],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PaginationNode:
        fragments = data.get("fragments") or ()
        return cls(
            node_id=str(data["nodeId"]),
            story=data.get("story"),
            sections=tuple(str(s) for s in (data.get("sections") or ())),
            page_start=_optional_int(data.get("pageStart")),
            page_end=_optional_int(data.get("pageEnd")),
            range=data.get("range"),
            fragments=tuple(PaginationFragment.from_dict(f) for f in fragments),
        )


def _optional_int(value: Any) -> Optional[int]:
    return None if value is None else int(value)


@dataclass(frozen=True)
class PaginationMap:
    """Stub pagination map for M2; real mapping arrives in M5."""

    version: str = "M2-stub"
    nodes: Tuple[PaginationNode, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "nodes": [n.to_dict() for n in self.nodes],
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> PaginationMap:
        nodes = data.get("nodes") or ()
        return cls(
            version=str(data.get("version", "M2-stub")),
            nodes=tuple(PaginationNode.from_dict(n) for n in nodes),
        )


# ---------------------------------------------------------------------------
# Execution issues and outcome
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ExecutionIssue:
    """A deterministic, serializable runtime issue from an executor."""

    code: str
    message: str
    placement: str = "document"
    node_id: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "code": self.code,
            "message": self.message,
            "placement": self.placement,
        }
        if self.node_id is not None:
            result["nodeId"] = self.node_id
        return result

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ExecutionIssue:
        return cls(
            code=str(data["code"]),
            message=str(data["message"]),
            placement=str(data.get("placement", "document")),
            node_id=data.get("nodeId"),
        )


@dataclass(frozen=True)
class ExecutionOutcome:
    """Result of executing a long-form generation plan."""

    staged_artifact: str
    issues: Tuple[ExecutionIssue, ...] = ()
    pagination_map: PaginationMap = field(default_factory=PaginationMap)

    def to_dict(self) -> dict[str, Any]:
        return {
            "stagedArtifact": self.staged_artifact,
            "issues": [i.to_dict() for i in self.issues],
            "paginationMap": self.pagination_map.to_dict(),
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ExecutionOutcome:
        return cls(
            staged_artifact=str(data["stagedArtifact"]),
            issues=tuple(ExecutionIssue.from_dict(i) for i in (data.get("issues") or ())),
            pagination_map=PaginationMap.from_dict(data.get("paginationMap") or {}),
        )


# ---------------------------------------------------------------------------
# LongformExecutor protocol
# ---------------------------------------------------------------------------

@runtime_checkable
class LongformExecutor(Protocol):
    """Protocol for a platform-independent long-form executor."""

    def execute(
        self,
        plan: GenerationPlan,
        resources: Tuple[GenerationResource, ...],
        deadline: Optional[float] = None,
    ) -> ExecutionOutcome:
        ...


# ---------------------------------------------------------------------------
# Field-refresh convergence
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class FieldSnapshot:
    """Snapshot of one field at the end of a finalize_fields round."""

    stable_key: Tuple[str, str, int]
    field_category: str
    result_hash: str
    toc_page_count: int
    figure_index_page_count: int
    table_index_page_count: int
    total_pages: int

    def __lt__(self, other: FieldSnapshot) -> bool:
        return self.stable_key < other.stable_key

    def to_dict(self) -> dict[str, Any]:
        return {
            "stableKey": list(self.stable_key),
            "fieldCategory": self.field_category,
            "resultHash": self.result_hash,
            "tocPageCount": self.toc_page_count,
            "figureIndexPageCount": self.figure_index_page_count,
            "tableIndexPageCount": self.table_index_page_count,
            "totalPages": self.total_pages,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> FieldSnapshot:
        stable_key = data["stableKey"]
        return cls(
            stable_key=(str(stable_key[0]), str(stable_key[1]), int(stable_key[2])),
            field_category=str(data["fieldCategory"]),
            result_hash=str(data["resultHash"]),
            toc_page_count=int(data["tocPageCount"]),
            figure_index_page_count=int(data["figureIndexPageCount"]),
            table_index_page_count=int(data["tableIndexPageCount"]),
            total_pages=int(data["totalPages"]),
        )


@dataclass(frozen=True)
class ConvergenceResult:
    """Result of the field-refresh convergence loop."""

    snapshot: Tuple[FieldSnapshot, ...] = ()
    issues: Tuple[ExecutionIssue, ...] = ()
    rounds: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "snapshot": [s.to_dict() for s in self.snapshot],
            "issues": [i.to_dict() for i in self.issues],
            "rounds": self.rounds,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConvergenceResult:
        return cls(
            snapshot=tuple(FieldSnapshot.from_dict(s) for s in (data.get("snapshot") or ())),
            issues=tuple(ExecutionIssue.from_dict(i) for i in (data.get("issues") or ())),
            rounds=int(data.get("rounds", 0)),
        )


def _unstable_message(snapshot: Tuple[FieldSnapshot, ...], rounds: int) -> str:
    """Build the deterministic FIELD_REFRESH_UNSTABLE issue message."""
    field_count = len(snapshot)
    if snapshot:
        representative = snapshot[-1]
        toc_pages = representative.toc_page_count
        figure_pages = representative.figure_index_page_count
        table_pages = representative.table_index_page_count
        total_pages = representative.total_pages
    else:
        toc_pages = figure_pages = table_pages = total_pages = 0
    return (
        f"Field refresh did not converge after {rounds} rounds; "
        f"fields={field_count}, toc_pages={toc_pages}, "
        f"figure_index_pages={figure_pages}, table_index_pages={table_pages}, "
        f"total_pages={total_pages}"
    )


def _normalize_snapshot(snapshot: Tuple[FieldSnapshot, ...]) -> Tuple[FieldSnapshot, ...]:
    """Return a deterministically ordered snapshot for comparison."""
    return tuple(sorted(snapshot))


def finalize_fields_with_convergence(executor: Any, max_rounds: int = 3) -> ConvergenceResult:
    """Run the fixed 5-step field-refresh sequence until convergence or a final round."""
    refresh = getattr(executor, "refresh_fields", None)
    if not callable(refresh):
        return ConvergenceResult(
            snapshot=(),
            issues=(
                ExecutionIssue(
                    code="EXECUTOR_CAPABILITY_MISSING",
                    message="Executor does not provide a callable refresh_fields method",
                    placement="document",
                ),
            ),
            rounds=0,
        )

    if max_rounds < 1:
        max_rounds = 1

    previous: Optional[Tuple[FieldSnapshot, ...]] = None

    for round_index in range(max_rounds):
        raw = refresh(round_index)
        if raw is None:
            raw = ()
        try:
            snapshot = _normalize_snapshot(tuple(raw))
        except TypeError:
            return ConvergenceResult(
                snapshot=(),
                issues=(
                    ExecutionIssue(
                        code="FIELD_REFRESH_SNAPSHOT_INVALID",
                        message=f"refresh_fields returned non-comparable snapshot at round {round_index}",
                        placement="document",
                    ),
                ),
                rounds=round_index + 1,
            )
        if previous is not None and snapshot == previous:
            return ConvergenceResult(
                snapshot=snapshot,
                issues=(),
                rounds=round_index + 1,
            )
        previous = snapshot

    # Deterministic 4th round (when max_rounds == 3) that freezes the result.
    final_round_index = max_rounds
    final_raw = refresh(final_round_index)
    if final_raw is None:
        final_raw = ()
    try:
        final_snapshot = _normalize_snapshot(tuple(final_raw))
    except TypeError:
        final_snapshot = previous if previous is not None else ()

    issue = ExecutionIssue(
        code=FIELD_REFRESH_UNSTABLE,
        message=_unstable_message(final_snapshot, max_rounds + 1),
        placement="document",
    )
    return ConvergenceResult(
        snapshot=final_snapshot,
        issues=(issue,),
        rounds=max_rounds + 1,
    )


# ---------------------------------------------------------------------------
# Recording executor
# ---------------------------------------------------------------------------

class RecordingLongformExecutor:
    """Deterministic executor that records calls and returns stub outcomes."""

    def __init__(
        self,
        *,
        artifact: str = "staged.docx",
        outcomes: Optional[Tuple[ExecutionOutcome, ...]] = None,
        snapshots: Optional[Tuple[Tuple[FieldSnapshot, ...], ...]] = None,
    ) -> None:
        self.artifact = artifact
        self._outcomes = outcomes
        self._snapshots = snapshots or ()
        self.calls: list[Tuple[GenerationPlan, Tuple[Any, ...], Optional[float]]] = []
        self.refresh_calls: list[int] = []
        self._outcome_index = 0

    def execute(
        self,
        plan: GenerationPlan,
        resources: Tuple[Any, ...] = (),
        deadline: Optional[float] = None,
    ) -> ExecutionOutcome:
        self.calls.append((plan, resources, deadline))
        if self._outcomes:
            outcome = self._outcomes[self._outcome_index]
            self._outcome_index = (self._outcome_index + 1) % len(self._outcomes)
            return outcome
        return ExecutionOutcome(
            staged_artifact=self.artifact,
            issues=(),
            pagination_map=PaginationMap(
                version="M2-stub",
                nodes=(
                    PaginationNode(
                        node_id="doc:finalize",
                        page_start=1,
                        page_end=1,
                        fragments=(PaginationFragment(page=1),),
                    ),
                ),
            ),
        )

    def refresh_fields(self, round_index: int) -> Tuple[FieldSnapshot, ...]:
        self.refresh_calls.append(round_index)
        if self._snapshots:
            return self._snapshots[round_index % len(self._snapshots)]
        return (
            FieldSnapshot(
                stable_key=("doc:finalize", "PAGE", 0),
                field_category="page",
                result_hash="stable",
                toc_page_count=0,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=1,
            ),
        )


__all__ = [
    "FIELD_REFRESH_UNSTABLE",
    "ConvergenceResult",
    "ExecutionIssue",
    "ExecutionOutcome",
    "FieldSnapshot",
    "LongformExecutor",
    "PaginationFragment",
    "PaginationMap",
    "PaginationNode",
    "RecordingLongformExecutor",
    "finalize_fields_with_convergence",
]
