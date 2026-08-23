"""Shared long-form executor interface and field-refresh convergence loop.

This module is pure: it does not import WPS/COM/JSAPI/subprocess modules.
Concrete platform executors (Windows, macOS) live in sibling modules.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional, Protocol, Tuple, runtime_checkable

from ..generation_plan import GenerationPlan
from .degradation import controlled_token, redact_private_text
from .resources import PreparedLongformResource


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
    stage: Optional[str] = None
    fallback: Optional[str] = None
    recoverable: Optional[bool] = None

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "code": self.code,
            "message": redact_private_text(self.message),
            "placement": self.placement,
        }
        if self.node_id is not None:
            result["nodeId"] = redact_private_text(self.node_id)
        stage = controlled_token(self.stage)
        fallback = controlled_token(self.fallback)
        if stage is not None:
            result["stage"] = stage
        if fallback is not None:
            result["fallback"] = fallback
        if type(self.recoverable) is bool:
            result["recoverable"] = self.recoverable
        return result

    def __repr__(self) -> str:
        return f"ExecutionIssue({self.to_dict()!r})"

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ExecutionIssue:
        return cls(
            code=str(data["code"]),
            message=str(data["message"]),
            placement=str(data.get("placement", "document")),
            node_id=data.get("nodeId"),
            stage=controlled_token(data.get("stage")),
            fallback=controlled_token(data.get("fallback")),
            recoverable=(
                data.get("recoverable")
                if type(data.get("recoverable")) is bool
                else None
            ),
        )


@dataclass(frozen=True)
class ExecutionOutcome:
    """Result of executing a long-form generation plan."""

    staged_artifact: str
    issues: Tuple[ExecutionIssue, ...] = ()
    pagination_map: PaginationMap = field(default_factory=PaginationMap)
    applied_operations: Optional[int] = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "stagedArtifact": self.staged_artifact,
            "issues": [i.to_dict() for i in self.issues],
            "paginationMap": self.pagination_map.to_dict(),
            "appliedOperations": self.applied_operations,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ExecutionOutcome:
        return cls(
            staged_artifact=str(data["stagedArtifact"]),
            issues=tuple(ExecutionIssue.from_dict(i) for i in (data.get("issues") or ())),
            pagination_map=PaginationMap.from_dict(data.get("paginationMap") or {}),
            applied_operations=data.get("appliedOperations"),
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
        resources: Tuple[PreparedLongformResource, ...],
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


def finalize_fields_with_convergence(executor: Any, max_rounds: int = 3) -> ConvergenceResult:
    """Compatibility facade delegating M2 refreshers to the M3 field contract.

    M3 native executors should call ``finalize_native_fields`` with a complete
    five-phase adapter.  The legacy facade retains M2's ``refresh_fields``
    shape and outcome behavior while sharing the bounded convergence engine.
    """
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

    from .field_contract import _finalize_native_fields
    upsert_notice = getattr(executor, "upsert_document_quality_notice", None)

    class _LegacyRefreshAdapter:
        _wpsc_legacy_snapshot = True

        def __init__(self) -> None:
            self.round_index = 0
            self.cached_snapshot: Tuple[FieldSnapshot, ...] = ()

        def repaginate_and_update_numbering(self) -> None:
            raw = refresh(self.round_index)
            self.round_index += 1
            self.cached_snapshot = tuple(raw or ())

        def refresh_bookmarks_and_references(self) -> None:
            return None

        def refresh_indexes(self) -> None:
            return None

        def repaginate_and_update_page_fields(self) -> None:
            return None

        def snapshot_fields(self) -> Tuple[FieldSnapshot, ...]:
            return self.cached_snapshot

        def upsert_document_quality_notice(self, issue: ExecutionIssue) -> None:
            if not callable(upsert_notice):
                raise RuntimeError("visible quality-notice API unavailable")
            upsert_notice(issue)

    return _finalize_native_fields(
        _LegacyRefreshAdapter(),
        max_rounds,
        allow_legacy_hashes=True,
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
        self.quality_notices: list[ExecutionIssue] = []
        self._outcome_index = 0

    def execute(
        self,
        plan: GenerationPlan,
        resources: Tuple[PreparedLongformResource, ...] = (),
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

    def upsert_document_quality_notice(self, issue: ExecutionIssue) -> None:
        """Record the visible notice mutation required by unstable convergence."""

        self.quality_notices.append(issue)


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
