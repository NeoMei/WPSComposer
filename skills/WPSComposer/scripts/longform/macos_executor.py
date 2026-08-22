"""macOS WPS JSAPI executor for long-form generation plans.

This module is importable without pywin32. It delegates operation execution to
the WPS JSAPI add-in over a LoopbackBridge, mirroring the Windows executor's
outcome shape and failure policies while keeping all native WPS interaction in
the add-in JavaScript.
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Mapping, Optional, Tuple

from ..generation_plan import GenerationPlan, validate_generation_plan
from ..macos_probe.bridge import LoopbackBridge
from ..macos_probe.models import ProbeResult
from .executor import (
    ExecutionIssue,
    ExecutionOutcome,
    FieldSnapshot,
    LongformExecutor,
    PaginationFragment,
    PaginationMap,
    PaginationNode,
)
from .field_contract import (
    NativeFieldContractError,
    evaluate_field_snapshot_history,
)
from .resources import PreparedLongformResource


MACOS_DEDICATED_HOST_UNAVAILABLE = "MACOS_DEDICATED_HOST_UNAVAILABLE"
EXECUTION_FAILED = "EXECUTION_FAILED"
EXECUTION_ABORTED = "EXECUTION_ABORTED"


class MacOSDedicatedHostUnavailableError(Exception):
    """Raised when the executor has no bridge to reach a dedicated WPS host."""

    def __init__(self, message: str = "Unable to reach a dedicated WPS host") -> None:
        super().__init__(message)
        self.code = MACOS_DEDICATED_HOST_UNAVAILABLE


class MacOSLongformExecutorError(Exception):
    """Engine-level error raised when execution must stop."""

    def __init__(self, message: str) -> None:
        super().__init__(message)


class MacOSLongformExecutor(LongformExecutor):
    """Execute a protocol v2 long-form generation plan on macOS WPS via JSAPI."""

    def __init__(
        self,
        *,
        bridge: Optional[LoopbackBridge] = None,
        staging_dir: Optional[str] = None,
    ) -> None:
        self._bridge = bridge
        self._staging_dir = staging_dir or tempfile.gettempdir()

    # ----------------------------------------------------------------------
    # Public interface
    # ----------------------------------------------------------------------
    def execute(
        self,
        plan: GenerationPlan,
        resources: Tuple[PreparedLongformResource, ...] = (),
        deadline: Optional[float] = None,
    ) -> ExecutionOutcome:
        validate_generation_plan(plan.to_dict(), component="writer")
        if self._bridge is None:
            raise MacOSDedicatedHostUnavailableError(
                "No LoopbackBridge available for macOS WPS execution"
            )

        paths = self._resolve_paths()
        staged_resources = self._stage_resources(resources, paths.staged_docx)
        resource_map = self._build_resource_map(staged_resources)
        params = {
            "plan": plan.to_dict(),
            "outputPath": paths.staged_docx,
            "resources": resource_map,
        }

        try:
            command = self._bridge.issue(
                "writer", "generate_longform_document", params
            )
            result = self._bridge.wait_result(command.id, timeout=300.0)
        except Exception as exc:
            return ExecutionOutcome(
                staged_artifact=paths.staged_docx,
                issues=(
                    ExecutionIssue(
                        code=EXECUTION_FAILED,
                        message=f"Bridge command failed: {exc}",
                        placement="document",
                    ),
                ),
                pagination_map=_build_pagination_map(plan.operations),
            )

        return self._build_outcome(result, paths.staged_docx, plan)

    # ----------------------------------------------------------------------
    # Helpers
    # ----------------------------------------------------------------------
    def _resolve_paths(self) -> "_ResolvedPaths":
        os.makedirs(self._staging_dir, exist_ok=True)
        base = tempfile.NamedTemporaryFile(
            prefix="wpsc-longform-",
            suffix="",
            dir=self._staging_dir,
            delete=False,
        ).name
        return _ResolvedPaths(staged_docx=base + ".docx")

    def _stage_resources(
        self,
        resources: Tuple[PreparedLongformResource, ...],
        staged_docx: str,
    ) -> Tuple[Tuple[str, Path], ...]:
        """Write normalized bytes into the private staging directory."""
        if not resources:
            return ()
        staging_dir = Path(staged_docx).parent
        staged: List[Tuple[str, Path]] = []
        for idx, resource in enumerate(resources):
            suffix = {
                "image/png": ".png",
                "image/jpeg": ".jpg",
                "image/tiff": ".tiff",
                "image/bmp": ".bmp",
                "image/gif": ".gif",
                "image/svg+xml": ".svg",
            }[resource.media_type]
            target = staging_dir / f"resource-{resource.id}-{idx}{suffix}"
            target.write_bytes(resource.payload_bytes)
            staged.append((resource.id, target))
        return tuple(staged)

    def _build_resource_map(
        self, resources: Tuple[Tuple[str, Path], ...]
    ) -> Mapping[str, str]:
        return {resource_id: str(path) for resource_id, path in resources}

    def _build_outcome(
        self,
        result: ProbeResult,
        staged_docx: str,
        plan: GenerationPlan,
    ) -> ExecutionOutcome:
        if not result.ok:
            error = result.error or {}
            code = error.get("code") or EXECUTION_FAILED
            message = error.get("message") or "WPS JSAPI command failed"
            if code == EXECUTION_ABORTED:
                raise MacOSLongformExecutorError(
                    f"Execution aborted at {error.get('opName') or 'unknown'}: {message}"
                )
            return ExecutionOutcome(
                staged_artifact=staged_docx,
                issues=(
                    ExecutionIssue(
                        code=code,
                        message=message,
                        placement="document",
                    ),
                ),
                pagination_map=_build_pagination_map(plan.operations),
            )

        value = result.value or {}
        staged_artifact = value.get("outputPath") or staged_docx
        issue_codes = value.get("issueCodes") or []
        pagination_map = _parse_pagination_map(
            value.get("paginationMap") or {}
        )
        issues: List[ExecutionIssue] = [
            _execution_issue(item) for item in issue_codes
        ]

        raw_history = value.get("fieldSnapshots")
        if (
            raw_history is not None
            and any(op.op == "writer.finalize_fields" for op in plan.operations)
        ):
            max_rounds = self._extract_max_rounds(plan)
            history = _parse_field_snapshot_history(raw_history)
            convergence = evaluate_field_snapshot_history(
                history,
                max_rounds=max_rounds,
                allow_legacy_hashes=True,
            )
            # The validated saved-state history is authoritative.  Discard any
            # stale/duplicate remote instability marker, then add the shared
            # result exactly once.
            issues = [
                issue for issue in issues if issue.code != "FIELD_REFRESH_UNSTABLE"
            ]
            issues = _merge_execution_issues(issues, convergence.issues)

        return ExecutionOutcome(
            staged_artifact=staged_artifact,
            issues=tuple(issues),
            pagination_map=pagination_map,
            applied_operations=value.get("appliedOperations"),
        )

    def _extract_max_rounds(self, plan: GenerationPlan) -> int:
        for op in plan.operations:
            if op.op == "writer.finalize_fields":
                return op.args.get("maxRounds", 3)
        return 3


@dataclass(frozen=True)
class _ResolvedPaths:
    staged_docx: str


def _execution_issue(raw: Mapping[str, Any]) -> ExecutionIssue:
    return ExecutionIssue(
        code=str(raw.get("code") or EXECUTION_FAILED),
        message=str(raw.get("message") or ""),
        placement=str(raw.get("placement") or "document"),
        node_id=raw.get("nodeId"),
    )


def _parse_field_snapshot_history(raw: Any) -> Tuple[Tuple[FieldSnapshot, ...], ...]:
    failed = False
    history: Tuple[Tuple[FieldSnapshot, ...], ...] = ()
    try:
        parsed = []
        for round_value in tuple(raw or ()):
            if isinstance(round_value, Mapping):
                parsed.append((FieldSnapshot.from_dict(dict(round_value)),))
            else:
                parsed.append(
                    tuple(
                        FieldSnapshot.from_dict(dict(item))
                        for item in tuple(round_value)
                    )
                )
        history = tuple(parsed)
    except Exception:
        failed = True
    if failed:
        raise NativeFieldContractError("field_history", "is invalid") from None
    return history


def _merge_execution_issues(
    existing: List[ExecutionIssue],
    additions: Tuple[ExecutionIssue, ...],
) -> List[ExecutionIssue]:
    merged = list(existing)
    for issue in additions:
        identity = (issue.code, issue.placement, issue.node_id)
        merged = [
            prior
            for prior in merged
            if (prior.code, prior.placement, prior.node_id) != identity
        ]
        merged.append(issue)
    return merged


def _parse_pagination_map(raw: Mapping[str, Any]) -> PaginationMap:
    version = str(raw.get("version") or "M2-stub")
    nodes = []
    for node in raw.get("nodes") or []:
        fragments = []
        for fragment in node.get("fragments") or []:
            fragments.append(
                PaginationFragment(
                    page=int(fragment.get("page", 1)),
                    bounds=fragment.get("bounds"),
                )
            )
        nodes.append(
            PaginationNode(
                node_id=str(node.get("nodeId") or ""),
                story=node.get("story"),
                sections=tuple(str(s) for s in (node.get("sections") or ())),
                page_start=_optional_int(node.get("pageStart")),
                page_end=_optional_int(node.get("pageEnd")),
                range=node.get("range"),
                fragments=tuple(fragments),
            )
        )
    return PaginationMap(version=version, nodes=tuple(nodes))


def _optional_int(value: Any) -> Optional[int]:
    return None if value is None else int(value)


def _build_pagination_map(
    operations: Tuple[Any, ...],
) -> PaginationMap:
    """Build the M2-stub pagination map from node ids in the plan."""
    nodes = []
    for op in operations:
        if not op.node_id:
            continue
        nodes.append(
            PaginationNode(
                node_id=op.node_id,
                fragments=(PaginationFragment(page=1),),
            )
        )
    return PaginationMap(version="M2-stub", nodes=tuple(nodes))


__all__ = [
    "EXECUTION_FAILED",
    "EXECUTION_ABORTED",
    "MACOS_DEDICATED_HOST_UNAVAILABLE",
    "MacOSDedicatedHostUnavailableError",
    "MacOSLongformExecutor",
    "MacOSLongformExecutorError",
]
