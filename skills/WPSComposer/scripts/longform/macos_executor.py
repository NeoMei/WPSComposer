"""macOS WPS JSAPI executor for long-form generation plans.

This module is importable without pywin32. It delegates operation execution to
the WPS JSAPI add-in over a LoopbackBridge, mirroring the Windows executor's
outcome shape and failure policies while keeping all native WPS interaction in
the add-in JavaScript.
"""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List, Mapping, Optional, Tuple

from ..generation_plan import (
    GenerationPlan,
    _is_m3_shape,
    validate_generation_plan,
)
from ..macos_probe.bridge import LoopbackBridge
from ..macos_probe.models import (
    ProbeResult,
    ProtocolError,
    validate_longform_generation_request,
    validate_longform_generation_value,
    validate_longform_notice_patch_request,
    validate_longform_notice_patch_value,
)
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
from .degradation import controlled_token
from .privacy import redact_private_text
from .quality import QualityFinding


MACOS_DEDICATED_HOST_UNAVAILABLE = "MACOS_DEDICATED_HOST_UNAVAILABLE"
EXECUTION_FAILED = "EXECUTION_FAILED"
EXECUTION_ABORTED = "EXECUTION_ABORTED"
_FIXED_PUBLIC_ISSUE_CODES = frozenset({
    "BIBLIOGRAPHY_INSERT_FAILED",
    "CROSS_REFERENCE_FAILED",
    "DEGRADATION_INSERT_FAILED",
    "EQUATION_INSERT_FAILED",
    "FIELD_REFRESH_UNSTABLE",
    "IMAGE_INSERT_FAILED",
    "TABLE_INSERT_FAILED",
    "TABLE_MERGE_APPLY_FAILED",
    "TABLE_ROW_FORCED_SPLIT",
    "TABLE_STYLE_APPLY_FAILED",
    "UNKNOWN_OPERATION",
})


class MacOSDedicatedHostUnavailableError(Exception):
    """Raised when the executor has no bridge to reach a dedicated WPS host."""

    def __init__(self, message: str = "Unable to reach a dedicated WPS host") -> None:
        super().__init__(message)
        self.code = MACOS_DEDICATED_HOST_UNAVAILABLE


class MacOSLongformExecutorError(Exception):
    """Engine-level error raised when execution must stop."""

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.cleanup_failed = False


class MacOSLongformExecutor(LongformExecutor):
    """Execute a protocol v2 long-form generation plan on macOS WPS via JSAPI."""

    def __init__(
        self,
        *,
        bridge: Optional[LoopbackBridge] = None,
        staging_dir: Optional[str] = None,
        activation_document: Optional[str] = None,
    ) -> None:
        self._bridge = bridge
        self._staging_dir = str(
            Path(staging_dir or tempfile.gettempdir()).expanduser().resolve()
        )
        self._activation_document = self._validate_activation_document(
            activation_document
        )

    def _validate_activation_document(self, candidate: Optional[str]) -> Optional[str]:
        if candidate is None:
            return None
        path = Path(candidate).expanduser()
        staging = Path(self._staging_dir)
        try:
            if path.is_symlink() or not path.is_file():
                raise MacOSLongformExecutorError(
                    "Private activation document is invalid"
                )
            resolved = path.resolve(strict=True)
            if (
                not resolved.is_relative_to(staging)
                or resolved.suffix.lower() != ".docx"
            ):
                raise MacOSLongformExecutorError(
                    "Private activation document is invalid"
                )
        except OSError:
            raise MacOSLongformExecutorError(
                "Private activation document is invalid"
            ) from None
        return str(resolved)

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

        self._validate_resource_manifest(plan, resources)
        paths = self._resolve_paths()
        staged_resources: Tuple[Tuple[str, Path], ...] = ()
        bridge_result: Optional[ProbeResult] = None
        outcome: Optional[ExecutionOutcome] = None
        primary_error: Optional[Exception] = None
        cleanup_failed = False
        try:
            staged_resources = self._stage_resources(resources, paths.staged_docx)
            self._verify_staged_resources(staged_resources, resources)
            request = {
                "plan": plan.to_dict(),
                "outputPath": paths.staged_docx,
                "resources": self._build_resource_map(staged_resources),
            }
            if self._activation_document is not None:
                request["activationDocument"] = self._activation_document
            params = validate_longform_generation_request(request)
            self._activation_document = None
            command = self._bridge.issue(
                "writer", "generate_longform_document", params
            )
            timeout = 300.0
            if deadline is not None:
                timeout = max(0.0, min(timeout, deadline - time.monotonic()))
            bridge_result = self._bridge.wait_result(command.id, timeout=timeout)
        except MacOSLongformExecutorError as error:
            primary_error = error
        except ProtocolError:
            primary_error = MacOSLongformExecutorError("Bridge request was invalid")
        except Exception:
            primary_error = MacOSLongformExecutorError("Bridge command failed")
        else:
            try:
                if bridge_result is None:
                    raise MacOSLongformExecutorError(
                        "Bridge result was unavailable"
                    )
                outcome = self._build_outcome(
                    bridge_result, paths.staged_docx, plan
                )
            except (MacOSLongformExecutorError, NativeFieldContractError) as error:
                primary_error = error
            except Exception:
                primary_error = MacOSLongformExecutorError(
                    "Bridge result was invalid"
                )
        finally:
            try:
                self._cleanup_resources(staged_resources)
            except Exception:
                cleanup_failed = True

        if primary_error is not None:
            try:
                primary_error.cleanup_failed = bool(
                    getattr(primary_error, "cleanup_failed", False)
                    or cleanup_failed
                )
            except Exception:
                pass
            raise primary_error from None
        if cleanup_failed:
            error = MacOSLongformExecutorError("Private resource cleanup failed")
            error.cleanup_failed = True
            raise error from None
        if outcome is None:
            raise MacOSLongformExecutorError("Bridge result was unavailable") from None
        return outcome

    def patch_quality_notices(
        self,
        source_path: Path,
        notices: Tuple[QualityFinding, ...],
        bookmark_by_node: Mapping[str, str],
        deadline: Optional[float] = None,
    ) -> ExecutionOutcome:
        """Apply one closed notice-only patch and return fresh pagination."""

        if self._bridge is None:
            raise MacOSDedicatedHostUnavailableError(
                "No LoopbackBridge available for macOS WPS execution"
            )
        if not notices:
            raise ValueError("notice patch requires at least one notice")
        source = Path(source_path).expanduser().resolve()
        target = Path(self._resolve_paths().staged_docx).resolve()
        payload = []
        for notice in notices:
            bookmark = bookmark_by_node.get(notice.node_id or "")
            if notice.node_id and not bookmark:
                raise ValueError("mapped block notice requires a native bookmark")
            payload.append({
                "code": notice.code,
                "message": notice.message,
                "fallback": notice.repair_key or "notice-only",
                "placement": "block" if bookmark else "document",
                "nodeId": notice.node_id or "doc:quality",
                "page": notice.page or 1,
                "bookmarkName": bookmark,
            })
        request = validate_longform_notice_patch_request({
            "sourcePath": str(source),
            "outputPath": str(target),
            "bookmarks": [
                {"nodeId": node_id, "bookmarkName": bookmark}
                for node_id, bookmark in sorted(bookmark_by_node.items())
            ],
            "notices": payload,
        })
        try:
            command = self._bridge.issue(
                "writer", "patch_longform_quality_notices", request
            )
            timeout = 300.0
            if deadline is not None:
                timeout = max(0.0, min(timeout, deadline - time.monotonic()))
            result = self._bridge.wait_result(command.id, timeout=timeout)
        except Exception:
            raise MacOSLongformExecutorError("Bridge notice patch failed") from None
        if not result.ok:
            raise MacOSLongformExecutorError("Bridge notice patch failed") from None
        try:
            value = validate_longform_notice_patch_value(result.value or {})
        except ProtocolError:
            raise MacOSLongformExecutorError("Bridge result was invalid") from None
        if value["outputPath"] != str(target):
            raise MacOSLongformExecutorError("Bridge result was invalid") from None
        return ExecutionOutcome(
            staged_artifact=str(target),
            issues=tuple(_execution_issue(item) for item in value.get("issueCodes", ())),
            pagination_map=_parse_pagination_map(value.get("paginationMap") or {}),
            applied_operations=value.get("appliedNotices"),
        )

    # ----------------------------------------------------------------------
    # Helpers
    # ----------------------------------------------------------------------
    def _resolve_paths(self) -> "_ResolvedPaths":
        os.makedirs(self._staging_dir, exist_ok=True)
        descriptor, base = tempfile.mkstemp(
            prefix="wpsc-longform-",
            suffix="",
            dir=self._staging_dir,
        )
        os.close(descriptor)
        os.unlink(base)
        return _ResolvedPaths(staged_docx=base + ".docx")

    def _validate_resource_manifest(
        self,
        plan: GenerationPlan,
        resources: Tuple[PreparedLongformResource, ...],
    ) -> None:
        entries: List[dict[str, Any]] = []
        by_id: dict[str, PreparedLongformResource] = {}
        invalid = False
        try:
            for resource in resources:
                if resource.id in by_id:
                    invalid = True
                    break
                if (
                    hashlib.sha256(resource.payload_bytes).hexdigest()
                    != resource.payload_sha256
                ):
                    invalid = True
                    break
                by_id[resource.id] = resource
                entries.append({
                    "resourceId": resource.id,
                    "sourceSha256": resource.source_sha256,
                    "payloadSha256": resource.payload_sha256,
                    "byteLength": len(resource.payload_bytes),
                    "mediaType": resource.media_type,
                    "normalizerId": resource.normalizer_id,
                })
            envelope = {
                "version": "1",
                "entries": sorted(entries, key=lambda item: item["resourceId"]),
            }
            canonical = json.dumps(
                envelope,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
            digest = "sha256:" + hashlib.sha256(canonical).hexdigest()
            if digest != plan.resource_manifest_digest:
                invalid = True
            for operation in plan.operations:
                if operation.op == "writer.add_equation":
                    fallback = operation.args.get("fallbackResource") or {}
                    resource_id = fallback.get("fallbackResourceId")
                    if resource_id is not None and resource_id not in by_id:
                        invalid = True
                        break
                    continue
                if operation.op != "writer.add_captioned_figure":
                    continue
                for child in operation.args.get("children", ()):
                    resource_id = child.get("resourceId")
                    if resource_id is None:
                        continue
                    resource = by_id.get(resource_id)
                    if (
                        resource is None
                        or child.get("mediaType") != resource.media_type
                        or child.get("normalizerId") != resource.normalizer_id
                    ):
                        invalid = True
                        break
        except Exception:
            invalid = True
        if invalid:
            raise MacOSLongformExecutorError("Private resource validation failed") from None

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
        suffixes = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/tiff": ".tiff",
            "image/bmp": ".bmp",
            "image/gif": ".gif",
            "image/svg+xml": ".svg",
        }
        try:
            for resource in resources:
                suffix = suffixes.get(resource.media_type)
                if suffix is None:
                    raise MacOSLongformExecutorError(
                        "Private resource staging failed"
                    )
                descriptor, name = tempfile.mkstemp(
                    prefix="wpsc-resource-",
                    suffix=suffix,
                    dir=staging_dir,
                )
                target = Path(name)
                staged.append((resource.id, target))
                with os.fdopen(descriptor, "wb") as stream:
                    stream.write(resource.payload_bytes)
                    stream.flush()
                os.chmod(target, 0o600)
        except Exception:
            cleanup_failed = False
            try:
                self._cleanup_resources(tuple(staged))
            except Exception:
                cleanup_failed = True
            error = MacOSLongformExecutorError("Private resource staging failed")
            error.cleanup_failed = cleanup_failed
            raise error from None
        return tuple(staged)

    @staticmethod
    def _verify_staged_resources(
        staged: Tuple[Tuple[str, Path], ...],
        resources: Tuple[PreparedLongformResource, ...],
    ) -> None:
        expected = {resource.id: resource.payload_sha256 for resource in resources}
        staged_ids = [resource_id for resource_id, _path in staged]
        invalid = (
            len(staged_ids) != len(expected)
            or len(set(staged_ids)) != len(staged_ids)
            or set(staged_ids) != set(expected)
        )
        try:
            for resource_id, path in staged:
                if (
                    resource_id not in expected
                    or hashlib.sha256(path.read_bytes()).hexdigest()
                    != expected[resource_id]
                ):
                    invalid = True
                    break
        except Exception:
            invalid = True
        if invalid:
            raise MacOSLongformExecutorError("Private resource validation failed") from None

    @staticmethod
    def _cleanup_resources(resources: Tuple[Tuple[str, Path], ...]) -> None:
        failed = False
        for _resource_id, path in resources:
            try:
                path.unlink()
            except FileNotFoundError:
                pass
            except OSError:
                failed = True
        if failed:
            raise MacOSLongformExecutorError("Private resource cleanup failed") from None

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
            error = result.error if isinstance(result.error, Mapping) else {}
            code = _safe_issue_code(
                error.get("code"), fallback=EXECUTION_ABORTED
            )
            detail = _safe_bridge_error_detail(error.get("message"), code)
            raise MacOSLongformExecutorError(
                f"Execution aborted by WPS JSAPI ({code}{': ' + detail if detail else ''})"
            ) from None

        value = validate_longform_generation_value(result.value or {})
        if value.get("outputPath") != staged_docx:
            raise MacOSLongformExecutorError("Bridge result was invalid") from None
        self._validate_result_references(value, plan)
        staged_artifact = staged_docx
        issue_codes = value.get("issueCodes") or []
        pagination_map = _parse_pagination_map(
            value.get("paginationMap") or {}
        )
        issues: List[ExecutionIssue] = [_execution_issue(item) for item in issue_codes]

        raw_history = value.get("fieldSnapshots")
        has_finalizer = any(
            op.op == "writer.finalize_fields" for op in plan.operations
        )
        requires_m3_history = any(
            _is_m3_shape(op.op, op.args) for op in plan.operations
        )
        if requires_m3_history and has_finalizer and not raw_history:
            raise NativeFieldContractError(
                "field_history", "is missing"
            ) from None
        if (
            raw_history is not None
            and has_finalizer
        ):
            max_rounds = self._extract_max_rounds(plan)
            history = _parse_field_snapshot_history(raw_history)
            convergence = evaluate_field_snapshot_history(
                history,
                max_rounds=max_rounds,
                allow_legacy_hashes=False,
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

    @staticmethod
    def _validate_result_references(
        value: Mapping[str, Any], plan: GenerationPlan
    ) -> None:
        allowed_nodes = {
            operation.node_id
            for operation in plan.operations
            if operation.node_id is not None
        }
        child_order: List[str] = []
        expected_run_children: dict[str, Tuple[str, Optional[str]]] = {}
        for operation in plan.operations:
            if operation.op == "writer.add_captioned_figure":
                for child in operation.args.get("children", ()):
                    node_id = child.get("nodeId")
                    if isinstance(node_id, str):
                        child_order.append(node_id)
                continue
            if operation.op == "writer.add_cross_reference":
                for run in operation.args.get("runs", ()):
                    if run.get("type") not in {"citation", "degradation"}:
                        continue
                    node_id = run.get("nodeId")
                    if not isinstance(node_id, str):
                        continue
                    child_order.append(node_id)
                    if run.get("type") == "citation":
                        expected_run_children[node_id] = ("applied", None)
                    else:
                        expected_run_children[node_id] = (
                            "degraded", run.get("code")
                        )
        allowed_nodes.update(child_order)
        child_positions = {node_id: index for index, node_id in enumerate(child_order)}
        allowed_issue_codes = set(_FIXED_PUBLIC_ISSUE_CODES)
        for operation in plan.operations:
            policy = operation.failure_policy or {}
            allowed_issue_codes.update(
                code
                for code in policy.get("recoverableCodes", ())
                if isinstance(code, str)
            )
            _collect_plan_issue_codes(operation.args, allowed_issue_codes)
        returned_children = value.get("childResults", ())
        invalid = value.get("appliedOperations", 0) != len(plan.operations)
        if [child.get("nodeId") for child in returned_children] != child_order:
            invalid = True
        for issue in value.get("issueCodes", ()):
            node_id = issue.get("nodeId")
            if (
                issue.get("code") not in allowed_issue_codes
                or (node_id is not None and node_id not in allowed_nodes)
            ):
                invalid = True
        for child in returned_children:
            node_id = child.get("nodeId")
            if (
                node_id not in child_positions
                or (
                    child.get("issueCode") is not None
                    and child.get("issueCode") not in allowed_issue_codes
                )
            ):
                invalid = True
            expected = expected_run_children.get(node_id)
            if expected is not None and (
                child.get("status"), child.get("issueCode")
            ) != expected:
                invalid = True
        returned_issues = value.get("issueCodes", ())
        for node_id, (status, issue_code) in expected_run_children.items():
            matching = [
                issue for issue in returned_issues
                if issue.get("nodeId") == node_id
            ]
            if status == "degraded":
                if len(matching) != 1 or matching[0].get("code") != issue_code:
                    invalid = True
            elif matching:
                invalid = True
        seen_pagination: set[str] = set()
        for node in value.get("paginationMap", {}).get("nodes", ()):
            node_id = node.get("nodeId")
            if node_id not in allowed_nodes or node_id in seen_pagination:
                invalid = True
            seen_pagination.add(node_id)
        if invalid:
            raise MacOSLongformExecutorError("Bridge result was invalid") from None

    def _extract_max_rounds(self, plan: GenerationPlan) -> int:
        for op in plan.operations:
            if op.op == "writer.finalize_fields":
                return op.args.get("maxRounds", 3)
        return 3


@dataclass(frozen=True)
class _ResolvedPaths:
    staged_docx: str


def _execution_issue(raw: Mapping[str, Any]) -> ExecutionIssue:
    code = _safe_issue_code(raw.get("code"))
    return ExecutionIssue(
        code=code,
        message=f"Native operation reported {code}",
        placement=(
            str(raw["placement"])
            if raw.get("placement") in {"block", "inline", "document"}
            else "document"
        ),
        node_id=raw.get("nodeId"),
        stage=controlled_token(raw.get("stage")),
        fallback=controlled_token(raw.get("fallback")),
        recoverable=(
            raw.get("recoverable")
            if type(raw.get("recoverable")) is bool
            else None
        ),
    )


def _safe_issue_code(value: Any, *, fallback: str = EXECUTION_FAILED) -> str:
    code = str(value or fallback)
    if (
        not 3 <= len(code) <= 64
        or not code[0].isalpha()
        or any(
            character not in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_"
            for character in code
        )
        or redact_private_text(code) != code
    ):
        return fallback
    return code


def _safe_bridge_error_detail(value: Any, code: str) -> str:
    if not isinstance(value, str) or not value or len(value) > 160:
        return ""
    if redact_private_text(value) != value:
        return ""
    prefix = code + ":writer."
    suffix = value[len(prefix):] if value.startswith(prefix) else ""
    if not suffix or any(
        character not in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_"
        for character in suffix
    ):
        return ""
    return value


def _collect_plan_issue_codes(value: Any, target: set[str]) -> None:
    if isinstance(value, Mapping):
        for key, child in value.items():
            if key == "code" and isinstance(child, str):
                target.add(child)
            else:
                _collect_plan_issue_codes(child, target)
    elif isinstance(value, (list, tuple)):
        for child in value:
            _collect_plan_issue_codes(child, target)


def _parse_field_snapshot_history(raw: Any) -> Tuple[Tuple[FieldSnapshot, ...], ...]:
    failed = False
    history: Tuple[Tuple[FieldSnapshot, ...], ...] = ()
    try:
        parsed = []
        for round_value in tuple(raw or ()):
            if isinstance(round_value, Mapping):
                parsed.append((_parse_remote_field_snapshot(round_value),))
            else:
                parsed.append(
                    tuple(
                        _parse_remote_field_snapshot(item)
                        for item in tuple(round_value)
                    )
                )
        history = tuple(parsed)
    except Exception:
        failed = True
    if failed:
        raise NativeFieldContractError("field_history", "is invalid") from None
    return history


def _parse_remote_field_snapshot(raw: Any) -> FieldSnapshot:
    keys = {
        "stableKey",
        "fieldCategory",
        "resultHash",
        "tocPageCount",
        "figureIndexPageCount",
        "tableIndexPageCount",
        "totalPages",
    }
    if not isinstance(raw, Mapping) or set(raw) != keys:
        raise ValueError("invalid field snapshot")
    stable_key = raw["stableKey"]
    if (
        not isinstance(stable_key, (list, tuple))
        or len(stable_key) != 3
        or not isinstance(stable_key[0], str)
        or not isinstance(stable_key[1], str)
        or not isinstance(stable_key[2], int)
        or isinstance(stable_key[2], bool)
        or not isinstance(raw["fieldCategory"], str)
        or not isinstance(raw["resultHash"], str)
    ):
        raise ValueError("invalid field snapshot")
    count_keys = (
        "tocPageCount",
        "figureIndexPageCount",
        "tableIndexPageCount",
        "totalPages",
    )
    if any(
        not isinstance(raw[key], int) or isinstance(raw[key], bool)
        for key in count_keys
    ):
        raise ValueError("invalid field snapshot")
    return FieldSnapshot(
        stable_key=(stable_key[0], stable_key[1], stable_key[2]),
        field_category=raw["fieldCategory"],
        result_hash=raw["resultHash"],
        toc_page_count=raw["tocPageCount"],
        figure_index_page_count=raw["figureIndexPageCount"],
        table_index_page_count=raw["tableIndexPageCount"],
        total_pages=raw["totalPages"],
    )


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
