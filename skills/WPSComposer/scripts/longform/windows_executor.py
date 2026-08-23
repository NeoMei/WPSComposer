"""Windows COM executor for long-form generation plans.

This module is importable on macOS/Linux because all pywin32 imports are
performed lazily inside the dedicated-composer factory.
"""
from __future__ import annotations

import inspect
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile
import time
from dataclasses import dataclass
from typing import Any, Callable, List, Mapping, Optional, Tuple

from ..generation_plan import GenerationOperation, GenerationPlan, validate_generation_plan
from ..writer import WriterComposer
from .executor import (
    ExecutionIssue,
    ExecutionOutcome,
    LongformExecutor,
    PaginationFragment,
    PaginationMap,
    PaginationNode,
    finalize_fields_with_convergence,
)
from .resources import PreparedLongformResource
from .field_contract import NativeFieldAdapter, finalize_native_fields
from .degradation import (
    DegradationDescriptor,
    LocalRecoveryController,
    RecoveryFatalError,
)
from .privacy import redact_private_text
from .quality import QualityFinding


WINDOWS_DEDICATED_HOST_UNAVAILABLE = "WINDOWS_DEDICATED_HOST_UNAVAILABLE"
EXECUTION_FAILED = "EXECUTION_FAILED"
EXECUTION_ABORTED = "EXECUTION_ABORTED"
DEGRADATION_FALLBACK_FAILED = "DEGRADATION_FALLBACK_FAILED"
UNKNOWN_OPERATION = "UNKNOWN_OPERATION"

# M4 completes the last Windows-native object that was deferred in M2/M3.
_M2_DEFERRED_OPERATIONS: dict[str, tuple[str, str]] = {}

_M3_NATIVE_OPERATIONS = frozenset({
    "writer.add_captioned_figure",
    "writer.add_semantic_table",
    "writer.add_equation",
    "writer.add_cross_reference",
    "writer.insert_figure_index",
    "writer.insert_table_index",
})

_LEGACY_OBJECT_DEFERRED = {
    "writer.add_captioned_figure": (
        "IMAGE_INSERT_FAILED", "figure-child-stack-then-notice"
    ),
    "writer.add_semantic_table": ("TABLE_INSERT_FAILED", "grid-then-text"),
    "writer.add_equation": (
        "EQUATION_INSERT_FAILED", "explicit-image-then-source-notice"
    ),
    "writer.add_cross_reference": ("CROSS_REFERENCE_FAILED", "inline-fallback"),
}


class WindowsDedicatedHostUnavailableError(Exception):
    """Raised when the executor cannot acquire a dedicated WPS COM instance."""

    def __init__(self, message: str = "Unable to acquire a dedicated WPS host") -> None:
        super().__init__(message)
        self.code = WINDOWS_DEDICATED_HOST_UNAVAILABLE


class WindowsLongformExecutorError(Exception):
    """Engine-level error raised when execution must stop."""

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.cleanup_failed = False


def _is_host_com_error(exc: BaseException) -> bool:
    """True when the error (or its abort cause) is a COM/RPC failure."""
    try:
        import pywintypes  # pywin32
    except ImportError:  # non-Windows hosts run this module under tests
        return False
    seen: Optional[BaseException] = exc
    for _ in range(4):
        if seen is None:
            return False
        if isinstance(seen, pywintypes.com_error):
            return True
        seen = getattr(seen, "__cause__", None)
    return False


def _create_dedicated_composer(staging_dir: Optional[str] = None) -> WriterComposer:
    """Create a dedicated WriterComposer via DispatchEx only.

    This function lazily imports pywin32 so the module can be imported on macOS
    and Linux.  It never falls back to a shared Dispatch instance.
    """
    import pythoncom  # pywin32
    # Tests inject win32com.client via sys.modules; prefer the injected module.
    client = sys.modules.get("win32com.client")
    if client is None:
        import win32com.client as client

    pythoncom.CoInitialize()
    app: Any = None
    composer: Optional[WriterComposer] = None

    # WPS's single-process model can hand DispatchEx a proxy into a previous
    # instance that is still quitting, or an app object whose properties are
    # not ready yet (AttributeError / mid-run RPC death). Retry the whole
    # dedicated-host construction with backoff, mirroring _base.__enter__.
    last_error: Optional[Exception] = None
    for attempt in range(3):
        try:
            for progid in WriterComposer._progids:
                try:
                    app = client.DispatchEx(progid)
                    break
                except Exception:  # pragma: no cover - exercised via mocks
                    continue
            if app is None:
                raise WindowsDedicatedHostUnavailableError(
                    "Could not create dedicated WPS application"
                )
            # Readiness probe: touch hot properties the executor will use.
            int(app.Documents.Count)
            app.Selection
            composer = WriterComposer.__new__(WriterComposer)
            composer._app = app
            composer._doc = None
            composer._path = None
            composer._read_only = False
            composer._visible = False
            composer._owns_app = True
            composer._owns_doc = False
            composer._com_initialized = True
            composer._first_section_configured = False
            try:
                app.Visible = 0
                app.DisplayAlerts = 0
            except Exception:
                pass

            composer._doc = composer._create_doc(app)
            break
        except WindowsDedicatedHostUnavailableError:
            raise
        except Exception as exc:  # pragma: no cover - WPS reuse race
            last_error = exc
            app = None
            if composer is not None:
                try:
                    composer.close(save_changes=False)
                except Exception:
                    pass
                composer = None
            if attempt == 2:
                pythoncom.CoUninitialize()
                raise WindowsDedicatedHostUnavailableError(
                    f"Could not create dedicated WPS host: {exc}"
                ) from exc
            time.sleep(0.6)

    assert composer is not None and composer._doc is not None
    app = composer._app

    composer._owns_doc = True
    if staging_dir:
        composer._staging_dir = staging_dir
    return composer


@dataclass(frozen=True)
class _ResolvedPaths:
    staged_docx: str


@dataclass(frozen=True)
class _LocalCheckpoint:
    token: Any
    rollback: Callable[[Any], Any]


class WindowsLongformExecutor(LongformExecutor):
    """Execute a protocol v2 long-form generation plan on a dedicated WPS host."""

    def __init__(
        self,
        *,
        staging_dir: Optional[str] = None,
        composer_factory: Optional[Callable[..., WriterComposer]] = None,
    ) -> None:
        self._staging_dir = staging_dir or tempfile.gettempdir()
        self._composer_factory = composer_factory or _create_dedicated_composer
        self._issues: List[ExecutionIssue] = []
        self._toc_density: dict[str, Any] = {}
        self._front_matter: dict[str, Any] = {}
        self._resource_locators: dict[str, str] = {}
        self._recovery_controller = LocalRecoveryController()
        self._pagination_ranges: list[dict[str, Any]] = []

    # ----------------------------------------------------------------------
    # Public interface
    # ----------------------------------------------------------------------
    def execute(
        self,
        plan: GenerationPlan,
        resources: Tuple[PreparedLongformResource, ...] = (),
        deadline: Optional[float] = None,
    ) -> ExecutionOutcome:
        self._issues = []
        self._recovery_controller = LocalRecoveryController()
        self._pagination_ranges = []
        validate_generation_plan(plan.to_dict(), component="writer")
        paths = self._resolve_paths()
        staged_resources: Tuple[str, ...] = ()
        pending_cleanup: Tuple[str, ...] = ()
        cleanup_attempted = False
        composer: Optional[WriterComposer] = None
        primary_error: Optional[Exception] = None
        cleanup_failed = False
        pagination_map = _build_pagination_map(plan.operations)
        try:
            self._validate_resource_manifest(plan, resources)
            self._resource_locators, staged_resources = self._stage_resources(resources)
            host_attempts = 0
            while True:
                try:
                    composer = self._acquire_composer()
                    self._dispatch_all(composer, plan.operations)
                    if isinstance(composer, NativeFieldAdapter):
                        convergence = finalize_native_fields(
                            composer, max_rounds=_extract_max_rounds(plan.operations)
                        )
                    else:
                        convergence = finalize_fields_with_convergence(
                            composer, max_rounds=_extract_max_rounds(plan.operations)
                        )
                    self._extend_issues(convergence.issues)
                    snapshotter = getattr(composer, "pagination_map_for_ranges", None)
                    if callable(snapshotter):
                        pagination_map = PaginationMap.from_dict(
                            snapshotter(tuple(self._pagination_ranges))
                        )
                    composer.save_docx(paths.staged_docx)
                    break
                except Exception as exc:
                    # WPS's automation session can die mid-run ("interface
                    # unknown" / RPC errors) once instances accumulate. One
                    # fresh-host retry of the SAME logical generation; a
                    # content failure re-raised by the retry path is
                    # classified as usual.
                    if (
                        host_attempts == 0
                        and not isinstance(exc, WindowsLongformExecutorError)
                        and _is_host_com_error(exc)
                    ):
                        host_attempts += 1
                        if composer is not None:
                            try:
                                composer.close(save_changes=False)
                            except Exception:
                                pass
                            composer = None
                        self._issues = []
                        self._pagination_ranges = []
                        # Let the previous instance finish quitting so WPS's
                        # single-instance re-registration completes before
                        # the replacement dispatch.
                        time.sleep(1.5)
                        continue
                    raise
            # WPS may retain an image handle until the document closes.  Try
            # once while the host is alive, then retry only locked paths after
            # close in the finally block.
            pending_cleanup = self._cleanup_resources(
                staged_resources, strict=False
            )
            cleanup_attempted = True
        except _ExecutionAbort as exc:
            primary_error = WindowsLongformExecutorError(
                f"Execution aborted at {exc.op_name}"
            )
        except (
            WindowsLongformExecutorError,
            WindowsDedicatedHostUnavailableError,
        ) as exc:
            primary_error = exc
        except Exception:
            primary_error = WindowsLongformExecutorError(
                "Windows native execution failed"
            )
        finally:
            try:
                if composer is not None:
                    try:
                        composer.close(save_changes=False)
                    except Exception:
                        pass
            finally:
                try:
                    final_targets = (
                        pending_cleanup if cleanup_attempted else staged_resources
                    )
                    try:
                        self._cleanup_resources(final_targets, strict=True)
                    except Exception:
                        cleanup_failed = True
                finally:
                    # Locator state is private and execution-scoped.  Cleanup
                    # failure must never retain it on a reusable executor.
                    self._resource_locators = {}
        if primary_error is not None:
            setattr(
                primary_error,
                "cleanup_failed",
                bool(
                    getattr(primary_error, "cleanup_failed", False)
                    or cleanup_failed
                ),
            )
            primary_error.__cause__ = None
            primary_error.__context__ = None
            raise primary_error from None
        if cleanup_failed:
            cleanup_error = WindowsLongformExecutorError(
                "Private resource cleanup failed"
            )
            cleanup_error.cleanup_failed = True
            raise cleanup_error from None
        return ExecutionOutcome(
            staged_artifact=paths.staged_docx,
            issues=tuple(self._issues),
            pagination_map=pagination_map,
            applied_operations=len(plan.operations),
        )

    def patch_quality_notices(
        self,
        source_path: Path,
        notices: Tuple[QualityFinding, ...],
        bookmark_by_node: Mapping[str, str],
        deadline: Optional[float] = None,
    ) -> ExecutionOutcome:
        """Apply one notice-only patch in a dedicated Writer instance."""

        if not notices:
            raise ValueError("notice patch requires at least one notice")
        if deadline is not None and time.monotonic() >= deadline:
            raise WindowsLongformExecutorError("Notice patch deadline expired")
        source = Path(source_path).expanduser().resolve()
        target = Path(self._resolve_paths().staged_docx).resolve()
        composer: Optional[WriterComposer] = None
        try:
            composer = self._acquire_composer()
            opener = getattr(composer, "open_existing_for_patch", None)
            if callable(opener):
                opener(str(source))
            else:
                if getattr(composer, "_doc", None) is not None:
                    composer._doc.Close(False)
                composer._doc = composer._app.Documents.Open(str(source), False, False)
            for notice in notices:
                bookmark = bookmark_by_node.get(notice.node_id or "")
                if notice.node_id and not bookmark:
                    raise ValueError("mapped block notice requires a native bookmark")
                composer.add_quality_notice_at_bookmark(
                    code=notice.code,
                    message=notice.message,
                    fallback=notice.repair_key or "notice-only",
                    node_id=notice.node_id or "doc:quality",
                    page=notice.page or 1,
                    bookmark_name=bookmark,
                )
            convergence = finalize_fields_with_convergence(composer, max_rounds=3)
            # Snapshot every stable generation bookmark only after fields have
            # converged; returning notice nodes alone leaves the final issue
            # pages fresh while the rest of the document remains stale.
            nodes = [
                composer.pagination_fragment_for_bookmark(node_id, bookmark)
                for node_id, bookmark in sorted(bookmark_by_node.items())
            ]
            if not nodes:
                nodes = [
                    composer.pagination_fragment_for_bookmark(
                        "doc:quality", "wpsc_document_quality_anchor"
                    )
                ]
            composer.save_docx(str(target))
            return ExecutionOutcome(
                staged_artifact=str(target),
                issues=convergence.issues,
                pagination_map=PaginationMap.from_dict({
                    "version": "M5-v1", "nodes": nodes,
                }),
                applied_operations=len(notices),
            )
        except ValueError:
            raise
        except Exception:
            raise WindowsLongformExecutorError("Windows notice patch failed") from None
        finally:
            if composer is not None:
                try:
                    composer.close(save_changes=False)
                except Exception:
                    pass

    def _validate_resource_manifest(
        self,
        plan: GenerationPlan,
        resources: Tuple[PreparedLongformResource, ...],
    ) -> None:
        seen: set[str] = set()
        entries = []
        for resource in resources:
            if resource.id in seen:
                raise WindowsLongformExecutorError("Private resource validation failed")
            seen.add(resource.id)
            actual = hashlib.sha256(resource.payload_bytes).hexdigest()
            if actual != resource.payload_sha256:
                raise WindowsLongformExecutorError("Private resource validation failed")
            entries.append({
                "resourceId": resource.id,
                "sourceSha256": resource.source_sha256,
                "payloadSha256": resource.payload_sha256,
                "byteLength": len(resource.payload_bytes),
                "mediaType": resource.media_type,
                "normalizerId": resource.normalizer_id,
            })
        envelope = {"version": "1", "entries": sorted(entries, key=lambda item: item["resourceId"])}
        canonical = json.dumps(
            envelope, sort_keys=True, separators=(",", ":"), ensure_ascii=False
        ).encode("utf-8")
        digest = "sha256:" + hashlib.sha256(canonical).hexdigest()
        if digest != plan.resource_manifest_digest:
            raise WindowsLongformExecutorError("Private resource validation failed")

        by_id = {resource.id: resource for resource in resources}
        for operation in plan.operations:
            if operation.op == "writer.add_equation":
                fallback = operation.args.get("fallbackResource") or {}
                resource_id = fallback.get("fallbackResourceId")
                if resource_id is not None and resource_id not in by_id:
                    raise WindowsLongformExecutorError(
                        "Private resource validation failed"
                    )
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
                    raise WindowsLongformExecutorError("Private resource validation failed")

    def _stage_resources(
        self, resources: Tuple[PreparedLongformResource, ...]
    ) -> Tuple[dict[str, str], Tuple[str, ...]]:
        locators: dict[str, str] = {}
        paths: list[str] = []
        suffixes = {
            "image/png": ".png", "image/jpeg": ".jpg", "image/tiff": ".tiff",
            "image/bmp": ".bmp", "image/gif": ".gif", "image/svg+xml": ".svg",
        }
        staging_failed = False
        try:
            for index, resource in enumerate(resources):
                suffix = suffixes.get(resource.media_type)
                if suffix is None:
                    raise WindowsLongformExecutorError("Private resource validation failed")
                handle = tempfile.NamedTemporaryFile(
                    prefix=f"wpsc-resource-{index}-", suffix=suffix,
                    dir=self._staging_dir, delete=False,
                )
                # Register immediately: write, flush, or close can fail after
                # the OS has already created the private file.
                paths.append(handle.name)
                try:
                    handle.write(resource.payload_bytes)
                    handle.flush()
                finally:
                    self._close_staging_handle(handle)
                locators[resource.id] = handle.name
        except Exception:
            staging_failed = True
        if staging_failed:
            cleanup_failed = False
            try:
                self._cleanup_resources(tuple(paths), strict=True)
            except Exception:
                cleanup_failed = True
            staging_error = WindowsLongformExecutorError(
                "Private resource staging failed"
            )
            staging_error.cleanup_failed = cleanup_failed
            raise staging_error from None
        return locators, tuple(paths)

    @staticmethod
    def _close_staging_handle(handle: Any) -> None:
        """Close a newly-created private handle with one bounded retry."""
        last_error: Optional[Exception] = None
        for _attempt in range(2):
            try:
                handle.close()
                return
            except Exception as exc:
                last_error = exc
        assert last_error is not None
        raise last_error

    @staticmethod
    def _cleanup_resources(
        paths: Tuple[str, ...], *, strict: bool
    ) -> Tuple[str, ...]:
        remaining = []
        for path in paths:
            try:
                os.unlink(path)
            except FileNotFoundError:
                pass
            except OSError:
                remaining.append(path)
        if remaining and strict:
            cleanup_error = WindowsLongformExecutorError(
                "Private resource cleanup failed"
            )
            cleanup_error.cleanup_failed = True
            raise cleanup_error from None
        return tuple(remaining)

    # ----------------------------------------------------------------------
    # Composer lifecycle
    # ----------------------------------------------------------------------
    def _acquire_composer(self) -> WriterComposer:
        try:
            # Tests pass either a 0-argument factory (legacy) or a 1-argument one.
            sig = inspect.signature(self._composer_factory)
            if list(sig.parameters.values()):
                return self._composer_factory(self._staging_dir)
            return self._composer_factory()
        except WindowsDedicatedHostUnavailableError:
            raise
        except Exception as exc:
            raise WindowsDedicatedHostUnavailableError(
                f"Could not acquire dedicated WPS host: {exc}"
            ) from exc

    def _resolve_paths(self) -> _ResolvedPaths:
        os.makedirs(self._staging_dir, exist_ok=True)
        descriptor, base = tempfile.mkstemp(
            prefix="wpsc-longform-",
            suffix="",
            dir=self._staging_dir,
        )
        os.close(descriptor)
        os.unlink(base)
        return _ResolvedPaths(staged_docx=base + ".docx")

    # ----------------------------------------------------------------------
    # Operation dispatch
    # ----------------------------------------------------------------------
    def _dispatch_all(
        self,
        composer: WriterComposer,
        operations: Tuple[GenerationOperation, ...],
    ) -> None:
        active_role = "body"
        position = getattr(composer, "_native_position", None)
        range_factory = getattr(getattr(composer, "_doc", None), "Range", None)
        for op in operations:
            before = position() if callable(position) else None
            self._run_op(composer, op)
            after = position() if callable(position) else None
            if op.op == "writer.configure_section" and op.args.get("role"):
                active_role = str(op.args["role"])
            if (
                op.node_id
                and isinstance(before, int)
                and isinstance(after, int)
                and callable(range_factory)
            ):
                self._pagination_ranges.append({
                    "nodeId": op.node_id,
                    "op": op.op,
                    "role": active_role,
                    "range": range_factory(min(before, after), max(before, after)),
                })

    def _run_op(self, composer: WriterComposer, op: GenerationOperation) -> None:
        policy = op.failure_policy
        deferred = _M2_DEFERRED_OPERATIONS.get(op.op)
        if deferred is None and not _is_m3_operation(op):
            deferred = _LEGACY_OBJECT_DEFERRED.get(op.op)
        checkpoint = self._checkpoint_local(composer) if (
            deferred is not None
            or (policy is not None and policy.get("mode") == "degrade")
        ) else None

        try:
            if deferred is not None:
                # M2 deferred ops render deterministically through their declared
                # fallback path instead of a real primitive.
                code, fallback = deferred
                raise _DeferredOperationError(code, op.op, fallback)
            self._dispatch_one(composer, op)
        except _ExecutionAbort:
            raise
        except _DeferredOperationError as exc:
            # Deterministic M2 degradation.  Record the stable issue and apply
            # the declared fallback without treating it as a primitive failure.
            self._recover_after_failure(
                composer, op, exc, exc.code, exc.fallback, checkpoint=checkpoint
            )
        except Exception as exc:
            code = _error_code(exc)
            if policy is not None and policy.get("mode") == "degrade":
                if code in policy.get("recoverableCodes", []):
                    self._recover_after_failure(
                        composer, op, exc, code, policy["fallback"], checkpoint=checkpoint
                    )
                    return
            raise _ExecutionAbort(op.op, exc) from exc

    def _recover_after_failure(
        self,
        composer: WriterComposer,
        op: GenerationOperation,
        exc: Exception,
        code: str,
        fallback: str,
        checkpoint: Any,
    ) -> None:
        placement = _fallback_placement(fallback)
        descriptor = DegradationDescriptor(
            code=code,
            placement=placement,
            object_label=_operation_label(op.op),
            reason=f"{op.op} native rendering failed",
            fallback_text=_operation_fallback_text(op),
            fallback_kind=fallback,
        )
        try:
            self._recovery_controller.recover_after_failure(
                node_id=op.node_id or op.op,
                descriptor=descriptor,
                checkpoint_token=checkpoint,
                native_code=code,
                rollback=lambda token: self._rollback_local(composer, token),
                fallback_attempt=lambda item: self._apply_fallback(
                    composer, op, item.fallback_kind, code=item.code
                ),
                # The inline fallback or block box is itself the required
                # same-location visible notice; this callback verifies that
                # the insertion path completed before the issue is recorded.
                insert_notice=lambda node_id, item: None,
            )
        except RecoveryFatalError as recovery_error:
            raise _ExecutionAbort(op.op, recovery_error) from recovery_error
        self._record_issue(
            code=code,
            message=(
                f"{op.op} is deferred to fallback in M2"
                if isinstance(exc, _DeferredOperationError)
                else f"{op.op} used its declared native fallback"
            ),
            node_id=op.node_id,
            placement=placement,
            stage="native",
            fallback=fallback,
            recoverable=True,
        )

    @staticmethod
    def _checkpoint_local(composer: WriterComposer) -> Any:
        checkpoint = getattr(composer, "degradation_checkpoint", None)
        if not callable(checkpoint):
            raise _ExecutionAbort(
                "local-checkpoint", RuntimeError("checkpoint API unavailable")
            )
        rollback = getattr(composer, "rollback_degradation_checkpoint", None)
        if not callable(rollback):
            raise _ExecutionAbort(
                "local-checkpoint", RuntimeError("rollback API unavailable")
            )
        try:
            return _LocalCheckpoint(token=checkpoint(), rollback=rollback)
        except Exception as exc:
            raise _ExecutionAbort("local-checkpoint", exc) from exc

    @staticmethod
    def _rollback_local(composer: WriterComposer, checkpoint: Any) -> None:
        if not isinstance(checkpoint, _LocalCheckpoint):
            raise RuntimeError("rollback API unavailable")
        checkpoint.rollback(checkpoint.token)

    def _apply_fallback(
        self,
        composer: WriterComposer,
        op: GenerationOperation,
        fallback: str,
        *,
        code: Optional[str] = None,
    ) -> None:
        args = op.args
        if fallback == "figure-child-stack-then-notice":
            result = composer.add_captioned_figure_fallback(
                **dict(args),
                owner_node_id=op.node_id,
                resource_locators=dict(self._resource_locators),
                failure_code=code or "IMAGE_INSERT_FAILED",
            )
            self._consume_native_result(result, op)
        elif fallback == "grid-then-text":
            native_args = dict(args)
            native_args.pop("cellCitations", None)
            result = composer.add_semantic_table_fallback(
                **native_args,
                owner_node_id=op.node_id,
                failure_code=code or "TABLE_INSERT_FAILED",
            )
            self._consume_native_result(result, op)
        elif fallback == "inline-fallback" and op.op == "writer.add_cross_reference":
            composer.add_cross_reference_fallback(
                **dict(args), owner_node_id=op.node_id,
                failure_code=code or "CROSS_REFERENCE_FAILED",
            )
        elif fallback in {"inline", "inline-fallback"}:
            text = str(
                args.get("fallbackText")
                or args.get("source")
                or args.get("text")
                or _reference_fallback_text(args.get("runs", ()))
                or ""
            )
            composer.add_inline_degradation(
                code=code or _op_fallback_code(op),
                message=_op_fallback_message(op),
                fallback_text=text,
            )
        elif (
            fallback == "explicit-image-then-source-notice"
            and op.op == "writer.add_equation"
            and args.get("renderMode") == "native-m4"
        ):
            fallback_resource = args.get("fallbackResource") or {}
            resource_id = fallback_resource.get("fallbackResourceId")
            locator = (
                self._resource_locators.get(resource_id)
                if resource_id is not None
                else None
            )
            if resource_id is not None and locator is None:
                raise RecoveryFatalError(
                    "RESOURCE_HASH_MISMATCH", "formula resource binding unavailable"
                ) from None
            composer.add_equation_native_fallback(
                numbering=args["numbering"],
                bookmarkName=args.get("bookmarkName"),
                fallbackText=args.get("fallbackText", ""),
                fallback_resource_locator=locator,
                owner_node_id=op.node_id,
                failure_code=code or "EQUATION_INSERT_FAILED",
            )
        elif fallback in {"notice", "explicit-image-then-source-notice"}:
            text = str(
                args.get("fallbackText")
                or args.get("source")
                or args.get("text")
                or _bibliography_fallback_text(args)
                or _table_fallback_text(args)
                or args.get("caption")
                or ""
            )
            composer.add_degradation_notice(
                code=code or _op_fallback_code(op),
                message=_op_fallback_message(op),
                fallback_text=text,
                placement=args.get("placement", "block"),
            )
        else:
            raise RecoveryFatalError(
                DEGRADATION_FALLBACK_FAILED, "fallback kind is not implemented"
            ) from None

    def _dispatch_one(self, composer: WriterComposer, op: GenerationOperation) -> None:
        name = op.op
        args = op.args

        if name == "writer.reset":
            composer.reset()
            return

        if name == "writer.configure_page":
            composer.set_margins(
                args.get("marginTop", 72),
                args.get("marginBottom", 72),
                args.get("marginLeft", 90),
                args.get("marginRight", 90),
            )
            return

        if name == "writer.configure_front_matter":
            self._front_matter = dict(args)
            metadata_setter = getattr(composer, "set_document_metadata", None)
            if callable(metadata_setter):
                metadata_setter(
                    title=args.get("title", ""),
                    author=args.get("author", ""),
                )
            composer.set_page_role(args.get("role", "front_matter"))
            return

        if name == "writer.configure_section":
            margins = args.get("margins")
            if margins is None and any(
                key in args for key in ("marginTop", "marginBottom", "marginLeft", "marginRight")
            ):
                margins = {
                    "top": args.get("marginTop", 72),
                    "bottom": args.get("marginBottom", 72),
                    "left": args.get("marginLeft", 90),
                    "right": args.get("marginRight", 90),
                }
            composer.configure_section(
                role=args.get("role", "body"),
                # Historic plans omit portrait orientation.  Make that legacy
                # representation explicit only at the native execution edge so
                # a landscape section cannot leak into the following section.
                landscape=args.get("landscape", False),
                page_size=args.get("pageSize"),
                margins=margins,
                restart_page_numbering=args.get("restartPageNumbering"),
                page_number_format=args.get("pageNumberFormat", "continue"),
                start_page_number=args.get("startPageNumber"),
                header_text=args.get("headerText"),
                footer_text=args.get("footerText"),
                link_to_previous_header=args.get("linkToPreviousHeader"),
                link_to_previous_footer=args.get("linkToPreviousFooter"),
            )
            if args.get("role") == "cover" and args.get(
                "titlePage", self._front_matter.get("titlePage")
            ):
                # Parity with the macOS addin: render the cover from the
                # stashed front matter (centered, explicit sizes).
                composer.add_paragraph(
                    self._front_matter.get("title", ""),
                    size=24,
                    bold=True,
                    align=1,
                    space_after=24,
                )
                if self._front_matter.get("author"):
                    composer.add_paragraph(
                        self._front_matter["author"], size=14, align=1
                    )
                if self._front_matter.get("date"):
                    composer.add_paragraph(
                        self._front_matter["date"], size=12, align=1
                    )
            return

        if name == "writer.configure_toc_styles":
            # Capture density/style configuration so it can be applied when
            # insert_toc is dispatched.  Plans may either wrap density in a
            # "density" key or place the density keys at the top level.
            self._toc_density = (
                args.get("density")
                or {
                    "minFontSizePt": args.get("minFontSizePt", {}),
                    "minSpaceBeforePt": args.get("minSpaceBeforePt", {}),
                    "minSpaceAfterPt": args.get("minSpaceAfterPt", {}),
                }
                or {}
            )
            return

        if name == "writer.set_page_role":
            composer.set_page_role(args["role"])
            return

        if name == "writer.set_page_numbering":
            composer.set_page_numbering(
                format=args.get("format", "arabic"),
                start=args.get("start"),
                restart=args.get("restart"),
            )
            return

        if name == "writer.set_header_footer":
            composer.set_header_footer(
                header=args.get("headerText"),
                footer=args.get("footerText"),
                link_to_previous_header=args.get("linkToPreviousHeader"),
                link_to_previous_footer=args.get("linkToPreviousFooter"),
            )
            return

        if name == "writer.ensure_styles":
            styles = args.get("styles", {})
            if isinstance(styles, (list, tuple)):
                styles = {
                    str(style.get("name", idx)): dict(style)
                    for idx, style in enumerate(styles)
                }
            composer.ensure_styles(styles or {})
            return

        if name == "writer.add_page_break":
            composer.add_page_break()
            return

        if name == "writer.add_heading":
            heading_args = {
                "text": args["text"],
                "level": args.get("level", 1),
                "numbering": args.get("numbering"),
                "scheme": args.get("numberingScheme"),
            }
            # Keep compatibility with injected and older composer adapters.
            # The new argument is only required by the one bounded M5 relayout.
            if args.get("keepWithNext") is True:
                heading_args["keep_with_next"] = True
            if args.get("bookmarkName"):
                heading_args["bookmark_name"] = args["bookmarkName"]
            composer.add_heading_level_native(**heading_args)
            return

        if name == "writer.add_paragraph":
            if args.get("style"):
                composer.add_styled_paragraph(
                    args.get("text", ""), str(args["style"])
                )
            else:
                composer.add_paragraph(text=args.get("text", ""))
            return

        if name == "writer.add_captioned_figure" and "numbering" in args:
            result = composer.add_captioned_figure_native(
                **args,
                owner_node_id=op.node_id,
                resource_locators=dict(self._resource_locators),
                controller_owned=True,
            )
            self._consume_native_result(result, op)
            return

        if name == "writer.add_semantic_table" and "numbering" in args:
            native_args = dict(args)
            # Citation metadata is plan-only because table cell text is already
            # the resolved static [n]. Cell degradations remain native styling
            # metadata and must reach the Writer primitive.
            native_args.pop("cellCitations", None)
            result = composer.add_semantic_table_native(
                **native_args, owner_node_id=op.node_id, controller_owned=True
            )
            self._consume_native_result(result, op)
            return

        if (
            name == "writer.add_equation"
            and args.get("renderMode") == "native-m4"
        ):
            native_args = dict(args)
            fallback_resource = native_args.pop("fallbackResource", None) or {}
            if "plannedDegradation" in native_args.get("content", {}):
                resource_id = fallback_resource.get("fallbackResourceId")
                locator = (
                    self._resource_locators.get(resource_id)
                    if resource_id is not None
                    else None
                )
                if resource_id is not None and locator is None:
                    raise NativeWriterObjectError(
                        "RESOURCE_HASH_MISMATCH", "formula resource unavailable"
                    ) from None
                native_args["fallback_resource_locator"] = locator
            result = composer.add_equation_native(
                **native_args, owner_node_id=op.node_id, controller_owned=True
            )
            self._consume_native_result(result, op)
            planned = fallback_resource.get(
                "fallbackResourcePlannedDegradation"
            )
            if planned is not None:
                composer.add_degradation_notice(
                    code=planned["code"],
                    message=planned["reason"],
                    fallback_text=planned["fallbackText"],
                    placement=planned["placement"],
                )
                self._record_issue(
                    code=planned["code"],
                    message="Optional formula fallback image is unavailable",
                    node_id=op.node_id,
                    placement=planned["placement"],
                    stage="native",
                    fallback="none",
                    recoverable=True,
                )
            return

        if name == "writer.add_equation" and "numbering" in args:
            result = composer.add_equation_number_native(
                **args, owner_node_id=op.node_id
            )
            self._consume_native_result(result, op)
            return

        if name == "writer.add_cross_reference" and "runs" in args:
            native_method = (
                composer.add_citation_paragraph
                if any(
                    run.get("type") in {"citation", "degradation"}
                    for run in args.get("runs", ())
                )
                else composer.add_cross_reference_paragraph
            )
            result = native_method(
                **args, owner_node_id=op.node_id, controller_owned=True
            )
            self._consume_native_result(result, op)
            return

        if name == "writer.add_bibliography":
            if args.get("schemaVersion") == 1:
                result = composer.add_bibliography_native(
                    **args, owner_node_id=op.node_id, controller_owned=True
                )
            else:
                legacy_args = dict(args)
                legacy_args["entries"] = list(legacy_args.get("entries", ()))
                result = composer.add_bibliography_legacy(
                    **legacy_args, owner_node_id=op.node_id
                )
            self._consume_native_result(result, op)
            return

        if name == "writer.add_list":
            items = args.get("items", [])
            if args.get("ordered"):
                composer.add_numbered_list(items)
            else:
                composer.add_bullet_list(items, glyph=args.get("glyph", "•"))
            return

        if name == "writer.add_inline_degradation":
            composer.add_inline_degradation(
                code=args.get("code", "DEGRADATION"),
                message=args.get("message", ""),
                fallback_text=args.get("fallbackText", ""),
            )
            return

        if name == "writer.add_degradation_notice":
            composer.add_degradation_notice(
                code=args.get("code", "DEGRADATION"),
                message=args.get("message", ""),
                fallback_text=args.get("fallbackText", ""),
                placement=args.get("placement", "block"),
            )
            return

        if name == "writer.add_document_quality_notice":
            composer.add_document_quality_notice(args.get("notices", []))
            return

        if name == "writer.reserve_document_quality_anchor":
            composer.reserve_document_quality_anchor(
                title=args.get("title", "生成质量提示"),
                notices=args.get("notices", []),
            )
            return

        if name == "writer.insert_toc":
            composer.insert_toc_with_styles(
                title=args.get("title", "目录"),
                density=self._toc_density,
            )
            return

        if name == "writer.insert_figure_index":
            if "sequenceId" in args:
                composer.insert_caption_index_native(
                    title=args.get("title"),
                    sequence_id=args["sequenceId"],
                    title_style_id=args["titleStyleId"],
                    owner_node_id=op.node_id,
                )
                return
            composer.insert_figure_index(title=args.get("title"))
            return

        if name == "writer.insert_table_index":
            if "sequenceId" in args:
                composer.insert_caption_index_native(
                    title=args.get("title"),
                    sequence_id=args["sequenceId"],
                    title_style_id=args["titleStyleId"],
                    owner_node_id=op.node_id,
                )
                return
            composer.insert_table_index(title=args.get("title"))
            return

        if name == "writer.finalize_fields":
            # The outer executor owns convergence.  Dispatching this operation
            # must not pre-refresh fields or create a second convergence owner.
            if args.get("compactTerminalParagraph") is True:
                composer.compact_terminal_paragraph()
            return

        raise NativeWriterObjectError(
            UNKNOWN_OPERATION, "operation is not implemented"
        )

    def _consume_native_result(
        self, result: Any, op: GenerationOperation
    ) -> None:
        if not isinstance(result, dict):
            return
        for raw in result.get("issues", ()):
            if not isinstance(raw, dict):
                continue
            raw_node_id = raw.get("nodeId")
            issue_node_id = op.node_id
            if (
                isinstance(raw_node_id, str)
                and isinstance(op.node_id, str)
                and (
                    raw_node_id == op.node_id
                    or raw_node_id.startswith(op.node_id + "/")
                )
            ):
                issue_node_id = raw_node_id
            self._record_issue(
                code=str(raw.get("code") or EXECUTION_FAILED),
                message=str(raw.get("message") or "Native object reported an issue"),
                node_id=issue_node_id,
                placement=(
                    str(raw["placement"])
                    if raw.get("placement") in {"block", "inline", "document"}
                    else "document"
                ),
                stage=str(raw["stage"]) if raw.get("stage") else None,
                fallback=str(raw["fallback"]) if raw.get("fallback") else None,
                recoverable=(
                    raw.get("recoverable")
                    if isinstance(raw.get("recoverable"), bool)
                    else None
                ),
            )

    # ----------------------------------------------------------------------
    # Helpers
    # ----------------------------------------------------------------------
    def _record_issue(
        self,
        code: str,
        message: str,
        node_id: Optional[str] = None,
        placement: str = "document",
        stage: Optional[str] = None,
        fallback: Optional[str] = None,
        recoverable: Optional[bool] = None,
    ) -> None:
        issue = ExecutionIssue(
            code=code,
            message=message,
            placement=placement,
            node_id=node_id,
            stage=stage,
            fallback=fallback,
            recoverable=recoverable,
        )
        self._extend_issues((issue,))

    def _extend_issues(self, issues: Tuple[ExecutionIssue, ...]) -> None:
        for issue in issues:
            identity = (issue.code, issue.placement, issue.node_id)
            prior_index = next((
                index for index, prior in enumerate(self._issues)
                if (prior.code, prior.placement, prior.node_id) == identity
            ), None)
            if prior_index is None:
                self._issues.append(issue)
            elif (
                issue.recoverable is True
                and self._issues[prior_index].recoverable is not True
            ):
                self._issues[prior_index] = issue


class _ExecutionAbort(Exception):
    """Raised when an operation with mode=fail fails and generation must stop."""

    def __init__(self, op_name: str, cause: Exception) -> None:
        self.op_name = op_name
        self.cause = cause


class _DeferredOperationError(Exception):
    """Raised by the executor itself for M2-deferred ops."""

    def __init__(self, code: str, op_name: str, fallback: str) -> None:
        self.code = code
        self.op_name = op_name
        self.fallback = fallback


def _error_code(exc: Exception) -> str:
    """Extract a deterministic error code from an exception."""
    explicit = getattr(exc, "code", None)
    if isinstance(explicit, str) and explicit:
        return explicit
    cls_name = type(exc).__name__
    if cls_name == "COMError":
        return "COM_ERROR"
    return "EXECUTION_FAILED"


def _is_m3_operation(op: GenerationOperation) -> bool:
    if op.op not in _M3_NATIVE_OPERATIONS:
        return False
    markers = {
        "writer.add_captioned_figure": "numbering",
        "writer.add_semantic_table": "numbering",
        "writer.add_equation": "numbering",
        "writer.add_cross_reference": "runs",
        "writer.insert_figure_index": "sequenceId",
        "writer.insert_table_index": "sequenceId",
    }
    return markers[op.op] in op.args


def _reference_fallback_text(runs: Any) -> str:
    parts = []
    for run in runs or ():
        if not isinstance(run, Mapping):
            continue
        parts.append(str(run.get("text") if run.get("type") == "text" else run.get("fallbackText", "")))
    return "".join(parts)


def _table_fallback_text(args: Any) -> str:
    headers = args.get("headers") if isinstance(args, Mapping) else None
    rows = args.get("rows") if isinstance(args, Mapping) else None
    if not headers:
        return ""
    return "\n".join(" | ".join(str(cell) for cell in row) for row in [headers, *(rows or ())])


def _bibliography_fallback_text(args: Any) -> str:
    if not isinstance(args, Mapping) or args.get("schemaVersion") != 1:
        return ""
    return "\n".join(
        f"[{entry['number']}] {entry['text']}"
        for entry in args.get("entries", ())
    )


def _fallback_placement(fallback: str) -> str:
    if fallback in {"inline", "inline-fallback"}:
        return "inline"
    if fallback == "document-quality-notice":
        return "document"
    return "block"


def _operation_label(op_name: str) -> str:
    return {
        "writer.add_bibliography": "参考文献",
        "writer.add_captioned_figure": "图像",
        "writer.add_cross_reference": "引用",
        "writer.add_equation": "公式",
        "writer.add_semantic_table": "表格",
    }.get(op_name, "文档对象")


def _operation_fallback_text(op: GenerationOperation) -> str:
    args = op.args
    return redact_private_text(str(
        args.get("fallbackText")
        or args.get("source")
        or args.get("text")
        or _reference_fallback_text(args.get("runs", ()))
        or _bibliography_fallback_text(args)
        or _table_fallback_text(args)
        or args.get("caption")
        or ""
    ))


def _op_fallback_code(op: GenerationOperation) -> str:
    args = op.args
    if "code" in args:
        return str(args["code"])
    deferred = _M2_DEFERRED_OPERATIONS.get(op.op)
    if deferred is not None:
        return deferred[0]
    return "EXECUTION_FAILED"


def _op_fallback_message(op: GenerationOperation) -> str:
    args = op.args
    if "message" in args:
        return str(args["message"])
    return f"{op.op} could not be rendered"


def _extract_max_rounds(operations: Tuple[GenerationOperation, ...]) -> int:
    for operation in operations:
        if operation.op == "writer.finalize_fields":
            return operation.args.get("maxRounds", 3)
    return 3


def _build_pagination_map(
    operations: Tuple[GenerationOperation, ...],
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
    "EXECUTION_ABORTED",
    "EXECUTION_FAILED",
    "WINDOWS_DEDICATED_HOST_UNAVAILABLE",
    "WindowsDedicatedHostUnavailableError",
    "WindowsLongformExecutor",
    "WindowsLongformExecutorError",
]
