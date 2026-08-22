"""Windows COM executor for long-form generation plans.

This module is importable on macOS/Linux because all pywin32 imports are
performed lazily inside the dedicated-composer factory.
"""
from __future__ import annotations

import inspect
import hashlib
import json
import os
import sys
import tempfile
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


WINDOWS_DEDICATED_HOST_UNAVAILABLE = "WINDOWS_DEDICATED_HOST_UNAVAILABLE"
EXECUTION_FAILED = "EXECUTION_FAILED"
EXECUTION_ABORTED = "EXECUTION_ABORTED"
DEGRADATION_FALLBACK_FAILED = "DEGRADATION_FALLBACK_FAILED"
UNKNOWN_OPERATION = "UNKNOWN_OPERATION"

# Operations whose native rendering is intentionally deferred past M3.  The
# executor emits a deterministic degradation notice/inline fallback for them
# and records the stable issue code declared by the plan builder.
_M2_DEFERRED_OPERATIONS = {
    "writer.add_bibliography": ("BIBLIOGRAPHY_INSERT_FAILED", "notice"),
}

_M3_NATIVE_OPERATIONS = frozenset({
    "writer.add_captioned_figure",
    "writer.add_semantic_table",
    "writer.add_equation",
    "writer.add_cross_reference",
    "writer.insert_figure_index",
    "writer.insert_table_index",
})

_LEGACY_OBJECT_DEFERRED = {
    "writer.add_captioned_figure": ("IMAGE_INSERT_FAILED", "notice"),
    "writer.add_semantic_table": ("TABLE_INSERT_FAILED", "notice"),
    "writer.add_equation": ("EQUATION_INSERT_FAILED", "inline"),
    "writer.add_cross_reference": ("CROSS_REFERENCE_FAILED", "inline"),
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
    last_error: Optional[Exception] = None
    app: Any = None

    for progid in WriterComposer._progids:
        try:
            app = client.DispatchEx(progid)
            break
        except Exception as exc:  # pragma: no cover - exercised via mocks
            last_error = exc
            continue

    if app is None:
        pythoncom.CoUninitialize()
        detail = f": {last_error}" if last_error else ""
        raise WindowsDedicatedHostUnavailableError(
            f"Could not create dedicated WPS application{detail}"
        )

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

    try:
        composer._doc = composer._create_doc(app)
    except Exception as exc:
        composer.close(save_changes=False)
        raise WindowsDedicatedHostUnavailableError(
            f"Could not create dedicated WPS document: {exc}"
        ) from exc

    composer._owns_doc = True
    if staging_dir:
        composer._staging_dir = staging_dir
    return composer


@dataclass(frozen=True)
class _ResolvedPaths:
    staged_docx: str


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
        self._resource_locators: dict[str, str] = {}

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
        validate_generation_plan(plan.to_dict(), component="writer")
        paths = self._resolve_paths()
        staged_resources: Tuple[str, ...] = ()
        pending_cleanup: Tuple[str, ...] = ()
        cleanup_attempted = False
        composer: Optional[WriterComposer] = None
        primary_error: Optional[Exception] = None
        cleanup_failed = False
        try:
            self._validate_resource_manifest(plan, resources)
            self._resource_locators, staged_resources = self._stage_resources(resources)
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
            composer.save_docx(paths.staged_docx)
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
            pagination_map=_build_pagination_map(plan.operations),
            applied_operations=len(plan.operations),
        )

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
        for op in operations:
            self._run_op(composer, op)

    def _run_op(self, composer: WriterComposer, op: GenerationOperation) -> None:
        policy = op.failure_policy
        deferred = _M2_DEFERRED_OPERATIONS.get(op.op)
        if deferred is None and not _is_m3_operation(op):
            deferred = _LEGACY_OBJECT_DEFERRED.get(op.op)

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
            self._record_issue(
                code=exc.code,
                message=f"{op.op} is deferred to fallback in M2",
                node_id=op.node_id,
            )
            self._apply_fallback(composer, op, exc.fallback)
        except Exception as exc:
            code = _error_code(exc)
            if policy is not None and policy.get("mode") == "degrade":
                if code in policy.get("recoverableCodes", []):
                    self._degrade_op(composer, op, exc, code, policy["fallback"])
                    return
            if _is_m3_operation(op):
                raise _ExecutionAbort(op.op, exc) from exc
            self._record_issue(
                code=EXECUTION_FAILED,
                message=f"{op.op} failed",
                node_id=op.node_id,
            )
            if policy is not None and policy.get("mode") == "fail":
                raise _ExecutionAbort(op.op, exc) from exc
            # No explicit fail policy: record the issue and continue so the
            # document is not left empty.

    def _degrade_op(
        self,
        composer: WriterComposer,
        op: GenerationOperation,
        exc: Exception,
        code: str,
        fallback: str,
    ) -> None:
        self._record_issue(
            code=code,
            message=f"{op.op} used its declared native fallback",
            node_id=op.node_id,
        )
        try:
            self._apply_fallback(composer, op, fallback)
        except Exception as fb_exc:
            self._record_issue(
                code=DEGRADATION_FALLBACK_FAILED,
                message=f"Fallback for {op.op} failed",
                node_id=op.node_id,
            )
            raise _ExecutionAbort(op.op, fb_exc) from fb_exc

    def _apply_fallback(
        self,
        composer: WriterComposer,
        op: GenerationOperation,
        fallback: str,
    ) -> None:
        args = op.args
        if fallback in {"inline", "inline-fallback"}:
            text = str(
                args.get("fallbackText")
                or args.get("source")
                or args.get("text")
                or _reference_fallback_text(args.get("runs", ()))
                or ""
            )
            composer.add_inline_degradation(
                code=_op_fallback_code(op),
                message=_op_fallback_message(op),
                fallback_text=text,
            )
        elif fallback in {
            "notice", "figure-child-stack-then-notice", "grid-then-text"
        }:
            text = str(
                args.get("fallbackText")
                or args.get("source")
                or args.get("text")
                or _table_fallback_text(args)
                or args.get("caption")
                or ""
            )
            composer.add_degradation_notice(
                code=_op_fallback_code(op),
                message=_op_fallback_message(op),
                fallback_text=text,
                placement=args.get("placement", "block"),
            )
        else:
            composer.add_degradation_notice(
                code=UNKNOWN_OPERATION,
                message=f"No fallback implementation for {op.op}",
                fallback_text=f"[{UNKNOWN_OPERATION}] {op.op}",
                placement="block",
            )

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
                landscape=args.get("landscape"),
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
            composer.add_heading_level_native(
                text=args["text"],
                level=args.get("level", 1),
                numbering=args.get("numbering"),
                scheme=args.get("numberingScheme"),
            )
            return

        if name == "writer.add_paragraph":
            composer.add_paragraph(
                text=args.get("text", ""),
                style=args.get("style"),
            )
            return

        if name == "writer.add_captioned_figure" and "numbering" in args:
            result = composer.add_captioned_figure_native(
                **args,
                owner_node_id=op.node_id,
                resource_locators=dict(self._resource_locators),
            )
            self._consume_native_result(result, op.node_id)
            return

        if name == "writer.add_semantic_table" and "numbering" in args:
            result = composer.add_semantic_table_native(
                **args, owner_node_id=op.node_id
            )
            self._consume_native_result(result, op.node_id)
            return

        if name == "writer.add_equation" and "numbering" in args:
            result = composer.add_equation_number_native(
                **args, owner_node_id=op.node_id
            )
            self._consume_native_result(result, op.node_id)
            return

        if name == "writer.add_cross_reference" and "runs" in args:
            result = composer.add_cross_reference_paragraph(
                **args, owner_node_id=op.node_id
            )
            self._consume_native_result(result, op.node_id)
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
            return

        # Fallback for anything else that reaches the executor.
        self._apply_fallback(composer, op, "notice")

    def _consume_native_result(self, result: Any, node_id: Optional[str]) -> None:
        if not isinstance(result, dict):
            return
        for raw in result.get("issues", ()):
            if not isinstance(raw, dict):
                continue
            self._record_issue(
                code=str(raw.get("code") or EXECUTION_FAILED),
                message="Native object used its declared recovery",
                node_id=node_id,
                placement=(
                    str(raw["placement"])
                    if raw.get("placement") in {"block", "inline", "document"}
                    else "document"
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
    ) -> None:
        issue = ExecutionIssue(
            code=code,
            message=message,
            placement=placement,
            node_id=node_id,
        )
        self._extend_issues((issue,))

    def _extend_issues(self, issues: Tuple[ExecutionIssue, ...]) -> None:
        identities = {
            (issue.code, issue.message, issue.placement, issue.node_id)
            for issue in self._issues
        }
        for issue in issues:
            identity = (issue.code, issue.message, issue.placement, issue.node_id)
            if identity not in identities:
                self._issues.append(issue)
                identities.add(identity)


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
        if not isinstance(run, dict):
            continue
        parts.append(str(run.get("text") if run.get("type") == "text" else run.get("fallbackText", "")))
    return "".join(parts)


def _table_fallback_text(args: Any) -> str:
    headers = args.get("headers") if isinstance(args, Mapping) else None
    rows = args.get("rows") if isinstance(args, Mapping) else None
    if not headers:
        return ""
    return "\n".join(" | ".join(str(cell) for cell in row) for row in [headers, *(rows or ())])


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
