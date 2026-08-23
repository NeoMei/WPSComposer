"""Production platform adapters for the M5 long-form quality lifecycle."""

from __future__ import annotations

from dataclasses import replace
import os
from pathlib import Path
import shutil
import sys
import tempfile
import time
from typing import Any, Mapping, Optional, Sequence
from uuid import uuid4
import zipfile

from ..artifact_transport import (
    publish_artifact,
    validate_office_package,
    validate_pdf,
)
from ..generation_plan import GenerationOperation, GenerationPlan, validate_generation_plan
from ..macos_probe.bridge import LoopbackBridge
from ..macos_probe.conversion import ORIGINS, _wait_for_pdf_artifact
from ..macos_probe.longform_evidence import _wait_for_writer_registration
from ..macos_probe.models import PathPolicy, ProtocolError
from ..macos_probe.runtime import ProbeRuntime
from .executor import ExecutionOutcome
from .lifecycle import run_longform_lifecycle
from .macos_executor import MacOSLongformExecutor
from .pdf_quality import QualityPolicy, analyze_pdf
from .pipeline import LongformBuild, execute_longform_plan
from .quality import GenerationOutcome, PageRole, QualityFinding
from .relayout import RelayoutDirective
from .unicode_text import display_units, shorten_display_units
from .windows_executor import WindowsLongformExecutor, _create_dedicated_composer


_REPOSITORY_ROOT = Path(__file__).resolve().parents[4]
_PROBE_ROOT = _REPOSITORY_ROOT / "macos" / "wps-jsapi-probe"


def _bookmark_map(plan: GenerationPlan) -> dict[str, str]:
    result = {}
    for operation in plan.operations:
        bookmark = operation.args.get("bookmarkName")
        if operation.node_id and isinstance(bookmark, str):
            result[operation.node_id] = bookmark
    return result


def _body_width(plan: GenerationPlan) -> float:
    for operation in plan.operations:
        if operation.op == "writer.configure_page":
            left = float(operation.args.get("marginLeft", 72))
            right = float(operation.args.get("marginRight", 72))
            return max(72.0, 595.28 - left - right)
    return 451.28


def _thawed(value: Any) -> Any:
    if isinstance(value, Mapping):
        return {key: _thawed(item) for key, item in value.items()}
    if isinstance(value, tuple):
        return [_thawed(item) for item in value]
    return value


def _apply_relayout(
    plan: GenerationPlan, directives: Sequence[RelayoutDirective]
) -> GenerationPlan:
    """Derive the one allowed relayout plan without changing plan identity in M1."""

    by_node: dict[str, set[str]] = {}
    document_kinds = set()
    for directive in directives:
        if directive.node_id is None:
            document_kinds.add(directive.kind)
        else:
            by_node.setdefault(directive.node_id, set()).add(directive.kind)
    body_width = _body_width(plan)
    operations = []
    previous_operation: Optional[GenerationOperation] = None
    for operation in plan.operations:
        args = _thawed(operation.args)
        kinds = by_node.get(operation.node_id or "", set())
        if "shorten-header" in document_kinds or "shorten-header" in kinds:
            for key in ("header", "headerText"):
                if isinstance(args.get(key), str):
                    args[key] = shorten_display_units(args[key], 32)
        if operation.op == "writer.add_captioned_figure" and "fit-image" in kinds:
            args["widthMode"] = "explicit"
            args["explicitWidthPt"] = round(body_width, 4)
            children = []
            for child in args.get("children", []):
                child = dict(child)
                width = child.get("displayWidthPt")
                height = child.get("displayHeightPt")
                if isinstance(width, (int, float)) and width > body_width:
                    scale = body_width / float(width)
                    child["displayWidthPt"] = round(body_width, 4)
                    if isinstance(height, (int, float)):
                        child["displayHeightPt"] = round(float(height) * scale, 4)
                children.append(child)
            args["children"] = children
        if operation.op == "writer.add_semantic_table":
            if args.get("orientation") == "landscape":
                args["continuousExit"] = True
                if (
                    previous_operation is not None
                    and previous_operation.op == "writer.add_heading"
                ):
                    args["includePreviousHeading"] = True
            if "compress-table" in kinds:
                args["cellIndentPt"] = 0.0
                args["allowRowSplit"] = True
                args["m5Relayout"] = "compress-table"
            if "force-table-split" in kinds:
                args["merges"] = []
                args["allowRowSplit"] = True
                args["m5Relayout"] = "force-table-split"
        if operation.op == "writer.add_heading" and "keep-heading" in kinds:
            args["keepWithNext"] = True
        if (
            operation.op == "writer.finalize_fields"
            and "remove-unexpected-blank" in document_kinds
        ):
            args["compactTerminalParagraph"] = True
        if operation.op in {"writer.add_captioned_figure", "writer.add_semantic_table"} and "keep-caption" in kinds:
            key = "keepWithCaption" if operation.op.endswith("figure") else "keepCaptionWithFirstRow"
            args[key] = True
        operations.append(
            GenerationOperation(
                op=operation.op,
                args=args,
                node_id=operation.node_id,
                failure_policy=_thawed(operation.failure_policy) if operation.failure_policy else None,
            )
        )
        previous_operation = operation
    derived = GenerationPlan(
        component=plan.component,
        operations=tuple(operations),
        protocol_version=plan.protocol_version,
        semantic_version=plan.semantic_version,
        resource_manifest_version=plan.resource_manifest_version,
        resource_manifest_digest=plan.resource_manifest_digest,
    )
    validate_generation_plan(derived.to_dict(), component="writer")
    return derived


def _quality_context(build: LongformBuild, outcome: ExecutionOutcome):
    page_roles: dict[int, PageRole] = {}
    for node in outcome.pagination_map.nodes:
        raw_role = node.sections[0] if node.sections else "body"
        try:
            role = PageRole(raw_role)
        except ValueError:
            role = PageRole.BODY
        for fragment in node.fragments:
            page_roles.setdefault(fragment.page, role)

    node_kinds = []
    effective_dpi = []
    table_headers = []
    unresolved = []
    header_units = 0
    for operation in build.plan.operations:
        kind = {
            "writer.add_heading": "heading",
            "writer.add_captioned_figure": "image",
            "writer.add_semantic_table": "table",
            "writer.add_equation": "formula",
        }.get(operation.op)
        if operation.node_id and kind:
            node_kinds.append((operation.node_id, kind))
        if operation.op == "writer.add_captioned_figure" and operation.node_id:
            dpis = [
                float(child["effectiveDpi"])
                for child in operation.args.get("children", ())
                if isinstance(child.get("effectiveDpi"), (int, float))
                and float(child["effectiveDpi"]) > 0
                and child.get("mediaType") != "image/svg+xml"
            ]
            if dpis:
                effective_dpi.append((operation.node_id, min(dpis)))
        if operation.op == "writer.add_semantic_table" and operation.node_id:
            table_headers.append((operation.node_id, bool(operation.args.get("repeatHeader"))))
        for key in ("header", "headerText"):
            if isinstance(operation.args.get(key), str):
                header_units = max(header_units, display_units(operation.args[key]))
    for issue in outcome.issues:
        if issue.node_id and issue.code in {"CROSS_REFERENCE_FAILED", "REFERENCE_UNRESOLVED"}:
            unresolved.append(issue.node_id)
    policy = QualityPolicy(
        node_kinds=tuple(node_kinds),
        effective_dpi=tuple(effective_dpi),
        table_header_repeated=tuple(table_headers),
        unresolved_field_nodes=tuple(sorted(set(unresolved))),
        header_display_units=header_units or None,
    )
    return page_roles, policy


class _BaseAdapter:
    def __init__(self, build: LongformBuild) -> None:
        self.build = build
        self.bookmarks = _bookmark_map(build.plan)

    def analyze(self, pdf: Path, outcome: ExecutionOutcome, build: LongformBuild):
        if outcome.pagination_map.version != "M5-v1":
            raise RuntimeError("native executor did not return the M5 pagination contract")
        roles, policy = _quality_context(build, outcome)
        return analyze_pdf(pdf, outcome.pagination_map, roles, policy)

    @staticmethod
    def validate_patch(
        docx: Path,
        pdf: Path,
        notices: Sequence[QualityFinding],
        outcome: ExecutionOutcome,
    ) -> None:
        validate_office_package(docx, "docx")
        validate_pdf(pdf)
        if outcome.pagination_map.version != "M5-v1":
            raise ValueError("notice patch did not return fresh M5 pagination")
        with zipfile.ZipFile(docx) as package:
            visible = package.read("word/document.xml").decode("utf-8", "ignore")
        for notice in notices:
            if notice.code not in visible:
                raise ValueError("notice marker is absent from the patched document")

    @staticmethod
    def publish(staged: Path, output: Path, overwrite: bool, deadline: float) -> Path:
        validator = validate_pdf if output.suffix.lower() == ".pdf" else (
            lambda path: validate_office_package(path, "docx")
        )
        return publish_artifact(
            staged,
            output,
            overwrite=overwrite,
            validator=validator,
            deadline=deadline,
        )

    @staticmethod
    def cleanup(path: Path) -> None:
        Path(path).unlink(missing_ok=True)


class MacLongformAdapter(_BaseAdapter):
    def __init__(self, build: LongformBuild) -> None:
        super().__init__(build)
        self.runtime_root = Path(tempfile.mkdtemp(prefix="wpscomposer-longform-m5-"))
        os.chmod(self.runtime_root, 0o700)
        self.bridge: Optional[LoopbackBridge] = None
        self.runtime: Optional[ProbeRuntime] = None
        self.executor: Optional[MacOSLongformExecutor] = None

    def _ensure_started(self, deadline: float) -> None:
        if self.executor is not None:
            return
        self.bridge = LoopbackBridge(ORIGINS)
        self.bridge.__enter__()
        self.runtime = ProbeRuntime(
            _PROBE_ROOT,
            self.runtime_root / "runtime",
            self.bridge.url,
            self.bridge.token,
            deadline=deadline,
        )
        self.runtime.__enter__()
        self.runtime.prepare_profiles()
        self.runtime.start_servers(deadline=deadline)
        self.runtime.activate_component("writer", isolated=True, deadline=deadline)
        remaining = max(0.0, deadline - time.monotonic())
        _wait_for_writer_registration(
            self.bridge, self.runtime, timeout=min(60.0, remaining)
        )
        if self.runtime.staging_dir is None:
            raise RuntimeError("macOS WPS staging is unavailable")
        self.executor = MacOSLongformExecutor(
            bridge=self.bridge, staging_dir=str(self.runtime.staging_dir)
        )

    def execute(self, build, directives, deadline):
        self._ensure_started(deadline)
        assert self.executor is not None
        derived = replace(build, plan=_apply_relayout(build.plan, directives))
        return execute_longform_plan(derived, self.executor, deadline=deadline)

    def export_pdf(self, docx: Path, deadline: float) -> Path:
        self._ensure_started(deadline)
        assert self.bridge is not None and self.runtime is not None
        assert self.runtime.staging_dir is not None
        policy = PathPolicy((self.runtime.staging_dir,))
        source = policy.require_allowed(docx)
        target = policy.require_allowed(
            self.runtime.staging_dir / f"wpsc-quality-{uuid4().hex}.pdf"
        )
        command = self.bridge.issue(
            "writer", "convert_writer_pdf",
            {"sourcePath": str(source), "outputPath": str(target)},
        )
        result = self.bridge.wait_result(
            command.id, timeout=max(0.0, deadline - time.monotonic())
        )
        if not result.ok:
            raise RuntimeError("macOS WPS PDF export failed")
        try:
            reported = policy.require_allowed(str(result.value.get("path", "")))
        except ProtocolError:
            raise RuntimeError("macOS WPS PDF export returned an invalid path") from None
        if reported != target:
            raise RuntimeError("macOS WPS PDF export returned an unexpected path")
        # WPS can acknowledge SaveAs before the filesystem has flushed the
        # complete PDF.  Reuse the deadline-bound readiness validator instead
        # of opening a new timeout window or racing on the first read.
        _wait_for_pdf_artifact(target, deadline=deadline)
        return target

    def patch_quality_notices(self, docx, notices, deadline):
        self._ensure_started(deadline)
        assert self.executor is not None
        return self.executor.patch_quality_notices(
            docx, tuple(notices), self.bookmarks, deadline=deadline
        )

    def close(self) -> None:
        runtime = self.runtime
        try:
            if runtime is not None:
                runtime.__exit__(None, None, None)
        finally:
            if self.bridge is not None:
                self.bridge.__exit__(None, None, None)
        if runtime is None or getattr(runtime, "registration_restored", True):
            shutil.rmtree(self.runtime_root, ignore_errors=True)


class WindowsLongformAdapter(_BaseAdapter):
    def __init__(self, build: LongformBuild) -> None:
        super().__init__(build)
        self.staging_root = Path(tempfile.mkdtemp(prefix="wpscomposer-longform-m5-win-"))
        os.chmod(self.staging_root, 0o700)
        self.executor = WindowsLongformExecutor(staging_dir=str(self.staging_root))

    def execute(self, build, directives, deadline):
        derived = replace(build, plan=_apply_relayout(build.plan, directives))
        return execute_longform_plan(derived, self.executor, deadline=deadline)

    def export_pdf(self, docx: Path, deadline: float) -> Path:
        if time.monotonic() >= deadline:
            raise TimeoutError("PDF export deadline expired")
        target = self.staging_root / f"wpsc-quality-{uuid4().hex}.pdf"
        composer = _create_dedicated_composer(str(self.staging_root))
        try:
            composer._doc.Close(False)
            composer._doc = composer._app.Documents.Open(str(docx), False, True)
            composer.export_pdf(str(target))
        finally:
            composer.close(save_changes=False)
        validate_pdf(target)
        return target

    def patch_quality_notices(self, docx, notices, deadline):
        return self.executor.patch_quality_notices(
            docx, tuple(notices), self.bookmarks, deadline=deadline
        )

    def close(self) -> None:
        shutil.rmtree(self.staging_root, ignore_errors=True)


def generate_longform(
    build: LongformBuild,
    *,
    format_name: str,
    output: Path,
    timeout: float,
    overwrite: bool,
) -> GenerationOutcome:
    """Run the public default DOCX/PDF route on the current native platform."""

    if sys.platform == "darwin":
        adapter: Any = MacLongformAdapter(build)
    elif sys.platform == "win32":
        adapter = WindowsLongformAdapter(build)
    else:
        raise RuntimeError("WPS long-form generation requires macOS or Windows")
    try:
        return run_longform_lifecycle(
            build,
            adapter,
            Path(output),
            format_name=format_name,
            timeout=timeout,
            overwrite=overwrite,
            quality_analyzer=adapter.analyze,
        )
    finally:
        adapter.close()


__all__ = [
    "MacLongformAdapter",
    "WindowsLongformAdapter",
    "generate_longform",
]
