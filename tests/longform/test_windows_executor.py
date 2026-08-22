"""Tests for the Windows COM long-form executor using mocks.

These tests never start WPS. They patch win32com.client.DispatchEx and the
WriterComposer COM primitives to verify operation dispatch and convergence.
"""

from __future__ import annotations

import subprocess
import sys
import types
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    GenerationOperation,
    GenerationPlan,
)
from skills.WPSComposer.scripts.longform.executor import (
    ExecutionIssue,
    ExecutionOutcome,
    FIELD_REFRESH_UNSTABLE,
    FieldSnapshot,
    LongformExecutor,
)
from skills.WPSComposer.scripts.longform.windows_executor import (
    WINDOWS_DEDICATED_HOST_UNAVAILABLE,
    WindowsDedicatedHostUnavailableError,
    WindowsLongformExecutor,
    WindowsLongformExecutorError,
)


# -----------------------------------------------------------------------------
# Mocks
# -----------------------------------------------------------------------------

class _FakePythoncomModule:
    def CoInitialize(self) -> None:
        pass

    def CoUninitialize(self) -> None:
        pass


class _FakeDispatchModule:
    """Replacement for win32com.client used by WindowsLongformExecutor."""

    def __init__(self, *, fail: bool = False) -> None:
        self._fail = fail
        self._calls: List[Tuple[str, ...]] = []

    def DispatchEx(self, progid: str) -> object:
        self._calls.append(("DispatchEx", progid))
        if self._fail:
            raise RuntimeError("WPS not available")
        return object()

    def Dispatch(self, progid: str) -> object:
        self._calls.append(("Dispatch", progid))
        raise RuntimeError("shared Dispatch should not be used")


# -----------------------------------------------------------------------------
# Fake WriterComposer that records every long-form primitive
# -----------------------------------------------------------------------------

@dataclass
class _PrimitiveCall:
    name: str
    args: Tuple[Any, ...]
    kwargs: Dict[str, Any]


class FakeWriterComposer:
    """Records long-form COM primitive calls and returns deterministic pagination."""

    _progids = ("KWps.Application",)

    def __init__(self) -> None:
        self.primitives: List[_PrimitiveCall] = []
        self.saved_path: Optional[str] = None
        self.closed = False
        self._snapshots: List[Tuple[FieldSnapshot, ...]] = []
        self._snapshot_index = 0

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close(save_changes=False)
        return False

    # -- lifecycle --
    def close(self, save_changes: bool = False) -> None:
        self.closed = True

    def save(self, path: str, fmt: Optional[int] = None) -> str:
        self.saved_path = str(path)
        return self.saved_path

    def save_docx(self, path: str) -> str:
        return self.save(path)

    # -- field refresh for convergence --
    def refresh_fields(self, round_index: int) -> Tuple[FieldSnapshot, ...]:
        if self._snapshots:
            snap = self._snapshots[self._snapshot_index % len(self._snapshots)]
            self._snapshot_index += 1
            return snap
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

    # -- existing primitives --
    def reset(self) -> None:
        self._record("reset")

    def set_margins(self, *args) -> None:
        self._record("set_margins", *args)

    def set_page_size(self, *args) -> None:
        self._record("set_page_size", *args)

    def ensure_styles(self, styles_dict) -> None:
        self._record("ensure_styles", styles_dict)

    def add_paragraph(self, **kwargs) -> None:
        self._record("add_paragraph", **kwargs)

    def add_numbered_list(self, **kwargs) -> None:
        self._record("add_numbered_list", **kwargs)

    def add_bullet_list(self, **kwargs) -> None:
        self._record("add_bullet_list", **kwargs)

    def add_page_break(self) -> None:
        self._record("add_page_break")

    # -- new M2 primitives --
    def set_page_role(self, role: str) -> None:
        self._record("set_page_role", role=role)

    def set_page_numbering(
        self,
        format: str,
        start: Optional[int] = None,
        restart: Optional[bool] = None,
    ) -> None:
        self._record("set_page_numbering", format=format, start=start, restart=restart)

    def set_header_footer(
        self,
        header: Optional[str] = None,
        footer: Optional[str] = None,
        link_to_previous_header: Optional[bool] = None,
        link_to_previous_footer: Optional[bool] = None,
    ) -> None:
        self._record(
            "set_header_footer",
            header=header,
            footer=footer,
            link_to_previous_header=link_to_previous_header,
            link_to_previous_footer=link_to_previous_footer,
        )

    def configure_section(self, **kwargs) -> None:
        self._record("configure_section", **kwargs)

    def insert_toc_with_styles(self, title: str, density: Dict[str, Any]) -> None:
        self._record("insert_toc_with_styles", title=title, density=density)

    def insert_figure_index(self, title: Optional[str] = None) -> None:
        self._record("insert_figure_index", title=title)

    def insert_table_index(self, title: Optional[str] = None) -> None:
        self._record("insert_table_index", title=title)

    def add_degradation_notice(
        self,
        code: str,
        message: str,
        fallback_text: str,
        placement: str = "block",
    ) -> None:
        self._record(
            "add_degradation_notice",
            code=code,
            message=message,
            fallback_text=fallback_text,
            placement=placement,
        )

    def add_inline_degradation(
        self,
        code: str,
        message: str,
        fallback_text: str,
    ) -> None:
        self._record(
            "add_inline_degradation",
            code=code,
            message=message,
            fallback_text=fallback_text,
        )

    def add_document_quality_notice(self, notices: list) -> None:
        self._record("add_document_quality_notice", notices=notices)

    def add_heading_level_native(
        self,
        text: str,
        level: int,
        numbering: Optional[bool] = None,
        scheme: Optional[str] = None,
    ) -> None:
        self._record(
            "add_heading_level_native",
            text=text,
            level=level,
            numbering=numbering,
            scheme=scheme,
        )

    def finalize_fields(self, *, max_rounds: int = 3) -> None:
        self._record("finalize_fields", max_rounds=max_rounds)

    # -- helpers --
    def _record(self, name: str, *args, **kwargs) -> None:
        self.primitives.append(_PrimitiveCall(name, args, kwargs))


# -----------------------------------------------------------------------------
# Fixtures
# -----------------------------------------------------------------------------

@pytest.fixture
def project_root() -> Path:
    return Path(__file__).parents[2]


@pytest.fixture
def fake_composer():
    return FakeWriterComposer()


@pytest.fixture
def executor(fake_composer):
    return WindowsLongformExecutor(composer_factory=lambda: fake_composer)


@pytest.fixture
def simple_plan() -> GenerationPlan:
    """A minimal protocol v2 plan with the key M2 operations."""
    ops = [
        {"op": "writer.reset", "args": {}},
        {
            "op": "writer.configure_page",
            "args": {
                "marginTop": 72,
                "marginBottom": 72,
                "marginLeft": 90,
                "marginRight": 90,
            },
        },
        {
            "op": "writer.configure_section",
            "args": {
                "role": "front_matter",
                "pageNumberFormat": "roman",
                "restartPageNumbering": True,
                "startPageNumber": 1,
                "headerText": "",
                "footerText": "",
                "linkToPreviousHeader": False,
                "linkToPreviousFooter": False,
            },
            "nodeId": "doc:section-0:front_matter",
        },
        {
            "op": "writer.configure_toc_styles",
            "args": {
                "tocTitle": "目录",
                "levels": 3,
                "includeFigureIndex": False,
                "includeTableIndex": False,
                "minFontSizePt": {"toc1": 10.5, "toc2": 10.0, "toc3": 10.0},
                "minSpaceBeforePt": {"toc1": 0.0, "toc2": 0.0, "toc3": 0.0},
                "minSpaceAfterPt": {"toc1": 0.0, "toc2": 0.0, "toc3": 0.0},
            },
            "nodeId": "doc:toc",
        },
        {
            "op": "writer.insert_toc",
            "args": {"title": "目录"},
            "nodeId": "doc:toc",
        },
        {
            "op": "writer.configure_section",
            "args": {
                "role": "body",
                "pageNumberFormat": "arabic",
                "restartPageNumbering": True,
                "startPageNumber": 1,
                "headerText": "Chapter Header",
                "footerText": "",
                "linkToPreviousHeader": False,
                "linkToPreviousFooter": False,
            },
            "nodeId": "doc:section-1:body",
        },
        {
            "op": "writer.add_heading",
            "args": {
                "text": "Introduction",
                "level": 1,
                "numbering": True,
                "numberingScheme": "decimal",
            },
            "nodeId": "sec:1",
        },
        {
            "op": "writer.add_paragraph",
            "args": {"text": "Body text.", "style": "Body Text"},
            "nodeId": "sec:1:p1",
        },
        {
            "op": "writer.finalize_fields",
            "args": {"maxRounds": 3},
            "nodeId": "doc:finalize",
        },
    ]
    operations = []
    for item in ops:
        operations.append(
            GenerationOperation(
                op=item["op"],
                args=item["args"],
                node_id=item.get("nodeId"),
                failure_policy=item.get("failurePolicy"),
            )
        )
    return GenerationPlan(
        component="writer",
        operations=tuple(operations),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )


# -----------------------------------------------------------------------------
# Purity: module imports on macOS
# -----------------------------------------------------------------------------

def test_windows_executor_module_imports_without_pywin32(project_root: Path):
    """Importing the module on a non-Windows host must not require pywin32."""
    script = """
import sys
sys.modules.pop('win32com', None)
sys.modules.pop('win32com.client', None)
sys.modules.pop('pythoncom', None)
from skills.WPSComposer.scripts.longform.windows_executor import WindowsLongformExecutor
print('imported')
"""
    result = subprocess.run(
        [sys.executable, "-c", script],
        capture_output=True,
        text=True,
        cwd=str(project_root),
    )
    assert result.returncode == 0, result.stderr
    assert "imported" in result.stdout


def test_windows_executor_is_longform_executor(executor):
    assert isinstance(executor, LongformExecutor)


# -----------------------------------------------------------------------------
# Operation dispatch
# -----------------------------------------------------------------------------

def test_execute_dispatches_section_and_toc_and_heading(
    executor, fake_composer, simple_plan
):
    outcome = executor.execute(simple_plan, ())

    assert isinstance(outcome, ExecutionOutcome)
    assert outcome.staged_artifact.endswith(".docx")
    assert Path(outcome.staged_artifact).name.startswith("wpsc-longform-")

    names = [call.name for call in fake_composer.primitives]
    assert "configure_section" in names
    assert "insert_toc_with_styles" in names
    assert "add_heading_level_native" in names
    assert "add_paragraph" in names
    assert "finalize_fields" in names


def test_configure_section_carries_roman_and_arabic_args(
    executor, fake_composer, simple_plan
):
    executor.execute(simple_plan, ())

    section_calls = [c for c in fake_composer.primitives if c.name == "configure_section"]
    assert len(section_calls) == 2

    front = section_calls[0].kwargs
    assert front["page_number_format"] == "roman"
    assert front["restart_page_numbering"] is True
    assert front["start_page_number"] == 1

    body = section_calls[1].kwargs
    assert body["page_number_format"] == "arabic"
    assert body["restart_page_numbering"] is True
    assert body["start_page_number"] == 1
    assert body["header_text"] == "Chapter Header"


def test_toc_density_styles_passed_to_primitive(executor, fake_composer, simple_plan):
    executor.execute(simple_plan, ())
    toc_calls = [c for c in fake_composer.primitives if c.name == "insert_toc_with_styles"]
    assert len(toc_calls) == 1
    density = toc_calls[0].kwargs["density"]
    assert density["minFontSizePt"]["toc1"] == 10.5


def test_heading_native_numbering_args(executor, fake_composer, simple_plan):
    executor.execute(simple_plan, ())
    heading = [c for c in fake_composer.primitives if c.name == "add_heading_level_native"][0]
    assert heading.kwargs["text"] == "Introduction"
    assert heading.kwargs["level"] == 1
    assert heading.kwargs["numbering"] is True
    assert heading.kwargs["scheme"] == "decimal"


# -----------------------------------------------------------------------------
# Header / footer and link-to-previous
# -----------------------------------------------------------------------------

def test_header_footer_is_set_by_configure_section(
    executor, fake_composer, simple_plan
):
    executor.execute(simple_plan, ())
    section_calls = [c for c in fake_composer.primitives if c.name == "configure_section"]
    body = section_calls[1].kwargs
    assert body["header_text"] == "Chapter Header"
    assert body["footer_text"] == ""
    assert body["link_to_previous_header"] is False


# -----------------------------------------------------------------------------
# Convergence loop
# -----------------------------------------------------------------------------

def test_convergence_uses_composer_refresh_fields(
    executor, fake_composer, simple_plan
):
    fake_composer._snapshots = [
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="hash1",
                toc_page_count=2,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=5,
            ),
        ),
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="hash1",
                toc_page_count=2,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=5,
            ),
        ),
    ]
    outcome = executor.execute(simple_plan, ())
    assert outcome.issues == ()


def test_convergence_emits_field_refresh_unstable(
    executor, fake_composer, simple_plan
):
    fake_composer._snapshots = [
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="hash1",
                toc_page_count=2,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=5,
            ),
        ),
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="hash2",
                toc_page_count=2,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=5,
            ),
        ),
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="hash3",
                toc_page_count=2,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=5,
            ),
        ),
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="hash4",
                toc_page_count=2,
                figure_index_page_count=0,
                table_index_page_count=0,
                total_pages=5,
            ),
        ),
    ]
    outcome = executor.execute(simple_plan, ())
    assert len(outcome.issues) == 1
    assert outcome.issues[0].code == FIELD_REFRESH_UNSTABLE


# -----------------------------------------------------------------------------
# Dedicated host ownership
# -----------------------------------------------------------------------------

def test_dedicated_host_unavailable_without_dispatch(monkeypatch):
    """If DispatchEx fails for every ProgID, raise the dedicated-host error."""
    fake_win32 = _FakeDispatchModule(fail=True)
    fake_pythoncom = _FakePythoncomModule()
    monkeypatch.setitem(sys.modules, "win32com", types.ModuleType("win32com"))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    executor = WindowsLongformExecutor()
    plan = GenerationPlan(
        component="writer",
        operations=(GenerationOperation(op="writer.reset", args={}),),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    with pytest.raises(WindowsDedicatedHostUnavailableError) as exc_info:
        executor.execute(plan, ())

    assert exc_info.value.code == WINDOWS_DEDICATED_HOST_UNAVAILABLE
    assert any(c[0] == "DispatchEx" for c in fake_win32._calls)


def test_no_shared_dispatch_fallback(monkeypatch):
    """Dispatch is never used; only DispatchEx is attempted."""
    fake_win32 = types.ModuleType("win32com.client")
    calls: List[str] = []

    def fake_dispatch(progid: str) -> Any:
        calls.append(f"Dispatch:{progid}")
        return object()

    def fake_dispatch_ex(progid: str) -> Any:
        calls.append(f"DispatchEx:{progid}")
        raise RuntimeError("dedicated unavailable")

    fake_win32.Dispatch = fake_dispatch
    fake_win32.DispatchEx = fake_dispatch_ex
    fake_pythoncom = _FakePythoncomModule()

    monkeypatch.setitem(sys.modules, "win32com", types.ModuleType("win32com"))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_win32)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    executor = WindowsLongformExecutor()
    plan = GenerationPlan(
        component="writer",
        operations=(GenerationOperation(op="writer.reset", args={}),),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    with pytest.raises(WindowsDedicatedHostUnavailableError):
        executor.execute(plan, ())

    assert any("DispatchEx" in c for c in calls)
    assert not any("Dispatch:" in c for c in calls)


# -----------------------------------------------------------------------------
# Resource staging hygiene
# -----------------------------------------------------------------------------

def test_staged_artifact_path_is_inside_staging_dir(tmp_path, fake_composer):
    executor = WindowsLongformExecutor(
        staging_dir=str(tmp_path),
        composer_factory=lambda: fake_composer,
    )
    plan = GenerationPlan(
        component="writer",
        operations=(GenerationOperation(op="writer.reset", args={}),),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    outcome = executor.execute(plan, ())
    assert Path(outcome.staged_artifact).parent == tmp_path


# -----------------------------------------------------------------------------
# Errors are surfaced as issues, not hidden
# -----------------------------------------------------------------------------

def test_operation_failure_is_recorded_as_execution_issue(
    executor, fake_composer, simple_plan
):
    def boom(**kwargs):
        raise RuntimeError("boom")

    fake_composer.add_paragraph = boom
    outcome = executor.execute(simple_plan, ())

    assert any(issue.code == "EXECUTION_FAILED" for issue in outcome.issues)



def test_deferred_ops_emit_stable_issues_and_fallbacks(executor, fake_composer):
    """M2-deferred ops emit stable issues and deterministic fallback text."""
    ops = [
        GenerationOperation(
            op="writer.add_captioned_figure",
            args={"caption": "Figure 1", "children": [], "layout": "stack"},
            node_id="fig:1",
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["IMAGE_INSERT_FAILED"],
                "fallback": "notice",
            },
        ),
        GenerationOperation(
            op="writer.add_semantic_table",
            args={"caption": "Table 1", "headers": [], "rows": [], "alignments": []},
            node_id="tab:1",
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["TABLE_INSERT_FAILED"],
                "fallback": "notice",
            },
        ),
        GenerationOperation(
            op="writer.add_equation",
            args={"source": "E = mc^2", "fallbackText": "E = mc^2"},
            node_id="eq:1",
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["EQUATION_INSERT_FAILED"],
                "fallback": "inline",
            },
        ),
        GenerationOperation(
            op="writer.add_bibliography",
            args={"entries": ["[1] Foo", "[2] Bar"]},
            node_id="bib:1",
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
                "fallback": "notice",
            },
        ),
    ]
    plan = GenerationPlan(
        component="writer",
        operations=tuple(ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    outcome = executor.execute(plan, ())

    codes = {issue.code for issue in outcome.issues}
    assert "IMAGE_INSERT_FAILED" in codes
    assert "TABLE_INSERT_FAILED" in codes
    assert "EQUATION_INSERT_FAILED" in codes
    assert "BIBLIOGRAPHY_INSERT_FAILED" in codes

    names = {call.name for call in fake_composer.primitives}
    assert "add_degradation_notice" in names
    assert "add_inline_degradation" in names


def test_index_placeholders_are_inserted(executor, fake_composer):
    ops = [
        GenerationOperation(
            op="writer.insert_figure_index",
            args={"title": "图目录"},
            node_id="doc:figure-index",
        ),
        GenerationOperation(
            op="writer.insert_table_index",
            args={"title": "表目录"},
            node_id="doc:table-index",
        ),
    ]
    plan = GenerationPlan(
        component="writer",
        operations=tuple(ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    executor.execute(plan, ())
    names = [call.name for call in fake_composer.primitives]
    assert "insert_figure_index" in names
    assert "insert_table_index" in names


def test_degrade_policy_runs_fallback_and_records_issue(
    executor, fake_composer, simple_plan
):
    def boom(**kwargs):
        raise RuntimeError("boom")

    fake_composer.add_paragraph = boom
    # Make add_paragraph fail with a recoverable code.
    ops = list(simple_plan.operations)
    new_ops = []
    for op in ops:
        if op.op == "writer.add_paragraph":
            new_ops.append(
                GenerationOperation(
                    op=op.op,
                    args=dict(op.args),
                    node_id=op.node_id,
                    failure_policy={
                        "mode": "degrade",
                        "recoverableCodes": ["EXECUTION_FAILED"],
                        "fallback": "notice",
                    },
                )
            )
        else:
            new_ops.append(op)
    plan = GenerationPlan(
        component="writer",
        operations=tuple(new_ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest=simple_plan.resource_manifest_digest,
    )
    outcome = executor.execute(plan, ())
    assert any(issue.code == "EXECUTION_FAILED" for issue in outcome.issues)
    assert any(call.name == "add_degradation_notice" for call in fake_composer.primitives)


def test_fail_policy_aborts_execution(
    executor, fake_composer, simple_plan
):
    def boom(**kwargs):
        raise RuntimeError("boom")

    fake_composer.add_paragraph = boom
    ops = list(simple_plan.operations)
    new_ops = []
    for op in ops:
        if op.op == "writer.add_paragraph":
            new_ops.append(
                GenerationOperation(
                    op=op.op,
                    args=dict(op.args),
                    node_id=op.node_id,
                    failure_policy={
                        "mode": "fail",
                        "recoverableCodes": [],
                        "fallback": "notice",
                    },
                )
            )
        else:
            new_ops.append(op)
    plan = GenerationPlan(
        component="writer",
        operations=tuple(new_ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest=simple_plan.resource_manifest_digest,
    )
    with pytest.raises(WindowsLongformExecutorError):
        executor.execute(plan, ())


def test_explicit_degradation_notices_dispatch(executor, fake_composer):
    ops = [
        GenerationOperation(
            op="writer.add_degradation_notice",
            args={
                "code": "TEST_NOTICE",
                "message": "message",
                "fallbackText": "fallback",
                "placement": "block",
            },
            node_id="n:1",
        ),
        GenerationOperation(
            op="writer.add_inline_degradation",
            args={
                "code": "TEST_INLINE",
                "message": "message",
                "fallbackText": "fallback",
            },
            node_id="n:2",
        ),
        GenerationOperation(
            op="writer.add_document_quality_notice",
            args={
                "notices": [
                    {
                        "code": "QUALITY",
                        "message": "msg",
                        "fallbackText": "fallback",
                        "placement": "document",
                    }
                ]
            },
            node_id="n:3",
        ),
    ]
    plan = GenerationPlan(
        component="writer",
        operations=tuple(ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    outcome = executor.execute(plan, ())
    names = {call.name for call in fake_composer.primitives}
    assert "add_degradation_notice" in names
    assert "add_inline_degradation" in names
    assert "add_document_quality_notice" in names
    assert len(outcome.issues) == 0



def test_ensure_styles_converts_list_of_mappings_to_dict(executor, fake_composer):
    """Regression: ensure_styles receives a list from the plan builder; the executor
    must convert it to the dict WriterComposer.ensure_styles expects."""
    ops = [
        GenerationOperation(
            op="writer.ensure_styles",
            args={
                "styles": [
                    {
                        "name": "Body Text",
                        "type": "paragraph",
                        "fontName": "SimSun",
                        "fontSize": 12,
                    },
                    {
                        "name": "Title",
                        "type": "paragraph",
                        "fontName": "SimHei",
                        "fontSize": 18,
                    },
                ]
            },
            node_id="doc:styles",
        ),
    ]
    plan = GenerationPlan(
        component="writer",
        operations=tuple(ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    executor.execute(plan, ())

    style_calls = [c for c in fake_composer.primitives if c.name == "ensure_styles"]
    assert len(style_calls) == 1
    styles = style_calls[0].args[0]
    assert isinstance(styles, dict)
    assert "Body Text" in styles
    assert "Title" in styles
    assert styles["Body Text"]["fontSize"] == 12
    assert styles["Title"]["fontName"] == "SimHei"


def test_configure_section_without_margins_passes_none(executor, fake_composer):
    """Regression: configure_section should not override configure_page margins with
    hardcoded defaults when the plan does not explicitly specify margins."""
    ops = [
        GenerationOperation(
            op="writer.configure_page",
            args={
                "marginTop": 100,
                "marginBottom": 100,
                "marginLeft": 120,
                "marginRight": 120,
            },
            node_id="doc:page",
        ),
        GenerationOperation(
            op="writer.configure_section",
            args={
                "role": "body",
                "pageNumberFormat": "arabic",
                "restartPageNumbering": True,
                "headerText": "Header",
            },
            node_id="doc:section-body",
        ),
    ]
    plan = GenerationPlan(
        component="writer",
        operations=tuple(ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    executor.execute(plan, ())

    section_calls = [c for c in fake_composer.primitives if c.name == "configure_section"]
    assert len(section_calls) == 1
    assert section_calls[0].kwargs["margins"] is None


def test_configure_section_with_margins_passes_them_through(executor, fake_composer):
    # Bypass plan validation because the real plan builder does not attach
    # margins to configure_section; this is a direct unit test of the dispatcher.
    op = GenerationOperation(
        op="writer.configure_section",
        args={
            "role": "body",
            "margins": {
                "top": 50,
                "bottom": 50,
                "left": 60,
                "right": 60,
            },
        },
        node_id="doc:section-body",
    )
    executor._dispatch_one(fake_composer, op)

    section_calls = [c for c in fake_composer.primitives if c.name == "configure_section"]
    assert len(section_calls) == 1
    margins = section_calls[0].kwargs["margins"]
    assert margins == {"top": 50, "bottom": 50, "left": 60, "right": 60}


def test_set_header_footer_uses_schema_keys(executor, fake_composer):
    ops = [
        GenerationOperation(
            op="writer.set_header_footer",
            args={
                "headerText": "Header from plan",
                "footerText": "Footer from plan",
                "linkToPreviousHeader": False,
                "linkToPreviousFooter": True,
            },
            node_id="doc:hf",
        ),
    ]
    plan = GenerationPlan(
        component="writer",
        operations=tuple(ops),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    executor.execute(plan, ())
    call = [c for c in fake_composer.primitives if c.name == "set_header_footer"][0]
    assert call.kwargs["header"] == "Header from plan"
    assert call.kwargs["footer"] == "Footer from plan"
    assert call.kwargs["link_to_previous_header"] is False
    assert call.kwargs["link_to_previous_footer"] is True
