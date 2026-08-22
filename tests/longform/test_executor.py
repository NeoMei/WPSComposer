"""Tests for the shared long-form executor interface and field-refresh convergence loop.

These tests are pure and do not start WPS or import platform/COM/JSAPI modules.
"""

from __future__ import annotations

import json
import subprocess
import sys
from dataclasses import replace
from pathlib import Path
from typing import Any, Optional, Tuple

import pytest

from skills.WPSComposer.scripts.document_model import DocumentIssue
from skills.WPSComposer.scripts.generation_plan import GenerationPlan
from skills.WPSComposer.scripts.longform.executor import (
    FIELD_REFRESH_UNSTABLE,
    ConvergenceResult,
    ExecutionIssue,
    ExecutionOutcome,
    FieldSnapshot,
    LongformExecutor,
    PaginationFragment,
    PaginationMap,
    PaginationNode,
    RecordingLongformExecutor,
    finalize_fields_with_convergence,
)


@pytest.fixture
def project_root() -> Path:
    return Path(__file__).parents[2]


# ---------------------------------------------------------------------------
# Import purity
# ---------------------------------------------------------------------------

def test_executor_import_is_pure_in_subprocess(project_root: Path):
    """Importing executor.py must not load WPS/executor/platform modules."""
    script = """
import sys
import skills.WPSComposer.scripts.longform.executor

forbidden = [
    "skills.WPSComposer.scripts.writer",
    "skills.WPSComposer.scripts.slide",
    "skills.WPSComposer.scripts.sheet",
    "skills.WPSComposer.scripts.wps_engine",
    "skills.WPSComposer.scripts.conversion",
    "skills.WPSComposer.scripts.macos_probe",
    "skills.WPSComposer.scripts._dispatch",
    "skills.WPSComposer.scripts.windows_writer_worker",
    "subprocess",
]
loaded = [m for m in sys.modules if any(m.startswith(f) for f in forbidden)]
if loaded:
    print(loaded)
    sys.exit(1)
print("pure")
"""
    env = {"PYTHONPATH": str(project_root)}
    result = subprocess.run(
        [sys.executable, "-c", script],
        capture_output=True,
        text=True,
        env=env,
        cwd=str(project_root),
    )
    assert result.returncode == 0, f"Import-time purity failed: {result.stdout} {result.stderr}"
    assert "pure" in result.stdout

def test_executor_module_does_not_import_platform_executors(project_root: Path):
    """The executor source must not directly import WPS/executor modules."""
    executor_path = project_root / "skills" / "WPSComposer" / "scripts" / "longform" / "executor.py"
    source = executor_path.read_text("utf-8")
    forbidden = (
        "from ..writer",
        "from ..slide",
        "from ..sheet",
        "from .._dispatch",
        "from ..conversion",
        "from ..macos_probe",
        "from ..windows_writer_worker",
        "import skills.WPSComposer.scripts.writer",
        "import skills.WPSComposer.scripts.slide",
        "import skills.WPSComposer.scripts.sheet",
        "import skills.WPSComposer.scripts._dispatch",
        "import skills.WPSComposer.scripts.conversion",
        "import skills.WPSComposer.scripts.macos_probe",
        "import skills.WPSComposer.scripts.windows_writer_worker",
        "import subprocess",
    )
    for fragment in forbidden:
        assert fragment not in source, f"executor.py imports forbidden module via {fragment!r}"


# ---------------------------------------------------------------------------
# Data-model round-trips
# ---------------------------------------------------------------------------

def test_execution_outcome_defaults_and_roundtrip():
    outcome = ExecutionOutcome(staged_artifact="test.docx")
    assert outcome.staged_artifact == "test.docx"
    assert outcome.issues == ()
    assert outcome.pagination_map.version == "M2-stub"

    issue = ExecutionIssue(
        code=FIELD_REFRESH_UNSTABLE,
        message="unstable",
        placement="document",
        node_id="doc:finalize",
    )
    outcome2 = ExecutionOutcome(
        staged_artifact="test.docx",
        issues=(issue,),
        pagination_map=PaginationMap(
            version="M2-stub",
            nodes=(
                PaginationNode(
                    node_id="sec:1",
                    page_start=1,
                    page_end=2,
                    fragments=(PaginationFragment(page=1), PaginationFragment(page=2)),
                ),
            ),
        ),
    )
    data = outcome2.to_dict()
    restored = ExecutionOutcome.from_dict(data)
    assert restored == outcome2

def test_pagination_map_stub_roundtrip():
    pmap = PaginationMap(
        version="M2-stub",
        nodes=(
            PaginationNode(
                node_id="doc:section-0:body",
                story="body",
                sections=("body",),
                page_start=1,
                page_end=1,
                fragments=(PaginationFragment(page=1, bounds=None),),
            ),
        ),
    )
    data = pmap.to_dict()
    json.dumps(data)  # serializable
    restored = PaginationMap.from_dict(data)
    assert restored == pmap


# ---------------------------------------------------------------------------
# Recording executor
# ---------------------------------------------------------------------------

def test_recording_executor_records_execute_calls():
    executor = RecordingLongformExecutor(artifact="staged.docx")
    plan = GenerationPlan(component="writer", operations=())
    result = executor.execute(plan, (), deadline=12345.0)

    assert len(executor.calls) == 1
    assert executor.calls[0][0] is plan
    assert executor.calls[0][1] == ()
    assert executor.calls[0][2] == 12345.0
    assert isinstance(result, ExecutionOutcome)
    assert result.staged_artifact == "staged.docx"

def test_recording_executor_returns_deterministic_stub_pagination_map():
    executor = RecordingLongformExecutor()
    plan = GenerationPlan(component="writer", operations=())
    first = executor.execute(plan, ())
    second = executor.execute(plan, ())
    assert first.pagination_map == second.pagination_map
    assert first.pagination_map.nodes[0].node_id == "doc:finalize"

def test_recording_executor_is_a_longform_executor():
    executor = RecordingLongformExecutor()
    assert isinstance(executor, LongformExecutor)


# ---------------------------------------------------------------------------
# Field-snapshot stable-key ordering
# ---------------------------------------------------------------------------

def test_field_snapshot_orders_by_stable_key():
    snap = (
        FieldSnapshot(
            stable_key=("b", "PAGE", 0),
            field_category="page",
            result_hash="bb",
            toc_page_count=0,
            figure_index_page_count=0,
            table_index_page_count=0,
            total_pages=1,
        ),
        FieldSnapshot(
            stable_key=("a", "PAGE", 0),
            field_category="page",
            result_hash="aa",
            toc_page_count=0,
            figure_index_page_count=0,
            table_index_page_count=0,
            total_pages=1,
        ),
    )
    sorted_snap = tuple(sorted(snap))
    assert sorted_snap[0].stable_key == ("a", "PAGE", 0)
    assert sorted_snap[1].stable_key == ("b", "PAGE", 0)

def test_field_snapshot_hashes_do_not_leak_into_issue_message():
    snapshots = [
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="deadbeef" * 8,
                toc_page_count=2,
                figure_index_page_count=1,
                table_index_page_count=1,
                total_pages=10,
            ),
        ),
        (
            FieldSnapshot(
                stable_key=("doc:toc", "TOC", 0),
                field_category="index",
                result_hash="cafebabe" * 8,
                toc_page_count=2,
                figure_index_page_count=1,
                table_index_page_count=1,
                total_pages=10,
            ),
        ),
    ]
    executor = RecordingLongformExecutor(snapshots=snapshots)
    result = finalize_fields_with_convergence(executor, max_rounds=1)
    issue = result.issues[0]
    assert "deadbeef" not in issue.message
    assert "cafebabe" not in issue.message
    assert "2" in issue.message
    assert "10" in issue.message


# ---------------------------------------------------------------------------
# Convergence loop
# ---------------------------------------------------------------------------

class _ConstantSnapshotExecutor:
    """Executor that returns the same field snapshot every round."""

    def __init__(self, snapshot: Tuple[FieldSnapshot, ...]) -> None:
        self.calls: list[int] = []
        self.snapshot = snapshot

    def execute(self, plan: GenerationPlan, resources: Tuple[Any, ...], deadline: Optional[float] = None) -> ExecutionOutcome:
        return ExecutionOutcome(staged_artifact="x.docx")

    def refresh_fields(self, round_index: int) -> Tuple[FieldSnapshot, ...]:
        self.calls.append(round_index)
        return self.snapshot

def _make_snapshot(node_id: str, hash_value: str) -> Tuple[FieldSnapshot, ...]:
    return (
        FieldSnapshot(
            stable_key=(node_id, "PAGE", 0),
            field_category="page",
            result_hash=hash_value,
            toc_page_count=1,
            figure_index_page_count=0,
            table_index_page_count=0,
            total_pages=5,
        ),
    )

def test_finalize_converges_on_second_round():
    snap = _make_snapshot("doc:body", "hash1")
    executor = _ConstantSnapshotExecutor(snap)
    result = finalize_fields_with_convergence(executor, max_rounds=3)

    assert isinstance(result, ConvergenceResult)
    assert result.snapshot == snap
    assert result.issues == ()
    assert result.rounds == 2
    assert executor.calls == [0, 1]

def test_finalize_emits_unstable_after_three_changing_rounds():
    snapshots = [
        _make_snapshot("doc:body", "hash1"),
        _make_snapshot("doc:body", "hash2"),
        _make_snapshot("doc:body", "hash3"),
        _make_snapshot("doc:body", "hash4"),
    ]

    class _ChangingExecutor:
        def __init__(self) -> None:
            self.calls: list[int] = []

        def execute(self, plan: GenerationPlan, resources: Tuple[Any, ...], deadline: Optional[float] = None) -> ExecutionOutcome:
            return ExecutionOutcome(staged_artifact="x.docx")

        def refresh_fields(self, round_index: int) -> Tuple[FieldSnapshot, ...]:
            self.calls.append(round_index)
            return snapshots[round_index]

    executor = _ChangingExecutor()
    result = finalize_fields_with_convergence(executor, max_rounds=3)

    assert result.snapshot == snapshots[3]
    assert len(result.issues) == 1
    issue = result.issues[0]
    assert issue.code == FIELD_REFRESH_UNSTABLE
    assert issue.placement == "document"
    assert result.rounds == 4
    assert executor.calls == [0, 1, 2, 3]

def test_finalize_deterministic_ordering_of_stable_keys():
    """Snapshots with the same fields in different order compare equal after sorting."""
    a = FieldSnapshot(
        stable_key=("a", "PAGE", 0),
        field_category="page",
        result_hash="h",
        toc_page_count=0,
        figure_index_page_count=0,
        table_index_page_count=0,
        total_pages=1,
    )
    b = FieldSnapshot(
        stable_key=("b", "PAGE", 0),
        field_category="page",
        result_hash="h",
        toc_page_count=0,
        figure_index_page_count=0,
        table_index_page_count=0,
        total_pages=1,
    )

    class _JumbledExecutor:
        def refresh_fields(self, round_index: int) -> Tuple[FieldSnapshot, ...]:
            if round_index % 2 == 0:
                return (a, b)
            return (b, a)

        def execute(self, plan: GenerationPlan, resources: Tuple[Any, ...], deadline: Optional[float] = None) -> ExecutionOutcome:
            return ExecutionOutcome(staged_artifact="x.docx")

    executor = _JumbledExecutor()
    result = finalize_fields_with_convergence(executor, max_rounds=3)
    assert result.issues == ()
    assert result.rounds == 2


# ---------------------------------------------------------------------------
# Robustness / bad input
# ---------------------------------------------------------------------------

def test_finalize_never_raises_on_bad_executor():
    """finalize_fields_with_convergence must not raise for a non-conforming executor."""
    class _NoRefresh:
        def execute(self, plan: GenerationPlan, resources: Tuple[Any, ...], deadline: Optional[float] = None) -> ExecutionOutcome:
            return ExecutionOutcome(staged_artifact="x.docx")

    executor = _NoRefresh()
    result = finalize_fields_with_convergence(executor, max_rounds=3)
    assert len(result.issues) == 1
    assert result.issues[0].code == "EXECUTOR_CAPABILITY_MISSING"
    assert result.snapshot == ()

def test_recording_executor_can_inject_snapshots_for_convergence():
    snapshots = [
        _make_snapshot("doc:body", "h1"),
        _make_snapshot("doc:body", "h2"),
        _make_snapshot("doc:body", "h3"),
        _make_snapshot("doc:body", "h4"),
    ]
    executor = RecordingLongformExecutor(snapshots=snapshots)
    result = finalize_fields_with_convergence(executor, max_rounds=3)
    assert result.issues[0].code == FIELD_REFRESH_UNSTABLE
    assert result.snapshot == snapshots[3]
    assert len(executor.refresh_calls) == 4


# ---------------------------------------------------------------------------
# Protocol violations
# ---------------------------------------------------------------------------

def test_class_without_execute_is_not_longform_executor():
    class _MissingExecute:
        pass

    assert not isinstance(_MissingExecute(), LongformExecutor)

def test_class_with_wrong_execute_signature_matches_protocol_loosely():
    """Runtime_checkable Protocol only checks method existence, not full signature."""
    class _BadExecute:
        def execute(self, x):
            return None

    assert isinstance(_BadExecute(), LongformExecutor)
