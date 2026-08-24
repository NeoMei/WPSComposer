from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil

import pytest

from skills.WPSComposer.scripts.longform.executor import (
    ExecutionIssue,
    ExecutionOutcome,
    PaginationFragment,
    PaginationMap,
    PaginationNode,
)
from skills.WPSComposer.scripts.longform.lifecycle import (
    LongformLifecycleError,
    run_longform_lifecycle,
)
from skills.WPSComposer.scripts.longform.quality import (
    QualityConfidence,
    QualityFinding,
    QualityReport,
    QualitySeverity,
)
from tests._pdf_fixture import write_minimal_pdf


@dataclass
class _Build:
    plan: object = object()


def _finding(repair_key=None, *, severity="degraded", confidence="high"):
    return QualityFinding(
        code="QUALITY_ISSUE",
        severity=QualitySeverity(severity),
        confidence=QualityConfidence(confidence),
        message="Controlled quality issue",
        node_id="fig:one",
        page=1,
        repair_key=repair_key,
    )


def _outcome(path: Path):
    return ExecutionOutcome(
        staged_artifact=str(path),
        pagination_map=PaginationMap(
            version="M5-v1",
            nodes=(
                PaginationNode(
                    node_id="fig:one",
                    page_start=1,
                    page_end=1,
                    fragments=(PaginationFragment(page=1, bounds=(72, 72, 300, 300)),),
                ),
            ),
        ),
    )


class _Adapter:
    def __init__(self, root: Path):
        self.root = root
        self.execute_calls = []
        self.export_calls = []
        self.patch_calls = []
        self.validate_patch_calls = []
        self.publish_calls = []
        self.cleanup_calls = []
        self.fail_at = None

    def execute(self, build, directives, deadline):
        self.execute_calls.append((tuple(directives), deadline))
        if self.fail_at == "execute":
            raise RuntimeError("native execute failed")
        path = self.root / f"pass-{len(self.execute_calls)}.docx"
        path.write_bytes(b"PK\x03\x04" + b"x" * 300)
        return _outcome(path)

    def export_pdf(self, docx, deadline):
        self.export_calls.append((str(docx), deadline))
        if self.fail_at == "export":
            raise RuntimeError("native export failed")
        path = self.root / f"export-{len(self.export_calls)}.pdf"
        return write_minimal_pdf(path)

    def patch_quality_notices(self, docx, notices, deadline):
        self.patch_calls.append((str(docx), tuple(notices), deadline))
        if self.fail_at == "patch":
            raise RuntimeError("patch failed")
        Path(docx).write_bytes(Path(docx).read_bytes() + b"NOTICE")
        return _outcome(Path(docx))

    def validate_patch(self, docx, pdf, notices, outcome):
        self.validate_patch_calls.append((str(docx), str(pdf), tuple(notices), outcome))
        if self.fail_at == "validate-patch":
            raise RuntimeError("patch validation failed")

    def publish(self, staged, output, overwrite, deadline):
        self.publish_calls.append((str(staged), str(output), overwrite, deadline))
        if self.fail_at == "publish":
            raise RuntimeError("publish failed")
        output = Path(output)
        if output.exists() and not overwrite:
            raise FileExistsError(str(output))
        shutil.copy2(staged, output)
        return output.resolve()

    def cleanup(self, path):
        self.cleanup_calls.append(str(path))
        Path(path).unlink(missing_ok=True)


class _Analyzer:
    def __init__(self, *reports):
        self.reports = list(reports)
        self.calls = []

    def __call__(self, pdf, outcome, build):
        self.calls.append((str(pdf), outcome, build))
        return self.reports.pop(0)


def _run(tmp_path, adapter, analyzer, *, format_name="docx", timeout=600, gate=lambda: None):
    return run_longform_lifecycle(
        _Build(),
        adapter,
        tmp_path / f"final.{format_name}",
        format_name=format_name,
        timeout=timeout,
        overwrite=False,
        quality_analyzer=analyzer,
        dependency_gate=gate,
    )


def test_clean_first_pass_publishes_requested_docx_and_removes_internal_pdf(tmp_path):
    adapter = _Adapter(tmp_path)
    analyzer = _Analyzer(QualityReport((), 1))
    result = _run(tmp_path, adapter, analyzer)
    assert Path(result.path).is_file()
    assert result.path.endswith("final.docx")
    assert result.degraded is False
    assert len(adapter.execute_calls) == 1
    assert len(adapter.export_calls) == 1
    assert adapter.patch_calls == []
    assert not (tmp_path / "export-1.pdf").exists()


def test_requested_pdf_reuses_final_quality_pdf_without_extra_export(tmp_path):
    adapter = _Adapter(tmp_path)
    result = _run(
        tmp_path, adapter, _Analyzer(QualityReport((), 1)), format_name="pdf"
    )
    assert Path(result.path).read_bytes().startswith(b"%PDF-")
    assert len(adapter.export_calls) == 1
    assert adapter.publish_calls[0][0].endswith("export-1.pdf")


def test_high_confidence_closed_issue_causes_exactly_one_relayout(tmp_path):
    adapter = _Adapter(tmp_path)
    analyzer = _Analyzer(
        QualityReport((_finding("fit-image"),), 1),
        QualityReport((), 1),
    )
    _run(tmp_path, adapter, analyzer)
    assert len(adapter.execute_calls) == 2
    assert adapter.execute_calls[0][0] == ()
    assert [item.kind for item in adapter.execute_calls[1][0]] == ["fit-image"]
    assert len(adapter.export_calls) == 2
    assert adapter.patch_calls == []


def test_final_degraded_issue_causes_one_notice_patch_and_third_export(tmp_path):
    finding = _finding("fit-image")
    adapter = _Adapter(tmp_path)
    analyzer = _Analyzer(QualityReport((finding,), 1), QualityReport((finding,), 1))
    result = _run(tmp_path, adapter, analyzer)
    assert len(adapter.execute_calls) == 2
    assert len(adapter.patch_calls) == 1
    assert len(adapter.export_calls) == 3
    assert len(adapter.validate_patch_calls) == 1
    assert result.degraded is True


def test_notice_only_issue_skips_second_generation(tmp_path):
    finding = _finding("low-dpi-notice")
    adapter = _Adapter(tmp_path)
    _run(tmp_path, adapter, _Analyzer(QualityReport((finding,), 1)))
    assert len(adapter.execute_calls) == 1
    assert len(adapter.patch_calls) == 1
    assert len(adapter.export_calls) == 2


def test_notice_patch_refreshes_final_issue_page_from_fresh_pagination(tmp_path):
    finding = _finding("low-dpi-notice")
    adapter = _Adapter(tmp_path)

    def patch_on_page_three(docx, notices, deadline):
        adapter.patch_calls.append((str(docx), tuple(notices), deadline))
        return ExecutionOutcome(
            staged_artifact=str(docx),
            pagination_map=PaginationMap(
                version="M5-v1",
                nodes=(
                    PaginationNode(
                        node_id="fig:one",
                        page_start=3,
                        page_end=3,
                        fragments=(PaginationFragment(page=3),),
                    ),
                ),
            ),
        )

    adapter.patch_quality_notices = patch_on_page_three
    result = _run(tmp_path, adapter, _Analyzer(QualityReport((finding,), 1)))
    assert result.issues[0].page == 3


def test_native_execution_issues_are_preserved_in_private_outcome(tmp_path):
    adapter = _Adapter(tmp_path)
    original_execute = adapter.execute

    def execute_with_degradation(build, directives, deadline):
        outcome = original_execute(build, directives, deadline)
        return ExecutionOutcome(
            staged_artifact=outcome.staged_artifact,
            pagination_map=outcome.pagination_map,
            issues=(
                ExecutionIssue(
                    code="IMAGE_INSERT_FAILED",
                    message="Image insertion degraded to a visible notice",
                    placement="block",
                    node_id="fig:one",
                    fallback="notice",
                    recoverable=True,
                ),
            ),
        )

    adapter.execute = execute_with_degradation
    result = _run(tmp_path, adapter, _Analyzer(QualityReport((), 1)))
    assert [issue.code for issue in result.issues] == ["IMAGE_INSERT_FAILED"]
    assert result.degraded is True


@pytest.mark.parametrize(("severity", "confidence"), [("warning", "high"), ("info", "low"), ("degraded", "medium")])
def test_warning_info_and_non_high_findings_never_mutate(tmp_path, severity, confidence):
    finding = _finding("fit-image", severity=severity, confidence=confidence)
    adapter = _Adapter(tmp_path)
    result = _run(tmp_path, adapter, _Analyzer(QualityReport((finding,), 1)))
    assert len(adapter.execute_calls) == 1
    assert adapter.patch_calls == []
    assert result.degraded is (severity == "degraded")


def test_dependency_gate_runs_before_any_native_work(tmp_path):
    adapter = _Adapter(tmp_path)

    def fail_gate():
        raise RuntimeError("missing analyzer")

    with pytest.raises(LongformLifecycleError) as exc:
        _run(tmp_path, adapter, _Analyzer(), gate=fail_gate)
    assert exc.value.code == "QUALITY_DEPENDENCY_MISSING"
    assert adapter.execute_calls == []
    assert not (tmp_path / "final.docx").exists()


@pytest.mark.parametrize(
    ("stage", "code"),
    [
        ("execute", "NATIVE_GENERATION_FAILED"),
        ("export", "PDF_EXPORT_FAILED"),
        ("patch", "NOTICE_PATCH_FAILED"),
        ("validate-patch", "FINAL_ARTIFACT_INVALID"),
        ("publish", "ARTIFACT_PUBLISH_FAILED"),
    ],
)
def test_fatal_stage_failure_never_publishes_partial_artifact(tmp_path, stage, code):
    adapter = _Adapter(tmp_path)
    adapter.fail_at = stage
    report = QualityReport((_finding("low-dpi-notice"),), 1)
    with pytest.raises(LongformLifecycleError) as exc:
        _run(tmp_path, adapter, _Analyzer(report))
    assert exc.value.code == code
    assert not (tmp_path / "final.docx").exists()
    assert not list(tmp_path.glob("pass-*.docx"))
    assert not list(tmp_path.glob("export-*.pdf"))


def test_timeout_uses_one_absolute_deadline_and_cleans_staging(tmp_path):
    class Clock:
        value = 100.0

        def __call__(self):
            return self.value

    clock = Clock()
    adapter = _Adapter(tmp_path)
    original_export = adapter.export_pdf

    def export_and_expire(docx, deadline):
        result = original_export(docx, deadline)
        clock.value = 111.0
        return result

    adapter.export_pdf = export_and_expire
    with pytest.raises(LongformLifecycleError) as exc:
        run_longform_lifecycle(
            _Build(),
            adapter,
            tmp_path / "final.docx",
            format_name="docx",
            timeout=10,
            overwrite=False,
            quality_analyzer=_Analyzer(QualityReport((), 1)),
            dependency_gate=lambda: None,
            clock=clock,
        )
    assert exc.value.code == "GENERATION_TIMEOUT"
    assert not list(tmp_path.glob("pass-*.docx"))
    assert not list(tmp_path.glob("export-*.pdf"))


def test_invalid_format_is_rejected_before_dependency_or_native_work(tmp_path):
    adapter = _Adapter(tmp_path)
    called = []
    with pytest.raises(ValueError, match="docx or pdf"):
        run_longform_lifecycle(
            _Build(),
            adapter,
            tmp_path / "final.pptx",
            format_name="pptx",
            timeout=10,
            overwrite=False,
            quality_analyzer=_Analyzer(),
            dependency_gate=lambda: called.append(True),
        )
    assert called == []
    assert adapter.execute_calls == []
