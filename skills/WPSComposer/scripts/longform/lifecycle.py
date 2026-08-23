"""Bounded M5 generation, PDF-quality, relayout, and publication lifecycle."""

from __future__ import annotations

from pathlib import Path
import time
from typing import Any, Callable, Optional, Protocol, Sequence

from .executor import ExecutionOutcome
from .pdf_quality import require_quality_dependencies
from .quality import GenerationOutcome, QualityFinding, QualityReport
from .relayout import RelayoutDirective, build_relayout_directives


class LongformLifecycleError(RuntimeError):
    """Fatal lifecycle failure with one stable, non-private error code."""

    def __init__(self, code: str, message: str) -> None:
        self.code = code
        super().__init__(message)


class LongformLifecycleAdapter(Protocol):
    def execute(
        self,
        build: Any,
        directives: Sequence[RelayoutDirective],
        deadline: float,
    ) -> ExecutionOutcome:
        ...

    def export_pdf(self, docx: Path, deadline: float) -> Path:
        ...

    def patch_quality_notices(
        self,
        docx: Path,
        notices: Sequence[QualityFinding],
        deadline: float,
    ) -> ExecutionOutcome:
        ...

    def validate_patch(
        self,
        docx: Path,
        pdf: Path,
        notices: Sequence[QualityFinding],
        outcome: ExecutionOutcome,
    ) -> None:
        ...

    def publish(
        self,
        staged: Path,
        output: Path,
        overwrite: bool,
        deadline: float,
    ) -> Path:
        ...

    def cleanup(self, path: Path) -> None:
        ...


def _check_deadline(deadline: float, clock: Callable[[], float]) -> None:
    if clock() >= deadline:
        raise LongformLifecycleError(
            "GENERATION_TIMEOUT", "Long-form generation deadline expired"
        )


def _call_stage(
    code: str,
    message: str,
    operation: Callable[[], Any],
) -> Any:
    try:
        return operation()
    except (LongformLifecycleError, FileExistsError):
        raise
    except Exception:
        raise LongformLifecycleError(code, message) from None


def _as_staged_path(outcome: ExecutionOutcome) -> Path:
    path = Path(outcome.staged_artifact).expanduser().resolve()
    if path.suffix.lower() != ".docx":
        raise LongformLifecycleError(
            "FINAL_ARTIFACT_INVALID", "Native generation did not return a DOCX staging artifact"
        )
    return path


def run_longform_lifecycle(
    build: Any,
    adapter: LongformLifecycleAdapter,
    output: Path,
    *,
    format_name: str,
    timeout: float,
    overwrite: bool,
    quality_analyzer: Callable[[Path, ExecutionOutcome, Any], QualityReport],
    dependency_gate: Callable[[], Any] = require_quality_dependencies,
    clock: Callable[[], float] = time.monotonic,
) -> GenerationOutcome:
    """Run the long-form lifecycle within the closed M5 mutation/export caps."""

    normalized_format = str(format_name).lower().lstrip(".")
    if normalized_format not in {"docx", "pdf"}:
        raise ValueError("long-form lifecycle format must be docx or pdf")
    if not isinstance(timeout, (int, float)) or timeout <= 0:
        raise ValueError("timeout must be positive")
    output_path = Path(output).expanduser().resolve()
    if output_path.suffix.lower() != f".{normalized_format}":
        raise ValueError("output extension must match the requested long-form format")
    if output_path.exists() and not overwrite:
        raise FileExistsError(str(output_path))

    try:
        dependency_gate()
    except Exception:
        raise LongformLifecycleError(
            "QUALITY_DEPENDENCY_MISSING",
            "Required long-form PDF analysis components are unavailable",
        ) from None

    deadline = clock() + float(timeout)
    staged_paths: list[Path] = []
    published: Optional[Path] = None
    generation_count = 0
    export_count = 0
    patch_count = 0
    current_outcome: Optional[ExecutionOutcome] = None
    current_docx: Optional[Path] = None
    current_pdf: Optional[Path] = None
    final_report = QualityReport()

    def execute(directives: Sequence[RelayoutDirective]) -> ExecutionOutcome:
        nonlocal generation_count, current_docx
        if generation_count >= 2:
            raise LongformLifecycleError(
                "LIFECYCLE_BOUND_EXCEEDED", "Full generation pass limit exceeded"
            )
        _check_deadline(deadline, clock)
        outcome = _call_stage(
            "NATIVE_GENERATION_FAILED",
            "Native WPS long-form generation failed",
            lambda: adapter.execute(build, directives, deadline),
        )
        generation_count += 1
        current_docx = _as_staged_path(outcome)
        staged_paths.append(current_docx)
        return outcome

    def export(docx: Path) -> Path:
        nonlocal export_count
        if export_count >= 3:
            raise LongformLifecycleError(
                "LIFECYCLE_BOUND_EXCEEDED", "PDF export pass limit exceeded"
            )
        _check_deadline(deadline, clock)
        pdf = Path(
            _call_stage(
                "PDF_EXPORT_FAILED",
                "WPS PDF export failed",
                lambda: adapter.export_pdf(docx, deadline),
            )
        ).expanduser().resolve()
        export_count += 1
        staged_paths.append(pdf)
        return pdf

    def analyze(pdf: Path, outcome: ExecutionOutcome) -> QualityReport:
        _check_deadline(deadline, clock)
        report = _call_stage(
            "PDF_ANALYSIS_FAILED",
            "Long-form PDF quality analysis failed",
            lambda: quality_analyzer(pdf, outcome, build),
        )
        if not isinstance(report, QualityReport):
            raise LongformLifecycleError(
                "PDF_ANALYSIS_FAILED", "Long-form PDF analyzer returned an invalid result"
            )
        return report

    try:
        current_outcome = execute(())
        if current_docx is None:  # defensive type narrowing
            raise LongformLifecycleError("FINAL_ARTIFACT_INVALID", "DOCX staging artifact is unavailable")
        current_pdf = export(current_docx)
        final_report = analyze(current_pdf, current_outcome)

        decision = build_relayout_directives(final_report)
        if decision.requires_relayout:
            current_outcome = execute(decision.directives)
            if current_docx is None:
                raise LongformLifecycleError("FINAL_ARTIFACT_INVALID", "DOCX staging artifact is unavailable")
            current_pdf = export(current_docx)
            final_report = analyze(current_pdf, current_outcome)
            decision = build_relayout_directives(
                final_report, relayout_already_used=True
            )

        notices = decision.notice_findings
        if notices:
            if patch_count >= 1:
                raise LongformLifecycleError(
                    "LIFECYCLE_BOUND_EXCEEDED", "Notice patch limit exceeded"
                )
            if current_docx is None:
                raise LongformLifecycleError("FINAL_ARTIFACT_INVALID", "DOCX staging artifact is unavailable")
            _check_deadline(deadline, clock)
            current_outcome = _call_stage(
                "NOTICE_PATCH_FAILED",
                "WPS quality-notice patch failed",
                lambda: adapter.patch_quality_notices(
                    current_docx, notices, deadline
                ),
            )
            patch_count += 1
            patched_docx = _as_staged_path(current_outcome)
            if patched_docx != current_docx:
                current_docx = patched_docx
                staged_paths.append(current_docx)
            current_pdf = export(current_docx)
            _check_deadline(deadline, clock)
            _call_stage(
                "FINAL_ARTIFACT_INVALID",
                "Notice-patched artifact validation failed",
                lambda: adapter.validate_patch(
                    current_docx, current_pdf, notices, current_outcome
                ),
            )

        _check_deadline(deadline, clock)
        staged_public = current_pdf if normalized_format == "pdf" else current_docx
        if staged_public is None:
            raise LongformLifecycleError(
                "FINAL_ARTIFACT_INVALID", "Requested staging artifact is unavailable"
            )
        published = Path(
            _call_stage(
                "ARTIFACT_PUBLISH_FAILED",
                "Atomic artifact publication failed",
                lambda: adapter.publish(
                    staged_public, output_path, overwrite, deadline
                ),
            )
        ).expanduser().resolve()
        return GenerationOutcome(path=str(published), issues=final_report.findings)
    finally:
        # The public artifact is never one of the private staging files in the
        # production adapter.  The equality guard also keeps injected adapters safe.
        seen: set[Path] = set()
        for staged in reversed(staged_paths):
            if staged in seen or (published is not None and staged == published):
                continue
            seen.add(staged)
            try:
                adapter.cleanup(staged)
            except Exception:
                # Cleanup errors cannot safely replace an active fatal error.
                # Production adapters perform strict cleanup before returning.
                pass


__all__ = [
    "LongformLifecycleAdapter",
    "LongformLifecycleError",
    "run_longform_lifecycle",
]
