from __future__ import annotations

import hashlib
import json
import os
import platform
from datetime import datetime
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform_m4_evidence import (
    run_longform_m4_evidence,
    validate_m4_evidence_report,
)


ROOT = Path(__file__).resolve().parents[2]


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


@pytest.fixture(scope="module")
def real_m4_evidence() -> tuple[Path, dict]:
    if platform.system() != "Darwin":
        pytest.skip("M4 acceptance requires a real macOS WPS run")
    if os.environ.get("WPSCOMPOSER_RUN_REAL_WPS") != "1":
        pytest.fail("set WPSCOMPOSER_RUN_REAL_WPS=1 for native macOS WPS acceptance")
    configured = os.environ.get("WPSCOMPOSER_M4_EVIDENCE_DIR")
    output = Path(configured) if configured else ROOT / "build" / "longform-m4" / (
        "macos-native-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    )
    report_path = run_longform_m4_evidence(output, timeout=300.0)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    validate_m4_evidence_report(report)
    return report_path.parent, report


def test_real_macos_wps_m4_create_reopen_refresh_and_move(
    real_m4_evidence: tuple[Path, dict],
) -> None:
    _, report = real_m4_evidence
    initial = report["metrics"]["initial"]
    moved = report["metrics"]["moved"]

    native = initial["formulaCount"] == 12
    assert moved["formulaCount"] == initial["formulaCount"]
    if native:
        assert initial["editableFormulaCount"] == moved["editableFormulaCount"] == 12
        assert initial["degradedFormulaCount"] == moved["degradedFormulaCount"] == 0
    else:
        assert initial["formulaCount"] == moved["formulaCount"] == 0
        assert initial["editableFormulaCount"] == moved["editableFormulaCount"] == 0
        assert initial["degradedFormulaCount"] == moved["degradedFormulaCount"] == 12
        assert initial["formulaNoticeCount"] == moved["formulaNoticeCount"] == 12
        assert initial["formulaFallbackImageCount"] == moved["formulaFallbackImageCount"] == 1
        assert "EQUATION_INSERT_FAILED" in report["codes"]
    assert initial["centeredFormulaCount"] == initial["numberAlignedFormulaCount"] == 12
    assert moved["centeredFormulaCount"] == moved["numberAlignedFormulaCount"] == 12
    assert initial["equationSequenceCount"] == moved["equationSequenceCount"] == 12
    assert initial["referenceCount"] >= 2
    assert initial["citationCount"] == 3
    assert initial["bibliographyCount"] == initial["hangingBibliographyCount"] == 3
    assert report["rounds"]
    assert all(1 <= rounds <= 4 for rounds in report["rounds"])


def test_real_macos_wps_m4_recoverable_and_fatal_boundaries(
    real_m4_evidence: tuple[Path, dict],
) -> None:
    output, report = real_m4_evidence
    degradation = report["metrics"]["degradation"]
    assert degradation["inlineNoticeCount"] >= 1
    assert degradation["blockNoticeCount"] >= 3
    assert degradation["documentNoticeCount"] >= 1
    assert "ENGINE_LOST" in report["codes"]
    assert not (output / "m4-engine-fatal.docx").exists()


def test_real_macos_wps_m4_evidence_hashes_and_visual_artifacts(
    real_m4_evidence: tuple[Path, dict],
) -> None:
    output, report = real_m4_evidence
    assert report["screenshots"]
    assert len([item for item in report["artifacts"] if item["name"].endswith(".pdf")]) == 2
    for artifact in report["artifacts"]:
        path = output / artifact["name"]
        assert path.is_file()
        assert _sha256(path) == artifact["sha256"]
    for screenshot in report["screenshots"]:
        assert (output / screenshot).is_file()
