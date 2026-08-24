from __future__ import annotations

import json
import os
from pathlib import Path
import platform

import pytest

from skills.WPSComposer.scripts.longform_m5_evidence import (
    run_longform_m5_evidence,
    validate_m5_evidence_report,
)


ROOT = Path(__file__).resolve().parents[2]


@pytest.mark.skipif(
    os.environ.get("WPSCOMPOSER_RUN_REAL_WPS") != "1",
    reason="set WPSCOMPOSER_RUN_REAL_WPS=1 for native macOS WPS acceptance",
)
def test_real_macos_m5_six_fixtures_and_performance(tmp_path):
    assert platform.system() == "Darwin"
    report_path = run_longform_m5_evidence(tmp_path / "evidence", timeout=300)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    validate_m5_evidence_report(report)
    assert len(report["fixtures"]) == 6
    assert 50 <= report["performance"]["pageCount"] <= 100
    for entry in report["fixtures"] + [report["performance"]]:
        assert (report_path.parent / entry["artifact"]["name"]).is_file()
        assert entry["screenshots"]
