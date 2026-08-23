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


@pytest.mark.skipif(
    os.environ.get("WPSCOMPOSER_RUN_WINDOWS_M5") != "1",
    reason="set WPSCOMPOSER_RUN_WINDOWS_M5=1 for native Windows WPS acceptance",
)
def test_real_windows_m5_six_fixtures_and_performance(tmp_path: Path):
    assert platform.system() == "Windows"
    report_path = run_longform_m5_evidence(tmp_path / "evidence", timeout=300)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    validate_m5_evidence_report(report)
    assert report["environment"]["system"] == "Windows"
    assert len(report["fixtures"]) == 6
    assert 50 <= report["performance"]["pageCount"] <= 100
