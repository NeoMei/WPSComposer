from __future__ import annotations

import copy

import pytest

from skills.WPSComposer.scripts.longform_m5_evidence import (
    validate_m5_evidence_report,
)


def _entry(name: str, pages: int = 1) -> dict:
    return {
        "name": name,
        "pageCount": pages,
        "operationCount": 1,
        "counts": {"generation": 1, "export": 1, "patch": 0, "analysis": 1},
        "stageSeconds": {
            "generation": 1.0,
            "export": 1.0,
            "patch": 0.0,
            "analysis": 1.0,
            "total": 3.0,
        },
        "issueCodes": [],
        "artifact": {"name": f"{name}.pdf", "sha256": "a" * 64},
        "screenshots": [f"screenshots/{name}-1.png"],
    }


def _report(system: str) -> dict:
    return {
        "version": "M5",
        "environment": {
            "system": system,
            "machine": "AMD64" if system == "Windows" else "arm64",
            "wpsVersion": "12.1.0",
            "protocolVersion": 2,
            "semanticVersion": "longform-1",
        },
        "fixtures": [_entry(name) for name in (
            "academic", "toc_dense", "wide_objects", "degradation",
            "unicode", "plain_short",
        )],
        "performance": _entry("performance-60", 60),
    }


@pytest.mark.parametrize("system", ("Darwin", "Windows"))
def test_m5_evidence_contract_accepts_both_native_release_platforms(system):
    validate_m5_evidence_report(_report(system))


def test_m5_evidence_contract_rejects_unsupported_platform_and_missing_screenshots():
    with pytest.raises(ValueError, match="platform contract"):
        validate_m5_evidence_report(_report("Linux"))
    report = copy.deepcopy(_report("Windows"))
    report["fixtures"][0]["screenshots"] = []
    with pytest.raises(ValueError, match="screenshot evidence"):
        validate_m5_evidence_report(report)
