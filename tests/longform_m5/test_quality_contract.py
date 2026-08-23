from __future__ import annotations

import json
import math

import pytest

from skills.WPSComposer.scripts.longform.executor import (
    PaginationFragment,
    PaginationMap,
    PaginationNode,
)
from skills.WPSComposer.scripts.longform.quality import (
    GenerationOutcome,
    PageRole,
    QualityConfidence,
    QualityEvidence,
    QualityFinding,
    QualityReport,
    QualitySeverity,
)


def _finding(**overrides):
    values = {
        "code": "OBJECT_OVERFLOW",
        "severity": QualitySeverity.DEGRADED,
        "confidence": QualityConfidence.HIGH,
        "message": "Mapped object crosses the body boundary",
        "node_id": "fig:alpha",
        "page": 3,
        "bounds": (72.0, 90.0, 540.0, 700.0),
        "evidence": QualityEvidence.from_mapping(
            {"overflowPoints": 12.5, "pageRole": "body"}
        ),
        "repair_key": "fit-image",
    }
    values.update(overrides)
    return QualityFinding(**values)


def test_closed_quality_enums_and_page_roles():
    assert {item.value for item in QualitySeverity} == {
        "fatal", "degraded", "warning", "info"
    }
    assert {item.value for item in QualityConfidence} == {
        "high", "medium", "low"
    }
    assert PageRole.BIBLIOGRAPHY.value == "bibliography"
    with pytest.raises(ValueError):
        QualitySeverity("critical")


def test_finding_and_report_roundtrip_are_deterministic():
    first = _finding(code="Z_LAST", node_id="fig:z", page=4)
    second = _finding(code="A_FIRST", node_id="fig:a", page=2)
    report = QualityReport(findings=(first, second), page_count=6)

    payload = report.to_dict()
    assert [item["code"] for item in payload["findings"]] == [
        "A_FIRST", "Z_LAST"
    ]
    assert QualityReport.from_dict(json.loads(json.dumps(payload))) == report
    assert report.degraded is True


@pytest.mark.parametrize(
    "bounds",
    [
        (0, 0, 0, 1),
        (0, 0, 1, 0),
        (0, 0, math.inf, 1),
        (0, 0, math.nan, 1),
        (4, 0, 3, 1),
    ],
)
def test_finding_rejects_invalid_bounds(bounds):
    with pytest.raises(ValueError):
        _finding(bounds=bounds)


def test_finding_rejects_non_positive_page_and_uncontrolled_tokens():
    with pytest.raises(ValueError):
        _finding(page=0)
    with pytest.raises(ValueError):
        _finding(code="bad code")
    with pytest.raises(ValueError):
        _finding(repair_key="../../escape")


@pytest.mark.parametrize(
    "evidence",
    [
        {"sourcePath": "/Users/private/input.docx"},
        {"detail": r"C:\\Users\\private\\input.docx"},
        {"hash": "sha256:abcdef"},
        {"bodyText": "customer paragraph"},
        {"unknownMetric": 1},
    ],
)
def test_evidence_rejects_private_or_uncontrolled_values(evidence):
    with pytest.raises(ValueError):
        QualityEvidence.from_mapping(evidence)


def test_generation_outcome_aggregates_degraded_state_without_exposing_details():
    warning = _finding(
        code="LAST_PAGE_SPARSE",
        severity=QualitySeverity.WARNING,
        confidence=QualityConfidence.MEDIUM,
        repair_key=None,
    )
    degraded = _finding()
    outcome = GenerationOutcome(
        path="/tmp/report.docx", issues=(warning, degraded)
    )
    assert outcome.degraded is True
    assert outcome.to_dict() == {
        "path": "/tmp/report.docx",
        "degraded": True,
        "issues": [warning.to_dict(), degraded.to_dict()],
    }


def test_m5_pagination_bounds_roundtrip_and_legacy_dict_input():
    fragment = PaginationFragment(page=2, bounds=(1.0, 2.0, 30.0, 40.0))
    pmap = PaginationMap(
        version="M5-v1",
        nodes=(
            PaginationNode(
                node_id="fig:alpha",
                story="main",
                sections=("body",),
                page_start=2,
                page_end=2,
                range="10:20",
                fragments=(fragment,),
            ),
        ),
    )
    data = pmap.to_dict()
    assert data["nodes"][0]["fragments"][0]["bounds"] == [1.0, 2.0, 30.0, 40.0]
    assert PaginationMap.from_dict(data) == pmap

    legacy = PaginationFragment.from_dict(
        {"page": 1, "bounds": {"x0": 1, "y0": 2, "x1": 3, "y1": 4}}
    )
    assert legacy.bounds == (1.0, 2.0, 3.0, 4.0)


@pytest.mark.parametrize(
    "payload",
    [
        {"page": 0, "bounds": None},
        {"page": 1, "bounds": [0, 0, 0, 1]},
        {"page": 1, "bounds": [0, 0, 1]},
        {"page": 1, "bounds": [0, 0, float("inf"), 1]},
    ],
)
def test_pagination_fragment_rejects_invalid_geometry(payload):
    with pytest.raises((TypeError, ValueError)):
        PaginationFragment.from_dict(payload)


def test_legacy_m2_pagination_map_remains_accepted():
    restored = PaginationMap.from_dict(
        {
            "version": "M2-stub",
            "nodes": [
                {
                    "nodeId": "doc:finalize",
                    "fragments": [{"page": 1, "bounds": None}],
                }
            ],
        }
    )
    assert restored.version == "M2-stub"
    assert restored.nodes[0].fragments[0].bounds is None
