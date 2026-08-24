from __future__ import annotations

import pytest

from skills.WPSComposer.scripts.longform.quality import (
    QualityConfidence,
    QualityFinding,
    QualityReport,
    QualitySeverity,
)
from skills.WPSComposer.scripts.longform.relayout import build_relayout_directives


def _finding(repair_key, *, node_id="node:1", severity="degraded", confidence="high"):
    return QualityFinding(
        code="LAYOUT_ISSUE",
        severity=QualitySeverity(severity),
        confidence=QualityConfidence(confidence),
        message="Controlled layout issue",
        node_id=node_id,
        page=1,
        repair_key=repair_key,
    )


@pytest.mark.parametrize(
    ("repair_key", "directive", "args"),
    [
        ("compact-toc", "compact-toc", {"minimumOnly": True}),
        ("shorten-header", "shorten-header", {"maxDisplayUnits": 32}),
        ("fit-image", "fit-image", {"fitToBodyWidth": True}),
        ("compress-table", "compress-table", {"attempts": 1}),
        ("force-table-split", "force-table-split", {"removeVerticalMerge": True}),
        ("keep-heading", "keep-heading", {"minimumFollowingLines": 2}),
        ("keep-caption", "keep-caption", {"samePage": True}),
        ("remove-unexpected-blank", "remove-unexpected-blank", {"maximumPages": 1}),
    ],
)
def test_closed_relayout_matrix(repair_key, directive, args):
    decision = build_relayout_directives(QualityReport((_finding(repair_key),), 1))
    assert len(decision.directives) == 1
    assert decision.directives[0].kind == directive
    assert decision.directives[0].args == args
    assert decision.notice_findings == ()


def test_low_dpi_is_notice_only_and_does_not_trigger_generation():
    finding = _finding("low-dpi-notice")
    decision = build_relayout_directives(QualityReport((finding,), 1))
    assert decision.directives == ()
    assert decision.notice_findings == (finding,)


@pytest.mark.parametrize(
    ("severity", "confidence"),
    [("warning", "high"), ("info", "high"), ("degraded", "medium"), ("degraded", "low")],
)
def test_only_high_confidence_degraded_findings_can_mutate(severity, confidence):
    finding = _finding("fit-image", severity=severity, confidence=confidence)
    decision = build_relayout_directives(QualityReport((finding,), 1))
    assert decision.directives == ()
    assert decision.notice_findings == ()


def test_unknown_or_missing_repair_is_not_heuristically_mutated():
    unknown = _finding("future-repair")
    missing = _finding(None, node_id="node:2")
    decision = build_relayout_directives(QualityReport((unknown, missing), 1))
    assert decision.directives == ()
    assert decision.notice_findings == (unknown, missing)


def test_directives_are_deduplicated_and_sorted_by_closed_priority():
    findings = (
        _finding("keep-caption", node_id="caption:z"),
        _finding("fit-image", node_id="fig:a"),
        _finding("fit-image", node_id="fig:a"),
        _finding("compact-toc", node_id="doc:toc"),
    )
    decision = build_relayout_directives(QualityReport(findings, 1))
    assert [(item.kind, item.node_id) for item in decision.directives] == [
        ("compact-toc", "doc:toc"),
        ("fit-image", "fig:a"),
        ("keep-caption", "caption:z"),
    ]


def test_second_relayout_is_forbidden_and_degraded_findings_become_notices():
    findings = (
        _finding("fit-image", node_id="fig:a"),
        _finding("keep-heading", node_id="head:b"),
    )
    decision = build_relayout_directives(
        QualityReport(findings, 1), relayout_already_used=True
    )
    assert decision.directives == ()
    assert decision.notice_findings == tuple(sorted(findings, key=lambda item: item.sort_key()))


def test_directive_rejects_missing_node_except_document_level_repairs():
    with pytest.raises(ValueError, match="node_id"):
        build_relayout_directives(
            QualityReport((_finding("fit-image", node_id=None),), 1)
        )
    decision = build_relayout_directives(
        QualityReport((_finding("shorten-header", node_id=None),), 1)
    )
    assert decision.directives[0].node_id is None
