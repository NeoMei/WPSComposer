from __future__ import annotations

from pathlib import Path

from skills.WPSComposer.scripts.longform.executor import (
    PaginationFragment,
    PaginationMap,
    PaginationNode,
)
from skills.WPSComposer.scripts.longform.pdf_quality import (
    PdfPage,
    QualityPolicy,
    analyze_pages,
    analyze_pdf,
)
from skills.WPSComposer.scripts.longform.quality import (
    PageRole,
    QualityConfidence,
    QualitySeverity,
)
from tests._pdf_fixture import write_minimal_pdf


def _page(number=1, *, glyphs=((72.0, 72.0, 80.0, 84.0),), objects=()):
    return PdfPage(
        physical_page=number,
        width=612.0,
        height=792.0,
        rotation=0,
        media_box=(0.0, 0.0, 612.0, 792.0),
        crop_box=(0.0, 0.0, 612.0, 792.0),
        glyph_bounds=tuple(glyphs),
        object_bounds=tuple(objects),
    )


def _map(*nodes):
    return PaginationMap(version="M5-v1", nodes=tuple(nodes))


def _node(node_id, page=1, bounds=(72.0, 90.0, 540.0, 300.0)):
    return PaginationNode(
        node_id=node_id,
        story="main",
        sections=("body",),
        page_start=page,
        page_end=page,
        fragments=(PaginationFragment(page=page, bounds=bounds),),
    )


def _codes(report):
    return [finding.code for finding in report.findings]


def test_blank_body_page_is_high_confidence_but_explicit_break_is_exempt():
    body = analyze_pages(
        (_page(glyphs=()),), _map(), {1: PageRole.BODY}, QualityPolicy()
    )
    assert _codes(body) == ["UNEXPECTED_BLANK_PAGE"]
    assert body.findings[0].confidence is QualityConfidence.HIGH
    assert body.findings[0].repair_key == "remove-unexpected-blank"

    explicit = analyze_pages(
        (_page(glyphs=()),), _map(), {1: PageRole.EXPLICIT_BREAK}, QualityPolicy()
    )
    assert explicit.findings == ()


def test_header_and_footer_alone_do_not_hide_a_blank_body_page():
    report = analyze_pages(
        (_page(glyphs=((72.0, 20.0, 180.0, 35.0), (300.0, 760.0, 312.0, 775.0))),),
        _map(),
        {1: PageRole.BODY},
        QualityPolicy(),
    )
    assert _codes(report) == ["UNEXPECTED_BLANK_PAGE"]

    landscape = PdfPage(
        physical_page=1,
        width=841.89,
        height=595.28,
        rotation=0,
        media_box=(0.0, 0.0, 841.89, 595.28),
        crop_box=(0.0, 0.0, 841.89, 595.28),
        glyph_bounds=(
            (398.4, 46.58, 425.4, 55.59),
            (72.0, 539.09, 74.0, 548.10),
        ),
    )
    report = analyze_pages(
        (landscape,), _map(), {1: PageRole.BODY}, QualityPolicy()
    )
    assert _codes(report) == ["UNEXPECTED_BLANK_PAGE"]


def test_visual_overflow_maps_to_stable_node_and_closed_repair():
    pmap = _map(_node("fig:wide", bounds=(60.0, 80.0, 570.0, 400.0)))
    report = analyze_pages(
        (_page(),),
        pmap,
        {1: PageRole.BODY},
        QualityPolicy(node_kinds=(("fig:wide", "image"),)),
    )
    finding = report.findings[0]
    assert finding.code == "IMAGE_TOO_WIDE"
    assert finding.node_id == "fig:wide"
    assert finding.bounds == (60.0, 80.0, 570.0, 400.0)
    assert finding.repair_key == "fit-image"


def test_table_overflow_uses_single_compression_repair():
    report = analyze_pages(
        (_page(),),
        _map(_node("tab:wide", bounds=(60.0, 80.0, 570.0, 400.0))),
        {1: PageRole.BODY},
        QualityPolicy(node_kinds=(("tab:wide", "table"),)),
    )
    assert report.findings[0].code == "TABLE_TOO_WIDE"
    assert report.findings[0].repair_key == "compress-table"


def test_heading_orphan_and_caption_separation_are_high_confidence():
    report = analyze_pages(
        (_page(1), _page(2)),
        _map(
            _node("head:last", bounds=(72.0, 700.0, 300.0, 725.0)),
            _node("fig:object", page=1),
            _node("fig:caption", page=2),
        ),
        {1: PageRole.BODY, 2: PageRole.BODY},
        QualityPolicy(
            node_kinds=(
                ("head:last", "heading"),
                ("fig:object", "image"),
                ("fig:caption", "caption"),
            ),
            caption_targets=(("fig:caption", "fig:object"),),
        ),
    )
    assert set(_codes(report)) == {"CAPTION_SEPARATED", "HEADING_ORPHAN"}
    assert all(item.confidence is QualityConfidence.HIGH for item in report.findings)


def test_heading_orphan_uses_following_body_lines_not_only_page_position():
    heading = _node("head:middle", bounds=(72.0, 100.0, 300.0, 120.0))
    orphan = analyze_pages(
        (_page(glyphs=((72.0, 100.0, 200.0, 118.0),)),),
        _map(heading),
        {1: PageRole.BODY},
        QualityPolicy(node_kinds=(("head:middle", "heading"),)),
    )
    assert _codes(orphan) == ["HEADING_ORPHAN"]

    supported = analyze_pages(
        (_page(glyphs=(
            (72.0, 100.0, 200.0, 118.0),
            (72.0, 130.0, 400.0, 142.0),
            (72.0, 146.0, 400.0, 158.0),
        )),),
        _map(heading),
        {1: PageRole.BODY},
        QualityPolicy(node_kinds=(("head:middle", "heading"),)),
    )
    assert supported.findings == ()


def test_low_dpi_stays_readable_and_requests_notice_not_resize():
    report = analyze_pages(
        (_page(),),
        _map(_node("fig:low")),
        {1: PageRole.BODY},
        QualityPolicy(
            node_kinds=(("fig:low", "image"),),
            effective_dpi=(("fig:low", 72.0),),
            minimum_dpi=150.0,
        ),
    )
    finding = report.findings[0]
    assert finding.code == "IMAGE_LOW_DPI"
    assert finding.repair_key == "low-dpi-notice"
    assert finding.severity is QualitySeverity.DEGRADED


def test_header_toc_table_and_field_structural_checks():
    report = analyze_pages(
        (_page(),),
        _map(_node("tab:one"), _node("ref:bad")),
        {1: PageRole.TOC},
        QualityPolicy(
            header_display_units=40,
            toc_last_page_utilization=0.18,
            table_header_repeated=(("tab:one", False),),
            unresolved_field_nodes=("ref:bad",),
        ),
    )
    assert set(_codes(report)) == {
        "FIELD_UNRESOLVED",
        "HEADER_OVERFLOW",
        "TABLE_HEADER_NOT_REPEATED",
        "TOC_SPARSE_OVERFLOW",
    }
    by_code = {item.code: item for item in report.findings}
    assert by_code["HEADER_OVERFLOW"].repair_key == "shorten-header"
    assert by_code["TOC_SPARSE_OVERFLOW"].repair_key == "compact-toc"
    assert by_code["FIELD_UNRESOLVED"].repair_key is None


def test_medium_and_low_observations_never_request_mutation():
    report = analyze_pages(
        (_page(), _page(2)),
        _map(),
        {1: PageRole.BIBLIOGRAPHY, 2: PageRole.BODY},
        QualityPolicy(
            bibliography_spacing_ratio=1.8,
            final_page_utilization=0.10,
        ),
    )
    assert set(_codes(report)) == {"BIBLIOGRAPHY_SPACING", "LAST_PAGE_SPARSE"}
    assert all(item.repair_key is None for item in report.findings)
    assert {item.confidence for item in report.findings} == {QualityConfidence.MEDIUM}


def test_page_number_sequence_failure_is_degraded_without_heuristic_repair():
    report = analyze_pages(
        (_page(),),
        _map(),
        {1: PageRole.BODY},
        QualityPolicy(page_number_sequence_valid=False),
    )
    finding = report.findings[0]
    assert finding.code == "PAGE_NUMBER_SEQUENCE_INVALID"
    assert finding.severity is QualitySeverity.DEGRADED
    assert finding.repair_key is None


def test_analyze_pdf_public_entry_uses_validated_geometry(tmp_path: Path):
    pdf = write_minimal_pdf(tmp_path / "blank.pdf")
    report = analyze_pdf(pdf, _map(), {1: PageRole.BODY}, QualityPolicy())
    assert report.page_count == 1
    assert _codes(report) == ["UNEXPECTED_BLANK_PAGE"]


def test_report_serialization_is_stable_across_repeated_analysis():
    inputs = (
        (_page(),),
        _map(_node("fig:wide", bounds=(60.0, 80.0, 570.0, 400.0))),
        {1: PageRole.BODY},
        QualityPolicy(node_kinds=(("fig:wide", "image"),)),
    )
    assert analyze_pages(*inputs).to_dict() == analyze_pages(*inputs).to_dict()
