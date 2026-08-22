from __future__ import annotations

import json
import os
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.longform_m3_evidence import (
    _assert_native_document,
    _assert_stable_bookmarks,
    _validate_pdf_pages,
    run_longform_m3_evidence,
    validate_m3_evidence_report,
)


@pytest.mark.skipif(
    os.environ.get("WPSCOMPOSER_RUN_REAL_WPS") != "1",
    reason="set WPSCOMPOSER_RUN_REAL_WPS=1 for native macOS WPS acceptance",
)
def test_real_macos_wps_m3_mutation_evidence() -> None:
    raw_output = os.environ.get("WPSCOMPOSER_M3_EVIDENCE_DIR")
    if not raw_output:
        pytest.fail("WPSCOMPOSER_M3_EVIDENCE_DIR must name a unique evidence directory")
    report_path = run_longform_m3_evidence(Path(raw_output), timeout=300.0)
    report = json.loads(report_path.read_text(encoding="utf-8"))
    validate_m3_evidence_report(report)
    assert report["status"] == "passed"
    assert report["mutationKinds"] == ["move", "insert", "delete"]
    assert len([item for item in report["artifacts"] if item["kind"] == "pdf"]) == 2
    assert report["screenshots"]


def test_m3_evidence_validator_rejects_private_or_open_ended_data() -> None:
    valid = {
        "status": "passed",
        "wpsVersion": "12.1",
        "capabilities": {
            "chapter": "passed", "global": "passed", "mutation": "passed",
            "pdf": "passed", "visual": "passed",
        },
        "artifacts": [{"name": "chapter.docx", "sha256": "a" * 64, "kind": "docx"}],
        "counts": {
            label: {
                "fieldCount": 1, "figureSequenceCount": 1,
                "tableSequenceCount": 1, "equationSequenceCount": 1,
                "referenceCount": 1, "indexCount": 2, "bookmarkCount": 3,
                "imageCount": 1, "tableCount": 1, "mergeCount": 0,
                "columnContainerCount": 1, "columnGeometryCount": 1,
                "borderlessColumnContainerCount": 1,
            }
            for label in ("chapterInitial", "chapterMutated", "globalInitial", "globalMutated")
        },
        "refreshRounds": [2] * 10,
        "mutationKinds": ["move", "insert", "delete"],
        "screenshots": ["screenshots/chapter-1.png"],
    }
    validate_m3_evidence_report(valid)
    private = dict(valid, fieldResults=["secret"])
    with pytest.raises(ValueError, match="private field"):
        validate_m3_evidence_report(private)
    private = dict(valid)
    private["artifacts"] = [{"name": "/Users/private/chapter.docx", "sha256": "a" * 64, "kind": "docx"}]
    with pytest.raises(ValueError, match="relative"):
        validate_m3_evidence_report(private)
    nested = json.loads(json.dumps(valid))
    nested["counts"]["chapterInitial"]["fieldResults"] = ["secret"]
    with pytest.raises(ValueError, match="private field"):
        validate_m3_evidence_report(nested)
    nested = json.loads(json.dumps(valid))
    nested["counts"]["chapterInitial"]["payload"] = "secret"
    with pytest.raises(ValueError, match="private field"):
        validate_m3_evidence_report(nested)
    for forbidden in ("bookmarkNames", "resourceHash", "paths"):
        nested = json.loads(json.dumps(valid))
        nested["counts"]["chapterInitial"][forbidden] = "secret"
        with pytest.raises(ValueError, match="private field"):
            validate_m3_evidence_report(nested)
    nested = json.loads(json.dumps(valid))
    nested["counts"]["chapterInitial"]["unexpected"] = 1
    with pytest.raises(ValueError, match="count shape"):
        validate_m3_evidence_report(nested)
    nested = json.loads(json.dumps(valid))
    nested["screenshots"] = ["/Users/private/chapter-1.png"]
    with pytest.raises(ValueError, match="screenshot"):
        validate_m3_evidence_report(nested)


def test_native_gate_rejects_localized_styleref_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    broken = {
        "fieldCount": 12,
        "figureSequenceCount": 2,
        "tableSequenceCount": 3,
        "equationSequenceCount": 1,
        "referenceCount": 3,
        "indexCount": 2,
        "bookmarkCount": 8,
        "imageCount": 2,
        "tableCount": 3,
        "mergeCount": 1,
        "columnContainerCount": 1,
        "columnGeometryCount": 1,
        "borderlessColumnContainerCount": 1,
        "fieldResults": [
            ('STYLEREF "Heading 1" \\s', "关联章节样式未自动编号"),
            ("SEQ WPSC_FIG \\* ARABIC \\s 1", "1"),
            ("SEQ WPSC_FIG \\* ARABIC \\s 1", "2"),
            ("SEQ WPSC_TAB \\* ARABIC \\s 1", "1"),
            ("SEQ WPSC_TAB \\* ARABIC \\s 1", "2"),
            ("SEQ WPSC_TAB \\* ARABIC \\s 1", "3"),
            ("SEQ WPSC_EQ \\* ARABIC \\s 1", "1"),
        ],
        "bookmarkNames": ["wpsc_" + str(index) for index in range(8)],
        "paragraphTexts": [
            "图 关联章节样式未自动编号-1 Chapter diagram",
            "[HEADING_PREFIX_AMBIGUOUS] HEADING_PREFIX_AMBIGUOUS",
        ],
        "xml": "<w:body><w:p></w:p><w:sectPr/></w:body>",
    }
    monkeypatch.setattr(
        "skills.WPSComposer.scripts.longform_m3_evidence.inspect_native_docx",
        lambda _path: broken,
    )

    with pytest.raises(AssertionError, match="unresolved native field result"):
        _assert_native_document(tmp_path / "broken.docx", mode="chapter", mutated=True)


def test_native_gate_requires_exact_chapter_labels(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    structurally_plausible_but_wrong = {
        "fieldCount": 12,
        "figureSequenceCount": 2,
        "tableSequenceCount": 3,
        "equationSequenceCount": 1,
        "referenceCount": 3,
        "indexCount": 2,
        "bookmarkCount": 8,
        "imageCount": 2,
        "tableCount": 3,
        "mergeCount": 1,
        "columnContainerCount": 1,
        "columnGeometryCount": 1,
        "borderlessColumnContainerCount": 1,
        "fieldResults": [
            ('STYLEREF "Heading 1" \\s', "2"),
            ("SEQ WPSC_FIG \\* ARABIC \\s 1", "1"),
            ("SEQ WPSC_FIG \\* ARABIC \\s 1", "2"),
            ("SEQ WPSC_TAB \\* ARABIC \\s 1", "1"),
            ("SEQ WPSC_TAB \\* ARABIC \\s 1", "2"),
            ("SEQ WPSC_TAB \\* ARABIC \\s 1", "3"),
            ("SEQ WPSC_EQ \\* ARABIC \\s 1", "1"),
            ("REF wpsc_fig_a", "2-2"),
            ("REF wpsc_tab_b", "2-1"),
            ("REF wpsc_eq_c", "2-1"),
        ],
        "bookmarkNames": ["wpsc_" + str(index) for index in range(8)],
        "paragraphTexts": [
            "图 2-1 Two-column comparison",
            "图 2-2 Chapter diagram",
            "表 2-1 Three-line grouped data",
            "(2-1)",
            "See 图 2-2, 表 2-1, and (2-1).",
        ],
        "xml": "<w:body><w:p/><w:sectPr/></w:body>",
    }
    monkeypatch.setattr(
        "skills.WPSComposer.scripts.longform_m3_evidence.inspect_native_docx",
        lambda _path: structurally_plausible_but_wrong,
    )

    with pytest.raises(AssertionError, match="exact visible labels"):
        _assert_native_document(tmp_path / "wrong.docx", mode="chapter", mutated=True)


def test_native_gate_rejects_stacked_images_and_visible_column_container(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    inspection = {
        "fieldCount": 12,
        "figureSequenceCount": 2,
        "tableSequenceCount": 3,
        "equationSequenceCount": 1,
        "referenceCount": 3,
        "indexCount": 2,
        "bookmarkCount": 8,
        "imageCount": 4,
        "tableCount": 4,
        "mergeCount": 1,
        "columnContainerCount": 0,
        "columnGeometryCount": 0,
        "borderlessColumnContainerCount": 0,
        "fieldResults": [
            ('TOC \\h \\z \\c "WPSC_FIG"', "Two-column comparison Chapter diagram"),
            ('TOC \\h \\z \\c "WPSC_TAB"', "Grid data Three-line grouped data Inserted chapter table"),
            ("SEQ WPSC_FIG", "1"), ("SEQ WPSC_FIG", "2"),
            ("SEQ WPSC_TAB", "1"), ("SEQ WPSC_TAB", "2"), ("SEQ WPSC_TAB", "3"),
            ("SEQ WPSC_EQ", "1"),
            ("REF wpsc_fig_a", "1-2"),
            ("REF wpsc_tab_b", "1-1"),
            ("REF wpsc_eq_c", "1-1"),
        ],
        "bookmarkNames": ["wpsc_" + str(index) for index in range(8)],
        "paragraphTexts": [
            "图 1-1 Two-column comparison", "表 1-1 Three-line grouped data",
            "(1-1)", "See 图 1-2, 表 1-2, and (1-1).",
            "Unnumbered appendix note", "表 1-3 Inserted chapter table",
        ],
        "xml": "<w:body><w:p></w:p><w:sectPr/></w:body>",
    }
    monkeypatch.setattr(
        "skills.WPSComposer.scripts.longform_m3_evidence.inspect_native_docx",
        lambda _path: inspection,
    )

    with pytest.raises(AssertionError, match="two-column geometry"):
        _assert_native_document(tmp_path / "stacked.docx", mode="chapter", mutated=True)


def test_pdf_visual_gate_rejects_vertical_stack_and_empty_cover() -> None:
    pages = [
        {"width": 595.0, "height": 842.0, "text": "1", "words": [], "images": []},
        {
            "width": 595.0,
            "height": 842.0,
            "text": "图 1-1 Two-column comparison",
            "words": [],
            "images": [
                {"x0": 190.0, "x1": 410.0, "top": 100.0, "bottom": 250.0},
                {"x0": 190.0, "x1": 410.0, "top": 270.0, "bottom": 420.0},
            ],
        },
    ]
    with pytest.raises(AssertionError, match="cover page"):
        _validate_pdf_pages(pages, mode="chapter")

    pages[0]["text"] = "M3 chapter evidence\nWPSComposer"
    with pytest.raises(AssertionError, match="side-by-side"):
        _validate_pdf_pages(pages, mode="chapter")


def test_bookmark_gate_rejects_renamed_target_after_mutation() -> None:
    initial = {"bookmarkNames": ["wpsc_fig_original", "wpsc_tab_deleted"]}
    mutated = {"bookmarkNames": ["wpsc_fig_renamed", "wpsc_tab_inserted"]}
    with pytest.raises(AssertionError, match="bookmark identity"):
        _assert_stable_bookmarks(
            initial,
            mutated,
            deleted="wpsc_tab_deleted",
            inserted="wpsc_tab_inserted",
        )
