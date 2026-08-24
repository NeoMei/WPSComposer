from __future__ import annotations

import json
from pathlib import Path
import zipfile

import pytest

from skills.WPSComposer.scripts.longform_m4_evidence import (
    _assert_pdf,
    _assert_formula_pdf,
    _assert_formula_rendering_mode,
    build_acceptance_snapshot,
    inspect_m4_docx,
    validate_m4_evidence_report,
)


HERE = Path(__file__).resolve().parent
FIXTURES = HERE / "fixtures"
SNAPSHOT = HERE / "snapshots" / "acceptance.json"


def _write_document_xml(path: Path, body: str) -> None:
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    )
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("word/document.xml", document)


def _complex_field(code: str, result: str) -> str:
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText>{code}</w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:t>{result}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def test_offline_acceptance_snapshot_is_exact_deterministic_and_private() -> None:
    first = build_acceptance_snapshot(FIXTURES)
    second = build_acceptance_snapshot(FIXTURES)

    assert first == second
    assert json.dumps(first, ensure_ascii=False, sort_keys=True, separators=(",", ":")) == (
        json.dumps(
            json.loads(SNAPSHOT.read_text(encoding="utf-8")),
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )
    )
    serialized = json.dumps(first, ensure_ascii=False, sort_keys=True)
    assert str(FIXTURES) not in serialized
    assert "/Users/" not in serialized
    assert "fallback_image" not in serialized


def test_snapshot_covers_m4_acceptance_surface() -> None:
    snapshot = build_acceptance_snapshot(FIXTURES)
    chapter = snapshot["fixtures"]["acceptance_chapter"]
    degradation = snapshot["fixtures"]["degradation"]

    assert len(chapter["formulas"]) == 12
    assert {item["numbering"]["mode"] for item in chapter["formulas"]} == {"chapter"}
    assert snapshot["fixtures"]["acceptance_global"]["formulas"][0]["numbering"]["mode"] == "global"
    assert chapter["citations"] == [
        {"targetId": "b", "number": 1},
        {"targetId": "b", "number": 1},
        {"targetId": "a", "number": 2},
    ]
    assert [(item["id"], item["cited"]) for item in chapter["bibliography"]] == [
        ("b", True), ("a", True), ("c", False),
    ]
    assert chapter["issueCodes"] == ["BIBLIOGRAPHY_ENTRY_MALFORMED", "REFERENCE_UNRESOLVED"]
    assert degradation["noticePlacements"] == [
        "block", "block", "block", "document", "inline",
    ]
    assert degradation["recovery"] == [
        {
            "op": "writer.add_equation",
            "recoverableCodes": ["EQUATION_INSERT_FAILED"],
            "fallback": "explicit-image-then-source-notice",
        },
        {
            "op": "writer.add_cross_reference",
            "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
        {
            "op": "writer.add_bibliography",
            "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
            "fallback": "notice",
        },
    ]


def test_evidence_report_is_closed_hashed_and_privacy_safe() -> None:
    valid = {
        "version": "M4",
        "wpsVersion": "12.1",
        "artifacts": [
            {"name": "acceptance.docx", "sha256": "a" * 64},
            {"name": "acceptance.pdf", "sha256": "b" * 64},
        ],
        "metrics": {
            label: {
                "formulaCount": 12,
                "editableFormulaCount": 12,
                "degradedFormulaCount": 0,
                "formulaFallbackImageCount": 0,
                "formulaNoticeCount": 0,
                "equationSequenceCount": 12,
                "referenceCount": 2,
                "citationCount": 3,
                "bibliographyCount": 3,
                "hangingBibliographyCount": 3,
                "inlineNoticeCount": 1,
                "blockNoticeCount": 1,
                    "documentNoticeCount": 1,
                    "pageCount": 2,
                    "citationOrderValid": 1,
                    "bibliographyOrderValid": 1,
                    "formulaOrderValid": 1,
                    "noticePlacementValid": 1,
                    "centeredFormulaCount": 12,
                    "numberAlignedFormulaCount": 12,
            }
            for label in ("initial", "moved", "degradation", "runtimeRecovery")
        },
        "codes": ["REFERENCE_UNRESOLVED"],
        "rounds": [2, 2, 2, 2],
        "screenshots": ["screenshots/acceptance-1.png"],
    }
    validate_m4_evidence_report(valid)

    for invalid in (
        {**valid, "paths": ["/private/output.docx"]},
        {**valid, "codes": ["Traceback: /Users/alice/private.py"]},
        {**valid, "artifacts": [{"name": "../secret.docx", "sha256": "a" * 64}]},
        {**valid, "screenshots": ["/Users/alice/screen.png"]},
        {**valid, "screenshots": ["screenshots/data:abc.png"]},
        {**valid, "rounds": [0]},
    ):
        with pytest.raises(ValueError):
            validate_m4_evidence_report(invalid)


def test_docx_inspector_requires_exact_math_and_fields_outside_math(tmp_path: Path) -> None:
    valid = tmp_path / "valid.docx"
    invalid = tmp_path / "invalid.docx"
    formula = '<m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath>'
    sequence = _complex_field(" SEQ WPSC_EQ ", "1")
    reference = _complex_field(" REF wpsc_eq_a \\h ", "1-1")
    _write_document_xml(valid, f"<w:p>{formula}<w:r><w:tab/></w:r>{sequence}</w:p><w:p>{reference}</w:p>")
    _write_document_xml(
        invalid,
        (
            '<w:p><m:oMath><m:r><m:t>x+y SEQ WPSC_EQ STYLEREF</m:t></m:r>'
            '<w:r><w:tab/></w:r>'
            f'{sequence}</m:oMath></w:p><w:p>{reference}{formula}</w:p>'
            '<w:p><w:r><w:rPr><w:oMath/></w:rPr><w:instrText> SEQ WPSC_EQ </w:instrText></w:r></w:p>'
        ),
    )

    valid_metrics, valid_private = inspect_m4_docx(valid)
    invalid_metrics, invalid_private = inspect_m4_docx(invalid)

    assert valid_metrics["formulaCount"] == 1
    assert valid_metrics["equationSequenceCount"] == 1
    assert valid_private["formulaStructureValid"] is True
    assert valid_private["mathFieldLeakCount"] == 0
    assert valid_private["mathRunPropertyLeakCount"] == 0
    assert valid_private["unexpectedNoticeStyleParagraphCount"] == 0
    assert valid_private["formulaStructuralTagCounts"]["sSup"] == 0
    assert invalid_metrics["formulaCount"] == 2
    assert invalid_private["formulaStructureValid"] is False
    assert invalid_private["mathFieldLeakCount"] == 1
    assert invalid_private["mathRunPropertyLeakCount"] == 1


def test_docx_inspector_detects_notice_style_and_native_math_structure(tmp_path: Path) -> None:
    path = tmp_path / "structured.docx"
    math = (
        '<m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>'
        '<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>'
    )
    styled = (
        '<w:p><w:pPr><w:rPr><w:i/><w:color w:val="9C0006"/>'
        '<w:shd w:fill="FCE8E6"/></w:rPr></w:pPr>'
        f"{math}</w:p>"
    )
    _write_document_xml(path, styled)

    _, private = inspect_m4_docx(path)

    assert private["formulaStructuralTagCounts"]["sSup"] == 1
    assert private["unexpectedNoticeStyleParagraphCount"] == 1


def test_pdf_gate_rejects_notice_text_overlapping_an_image(tmp_path: Path) -> None:
    from PIL import Image
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    image = Image.new("RGB", (20, 20), "white")
    target = tmp_path / "overlap.pdf"
    document = canvas.Canvas(str(target))
    document.drawImage(ImageReader(image), 72, 500, width=300, height=100)
    document.drawString(80, 540, "[FORMULA_MALFORMED] visible fallback")
    document.save()

    with pytest.raises(AssertionError, match="overlaps an image"):
        _assert_pdf(target)


def test_docx_inspector_accepts_honest_all_formula_degradation(tmp_path: Path) -> None:
    path = tmp_path / "degraded.docx"
    styled_source = (
        '<w:r><w:rPr><w:i/><w:color w:val="9C0006"/>'
        '<w:shd w:fill="FCE8E6"/></w:rPr>'
        '<w:t>[EQUATION_INSERT_FAILED: x^2]</w:t></w:r>'
    )
    tabs = (
        '<w:pPr><w:tabs><w:tab w:val="center" w:pos="3600"/>'
        '<w:tab w:val="right" w:pos="7200"/></w:tabs></w:pPr>'
    )
    source_number = (
        '<w:bookmarkStart w:id="1" w:name="wpsc_eq_source"/>'
        + _complex_field(" SEQ WPSC_EQ ", "1")
        + '<w:bookmarkEnd w:id="1"/>'
    )
    image_number = (
        '<w:bookmarkStart w:id="2" w:name="wpsc_eq_image"/>'
        + _complex_field(" SEQ WPSC_EQ ", "2")
        + '<w:bookmarkEnd w:id="2"/>'
    )
    image = '<w:r><w:drawing><w:inline/></w:drawing></w:r>'
    image_notice = (
        '<w:p><w:r><w:rPr><w:i/><w:color w:val="9C0006"/>'
        '<w:shd w:fill="FCE8E6"/></w:rPr>'
        '<w:t>[EQUATION_INSERT_FAILED: formula image fallback]</w:t></w:r></w:p>'
    )
    _write_document_xml(
        path,
        f'<w:p>{tabs}<w:r><w:tab/></w:r>{styled_source}<w:r><w:tab/></w:r>{source_number}</w:p>'
        f'<w:p>{tabs}<w:r><w:tab/></w:r>{image}<w:r><w:tab/></w:r>{image_number}</w:p>'
        + image_notice,
    )

    metrics, private = inspect_m4_docx(path)
    mode = _assert_formula_rendering_mode(
        metrics,
        private,
        [
            {"bookmarkName": "wpsc_eq_source", "fallbackText": "x^2", "hasImage": False},
            {"bookmarkName": "wpsc_eq_image", "fallbackText": "E = mc^2", "hasImage": True},
        ],
    )

    assert mode == "degraded"
    assert metrics["formulaCount"] == 0
    assert metrics["degradedFormulaCount"] == 2
    assert metrics["formulaFallbackImageCount"] == 1
    assert metrics["formulaNoticeCount"] == 2
    assert metrics["centeredFormulaCount"] == metrics["numberAlignedFormulaCount"] == 2
    assert private["unexpectedNoticeStyleParagraphCount"] == 0


def test_formula_gate_rejects_mixed_pseudo_native_and_degraded_output(tmp_path: Path) -> None:
    path = tmp_path / "mixed.docx"
    native = '<m:oMath><m:r><m:t>x^2</m:t></m:r></m:oMath>'
    first_number = (
        '<w:bookmarkStart w:id="1" w:name="wpsc_eq_native"/>'
        + _complex_field(" SEQ WPSC_EQ ", "1")
        + '<w:bookmarkEnd w:id="1"/>'
    )
    second_number = (
        '<w:bookmarkStart w:id="2" w:name="wpsc_eq_source"/>'
        + _complex_field(" SEQ WPSC_EQ ", "2")
        + '<w:bookmarkEnd w:id="2"/>'
    )
    _write_document_xml(
        path,
        f'<w:p>{native}{first_number}</w:p>'
        f'<w:p><w:r><w:t>[EQUATION_INSERT_FAILED: a/b]</w:t></w:r>{second_number}</w:p>',
    )
    metrics, private = inspect_m4_docx(path)

    with pytest.raises(AssertionError, match="mixed native and degraded"):
        _assert_formula_rendering_mode(
            metrics,
            private,
            [
                {"bookmarkName": "wpsc_eq_native", "fallbackText": "x^2", "hasImage": False},
                {"bookmarkName": "wpsc_eq_source", "fallbackText": "a/b", "hasImage": False},
            ],
        )


def test_formula_pdf_gate_requires_exact_degraded_sources_and_explicit_images(tmp_path: Path) -> None:
    from PIL import Image
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas

    target = tmp_path / "degraded.pdf"
    image = Image.new("RGB", (20, 20), "white")
    document = canvas.Canvas(str(target))
    document.drawString(72, 720, "[EQUATION_INSERT_FAILED: x^2]")
    document.drawImage(ImageReader(image), 72, 620, width=20, height=20)
    document.drawString(72, 580, "[EQUATION_INSERT_FAILED: formula image fallback]")
    document.save()
    expectations = [
        {"fallbackText": "x^2", "hasImage": False},
        {"fallbackText": "E = mc^2", "hasImage": True},
    ]

    _assert_formula_pdf(target, "degraded", expectations)
    with pytest.raises(AssertionError, match="source fallback"):
        _assert_formula_pdf(
            target,
            "degraded",
            [{"fallbackText": "missing/source", "hasImage": False}, expectations[1]],
        )
