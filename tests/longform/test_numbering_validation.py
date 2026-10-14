"""Structural numbering contract tests (synthetic package fixtures only).

These tests exercise the read-only XML validator.  They are deliberately not
native Word/WPS acceptance tests; those still require an archived application
output and an edit/reopen round trip.
"""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from skills.WPSComposer.scripts.artifact_transport import ArtifactValidationError
from skills.WPSComposer.scripts.longform.numbering_validation import (
    heading_expectations,
    validate_heading_numbering,
)


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = f'xmlns:w="{W}"'


def _heading(bookmark: str, text: str, level: int, *, num_id: str | None = "7", ilvl: str | None = None) -> str:
    ppr = f'<w:pStyle w:val="Heading{level}"/>'
    if num_id is not None:
        ppr += '<w:numPr>'
        if ilvl is not None:
            ppr += f'<w:ilvl w:val="{ilvl}"/>'
        ppr += f'<w:numId w:val="{num_id}"/></w:numPr>'
    return (
        f"<w:p><w:pPr>{ppr}</w:pPr>"
        f'<w:bookmarkStart w:id="{bookmark}" w:name="{bookmark}"/>'
        f"<w:r><w:t>{text}</w:t></w:r>"
        f'<w:bookmarkEnd w:id="{bookmark}"/></w:p>'
    )


def _docx(
    path: Path,
    paragraphs: str,
    *,
    levels: str = "",
    styles: str = "",
    nums: str = "",
) -> Path:
    with ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", f"<w:document {NS}><w:body>{paragraphs}</w:body></w:document>")
        archive.writestr("word/styles.xml", f"<w:styles {NS}>{styles}</w:styles>")
        if nums:
            archive.writestr("word/numbering.xml", f"<w:numbering {NS}>{nums}</w:numbering>")
    return path


def _decimal_defs(*, num_id: str = "7", abstract_id: str = "3", split: bool = False, restart: str = "1") -> str:
    num = abstract_id if not split else "99"
    levels = "".join(
        f'<w:lvl w:ilvl="{index}"><w:start w:val="{restart}"/>'
        f'<w:pStyle w:val="Heading{index + 1}"/><w:numFmt w:val="decimal"/>'
        f'<w:lvlText w:val="{placeholder}"/></w:lvl>'
        for index, placeholder in enumerate(("%1.", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4"))
    )
    return f'<w:abstractNum w:abstractNumId="{abstract_id}">{levels}</w:abstractNum><w:num w:numId="{num_id}"><w:abstractNumId w:val="{num}"/></w:num>'


def _expect(*items: dict) -> list[dict]:
    return list(items)


def _valid_expectations() -> list[dict]:
    return _expect(
        {"bookmarkName": "h1", "text": "章节", "level": 1, "numbering": True, "numberingScheme": "decimal", "sequenceTransparent": False},
        {"bookmarkName": "h2", "text": "小节", "level": 2, "numbering": True, "numberingScheme": "decimal", "sequenceTransparent": False},
    )


def test_heading_expectations_extracts_plain_contract_fields():
    class Plan:
        operations = [
            {"op": "writer.add_paragraph", "args": {"text": "body"}},
            {"op": "writer.add_heading", "args": {"bookmarkName": "h1", "text": "章节", "level": 1, "numbering": True, "numberingScheme": "decimal", "sequenceTransparent": False, "private": "ignored"}},
        ]

    assert heading_expectations(Plan()) == [{
        "bookmarkName": "h1", "text": "章节", "level": 1, "numbering": True,
        "numberingScheme": "decimal", "sequenceTransparent": False,
    }]


def test_valid_shared_numbering_passes(tmp_path: Path):
    path = _docx(
        tmp_path / "valid.docx",
        _heading("h1", "章节", 1, ilvl="0") + _heading("h2", "小节", 2, ilvl="1"),
        nums=_decimal_defs(),
    )
    assert validate_heading_numbering(path, _valid_expectations()) is True


@pytest.mark.parametrize(
    ("case", "expected_code"),
    [
        ("split", "WPSC_HEADING_NUMBERING_SHARED_SEQUENCE"),
        ("zero", "WPSC_HEADING_NUMBERING_NUMID_ZERO"),
        ("missing-abstract", "WPSC_HEADING_NUMBERING_ABSTRACT_MISSING"),
        ("wrong-level", "WPSC_HEADING_NUMBERING_LEVEL_MISMATCH"),
        ("restart", "WPSC_HEADING_NUMBERING_START_INVALID"),
    ],
)
def test_corrupt_numbering_is_rejected_with_stable_code(tmp_path: Path, case: str, expected_code: str):
    paragraphs = _heading("h1", "章节", 1, ilvl="0") + _heading("h2", "小节", 2, ilvl="1")
    nums = _decimal_defs(split=case == "split", restart="9" if case == "restart" else "1")
    if case == "split":
        paragraphs = _heading("h1", "章节", 1, ilvl="0", num_id="7") + _heading("h2", "小节", 2, ilvl="1", num_id="8")
        nums = _decimal_defs() + '<w:num w:numId="8"><w:abstractNumId w:val="3"/></w:num>'
    if case == "zero":
        paragraphs = _heading("h1", "章节", 1, num_id="0", ilvl="0") + _heading("h2", "小节", 2, ilvl="1")
    if case == "missing-abstract":
        nums = _decimal_defs().replace('<w:abstractNum w:abstractNumId="3">', '<w:abstractNum w:abstractNumId="8">')
    if case == "wrong-level":
        paragraphs = _heading("h1", "章节", 1, ilvl="1") + _heading("h2", "小节", 2, ilvl="1")
    path = _docx(tmp_path / f"{case}.docx", paragraphs, nums=nums)
    with pytest.raises(ArtifactValidationError) as raised:
        validate_heading_numbering(path, _valid_expectations())
    assert getattr(raised.value, "code", None) == expected_code
    assert "章节" not in str(raised.value)


def test_unnumbered_heading_must_not_inherit_sequence(tmp_path: Path):
    path = _docx(
        tmp_path / "unnumbered.docx",
        _heading("h1", "题名", 1, num_id="7", ilvl="0"),
        nums=_decimal_defs(),
    )
    expectation = [{"bookmarkName": "h1", "text": "题名", "level": 1, "numbering": False, "numberingScheme": None, "sequenceTransparent": True}]
    with pytest.raises(ArtifactValidationError) as raised:
        validate_heading_numbering(path, expectation)
    assert getattr(raised.value, "code", None) == "WPSC_UNNUMBERED_HEADING_BOUND"


def test_duplicate_literal_prefix_and_missing_bookmark_fail_without_text_fallback(tmp_path: Path):
    path = _docx(
        tmp_path / "duplicate.docx",
        _heading("other", "1. 章节", 1, ilvl="0"),
        nums=_decimal_defs(),
    )
    expectation = [{"bookmarkName": "h1", "text": "章节", "level": 1, "numbering": True, "numberingScheme": "decimal", "sequenceTransparent": False}]
    with pytest.raises(ArtifactValidationError) as raised:
        validate_heading_numbering(path, expectation)
    assert getattr(raised.value, "code", None) == "WPSC_HEADING_BOOKMARK_MISSING"
    assert "章节" not in str(raised.value)


def test_duplicate_literal_prefix_is_rejected(tmp_path: Path):
    path = _docx(
        tmp_path / "literal-prefix.docx",
        _heading("h1", "1. 章节", 1, ilvl="0"),
        nums=_decimal_defs(),
    )
    expectation = [{"bookmarkName": "h1", "text": "1. 章节", "level": 1, "numbering": True, "numberingScheme": "decimal", "sequenceTransparent": False}]
    with pytest.raises(ArtifactValidationError) as raised:
        validate_heading_numbering(path, expectation)
    assert getattr(raised.value, "code", None) == "WPSC_HEADING_NUMBERING_LITERAL_PREFIX"


def test_bookmark_boundaries_can_cross_paragraphs(tmp_path: Path):
    paragraphs = (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>'
        '<w:bookmarkStart w:id="h1" w:name="h1"/><w:r><w:t>章节</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>续</w:t></w:r><w:bookmarkEnd w:id="h1"/></w:p>'
    )
    path = _docx(tmp_path / "split-bookmark.docx", paragraphs, nums=_decimal_defs())
    expectation = [{"bookmarkName": "h1", "text": "章节", "level": 1, "numbering": True, "numberingScheme": "decimal", "sequenceTransparent": False}]
    assert validate_heading_numbering(path, expectation) is True
