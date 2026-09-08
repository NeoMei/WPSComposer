"""Check saved-artifact validation against copies of real native Word outputs."""
from __future__ import annotations

import importlib.util
import json
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import pytest

fitz = pytest.importorskip("fitz", reason="Native artifact validation requires PyMuPDF")
pytest.importorskip("PIL")
ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "fixtures/msoffice_spike/validate_windows_artifacts.py"
EVIDENCE = ROOT / "docs/verification/windows-word-spike/final-cleanup-review/windows-run-09"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"


@pytest.fixture
def validator(monkeypatch):
    monkeypatch.syspath_prepend(str(SCRIPT.parent))
    spec = importlib.util.spec_from_file_location("msoffice_artifact_validator", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.fixture
def artifacts(tmp_path):
    directory = tmp_path / "run"
    directory.mkdir()
    for name in ("probe.docx", "probe.pdf", "result.json", "artifact-validation.json"):
        shutil.copy2(EVIDENCE / name, directory / name)
    return directory


def edit_xml(directory, edit, member="word/document.xml"):
    path = directory / "probe.docx"
    with zipfile.ZipFile(path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(files[member])
    edit(root)
    files[member] = ET.tostring(root)
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in files.items():
            archive.writestr(name, data)


def report(directory):
    return json.loads((directory / "artifact-validation.json").read_text("utf-8"))


def heading(root):
    return next(p for p in root.findall("w:body/w:p", NS)
                if "".join(n.text or "" for n in p.findall(".//w:t", NS)) == "原生办公文档验证")


def set_paragraph_number(root, num_id):
    props = heading(root).find("w:pPr", NS)
    num_props = props.find("w:numPr", NS)
    if num_props is None:
        num_props = ET.SubElement(props, W + "numPr")
    num = num_props.find("w:numId", NS)
    if num is None:
        num = ET.SubElement(num_props, W + "numId")
    num.set(W + "val", num_id)


def test_native_artifacts_pass(validator, artifacts):
    assert validator.validate(artifacts)
    assert report(artifacts)["all_checks_passed"]


@pytest.mark.parametrize("value", ["0", "false", "off"])
def test_disabled_repeating_header_is_rejected(validator, artifacts, value):
    edit_xml(artifacts, lambda root: root.find(".//w:tblHeader", NS).set(W + "val", value))
    assert not validator.validate(artifacts)
    assert report(artifacts)["checks"]["repeat_header_xml"] is False


@pytest.mark.parametrize("value", ["1", "true", "on"])
def test_enabled_repeating_header_is_accepted(validator, artifacts, value):
    edit_xml(artifacts, lambda root: root.find(".//w:tblHeader", NS).set(W + "val", value))
    assert validator.validate(artifacts)


@pytest.mark.parametrize("num_id", ["0", "999"])
def test_paragraph_numbering_override_is_checked(validator, artifacts, num_id):
    edit_xml(artifacts, lambda root: set_paragraph_number(root, num_id))
    assert not validator.validate(artifacts)
    assert report(artifacts)["all_checks_passed"] is False


def test_valid_paragraph_numbering_override_is_accepted(validator, artifacts):
    edit_xml(artifacts, lambda root: set_paragraph_number(root, "2"))
    assert validator.validate(artifacts)


@pytest.mark.parametrize("damage", ["missing_table", "missing_heading", "broken_docx", "broken_pdf", "missing_pdf", "broken_result"])
def test_malformed_artifacts_replace_stale_pass(validator, artifacts, damage):
    if damage == "missing_table":
        edit_xml(artifacts, lambda root: root.find("w:body", NS).remove(root.find("w:body/w:tbl", NS)))
    elif damage == "missing_heading":
        edit_xml(artifacts, lambda root: root.find("w:body", NS).remove(heading(root)))
    elif damage == "missing_pdf":
        (artifacts / "probe.pdf").unlink()
    else:
        name = {"broken_docx": "probe.docx", "broken_pdf": "probe.pdf", "broken_result": "result.json"}[damage]
        (artifacts / name).write_bytes(b"invalid artifact")
    assert report(artifacts)["all_checks_passed"] is True
    assert validator.validate(artifacts) is False
    failed = report(artifacts)
    assert failed["all_checks_passed"] is False
    assert failed["errors"] and failed["errors"][0]["message"]


def test_cli_failure_updates_report_and_exits_nonzero(artifacts):
    (artifacts / "probe.docx").write_bytes(b"invalid zip")
    process = subprocess.run([sys.executable, str(SCRIPT), str(artifacts)], capture_output=True, text=True)
    assert process.returncode != 0
    assert report(artifacts)["all_checks_passed"] is False
    assert report(artifacts)["errors"][0]["type"] == "BadZipFile"


def remove_pdf_number(directory, occurrence):
    path = directory / "probe.pdf"
    with fitz.open(path) as pdf:
        matches = pdf[0].search_for("1.1.1")
        assert len(matches) == 2, "The native fixture contains a TOC entry and a body heading"
        pdf[0].add_redact_annot(matches[occurrence], fill=(1, 1, 1))
        pdf[0].apply_redactions()
        pdf.save(directory / "changed.pdf")
    (directory / "changed.pdf").replace(path)


def test_pdf_heading_number_is_required(validator, artifacts):
    remove_pdf_number(artifacts, -1)
    assert validator.validate(artifacts) is False
    assert report(artifacts)["checks"]["pdf_numbered_headings_and_toc_pages"] is False


def test_pdf_toc_number_is_required(validator, artifacts):
    remove_pdf_number(artifacts, 0)
    assert validator.validate(artifacts) is False
    assert report(artifacts)["checks"]["pdf_numbered_headings_and_toc_pages"] is False


def test_pdf_toc_page_number_matches_heading_page(validator, artifacts):
    path = artifacts / "probe.pdf"
    with fitz.open(path) as pdf:
        matches = pdf[0].search_for("1", clip=fitz.Rect(490, 70, 510, 90))
        assert len(matches) == 1
        box = matches[0]
        pdf[0].add_redact_annot(box, fill=(1, 1, 1))
        pdf[0].apply_redactions()
        pdf[0].insert_text((box.x0, box.y1 - 2), "9", fontsize=10)
        pdf.save(artifacts / "changed.pdf")
    (artifacts / "changed.pdf").replace(path)
    assert validator.validate(artifacts) is False
    assert report(artifacts)["checks"]["pdf_numbered_headings_and_toc_pages"] is False


def test_numbering_inherited_from_base_style_is_accepted(validator, artifacts):
    def inherit(root):
        style = next(s for s in root if s.get(W + "styleId") == "1")
        props = style.find("w:pPr", NS)
        num = props.find("w:numPr", NS)
        props.remove(num)
        parent = ET.SubElement(root, W + "style", {W + "type": "paragraph", W + "styleId": "AuditBase"})
        ET.SubElement(parent, W + "pPr").append(num)
        based_on = style.find("w:basedOn", NS)
        if based_on is None:
            based_on = ET.SubElement(style, W + "basedOn")
        based_on.set(W + "val", "AuditBase")
    edit_xml(artifacts, inherit, "word/styles.xml")
    assert validator.validate(artifacts)


@pytest.mark.parametrize("damage", ["none_format", "roman_format", "start_99", "instance_start_99", "parent_start_99", "level_override_start_99", "level_override_none"])
def test_numbering_format_and_start_must_match_fixture(validator, artifacts, damage):
    def mutate(root):
        number = next(n for n in root.findall("w:num", NS) if n.get(W + "numId") == "4")
        abstract = next(a for a in root.findall("w:abstractNum", NS) if a.get(W + "abstractNumId") == "3")
        level = abstract.find("w:lvl[@w:ilvl='2']", NS)
        if damage in ("none_format", "roman_format"):
            level.find("w:numFmt", NS).set(W + "val", "none" if damage == "none_format" else "upperRoman")
        elif damage == "start_99":
            level.find("w:start", NS).set(W + "val", "99")
        else:
            overridden_level = "0" if damage == "parent_start_99" else "2"
            override = ET.SubElement(number, W + "lvlOverride", {W + "ilvl": overridden_level})
            if damage in ("instance_start_99", "parent_start_99"):
                ET.SubElement(override, W + "startOverride", {W + "val": "99"})
            else:
                replacement = ET.fromstring(ET.tostring(level))
                override.append(replacement)
                prop, value = ("start", "99") if damage == "level_override_start_99" else ("numFmt", "none")
                replacement.find("w:" + prop, NS).set(W + "val", value)
    edit_xml(artifacts, mutate, "word/numbering.xml")
    assert not validator.validate(artifacts)
    assert report(artifacts)["checks"]["heading_styles_sizes_native_numbering_links"] is False


def test_valid_instance_start_override_takes_precedence(validator, artifacts):
    def override(root):
        number = next(n for n in root.findall("w:num", NS) if n.get(W + "numId") == "4")
        abstract = next(a for a in root.findall("w:abstractNum", NS) if a.get(W + "abstractNumId") == "3")
        level = abstract.find("w:lvl[@w:ilvl='2']", NS)
        level.find("w:start", NS).set(W + "val", "99")
        instance = ET.SubElement(number, W + "lvlOverride", {W + "ilvl": "2"})
        ET.SubElement(instance, W + "startOverride", {W + "val": "1"})
    edit_xml(artifacts, override, "word/numbering.xml")
    assert validator.validate(artifacts)
