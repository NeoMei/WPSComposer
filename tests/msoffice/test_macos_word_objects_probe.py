"""Contract tests for the production-independent macOS Word object probe."""
from __future__ import annotations

import importlib.util
from pathlib import Path
import shutil
import subprocess
import sys
from zipfile import ZIP_DEFLATED, ZipFile

import pytest


ROOT = Path(__file__).resolve().parents[2]
FIXTURE = ROOT / "fixtures/microsoft_parity/macos_word_objects_probe.py"


def _module():
    spec = importlib.util.spec_from_file_location("macos_word_objects_probe", FIXTURE)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def test_probe_import_has_no_native_or_filesystem_side_effects(monkeypatch):
    from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession

    def forbidden(*args, **kwargs):
        raise AssertionError("fixture performed native work during import")

    monkeypatch.setattr(MacWordSession, "_prepare", forbidden)
    module = _module()
    assert callable(module.run)
    assert callable(module.main)
    assert callable(module.build_probe_commands)


def test_probe_compiler_uses_exact_bound_objects_and_no_global_commands(tmp_path):
    module = _module()
    image = tmp_path / "private-image.png"
    image.write_bytes(b"private")
    source = "\n".join(module.build_probe_commands(image))

    assert "boundDoc" in source
    assert "make new table" in source
    assert "make new inline picture" in source
    assert "make new picture" in source
    assert "make new text box" in source
    assert "merge cell" in source
    assert "heading format of row 1" in source
    assert "allow break across pages" in source
    assert "alternative text" in source
    assert 'set alternative text of floatingPicture' not in source
    assert "wrap type of wrap format" in source
    assert "中文😀" in source
    from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
    assert apple_string(str(image)) in source
    lowered = source.lower()
    for forbidden in ("quit", "kill", "clipboard", "normal template", "macro security"):
        assert forbidden not in lowered


@pytest.mark.skipif(
    sys.platform != "darwin" or not Path("/Applications/Microsoft Word.app").is_dir(),
    reason="Installed macOS Word dictionary required for syntax-only compilation",
)
def test_every_probe_a_branch_compiles_without_execution(tmp_path):
    module = _module()
    image = tmp_path / "private-image.png"
    image.write_bytes(b"private")
    source = (
        'tell application "/Applications/Microsoft Word.app"\n'
        "set boundDoc to document 1\n"
        + "\n".join(module.build_probe_commands(image))
        + "\nend tell\n"
    )
    script = tmp_path / "objects.applescript"
    script.write_text(source, encoding="utf-8")
    result = subprocess.run(
        ["/usr/bin/osacompile", "-o", str(tmp_path / "objects.scpt"), str(script)],
        capture_output=True,
        text=True,
        timeout=20,
    )
    assert result.returncode == 0, result.stderr


def test_artifact_verifier_requires_native_objects_and_embedded_media(tmp_path):
    module = _module()
    with pytest.raises(FileNotFoundError):
        module.verify_artifacts(tmp_path)


def test_native_verifier_rejects_short_geometry_and_wrong_aspect_height():
    module = _module()
    import json
    evidence = ROOT / "docs/verification/microsoft-parity/macos-word-objects-probe/run-05/report.json"
    rows = json.loads(evidence.read_text())["native_reopen_rows"]
    assert all(module.verify_native_rows(rows).values())
    cases = [
        (("table1", "widths"), [], "native_table_widths"),
        (("floating", "geometry"), [300, 120], "native_floating_geometry"),
        (("textbox", "geometry"), [], "native_textbox_geometry"),
        (("inline1", "height"), 1, "native_inline_one_dimension"),
    ]
    for key, replacement, check in cases:
        changed = [list(row) for row in rows]
        for row in changed:
            if (row[0], row[2]) == key:
                row[3] = replacement
        assert module.verify_native_rows(changed)[check] is False


def test_artifact_verifier_rejects_missing_image_relationship_target(tmp_path):
    module = _module()
    source = ROOT / "docs/verification/microsoft-parity/macos-word-objects-probe/run-05"
    shutil.copy2(source / "word-objects.pdf", tmp_path / "word-objects.pdf")
    shutil.copy2(source / "probe-source.png", tmp_path / "probe-source.png")
    with ZipFile(source / "word-objects.docx") as incoming, ZipFile(tmp_path / "word-objects.docx", "w", ZIP_DEFLATED) as outgoing:
        for info in incoming.infolist():
            payload = incoming.read(info.filename)
            if info.filename == "word/_rels/document.xml.rels":
                payload = payload.replace(b"media/image1.png", b"media/missing.png")
            outgoing.writestr(info, payload)
    checks = module.verify_artifacts(tmp_path)
    assert checks["ooxml_embedded_media"] is False


def test_artifact_verifier_rejects_arbitrary_pdf_image_layout(tmp_path):
    module = _module()
    import fitz
    source = ROOT / "docs/verification/microsoft-parity/macos-word-objects-probe/run-05"
    shutil.copy2(source / "word-objects.docx", tmp_path / "word-objects.docx")
    shutil.copy2(source / "probe-source.png", tmp_path / "probe-source.png")
    doc = fitz.open()
    page = doc.new_page()
    page.insert_text((40, 40), "Header Centered Right aligned Horizontal merge M3 center Textbox Following")
    page.insert_image(fitz.Rect(40, 60, 80, 80), filename=str(source / "probe-source.png"))
    page.insert_image(fitz.Rect(100, 60, 130, 75), filename=str(source / "probe-source.png"))
    doc.save(tmp_path / "word-objects.pdf")
    doc.close()
    checks = module.verify_artifacts(tmp_path)
    assert checks["pdf_contains_images"] is False


def test_merge_verifier_rejects_extra_merge_topology():
    module = _module()
    from fixtures.microsoft_parity.word_object_artifact_verifier import W, verify_second_table_merges
    import xml.etree.ElementTree as ET
    source = ROOT / "docs/verification/microsoft-parity/macos-word-objects-probe/run-05/word-objects.docx"
    with ZipFile(source) as package:
        document = ET.fromstring(package.read("word/document.xml"))
    table = document.findall(".//" + W + "tbl")[1]
    rows = table.findall("./" + W + "tr")
    for row, value in ((rows[1], "restart"), (rows[2], None)):
        cell = row.findall("./" + W + "tc")[0]
        properties = cell.find("./" + W + "tcPr")
        merge = ET.SubElement(properties, W + "vMerge")
        if value:
            merge.set(W + "val", value)
    assert verify_second_table_merges(
        document, "Horizontal merge", "Vertical merge",
        ("M1 outside", "M2 outside", "M3 outside"),
    ) is False
