from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pytest

from skills.WPSComposer.scripts import orchestrator
from skills.WPSComposer.scripts.longform.quality import GenerationOutcome


def test_longform_preserves_native_numbering_after_validated_publication(monkeypatch, tmp_path):
    """The legacy post-pass must not rewrite M5's already validated list links."""
    native = tmp_path / "native.docx"
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    with ZipFile(native, "w") as archive:
        archive.writestr(
            "word/document.xml",
            f'<w:document xmlns:w="{w}"><w:body><w:p>'
            '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'
            '<w:r><w:t>方案定位</w:t></w:r></w:p></w:body></w:document>',
        )
        archive.writestr(
            "word/styles.xml",
            f'<w:styles xmlns:w="{w}"><w:style w:type="paragraph" w:styleId="Heading1">'
            '<w:name w:val="Heading 1"/><w:pPr><w:numPr>'
            '<w:ilvl w:val="0"/><w:numId w:val="7"/>'
            '</w:numPr></w:pPr></w:style></w:styles>',
        )
        archive.writestr(
            "word/numbering.xml",
            f'<w:numbering xmlns:w="{w}"><w:abstractNum w:abstractNumId="3">'
            '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="chineseCounting"/>'
            '<w:lvlText w:val="%1、"/></w:lvl></w:abstractNum>'
            '<w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num></w:numbering>',
        )
        archive.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        )
    validated_bytes = native.read_bytes()

    def native_generation(build, format_name, output, timeout, overwrite):
        Path(output).write_bytes(validated_bytes)
        return GenerationOutcome(str(output))

    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    monkeypatch.setattr(orchestrator, "resolve_engine", lambda *args: "wps")
    monkeypatch.setattr(orchestrator, "_generate_longform_outcome", native_generation)
    output = tmp_path / "report.docx"
    orchestrator.generate(
        "# 方案\n\n## 一、方案定位\n\n正文。",
        source_is_text=True,
        output=str(output),
    )
    with ZipFile(output) as archive:
        styles = ET.fromstring(archive.read("word/styles.xml"))
    assert styles.find(f".//{{{w}}}numId").get(f"{{{w}}}val") == "7"
    assert output.read_bytes() == validated_bytes


def test_docx_defaults_to_longform_private_outcome(monkeypatch, tmp_path):
    calls = []

    def fake(build, format_name, output, timeout, overwrite):
        calls.append((build, format_name, output, timeout, overwrite))
        Path(output).write_bytes(b"PK\x03\x04" + b"x" * 300)
        return GenerationOutcome(str(Path(output).resolve()))

    monkeypatch.setattr(orchestrator, "_generate_longform_outcome", fake)
    output = tmp_path / "report.docx"
    result = orchestrator.generate(
        "# Report\n\nBody", source_is_text=True, output=str(output)
    )
    assert result == str(output.resolve())
    assert calls[0][1:] == ("docx", output.resolve(), 600, False)
    assert calls[0][0].semantic.config.layout_engine == "longform"


def test_pdf_defaults_to_longform_and_returns_only_requested_artifact(monkeypatch, tmp_path):
    def fake(build, format_name, output, timeout, overwrite):
        assert format_name == "pdf"
        Path(output).write_bytes(b"%PDF-1.4" + b"x" * 300)
        return GenerationOutcome(str(Path(output).resolve()))

    monkeypatch.setattr(orchestrator, "_generate_longform_outcome", fake)
    output = tmp_path / "report.pdf"
    assert orchestrator.generate(
        "# Report", format="pdf", source_is_text=True, output=str(output)
    ) == str(output.resolve())
    assert not (tmp_path / "report.docx").exists()


def test_explicit_legacy_uses_old_writer_route(monkeypatch, tmp_path):
    longform_calls = []
    legacy_calls = []
    monkeypatch.setattr(
        orchestrator,
        "_generate_longform_outcome",
        lambda *args: longform_calls.append(args),
    )

    def fake_legacy(doc, format_name, output, preset, **kwargs):
        legacy_calls.append((format_name, output, kwargs))
        Path(output).write_bytes(b"PK\x03\x04" + b"x" * 300)
        return output

    monkeypatch.setattr(orchestrator, "generate_macos", fake_legacy)
    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    output = tmp_path / "legacy.docx"
    markdown = "---\nlayout_engine: legacy\n---\n# Legacy\n"
    result = orchestrator.generate(
        markdown, source_is_text=True, output=str(output)
    )
    assert result == str(output.resolve())
    assert longform_calls == []
    assert legacy_calls[0][0] == "docx"


def test_public_preset_override_reaches_longform_plan(monkeypatch, tmp_path):
    captured = []

    def fake(build, format_name, output, timeout, overwrite):
        captured.append(build)
        Path(output).write_bytes(b"PK\x03\x04" + b"x" * 300)
        return GenerationOutcome(str(Path(output).resolve()))

    monkeypatch.setattr(orchestrator, "_generate_longform_outcome", fake)
    orchestrator.generate(
        "# Report",
        source_is_text=True,
        preset="business",
        output=str(tmp_path / "business.docx"),
    )
    assert captured[0].semantic.document.metadata["design"] == "business"


@pytest.mark.parametrize("format_name", ["pptx", "xlsx"])
def test_pptx_xlsx_do_not_enter_longform(monkeypatch, tmp_path, format_name):
    monkeypatch.setattr(
        orchestrator,
        "_generate_longform_outcome",
        lambda *args: pytest.fail("non-writer format entered longform"),
    )

    def fake_legacy(doc, actual_format, output, preset, **kwargs):
        assert actual_format == format_name
        Path(output).write_bytes(b"PK\x03\x04" + b"x" * 300)
        return output

    monkeypatch.setattr(orchestrator, "generate_macos", fake_legacy)
    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    output = tmp_path / f"artifact.{format_name}"
    orchestrator.generate(
        "# Content", format=format_name, source_is_text=True, output=str(output)
    )
