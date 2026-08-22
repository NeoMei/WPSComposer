from __future__ import annotations

import json
import zipfile
from pathlib import Path
from typing import Any, Optional

import pytest

from skills.WPSComposer.scripts.longform.executor import ExecutionOutcome, PaginationMap
from skills.WPSComposer.scripts.macos_probe.longform_evidence import (
    inspect_docx,
    run_longform_m2_evidence,
)
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult


class FakeBridge:
    """Recording bridge for unit tests."""

    def __init__(self, artifact_path: Path) -> None:
        self._artifact = artifact_path
        self._commands: list[dict[str, Any]] = []
        self._output_paths: dict[str, str] = {}
        self._registered: set[str] = set()
        self.url = "http://127.0.0.1:3889"
        self.token = "test-token"
        self.session_nonce = "test-nonce"

    def issue(self, component: str, method: str, params: dict[str, Any]):
        command_id = f"cmd-{len(self._commands)}"
        self._commands.append({"component": component, "method": method, "params": params, "id": command_id})
        self._output_paths[command_id] = str(params.get("outputPath", self._artifact))
        class Command:
            id = command_id
        return Command()

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        output_path = self._output_paths.get(command_id, str(self._artifact))
        _make_minimal_docx(Path(output_path))
        return ProbeResult(
            id=command_id,
            ok=True,
            value={"outputPath": output_path, "appliedOperations": 12, "issueCodes": [], "paginationMap": {"version": "M2-stub", "nodes": []}},
            error=None,
        )

    def wait_registered(self, expected: set[str], timeout: float) -> None:
        pass

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


class FakeRuntime:
    """Recording runtime for unit tests."""

    def __init__(self, staging_dir: Path) -> None:
        self.staging_dir = staging_dir
        self.logs: dict[str, Path] = {}
        self.registration_restored = True
        self._activated: list[str] = []
        self._profiles: list[str] = []
        self._servers_started = False

    def prepare_profiles(self) -> dict[str, Path]:
        self._profiles.append("writer")
        return {"writer": self.staging_dir}

    def start_servers(self) -> None:
        self._servers_started = True

    def activate_component(self, component: str) -> Path:
        self._activated.append(component)
        return self.staging_dir / f"{component}.docx"

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return False


def _make_minimal_docx(path: Path, text: str = "Hello") -> None:
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        package.writestr(
            "word/document.xml",
            f"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            f"<w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p>"
            f"<w:sectPr><w:pgNumType w:fmt=\"decimal\" w:start=\"1\"/></w:sectPr>"
            f"</w:body></w:document>",
        )
        package.writestr("word/styles.xml", "<w:styles />")


def test_inspect_docx_reads_structural_evidence(tmp_path: Path) -> None:
    docx = tmp_path / "test.docx"
    _make_minimal_docx(docx)
    result = inspect_docx(docx)
    assert result["sectPrCount"] == 1
    assert result["roles"]["unknown"][0]["fmt"] == "decimal"
    assert result["numberingXmlPresent"] is False
    assert result["contentTypes"] is True


def test_run_longform_m2_evidence_with_mocks(tmp_path: Path) -> None:
    fixtures_dir = tmp_path / "fixtures"
    fixtures_dir.mkdir()
    (fixtures_dir / "plain_short.md").write_text("# Test\n\nHello.\n", encoding="utf-8")

    output_dir = tmp_path / "output"
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    fake_artifact = staging_dir / "plain_short.docx"
    _make_minimal_docx(fake_artifact)

    bridge = FakeBridge(fake_artifact)
    runtime = FakeRuntime(staging_dir)

    def runtime_factory(probe_root, runtime_dir, url, token, node_override=None):
        return runtime

    def bridge_factory(origins):
        return bridge

    report_path = run_longform_m2_evidence(
        output_dir,
        timeout=1.0,
        fixtures_dir=fixtures_dir,
        bridge_factory=bridge_factory,
        runtime_factory=runtime_factory,
    )

    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["status"] == "passed"
    assert len(report["fixtures"]) == 1
    fixture = report["fixtures"][0]
    assert fixture["fixture"] == "plain_short"
    assert fixture["status"] == "passed"
    assert fixture["appliedOperations"] == 12
    assert Path(fixture["artifact"]).is_file()
    assert report["wpsVersion"] != "unknown"
    assert runtime._activated == ["writer"]
    assert runtime._servers_started is True


def test_run_longform_m2_evidence_records_pdf_conversion(tmp_path: Path) -> None:
    fixtures_dir = tmp_path / "fixtures"
    fixtures_dir.mkdir()
    (fixtures_dir / "academic.md").write_text("# Test\n\nHello.\n", encoding="utf-8")

    output_dir = tmp_path / "output"
    staging_dir = tmp_path / "staging"
    staging_dir.mkdir()
    fake_artifact = staging_dir / "academic.docx"
    _make_minimal_docx(fake_artifact)
    pdf_path = output_dir / "academic.pdf"
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    pdf_path.write_bytes(b"%PDF-1.4 test")

    def pdf_converter(docx: Path) -> Path:
        return pdf_path

    bridge = FakeBridge(fake_artifact)
    runtime = FakeRuntime(staging_dir)

    def runtime_factory(probe_root, runtime_dir, url, token, node_override=None):
        return runtime

    def bridge_factory(origins):
        return bridge

    report_path = run_longform_m2_evidence(
        output_dir,
        timeout=1.0,
        fixtures_dir=fixtures_dir,
        bridge_factory=bridge_factory,
        runtime_factory=runtime_factory,
        pdf_converter=pdf_converter,
    )

    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["status"] == "passed"
    assert report["pdfConversion"]["status"] == "passed"
    assert report["pdfConversion"]["fixture"] == "academic"
