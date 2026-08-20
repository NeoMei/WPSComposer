from __future__ import annotations

import json
from pathlib import Path
import shutil
from types import SimpleNamespace
import zipfile

import pytest
from pypdf import PdfWriter

from skills.WPSComposer.scripts.macos_probe.models import ProbeResult
from skills.WPSComposer.scripts.longform_m0 import host_checks
from skills.WPSComposer.scripts.longform_m0.macos import (
    EMPTY_MANIFEST,
    MacosM0Failed,
    run_macos_probe,
)


def write_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        package.writestr(
            "word/document.xml",
            "<document>" + ("x" * 2048) + "</document>",
        )


def write_pdf(path: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)
    with path.open("wb") as stream:
        writer.write(stream)


def capability_rows(status: str = "passed") -> list[dict]:
    rows = []
    for capability_id in range(1, 16):
        checks = ["native"]
        if capability_id >= 3:
            checks.extend(("reopened", "refreshed"))
        if capability_id == 2:
            checks = ["not-applicable-macos"]
        rows.append(
            {
                "id": capability_id,
                "status": status,
                "checks": checks,
                "metrics": {},
            }
        )
    return rows


class FakeBridge:
    instances = []
    mode = "success"
    capabilities = capability_rows()

    def __init__(self, origins):
        self.origins = origins
        self.url = "http://127.0.0.1:45678"
        self.token = "private-token"
        self.params = None
        self.closed = False
        type(self).instances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.closed = True

    def wait_registered(self, expected, timeout):
        assert expected == {"writer"}
        assert timeout > 0

    def issue(self, component, method, params):
        assert component == "writer"
        assert method == "probe_longform_m0"
        self.params = params
        return SimpleNamespace(id="command-1")

    def wait_result(self, command_id, timeout):
        assert command_id == "command-1"
        assert timeout > 0
        if self.mode == "timeout":
            raise TimeoutError("native command timed out")
        if self.mode == "interrupt":
            raise KeyboardInterrupt()
        if self.mode == "remote-failure":
            return ProbeResult(
                command_id,
                False,
                {},
                {"code": "PROTOCOL_MISMATCH", "message": "bad protocol"},
            )
        assert self.params is not None
        docx = Path(self.params["stagedDocxPath"])
        pdf = Path(self.params["stagedPdfPath"])
        write_docx(docx)
        write_pdf(pdf)
        docx_result = str(docx)
        if self.mode == "wrong-path":
            docx_result = str(docx.parent / "other.docx")
        capabilities = json.loads(json.dumps(self.capabilities))
        return ProbeResult(
            command_id,
            True,
            {
                "probeVersion": "0.8.0-m0.1",
                "protocolVersion": 2,
                "resourceManifestVersion": 1,
                "platform": "macos",
                "wpsVersion": "12.1.fake",
                "capabilities": capabilities,
                "failures": [],
                "docxPath": docx_result,
                "pdfPath": str(pdf),
            },
            None,
        )


class FakeRuntime:
    instances = []

    def __init__(self, probe_root, runtime_dir, bridge_url, token, deadline=None):
        self.probe_root = probe_root
        self.runtime_dir = runtime_dir
        self.bridge_url = bridge_url
        self.token = token
        self.deadline = deadline
        self.staging_dir = None
        self.registration_restored = True
        self.prepared = False
        self.started = False
        self.activated = False
        self.exited = False
        type(self).instances.append(self)

    def __enter__(self):
        self.runtime_dir.mkdir(parents=True)
        self.staging_dir = self.runtime_dir / "staging"
        self.staging_dir.mkdir()
        return self

    def __exit__(self, exc_type, exc, traceback):
        assert self.staging_dir is not None
        shutil.rmtree(self.staging_dir)
        self.registration_restored = True
        self.exited = True

    def prepare_profiles(self):
        self.prepared = True

    def start_servers(self, deadline=None):
        assert deadline == self.deadline
        self.started = True

    def activate_components(self):
        self.activated = True


@pytest.fixture(autouse=True)
def reset_fakes():
    FakeBridge.instances = []
    FakeBridge.mode = "success"
    FakeBridge.capabilities = capability_rows()
    FakeRuntime.instances = []


def run_fake(tmp_path: Path) -> Path:
    return run_macos_probe(
        tmp_path / "evidence",
        timeout=30,
        bridge_factory=FakeBridge,
        runtime_factory=FakeRuntime,
        version_reader=lambda: "12.1.fake",
    )


def test_empty_manifest_is_canonical_and_digest_bound():
    assert EMPTY_MANIFEST == {
        "version": 1,
        "entries": [],
        "digest": "a6a20076da005b27c9afc3a5d5b2457798c0ac817d1abc38b2fee4398ac3f133",
    }


def test_dependency_failure_writes_evidence_without_starting_runtime(
    monkeypatch, tmp_path: Path
):
    monkeypatch.setattr(
        host_checks,
        "_find_missing_dependencies",
        lambda: ["pypdf"],
    )

    with pytest.raises(MacosM0Failed) as caught:
        run_fake(tmp_path)

    assert FakeRuntime.instances == []
    evidence = json.loads(caught.value.evidence_path.read_text(encoding="utf-8"))
    assert evidence["artifacts"] == {}
    assert evidence["failures"][0]["code"] == "DEPENDENCY_UNAVAILABLE"


def test_success_publishes_two_artifacts_and_restores_runtime(tmp_path: Path):
    evidence_path = run_fake(tmp_path)

    evidence = json.loads(evidence_path.read_text(encoding="utf-8"))
    assert evidence["platform"] == "macos"
    assert evidence["artifacts"]["docx"]["name"] == "probe.docx"
    assert evidence["artifacts"]["pdf"]["name"] == "probe.pdf"
    assert (evidence_path.parent / "probe.docx").is_file()
    assert (evidence_path.parent / "probe.pdf").is_file()
    assert evidence["capabilities"][4]["metrics"]["hostPdfSnapshot"][
        "pageCount"
    ] == 1
    runtime = FakeRuntime.instances[0]
    assert runtime.prepared and runtime.started and runtime.activated
    assert runtime.exited and runtime.registration_restored
    assert runtime.staging_dir is not None and not runtime.staging_dir.exists()
    bridge = FakeBridge.instances[0]
    assert bridge.closed
    assert bridge.params["manifest"] == EMPTY_MANIFEST


@pytest.mark.parametrize(
    ("mode", "code"),
    [
        ("timeout", "COMMAND_TIMEOUT"),
        ("remote-failure", "PROTOCOL_MISMATCH"),
        ("wrong-path", "PROTOCOL_ERROR"),
    ],
)
def test_native_failure_keeps_only_redacted_failure_evidence(
    tmp_path: Path, mode: str, code: str
):
    FakeBridge.mode = mode

    with pytest.raises(MacosM0Failed) as caught:
        run_fake(tmp_path)

    evidence = json.loads(caught.value.evidence_path.read_text(encoding="utf-8"))
    assert evidence["artifacts"] == {}
    assert evidence["failures"] == [
        {"code": code, "message": "macOS long-form M0 probe failed"}
    ]
    assert not (caught.value.evidence_path.parent / "probe.docx").exists()
    assert not (caught.value.evidence_path.parent / "probe.pdf").exists()
    assert "/Users/" not in caught.value.evidence_path.read_text(encoding="utf-8")
    assert FakeRuntime.instances[0].exited


def test_private_remote_metrics_fail_closed(tmp_path: Path):
    FakeBridge.capabilities[4]["metrics"] = {
        "sourcePath": "/Users/alice/private.docx"
    }

    with pytest.raises(MacosM0Failed) as caught:
        run_fake(tmp_path)

    evidence = json.loads(caught.value.evidence_path.read_text(encoding="utf-8"))
    assert evidence["failures"][0]["code"] == "EVIDENCE_INVALID"
    assert evidence["artifacts"] == {}


def test_keyboard_interrupt_propagates_after_runtime_cleanup(tmp_path: Path):
    FakeBridge.mode = "interrupt"
    evidence_path = tmp_path / "evidence" / "platform-evidence.json"

    with pytest.raises(KeyboardInterrupt):
        run_fake(tmp_path)

    assert FakeRuntime.instances[0].exited
    assert not evidence_path.exists()
