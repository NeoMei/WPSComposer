from concurrent.futures import ThreadPoolExecutor
import json
from pathlib import Path
import shutil
import zipfile

import pytest

from skills.WPSComposer.scripts.artifact_transport import ArtifactTransportError
from skills.WPSComposer.scripts.macos_probe.bridge import LoopbackBridge
from skills.WPSComposer.scripts.macos_probe.models import PathPolicy, ProbeResult
from skills.WPSComposer.scripts.macos_probe import runner
from skills.WPSComposer.scripts.macos_probe.runner import (
    _execute_commands,
    _wait_for_component_registration,
    ArtifactError,
    validate_artifact,
    wait_for_artifact,
)


def test_validate_native_zip_artifacts(tmp_path: Path):
    members = {
        "docx": "word/document.xml",
        "pptx": "ppt/presentation.xml",
        "xlsx": "xl/workbook.xml",
    }
    for suffix, member in members.items():
        path = tmp_path / f"smoke.{suffix}"
        with zipfile.ZipFile(path, "w") as package:
            package.writestr("[Content_Types].xml", "<Types />")
            package.writestr(member, "x" * 2048)
        validate_artifact(path, suffix)


def test_runtime_allocator_returns_an_uncreated_child_directory():
    root, runtime_path = runner._allocate_runtime_dir()
    try:
        assert root.is_dir()
        assert runtime_path == root / "runtime"
        assert not runtime_path.exists()
    finally:
        shutil.rmtree(root)


def test_bridge_origins_match_wpsjs_sequential_default_ports():
    assert runner.ORIGINS == {
        "http://127.0.0.1:3889",
        "http://127.0.0.1:3890",
        "http://127.0.0.1:3891",
    }


def test_registration_retry_activates_only_missing_components():
    class FakeBridge:
        def __init__(self):
            self.waits = 0

        def wait_registered(self, expected, timeout):
            self.waits += 1
            if self.waits == 1:
                raise TimeoutError("writer missing")

        def registered_components(self):
            return {"presentation", "spreadsheet"}

    class FakeRuntime:
        def __init__(self):
            self.activated = []

        def activate_component(self, component):
            self.activated.append(component)

    bridge = FakeBridge()
    runtime_instance = FakeRuntime()

    _wait_for_component_registration(bridge, runtime_instance, timeout=2)

    assert runtime_instance.activated == ["writer"]
    assert bridge.waits == 2


def test_validate_pdf_signature(tmp_path: Path):
    path = tmp_path / "smoke.pdf"
    path.write_bytes(b"%PDF-1.7\n" + b"x" * 2048)
    validate_artifact(path, "pdf")


def test_validate_artifact_rejects_empty_or_wrong_signature(tmp_path: Path):
    path = tmp_path / "bad.docx"
    path.write_bytes(b"not a package" + b"x" * 2048)
    with pytest.raises(ArtifactError, match="Invalid DOCX ZIP package"):
        validate_artifact(path, "docx")


def test_wait_for_artifact_handles_delayed_writer_flush(tmp_path: Path):
    path = tmp_path / "delayed.docx"

    def write_later():
        import time

        time.sleep(0.05)
        with zipfile.ZipFile(path, "w") as package:
            package.writestr("[Content_Types].xml", "<Types />")
            package.writestr("word/document.xml", "x" * 2048)

    with ThreadPoolExecutor(max_workers=1) as pool:
        future = pool.submit(write_later)
        wait_for_artifact(path, "docx", timeout=1)
        future.result(timeout=1)


def fake_component(bridge, component, command_count):
    bridge.state.register(component)
    for _ in range(command_count):
        command = bridge.state.next(component, 2)
        assert command is not None
        if command.method == "probe_capabilities":
            value = {
                "component": component,
                "applicationName": f"Fake {component}",
                "capabilities": {
                    f"{component}.probe": {
                        "classification": "native",
                        "detail": "fake",
                    }
                },
            }
        else:
            if component == "writer":
                assert command.params["imageUrl"] == runner.WRITER_IMAGE_URL
            path = Path(command.params["outputPath"])
            path.parent.mkdir(parents=True, exist_ok=True)
            if path.suffix == ".pdf":
                path.write_bytes(b"%PDF-1.7\n" + b"x" * 2048)
            else:
                members = {
                    ".docx": "word/document.xml",
                    ".pptx": "ppt/presentation.xml",
                    ".xlsx": "xl/workbook.xml",
                }
                with zipfile.ZipFile(path, "w") as package:
                    package.writestr("[Content_Types].xml", "<Types />")
                    package.writestr(members[path.suffix], "x" * 2048)
            value = {"path": str(path), "capabilities": {}}
            if path.suffix == ".pdf":
                value["preset"] = "obsidian-reading"
        bridge.state.complete(ProbeResult(command.id, True, value, None))


def test_execute_commands_keeps_four_outputs_independent(tmp_path: Path):
    staging = tmp_path / "container" / "session"
    staging.mkdir(parents=True)
    output = tmp_path / "output"
    output.mkdir()
    image = tmp_path / "fixture.png"
    image.write_bytes(b"png")
    policy = PathPolicy((tmp_path, output))
    origins = {
        "http://127.0.0.1:3891",
        "http://127.0.0.1:3892",
        "http://127.0.0.1:3893",
    }

    with LoopbackBridge(origins) as bridge, ThreadPoolExecutor(
        max_workers=3
    ) as pool:
        futures = (
            pool.submit(fake_component, bridge, "writer", 3),
            pool.submit(fake_component, bridge, "presentation", 2),
            pool.submit(fake_component, bridge, "spreadsheet", 2),
        )
        bridge.wait_registered(
            {"writer", "presentation", "spreadsheet"}, 2
        )
        report = _execute_commands(
            bridge, policy, staging, output, image, timeout=2
        )
        for future in futures:
            future.result(timeout=2)

    assert [item["format"] for item in report["artifacts"]] == [
        "docx",
        "pptx",
        "xlsx",
        "pdf",
    ]
    assert report["status"] == "passed"
    assert report["pdfPreset"] == "obsidian-reading"
    assert report["artifactTransport"] == "wps-container-staging"
    assert {Path(item["path"]).parent for item in report["artifacts"]} == {
        output.resolve()
    }
    assert str(staging.resolve()) not in json.dumps(report)
    for _, _, filename, _ in runner.OUTPUT_COMMANDS:
        assert not (staging / filename).exists()
    assert "temporaryDocx" not in report


def test_execute_commands_reports_typed_publication_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    staging = tmp_path / "container" / "session"
    staging.mkdir(parents=True)
    output = tmp_path / "output"
    output.mkdir()
    image = tmp_path / "fixture.png"
    image.write_bytes(b"png")
    policy = PathPolicy((tmp_path, output))
    origins = {
        "http://127.0.0.1:3891",
        "http://127.0.0.1:3892",
        "http://127.0.0.1:3893",
    }
    original_publish = runner.publish_artifact
    attempts = 0

    def fail_once(*args, **kwargs):
        nonlocal attempts
        attempts += 1
        if attempts == 1:
            raise ArtifactTransportError(
                "ARTIFACT_PUBLISH_FAILED", "destination rejected"
            )
        return original_publish(*args, **kwargs)

    monkeypatch.setattr(runner, "publish_artifact", fail_once)
    with LoopbackBridge(origins) as bridge, ThreadPoolExecutor(
        max_workers=3
    ) as pool:
        futures = (
            pool.submit(fake_component, bridge, "writer", 3),
            pool.submit(fake_component, bridge, "presentation", 2),
            pool.submit(fake_component, bridge, "spreadsheet", 2),
        )
        bridge.wait_registered({"writer", "presentation", "spreadsheet"}, 2)
        report = _execute_commands(
            bridge, policy, staging, output, image, timeout=2
        )
        for future in futures:
            future.result(timeout=2)

    assert report["status"] == "failed"
    assert report["failures"][0]["code"] == "ARTIFACT_PUBLISH_FAILED"


def test_publish_writer_image_copies_only_into_writer_profile(tmp_path: Path):
    source = tmp_path / "source.png"
    source.write_bytes(b"png-data")
    writer = tmp_path / "profiles" / "writer"
    writer.mkdir(parents=True)

    url = runner._publish_writer_image({"writer": writer}, source)

    assert url == "http://127.0.0.1:3889/fixture.png"
    assert (writer / "fixture.png").read_bytes() == b"png-data"
