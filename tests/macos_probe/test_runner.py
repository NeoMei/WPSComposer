from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.macos_probe.bridge import LoopbackBridge
from skills.WPSComposer.scripts.macos_probe.models import PathPolicy, ProbeResult
from skills.WPSComposer.scripts.macos_probe.runner import (
    _execute_commands,
    ArtifactError,
    validate_artifact,
)


def test_validate_native_zip_artifacts(tmp_path: Path):
    for suffix in ("docx", "pptx", "xlsx"):
        path = tmp_path / f"smoke.{suffix}"
        path.write_bytes(b"PK\x03\x04" + b"x" * 2048)
        validate_artifact(path, suffix)


def test_validate_pdf_signature(tmp_path: Path):
    path = tmp_path / "smoke.pdf"
    path.write_bytes(b"%PDF-1.7\n" + b"x" * 2048)
    validate_artifact(path, "pdf")


def test_validate_artifact_rejects_empty_or_wrong_signature(tmp_path: Path):
    path = tmp_path / "bad.docx"
    path.write_bytes(b"not a package")
    with pytest.raises(ArtifactError, match="invalid DOCX signature"):
        validate_artifact(path, "docx")


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
            path = Path(command.params["outputPath"])
            signature = (
                b"%PDF-1.7\n" if path.suffix == ".pdf" else b"PK\x03\x04"
            )
            path.write_bytes(signature + b"x" * 2048)
            value = {"path": str(path), "capabilities": {}}
            if path.suffix == ".pdf":
                value["preset"] = "obsidian-reading"
        bridge.state.complete(ProbeResult(command.id, True, value, None))


def test_execute_commands_keeps_four_outputs_independent(tmp_path: Path):
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
            bridge, policy, output, image, timeout=2
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
    assert "temporaryDocx" not in report
