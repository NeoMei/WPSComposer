from __future__ import annotations

from pathlib import Path
import shutil
from types import SimpleNamespace
import zipfile

import pytest

from skills.WPSComposer.scripts.macos_probe import inspection
from skills.WPSComposer.scripts.macos_probe.inspection import (
    InspectionError,
    _run_edit,
    edit_macos,
)
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult


def _write_pptx(path: Path, marker: str = "valid") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", "<Types />")
        package.writestr(
            "ppt/presentation.xml",
            f"<presentation><marker>{marker}</marker></presentation>",
        )
    return path


class FakeRuntime:
    def __init__(self, staging_dir: Path):
        self.staging_dir = staging_dir.resolve()
        self.staging_dir.mkdir(parents=True)

    def prepare_profiles(self):
        pass

    def start_servers(self, *, deadline):
        pass

    def activate_component(self, component, *, deadline):
        pass


class FakeBridge:
    def __init__(
        self,
        *,
        reports=None,
        returned_path="expected",
        corrupt_output=False,
    ):
        self.reports = reports if reports is not None else [
            {"target": "slide:1", "accepted": ["name"], "rejected": [], "ok": True}
        ]
        self.returned_path = returned_path
        self.corrupt_output = corrupt_output
        self.commands = []
        self.state = SimpleNamespace(cancel=lambda command_id: None)

    def wait_registered(self, expected, timeout):
        pass

    def issue(self, component, method, params):
        command = SimpleNamespace(
            id="command-1", component=component, method=method, params=dict(params)
        )
        self.commands.append(command)
        return command

    def wait_result(self, command_id, timeout):
        params = self.commands[0].params
        expected = Path(params["outputPath"])
        if self.corrupt_output:
            expected.write_bytes(b"not-an-ooxml-package")
        else:
            _write_pptx(expected, "edited")
        if self.returned_path == "expected":
            reported = expected
        elif self.returned_path == "source":
            reported = Path(params["sourcePath"])
        else:
            reported = Path(self.returned_path)
        return ProbeResult(
            command_id,
            True,
            {"path": str(reported), "patches": self.reports},
            None,
        )


def _run(
    tmp_path: Path,
    *,
    bridge=None,
    output_exists=False,
    overwrite=False,
    atomic=True,
    raise_on_error=False,
    patches=None,
):
    source = _write_pptx(tmp_path / "source.pptx", "source")
    output = tmp_path / "nested" / "output.pptx"
    if output_exists:
        _write_pptx(output, "original-output")
    runtime = FakeRuntime(tmp_path / "container" / "session")
    fake_bridge = bridge or FakeBridge()
    result = _run_edit(
        source,
        output,
        "presentation",
        "edit_presentation",
        patches or [{"target": "slide:1", "name": "Updated"}],
        fake_bridge,
        runtime,
        inspection.time.monotonic() + 2,
        atomic=atomic,
        raise_on_error=raise_on_error,
        overwrite=overwrite,
    )
    return result, source, output, fake_bridge


@pytest.mark.parametrize("suffix", ["ppt", "pptm", "pps", "ppsx", "ppsm"])
def test_macos_edit_rejects_unverified_presentation_inputs(tmp_path: Path, suffix: str):
    source = tmp_path / f"legacy.{suffix}"
    source.write_bytes(b"presentation")

    with pytest.raises(ValueError, match="not yet supported"):
        edit_macos(source, [{"target": "slide:1", "name": "Updated"}])


def test_macos_edit_requires_pptx_output(tmp_path: Path):
    source = _write_pptx(tmp_path / "source.pptx")

    with pytest.raises(ValueError, match="output must use '.pptx'"):
        edit_macos(
            source,
            [{"target": "slide:1", "name": "Updated"}],
            output=tmp_path / "output.pptm",
        )


def test_macos_edit_refuses_existing_output_without_overwrite(tmp_path: Path):
    source = _write_pptx(tmp_path / "source.pptx")
    output = _write_pptx(tmp_path / "output.pptx", "approved")
    original = output.read_bytes()

    with pytest.raises(FileExistsError, match="Output already exists"):
        edit_macos(
            source,
            [{"target": "slide:1", "name": "Updated"}],
            output=output,
        )

    assert output.read_bytes() == original


def test_run_edit_requires_exact_expected_reported_path(tmp_path: Path):
    bridge = FakeBridge(returned_path="source")

    with pytest.raises(InspectionError) as caught:
        _run(tmp_path, bridge=bridge)

    assert caught.value.code == "PROTOCOL_ERROR"


def test_run_edit_rejects_corrupt_staged_package_without_publishing(tmp_path: Path):
    bridge = FakeBridge(corrupt_output=True)

    with pytest.raises(InspectionError) as caught:
        _run(tmp_path, bridge=bridge)

    assert caught.value.code == "STAGED_ARTIFACT_INVALID"
    assert not (tmp_path / "nested" / "output.pptx").exists()


def test_run_edit_atomic_rejection_never_publishes(tmp_path: Path):
    reports = [
        {"target": "slide:1", "accepted": ["name"], "rejected": [], "ok": True},
        {"target": "slide:404", "accepted": [], "rejected": [], "ok": False,
         "error": "slide not found"},
    ]
    bridge = FakeBridge(reports=reports)

    result, _source, output, _bridge = _run(
        tmp_path,
        bridge=bridge,
        atomic=True,
        patches=[
            {"target": "slide:1", "name": "Updated"},
            {"target": "slide:404", "name": "Missing"},
        ],
    )

    assert result["saved"] is False
    assert result["path"] is None
    assert result["patches"] == reports
    assert not output.exists()


def test_run_edit_atomic_rejected_fields_override_inconsistent_ok(tmp_path: Path):
    reports = [
        {"target": "slide:1", "accepted": [], "rejected": ["font.size"], "ok": True},
    ]

    result, _source, output, _bridge = _run(
        tmp_path,
        bridge=FakeBridge(reports=reports),
        atomic=True,
    )

    assert result["saved"] is False
    assert not output.exists()


def test_run_edit_rejects_missing_patch_reports_without_publishing(tmp_path: Path):
    bridge = FakeBridge(reports=[])

    with pytest.raises(InspectionError) as caught:
        _run(tmp_path, bridge=bridge, atomic=True)

    assert caught.value.code == "PROTOCOL_ERROR"
    assert not (tmp_path / "nested" / "output.pptx").exists()


def test_run_edit_passes_atomic_and_raise_on_error_to_bridge(tmp_path: Path):
    result, _source, output, bridge = _run(
        tmp_path,
        atomic=False,
        raise_on_error=True,
    )

    assert result["saved"] is True
    assert Path(result["path"]) == output.resolve()
    assert bridge.commands[0].params["atomic"] is False
    assert bridge.commands[0].params["raiseOnError"] is True


def test_run_edit_overwrite_failure_preserves_original_destination(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    original_publish = inspection.publish_artifact

    def fail_after_validation(*args, **kwargs):
        kwargs["validator"](Path(args[0]))
        raise OSError("simulated destination I/O failure")

    monkeypatch.setattr(inspection, "publish_artifact", fail_after_validation)
    source = _write_pptx(tmp_path / "source.pptx", "source")
    output = _write_pptx(tmp_path / "output.pptx", "approved")
    original = output.read_bytes()
    runtime = FakeRuntime(tmp_path / "container" / "session")

    with pytest.raises(InspectionError):
        _run_edit(
            source,
            output,
            "presentation",
            "edit_presentation",
            [{"target": "slide:1", "name": "Updated"}],
            FakeBridge(),
            runtime,
            inspection.time.monotonic() + 2,
            atomic=True,
            raise_on_error=False,
            overwrite=True,
        )

    assert output.read_bytes() == original
    monkeypatch.setattr(inspection, "publish_artifact", original_publish)


def test_presentation_atomic_failure_guard_precedes_save_as():
    source = (
        Path(__file__).resolve().parents[2]
        / "macos/wps-jsapi-probe/addin/presentation.js"
    ).read_text()
    guard = source.index("if (atomic && hasFailures)")
    save = source.index("presentation.SaveAs(outputPath, 24)", guard)

    assert guard < save
