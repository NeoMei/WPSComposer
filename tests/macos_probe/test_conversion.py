from __future__ import annotations

from pathlib import Path
import shutil
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.conversion import (
    ConversionError,
    ConversionRequest,
)
from skills.WPSComposer.scripts.macos_probe import conversion as mac_conversion
from skills.WPSComposer.scripts.macos_probe.conversion import convert_macos
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult


class FakeBridge:
    def __init__(
        self,
        *,
        result_error=None,
        error_code="CONVERSION_COMMAND_FAILED",
        returned_path=None,
    ):
        self.url = "http://127.0.0.1:45678"
        self.token = "token"
        self.result_error = result_error
        self.error_code = error_code
        self.returned_path = returned_path
        self.commands = []
        self.registrations = []
        self.source_bytes = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return None

    def wait_registered(self, expected, timeout):
        self.registrations.append((set(expected), timeout))

    def registered_components(self):
        return self.registrations[-1][0] if self.registrations else set()

    def issue(self, component, method, params):
        command = SimpleNamespace(
            id=f"command-{len(self.commands) + 1}",
            component=component,
            method=method,
            params=dict(params),
        )
        self.commands.append(command)
        return command

    def wait_result(self, command_id, timeout):
        command = next(item for item in self.commands if item.id == command_id)
        self.source_bytes = Path(command.params["sourcePath"]).read_bytes()
        if self.result_error is not None:
            output = Path(command.params["outputPath"])
            message = self.result_error.replace(
                "{staging}", str(output.parent)
            )
            return ProbeResult(
                command_id,
                False,
                {},
                {"code": self.error_code, "message": message},
            )
        output = Path(command.params["outputPath"])
        output.write_bytes(b"%PDF-1.7\n" + b"x" * 2048)
        path = output if self.returned_path is None else self.returned_path
        return ProbeResult(command_id, True, {"path": str(path)}, None)


class FakeRuntime:
    def __init__(self, staging_dir: Path, calls: list):
        self.staging_dir = staging_dir.resolve()
        self.calls = calls
        self.registration_restored = True

    def __enter__(self):
        self.staging_dir.mkdir(parents=True)
        self.calls.append(("enter", self.staging_dir))
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.calls.append(("close",))
        shutil.rmtree(self.staging_dir, ignore_errors=True)

    def prepare_profiles(self):
        self.calls.append(("prepare_profiles",))

    def start_servers(self):
        self.calls.append(("start_servers",))

    def activate_component(self, component):
        self.calls.append(("activate_component", component))


def _request(tmp_path: Path, suffix: str, component: str) -> ConversionRequest:
    source = tmp_path / f"input.{suffix}"
    source.write_bytes(b"office-source")
    return ConversionRequest(
        source=source.resolve(),
        output=(tmp_path / "output" / f"{component}.pdf").resolve(),
        component=component,
        overwrite=False,
    )


def _run_with_fakes(
    request: ConversionRequest,
    tmp_path: Path,
    *,
    bridge: FakeBridge | None = None,
):
    calls = []
    fake_bridge = bridge or FakeBridge()
    runtime = FakeRuntime(tmp_path / "container" / "session-1", calls)
    result = convert_macos(
        request,
        enabled=True,
        bridge_factory=lambda origins: fake_bridge,
        runtime_factory=lambda *args, **kwargs: runtime,
        timeout=2,
    )
    return result, fake_bridge, calls, runtime


@pytest.mark.parametrize(
    ("suffix", "component", "method"),
    [
        ("docx", "writer", "convert_writer_pdf"),
        ("xlsx", "spreadsheet", "convert_workbook_pdf"),
        ("pptx", "presentation", "convert_presentation_pdf"),
    ],
)
def test_macos_copies_source_and_issues_only_staged_paths(
    tmp_path: Path, suffix: str, component: str, method: str
):
    request = _request(tmp_path, suffix, component)

    result, bridge, calls, runtime = _run_with_fakes(request, tmp_path)

    command = bridge.commands[0]
    staged_source = Path(command.params["sourcePath"])
    staged_output = Path(command.params["outputPath"])
    assert command.component == component
    assert command.method == method
    assert staged_source.parent == runtime.staging_dir
    assert staged_output.parent == runtime.staging_dir
    assert bridge.source_bytes == b"office-source"
    assert result == request.output
    assert result.is_file()
    assert not runtime.staging_dir.exists()
    assert ("activate_component", component) in calls


def test_macos_real_conversion_gate_is_enabled():
    assert mac_conversion.MACOS_CONVERSION_ENABLED is True


def test_macos_gate_can_be_explicitly_disabled(tmp_path: Path):
    request = _request(tmp_path, "docx", "writer")
    with pytest.raises(ConversionError) as caught:
        convert_macos(request, enabled=False)
    assert caught.value.code == "MACOS_GATE_NOT_PASSED"
    assert caught.value.backend == "mac-wps-jsapi"


def test_macos_command_failure_publishes_no_partial_pdf(tmp_path: Path):
    request = _request(tmp_path, "xlsx", "spreadsheet")
    bridge = FakeBridge(
        result_error="password or modal dialog required",
        error_code="INTERACTIVE_INPUT_REQUIRED",
    )

    with pytest.raises(ConversionError) as caught:
        _run_with_fakes(request, tmp_path, bridge=bridge)

    assert caught.value.code == "INTERACTIVE_INPUT_REQUIRED"
    assert caught.value.message == "password or modal dialog required"
    assert not request.output.exists()


def test_macos_rejects_result_path_outside_staging(tmp_path: Path):
    request = _request(tmp_path, "ppt", "presentation")
    bridge = FakeBridge(returned_path=tmp_path / "outside.pdf")

    with pytest.raises(ConversionError) as caught:
        _run_with_fakes(request, tmp_path, bridge=bridge)

    assert caught.value.code == "PROTOCOL_ERROR"
    assert not request.output.exists()


def test_macos_redacts_staging_path_from_command_error(tmp_path: Path):
    request = _request(tmp_path, "docx", "writer")
    bridge = FakeBridge(result_error="failed at {staging}/converted.pdf")

    with pytest.raises(ConversionError) as caught:
        _run_with_fakes(request, tmp_path, bridge=bridge)

    assert str((tmp_path / "container").resolve()) not in caught.value.message
    assert "<wps-staging>/converted.pdf" in caught.value.message


def test_registration_restore_failure_retains_durable_recovery(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    request = _request(tmp_path, "pptx", "presentation")
    runtime_root = tmp_path / "durable-runtime"
    runtime_root.mkdir()
    monkeypatch.setattr(
        mac_conversion.tempfile,
        "mkdtemp",
        lambda *args, **kwargs: str(runtime_root),
    )

    class RecoveryRuntime(FakeRuntime):
        def __init__(self, probe_root, runtime_dir, bridge_url, token):
            super().__init__(tmp_path / "container" / "session", [])
            self.runtime_dir = Path(runtime_dir)

        def __enter__(self):
            super().__enter__()
            recovery = self.runtime_dir / "registration-recovery"
            recovery.mkdir(parents=True)
            (recovery / "publish.xml.original").write_bytes(b"original")
            return self

        def __exit__(self, exc_type, exc, traceback):
            self.registration_restored = False
            shutil.rmtree(self.staging_dir, ignore_errors=True)
            raise RuntimeError("registration restore failed")

    with pytest.raises(ConversionError) as caught:
        convert_macos(
            request,
            enabled=True,
            bridge_factory=lambda origins: FakeBridge(),
            runtime_factory=RecoveryRuntime,
            timeout=2,
        )

    recovery = runtime_root / "runtime" / "registration-recovery"
    assert caught.value.code == "REGISTRATION_RESTORE_FAILED"
    assert str(recovery) in caught.value.message
    assert (recovery / "publish.xml.original").read_bytes() == b"original"
