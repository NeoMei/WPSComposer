from __future__ import annotations

from pathlib import Path
import shutil
from types import SimpleNamespace
import zipfile

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    GenerationOperation,
    GenerationPlan,
    GenerationResource,
    RecordedGeneration,
)
from skills.WPSComposer.scripts.macos_probe import generation as mac_generation
from skills.WPSComposer.scripts.macos_probe.generation import (
    GenerationError,
    GenerationRequest,
    execute_generation_plan,
)
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult


MARKER = "IN_PLACE_MARKER"
WRITER_MARKER_PLAN = GenerationPlan(
    "writer",
    (
        GenerationOperation("writer.reset", {}),
        GenerationOperation(
            "writer.add_paragraph",
            {"text": MARKER, "style": "Body Text"},
        ),
    ),
)
SHEET_MARKER_PLAN = GenerationPlan(
    "spreadsheet",
    (
        GenerationOperation("sheet.reset", {}),
        GenerationOperation(
            "sheet.write_table",
            {"startRow": 1, "startCol": 1, "values": [[MARKER]]},
        ),
    ),
)
SLIDE_MARKER_PLAN = GenerationPlan(
    "presentation",
    (
        GenerationOperation("slide.reset", {}),
        GenerationOperation("slide.add_title", {"title": MARKER}),
    ),
)


def _add_marker(path: Path, format_name: str) -> None:
    member = {
        "docx": "word/document.xml",
        "xlsx": "xl/workbook.xml",
        "pptx": "ppt/presentation.xml",
    }[format_name]
    with zipfile.ZipFile(path) as source:
        contents = {
            name: source.read(name) for name in source.namelist()
        }
    contents[member] += MARKER.encode("utf-8")
    with zipfile.ZipFile(path, "w") as destination:
        for name, data in contents.items():
            destination.writestr(name, data)


class FakeBridge:
    def __init__(
        self,
        *,
        result_error=None,
        error_code="GENERATION_COMMAND_FAILED",
        returned_path=None,
        applied_operations=2,
        write_marker=True,
    ):
        self.url = "http://127.0.0.1:45678"
        self.token = "token"
        self.result_error = result_error
        self.error_code = error_code
        self.returned_path = returned_path
        self.applied_operations = applied_operations
        self.write_marker = write_marker
        self.commands = []
        self.registrations = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        return None

    def wait_registered(self, expected, timeout):
        self.registrations.append((set(expected), timeout))

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
        staged = Path(command.params["stagedPath"])
        if self.result_error is not None:
            message = self.result_error.replace("{staging}", str(staged.parent))
            return ProbeResult(
                command_id,
                False,
                {},
                {"code": self.error_code, "message": message},
            )
        if self.write_marker:
            _add_marker(staged, command.params["formatName"])
        returned = staged if self.returned_path is None else self.returned_path
        return ProbeResult(
            command_id,
            True,
            {
                "path": str(returned),
                "appliedOperations": self.applied_operations,
            },
            None,
        )


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


def _run_with_fakes(
    tmp_path: Path,
    component: str,
    format_name: str,
    plan: GenerationPlan,
    *,
    bridge: FakeBridge | None = None,
    timeout: float = 2,
):
    calls = []
    fake_bridge = bridge or FakeBridge()
    runtime = FakeRuntime(tmp_path / "container" / "session-1", calls)
    request = GenerationRequest(
        output=tmp_path / f"final.{format_name}",
        component=component,
        format_name=format_name,
        overwrite=False,
    )
    result = execute_generation_plan(
        request,
        RecordedGeneration(plan, ()),
        enabled={format_name: True},
        bridge_factory=lambda origins: fake_bridge,
        runtime_factory=lambda *args, **kwargs: runtime,
        timeout=timeout,
    )
    return result, request, fake_bridge, runtime, calls


@pytest.mark.parametrize(
    ("component", "format_name", "method", "plan"),
    [
        ("writer", "docx", "generate_writer_document", WRITER_MARKER_PLAN),
        (
            "spreadsheet",
            "xlsx",
            "generate_spreadsheet_workbook",
            SHEET_MARKER_PLAN,
        ),
        (
            "presentation",
            "pptx",
            "generate_presentation_deck",
            SLIDE_MARKER_PLAN,
        ),
    ],
)
def test_generation_uses_only_staged_path_and_publishes_valid_package(
    tmp_path: Path,
    component: str,
    format_name: str,
    method: str,
    plan: GenerationPlan,
):
    result, request, bridge, runtime, calls = _run_with_fakes(
        tmp_path, component, format_name, plan
    )

    command = bridge.commands[0]
    assert command.component == component
    assert command.method == method
    assert set(command.params) == {"stagedPath", "formatName", "plan"}
    assert Path(command.params["stagedPath"]).parent == runtime.staging_dir
    assert command.params["formatName"] == format_name
    assert command.params["plan"] == plan.to_dict()
    assert "output" not in command.params
    assert result == request.output.resolve()
    assert result.is_file()
    assert not runtime.staging_dir.exists()
    assert ("activate_component", component) in calls


def test_generation_production_gates_remain_disabled(tmp_path: Path):
    assert mac_generation.MACOS_GENERATION_ENABLED == {
        "docx": False,
        "xlsx": False,
        "pptx": False,
        "pdf": False,
    }
    request = GenerationRequest(
        tmp_path / "final.docx", "writer", "docx", False
    )
    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request, RecordedGeneration(WRITER_MARKER_PLAN, ())
        )
    assert caught.value.code == "MACOS_GENERATION_GATE_NOT_PASSED"
    assert caught.value.backend == "mac-wps-jsapi"


def test_generation_rejects_component_format_and_plan_mismatch(tmp_path: Path):
    request = GenerationRequest(
        tmp_path / "final.xlsx", "spreadsheet", "xlsx", False
    )
    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            RecordedGeneration(WRITER_MARKER_PLAN, ()),
            enabled={"xlsx": True},
        )
    assert caught.value.code == "OPERATION_PLAN_INVALID"


def test_generation_rejects_host_resources_in_feasibility_backend(tmp_path: Path):
    resource_path = tmp_path / "image.png"
    resource_path.write_bytes(b"not staged")
    resource = GenerationResource("image-1", resource_path, "image/png")
    request = GenerationRequest(
        tmp_path / "final.docx", "writer", "docx", False
    )
    with pytest.raises(GenerationError) as caught:
        execute_generation_plan(
            request,
            RecordedGeneration(WRITER_MARKER_PLAN, (resource,)),
            enabled={"docx": True},
        )
    assert caught.value.code == "OPERATION_PLAN_INVALID"


@pytest.mark.parametrize(
    ("bridge", "expected_code"),
    [
        (
            FakeBridge(
                result_error="save failed at {staging}/generated.docx",
                error_code="GENERATION_COMMAND_FAILED",
            ),
            "GENERATION_COMMAND_FAILED",
        ),
        (
            FakeBridge(result_error="vendor failure", error_code="WPS_E_9001"),
            "GENERATION_COMMAND_FAILED",
        ),
        (FakeBridge(returned_path=Path("/tmp/outside.docx")), "PROTOCOL_ERROR"),
        (FakeBridge(applied_operations=1), "PROTOCOL_ERROR"),
        (FakeBridge(write_marker=False), "STAGED_ARTIFACT_INVALID"),
    ],
)
def test_generation_rejects_failed_or_invalid_results(
    tmp_path: Path, bridge: FakeBridge, expected_code: str
):
    with pytest.raises(GenerationError) as caught:
        _run_with_fakes(
            tmp_path,
            "writer",
            "docx",
            WRITER_MARKER_PLAN,
            bridge=bridge,
            timeout=0.02,
        )
    assert caught.value.code == expected_code
    assert not (tmp_path / "final.docx").exists()
    assert str((tmp_path / "container").resolve()) not in caught.value.message
