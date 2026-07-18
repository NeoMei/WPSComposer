"""Minimal macOS WPS JSAPI feasibility backend for in-place saves."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil
import tempfile
import time
from typing import Callable, Mapping, Optional
import zipfile
from xml.etree import ElementTree

from ..artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    publish_artifact,
    validate_office_package,
)
from ..generation_plan import (
    OperationPlanError,
    RecordedGeneration,
    validate_generation_plan,
)
from .bridge import LoopbackBridge
from .models import PathPolicy, ProtocolError
from .runtime import ProbeRuntime
from .templates import TemplateError, clone_template


ORIGINS = {
    "http://127.0.0.1:3889",
    "http://127.0.0.1:3890",
    "http://127.0.0.1:3891",
}
METHODS = {
    "writer": "generate_writer_document",
    "spreadsheet": "generate_spreadsheet_workbook",
    "presentation": "generate_presentation_deck",
}
FORMAT_COMPONENTS = {
    "docx": "writer",
    "xlsx": "spreadsheet",
    "pptx": "presentation",
}
REMOTE_GENERATION_ERROR_CODES = frozenset(
    {
        "GENERATION_COMMAND_FAILED",
        "INTERACTIVE_INPUT_REQUIRED",
        "MACOS_CAPABILITY_UNAVAILABLE",
        "OPERATION_PLAN_INVALID",
    }
)
MACOS_GENERATION_ENABLED = {
    "docx": False,
    "xlsx": False,
    "pptx": False,
    "pdf": False,
}
FEASIBILITY_MARKER = "IN_PLACE_MARKER"
FEASIBILITY_PLANS = {
    "writer": {
        "component": "writer",
        "operations": [
            {"op": "writer.reset", "args": {}},
            {
                "op": "writer.add_paragraph",
                "args": {
                    "text": FEASIBILITY_MARKER,
                    "style": "Body Text",
                },
            },
        ],
    },
    "spreadsheet": {
        "component": "spreadsheet",
        "operations": [
            {"op": "sheet.reset", "args": {}},
            {
                "op": "sheet.write_table",
                "args": {
                    "startRow": 1,
                    "startCol": 1,
                    "values": [[FEASIBILITY_MARKER]],
                },
            },
        ],
    },
    "presentation": {
        "component": "presentation",
        "operations": [
            {"op": "slide.reset", "args": {}},
            {
                "op": "slide.add_title",
                "args": {"title": FEASIBILITY_MARKER},
            },
        ],
    },
}
MARKER_CONTENT_MEMBERS = {
    "docx": frozenset({"word/document.xml"}),
    "xlsx": frozenset({"xl/sharedStrings.xml"}),
    "pptx": frozenset({"ppt/slides/slide1.xml"}),
}


@dataclass(frozen=True)
class GenerationRequest:
    output: Path
    component: str
    format_name: str
    overwrite: bool = False


class GenerationError(RuntimeError):
    def __init__(
        self, *, code, output, component, backend, message
    ):
        super().__init__(message)
        self.code = code
        self.output = output
        self.component = component
        self.backend = backend
        self.message = message


def _error(
    request: GenerationRequest, code: str, message: str
) -> GenerationError:
    return GenerationError(
        code=code,
        output=str(request.output),
        component=request.component,
        backend="mac-wps-jsapi",
        message=message,
    )


def _normalize_remote_error_code(code: object) -> str:
    value = str(code) if isinstance(code, str) else ""
    if value in REMOTE_GENERATION_ERROR_CODES:
        return value
    return "GENERATION_COMMAND_FAILED"


def _redact_staging(message: str, staging_dir: Path) -> str:
    return str(message).replace(str(staging_dir), "<wps-staging>")


def _validate_feasibility_recording(
    request: GenerationRequest, recorded: RecordedGeneration
) -> RecordedGeneration:
    try:
        validated = validate_generation_plan(
            recorded.plan.to_dict(), request.component
        )
    except (AttributeError, OperationPlanError, TypeError) as exc:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            str(exc) or "Generation plan is invalid",
        ) from exc
    if validated.to_dict() != FEASIBILITY_PLANS[request.component]:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Generation plan is not the exact in-place feasibility plan",
        )
    if recorded.resources:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Feasibility generation does not accept host resources",
        )
    return RecordedGeneration(validated, ())


def _wait_for_registration(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    component: str,
    timeout: float,
) -> None:
    deadline = time.monotonic() + timeout
    for attempt in range(4):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            bridge.wait_registered({component}, 0)
            return
        try:
            bridge.wait_registered({component}, min(10, remaining))
            return
        except TimeoutError:
            if attempt == 3:
                raise
            runtime.activate_component(component)


def _validate_marker_package(path: Path, format_name: str) -> None:
    validate_office_package(path, format_name)
    normalized = str(format_name).lower().lstrip(".")
    content_members = MARKER_CONTENT_MEMBERS[normalized]
    found = False
    try:
        with zipfile.ZipFile(path) as package:
            for name in package.namelist():
                lowered = name.lower()
                if not (lowered.endswith(".xml") or lowered.endswith(".rels")):
                    continue
                root = ElementTree.fromstring(package.read(name))
                if name in content_members:
                    text = "".join(root.itertext())
                    if FEASIBILITY_MARKER in text:
                        found = True
    except (
        ElementTree.ParseError,
        KeyError,
        OSError,
        RuntimeError,
        zipfile.BadZipFile,
    ) as exc:
        raise ArtifactValidationError(
            f"Invalid {normalized.upper()} XML package: {path}"
        ) from exc
    if not found:
        raise ArtifactValidationError(
            f"{normalized.upper()} package is missing {FEASIBILITY_MARKER} "
            f"in its generation content member: {path}"
        )


def _wait_for_marker(path: Path, format_name: str, timeout: float) -> None:
    deadline = time.monotonic() + timeout
    failure: Optional[ArtifactValidationError] = None
    while True:
        try:
            _validate_marker_package(path, format_name)
            return
        except ArtifactValidationError as exc:
            failure = exc
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            assert failure is not None
            raise failure
        time.sleep(min(0.05, remaining))


def _run_generation(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    method: str,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    probe_root: Path,
    timeout: float,
) -> Path:
    if runtime.staging_dir is None:
        raise _error(
            request,
            "STAGING_UNAVAILABLE",
            "WPS container staging session was not created",
        )
    runtime.prepare_profiles()
    runtime.start_servers()
    runtime.activate_component(request.component)
    _wait_for_registration(bridge, runtime, request.component, timeout)

    policy = PathPolicy((runtime.staging_dir,))
    try:
        staged = policy.require_allowed(
            clone_template(probe_root, runtime.staging_dir, request.component)
        )
    except (ArtifactValidationError, OSError, TemplateError) as exc:
        raise _error(
            request,
            "STAGING_SAVE_FAILED",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc

    command = bridge.issue(
        request.component,
        method,
        {
            "stagedPath": str(staged),
            "formatName": request.format_name,
            "plan": recorded.plan.to_dict(),
        },
    )
    try:
        result = bridge.wait_result(command.id, timeout)
    except TimeoutError as exc:
        raise _error(
            request,
            "GENERATION_COMMAND_FAILED",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    if not result.ok:
        details = dict(result.error or {})
        raise _error(
            request,
            _normalize_remote_error_code(details.get("code")),
            _redact_staging(
                str(details.get("message") or "WPS generation failed"),
                runtime.staging_dir,
            ),
        )
    try:
        reported = policy.require_allowed(str(result.value.get("path", "")))
    except ProtocolError as exc:
        raise _error(
            request,
            "PROTOCOL_ERROR",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    if reported != staged:
        raise _error(
            request,
            "PROTOCOL_ERROR",
            "WPS returned an unexpected staged generation path",
        )
    applied = result.value.get("appliedOperations")
    if (
        isinstance(applied, bool)
        or not isinstance(applied, int)
        or applied != len(recorded.plan.operations)
    ):
        raise _error(
            request,
            "PROTOCOL_ERROR",
            "WPS returned an unexpected applied operation count",
        )
    try:
        _wait_for_marker(staged, request.format_name, timeout)
    except ArtifactValidationError as exc:
        raise _error(
            request,
            "STAGED_ARTIFACT_INVALID",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    try:
        return publish_artifact(
            staged,
            request.output,
            overwrite=request.overwrite,
            validator=lambda path: _validate_marker_package(
                path, request.format_name
            ),
        )
    except ArtifactTransportError as exc:
        raise _error(
            request,
            exc.code,
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc


def execute_generation_plan(
    request: GenerationRequest,
    recorded: RecordedGeneration,
    *,
    enabled: Optional[Mapping[str, bool]] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 90,
) -> Path:
    """Execute the closed feasibility plan against a private template clone."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    gates = MACOS_GENERATION_ENABLED if enabled is None else enabled
    if not bool(gates.get(request.format_name, False)):
        raise _error(
            request,
            "MACOS_GENERATION_GATE_NOT_PASSED",
            "Mac WPS generation is disabled until the real acceptance gate passes",
        )
    expected_component = FORMAT_COMPONENTS.get(request.format_name)
    if expected_component != request.component:
        raise _error(
            request,
            "OPERATION_PLAN_INVALID",
            "Generation component and format do not match",
        )
    recorded = _validate_feasibility_recording(request, recorded)
    method = METHODS[request.component]

    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    runtime_root = Path(
        tempfile.mkdtemp(prefix="wpscomposer-macos-generate-")
    ).resolve()
    runtime = None
    try:
        with bridge_factory(ORIGINS) as bridge:
            runtime = runtime_factory(
                probe_root,
                runtime_root / "runtime",
                bridge.url,
                bridge.token,
            )
            try:
                with runtime:
                    return _run_generation(
                        request,
                        recorded,
                        method,
                        bridge,
                        runtime,
                        probe_root,
                        timeout,
                    )
            except GenerationError:
                raise
            except Exception as exc:
                if not getattr(runtime, "registration_restored", True):
                    recovery = runtime.runtime_dir / "registration-recovery"
                    raise _error(
                        request,
                        "REGISTRATION_RESTORE_FAILED",
                        "WPS registration restore failed; recovery retained at "
                        f"{recovery}",
                    ) from exc
                raise
    finally:
        if runtime is None or getattr(runtime, "registration_restored", True):
            shutil.rmtree(runtime_root, ignore_errors=True)
