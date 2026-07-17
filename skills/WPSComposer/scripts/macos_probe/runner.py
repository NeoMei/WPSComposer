from __future__ import annotations

import base64
import hashlib
import json
import locale
import shutil
import tempfile
import time
from pathlib import Path
from typing import Any, Optional

from .bridge import LoopbackBridge
from .models import COMPONENTS, PathPolicy, ProbeResult, ProtocolError
from .runtime import ProbeRuntime, read_wps_version

PNG_FIXTURE = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8"
    "/x8AAusB9Y9Zl1sAAAAASUVORK5CYII="
)
ORIGINS = {
    "http://127.0.0.1:3889",
    "http://127.0.0.1:3890",
    "http://127.0.0.1:3891",
}
WRITER_IMAGE_URL = "http://127.0.0.1:3889/fixture.png"
OUTPUT_COMMANDS = (
    ("writer", "smoke_docx", "smoke.docx", "docx"),
    ("presentation", "smoke_pptx", "smoke.pptx", "pptx"),
    ("spreadsheet", "smoke_xlsx", "smoke.xlsx", "xlsx"),
    ("writer", "smoke_pdf", "smoke.pdf", "pdf"),
)


class ArtifactError(RuntimeError):
    """Raised when WPS did not produce a structurally plausible artifact."""


class Phase0Failed(RuntimeError):
    """Raised after a failing probe report has been safely written."""

    def __init__(self, message: str, report_path: Path):
        super().__init__(message)
        self.report_path = report_path


def validate_artifact(path: Path, format_name: str) -> None:
    if not path.is_file():
        raise ArtifactError(
            f"{format_name.upper()} output is missing: {path}"
        )
    with path.open("rb") as stream:
        signature = stream.read(8)
    expected = b"%PDF-" if format_name == "pdf" else b"PK\x03\x04"
    if not signature.startswith(expected):
        raise ArtifactError(
            f"invalid {format_name.upper()} signature: {path}"
        )
    if path.stat().st_size < 1024:
        raise ArtifactError(
            f"{format_name.upper()} output is too small: {path}"
        )


def wait_for_artifact(path: Path, format_name: str, timeout: float) -> None:
    """Wait for WPS Writer's asynchronous filesystem flush to finish."""
    if timeout <= 0:
        raise ValueError("artifact timeout must be positive")
    deadline = time.monotonic() + timeout
    while True:
        try:
            validate_artifact(path, format_name)
            return
        except ArtifactError:
            if time.monotonic() >= deadline:
                raise
            time.sleep(min(0.1, max(0, deadline - time.monotonic())))


def _failure(
    component: str,
    method: str,
    code: str,
    message: str,
    error: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    return {
        "component": component,
        "method": method,
        "code": code,
        "message": message,
        "error": error,
        "logPath": None,
    }


def _wait_for_result(
    bridge: LoopbackBridge,
    component: str,
    method: str,
    command_id: str,
    timeout: float,
) -> tuple[Optional[ProbeResult], Optional[dict[str, Any]]]:
    try:
        result = bridge.wait_result(command_id, timeout)
    except TimeoutError as exc:
        bridge.state.cancel(command_id)
        return None, _failure(
            component,
            method,
            "COMMAND_TIMEOUT",
            str(exc),
            {"cancellation": "mapped"},
        )
    if not result.ok:
        error = None if result.error is None else dict(result.error)
        message = (
            str(error.get("message", "WPS JSAPI command failed"))
            if error
            else "WPS JSAPI command failed"
        )
        return None, _failure(
            component,
            method,
            "JSAPI_COMMAND_FAILED",
            message,
            error,
        )
    return result, None


def _execute_commands(
    bridge: LoopbackBridge,
    policy: PathPolicy,
    output_dir: Path,
    image_path: Path,
    timeout: float,
) -> dict[str, Any]:
    """Run independent component probes and all four typed smoke commands."""
    started = time.monotonic()
    failures: list[dict[str, Any]] = []
    artifacts: list[dict[str, Any]] = []
    capabilities: dict[str, dict[str, Any]] = {
        component: {} for component in COMPONENTS
    }

    for component in COMPONENTS:
        command = bridge.issue(component, "probe_capabilities", {})
        result, failure = _wait_for_result(
            bridge,
            component,
            "probe_capabilities",
            command.id,
            timeout,
        )
        if failure is not None:
            failures.append(failure)
            continue
        assert result is not None
        reported = result.value.get("capabilities") or {}
        if isinstance(reported, dict):
            capabilities[component].update(reported)

    allowed_image = policy.require_allowed(image_path)
    pdf_source = policy.require_allowed(
        image_path.parent / "smoke-pdf-source.docx"
    )
    for component, method, filename, format_name in OUTPUT_COMMANDS:
        output_path = policy.require_allowed(output_dir / filename)
        command = bridge.issue(
            component,
            method,
            {
                "outputPath": str(output_path),
                "imagePath": str(allowed_image),
                "imageUrl": WRITER_IMAGE_URL,
                "sourcePath": str(pdf_source),
            },
        )
        result, failure = _wait_for_result(
            bridge, component, method, command.id, timeout
        )
        if failure is not None:
            failures.append(failure)
            continue
        assert result is not None
        reported_capabilities = result.value.get("capabilities") or {}
        if isinstance(reported_capabilities, dict):
            capabilities[component].update(reported_capabilities)
        try:
            reported_path = policy.require_allowed(
                str(result.value.get("path", ""))
            )
            if reported_path != output_path:
                raise ArtifactError(
                    f"WPS reported {reported_path}, expected {output_path}"
                )
            wait_for_artifact(
                output_path,
                format_name,
                timeout=min(timeout, 10),
            )
            if format_name == "pdf" and result.value.get("preset") != (
                "obsidian-reading"
            ):
                raise ArtifactError(
                    "PDF result did not report obsidian-reading"
                )
        except (ArtifactError, ProtocolError) as exc:
            failures.append(
                _failure(
                    component,
                    method,
                    "ARTIFACT_VALIDATION_FAILED",
                    str(exc),
                )
            )
            continue
        artifacts.append(
            {
                "format": format_name,
                "path": str(output_path),
                "size": output_path.stat().st_size,
            }
        )

    return {
        "status": "passed" if not failures and len(artifacts) == 4 else "failed",
        "artifacts": artifacts,
        "failures": failures,
        "components": capabilities,
        "pdfPreset": "obsidian-reading",
        "timingsMs": {
            "commands": round((time.monotonic() - started) * 1000, 3)
        },
    }


def _write_json(path: Path, payload: object) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _smoke_spec_hash(addin_dir: Path) -> str:
    digest = hashlib.sha256()
    for path in sorted(item for item in addin_dir.iterdir() if item.is_file()):
        digest.update(path.name.encode("utf-8"))
        digest.update(path.read_bytes())
    return digest.hexdigest()


def _publish_writer_image(
    profiles: dict[str, Path], image_path: Path
) -> str:
    writer_profile = profiles["writer"]
    shutil.copy2(image_path, writer_profile / "fixture.png")
    return WRITER_IMAGE_URL


def _copy_logs(runtime: ProbeRuntime, output_dir: Path) -> dict[str, str]:
    copied: dict[str, str] = {}
    existing = {name: path for name, path in runtime.logs.items() if path.is_file()}
    if not existing:
        return copied
    logs_dir = output_dir / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    for component, source in existing.items():
        target = logs_dir / source.name
        shutil.copy2(source, target)
        copied[component] = str(target.resolve())
    return copied


def _ensure_output_targets_are_new(output_dir: Path) -> None:
    for _, _, filename, _ in OUTPUT_COMMANDS:
        target = output_dir / filename
        if target.exists():
            raise FileExistsError(
                f"Refusing to overwrite an existing Phase 0 artifact: {target}"
            )
    with tempfile.NamedTemporaryFile(dir=output_dir, prefix=".write-test-"):
        pass


def _allocate_runtime_dir() -> tuple[Path, Path]:
    root = Path(tempfile.mkdtemp(prefix="wpscomposer-phase0-"))
    return root, root / "runtime"


def _wait_for_component_registration(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    timeout: float,
) -> None:
    """Retry only WPS components that miss the initial registration window."""
    expected = set(COMPONENTS)
    deadline = time.monotonic() + timeout
    for attempt in range(4):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            bridge.wait_registered(expected, 0)
            return
        try:
            bridge.wait_registered(expected, min(10, remaining))
            return
        except TimeoutError:
            missing = expected - bridge.registered_components()
            if attempt == 3:
                raise
            for component in sorted(missing):
                runtime.activate_component(component)


def run_phase0(
    output_dir: Path,
    node: Optional[str] = None,
    timeout: float = 90,
) -> Path:
    """Run the installed-Mac feasibility gate and return its report path."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    started = time.monotonic()
    output = output_dir.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    report_path = output / "phase0-report.json"
    capabilities_path = output / "capabilities.json"
    runtime_root, runtime_dir = _allocate_runtime_dir()
    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    addin_dir = probe_root / "addin"
    capabilities: dict[str, dict[str, Any]] = {
        component: {} for component in COMPONENTS
    }
    report: dict[str, Any] = {
        "schemaVersion": 1,
        "status": "failed",
        "backend": "mac-wps-jsapi-probe",
        "platform": "macOS",
        "wpsVersion": "unknown",
        "wpsjsVersion": "2.2.3",
        "locale": locale.getlocale()[0] or "unknown",
        "requestedFormats": ["docx", "pptx", "xlsx", "pdf"],
        "renderPlanSchemaVersion": None,
        "smokeSpecHash": _smoke_spec_hash(addin_dir),
        "requestedFonts": ["PingFang SC", "Helvetica Neue"],
        "resolvedFonts": [],
        "pdfPreset": "obsidian-reading",
        "artifacts": [],
        "failures": [],
        "warnings": [],
        "logPaths": {},
        "timingsMs": {},
    }
    runtime: Optional[ProbeRuntime] = None
    try:
        _ensure_output_targets_are_new(output)
        report["wpsVersion"] = read_wps_version()
        with LoopbackBridge(ORIGINS) as bridge:
            runtime = ProbeRuntime(
                probe_root,
                runtime_dir,
                bridge.url,
                bridge.token,
                node_override=node,
            )
            with runtime:
                profiles = runtime.prepare_profiles()
                image_path = runtime_dir / "fixture.png"
                image_path.write_bytes(base64.b64decode(PNG_FIXTURE))
                _publish_writer_image(profiles, image_path)
                runtime.start_servers()
                runtime.activate_components()
                _wait_for_component_registration(bridge, runtime, timeout)
                policy = PathPolicy((runtime_dir, output))
                execution = _execute_commands(
                    bridge, policy, output, image_path, timeout
                )
                capabilities = execution.pop("components")
                report.update(execution)
    except Exception as exc:
        report["status"] = "failed"
        report["failures"].append(
            _failure(
                "runtime",
                "run_phase0",
                type(exc).__name__,
                str(exc),
            )
        )
    finally:
        if runtime is not None:
            report["logPaths"] = _copy_logs(runtime, output)
            for failure in report["failures"]:
                component = failure.get("component")
                if component in report["logPaths"]:
                    failure["logPath"] = report["logPaths"][component]
            if runtime.registration_restored:
                shutil.rmtree(runtime_root, ignore_errors=True)
            else:
                report["recoveryDirectory"] = str(runtime_dir)
        else:
            shutil.rmtree(runtime_root, ignore_errors=True)

    unsupported = []
    for component, entries in capabilities.items():
        for name, entry in entries.items():
            if isinstance(entry, dict) and entry.get("classification") == (
                "unsupported"
            ):
                unsupported.append(f"{component}:{name}")
    if unsupported:
        report["warnings"].append(
            "Unsupported capabilities: " + ", ".join(sorted(unsupported))
        )
    report["timingsMs"]["total"] = round(
        (time.monotonic() - started) * 1000, 3
    )
    _write_json(
        capabilities_path,
        {
            "schemaVersion": 1,
            "platform": "macOS",
            "wpsVersion": report["wpsVersion"],
            "components": capabilities,
        },
    )
    _write_json(report_path, report)
    if report["status"] != "passed":
        first = report["failures"][0] if report["failures"] else {}
        raise Phase0Failed(
            str(first.get("message", "Mac WPS Phase 0 failed")),
            report_path,
        )
    return report_path
