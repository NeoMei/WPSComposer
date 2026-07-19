from __future__ import annotations

import base64
import json
import time
from pathlib import Path
from typing import Any, Mapping

from .bridge import LoopbackBridge
from .models import PathPolicy, ProbeResult

PNG_FIXTURE = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8"
    "/x8AAusB9Y9Zl1sAAAAASUVORK5CYII="
)


class ArtifactError(RuntimeError):
    pass


def validate_artifact(path: Path, format_name: str) -> None:
    if not path.is_file():
        raise ArtifactError(f"{format_name.upper()} output is missing: {path}")
    signature = path.read_bytes()[:8]
    expected = b"%PDF-" if format_name == "pdf" else b"PK\x03\x04"
    if not signature.startswith(expected):
        raise ArtifactError(f"invalid {format_name.upper()} signature: {path}")
    if path.stat().st_size < 1024:
        raise ArtifactError(f"{format_name.upper()} output is too small: {path}")


def _write_json(path: Path, data: Any) -> None:
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )


def _execute_commands(
    bridge: LoopbackBridge,
    policy: PathPolicy,
    output_dir: Path,
    image_path: Path,
    timeout: float = 90,
    registered: set | None = None,
) -> dict[str, Any]:
    policy.require_allowed(image_path)
    if registered is None:
        registered = {"writer", "presentation", "spreadsheet"}
    capabilities: dict[str, Any] = {"writer": {}, "presentation": {}, "spreadsheet": {}}
    failures: list[dict[str, Any]] = []
    artifacts: list[dict[str, Any]] = []
    timings: dict[str, Any] = {}
    pdf_preset = ""

    for component in ("writer", "presentation", "spreadsheet"):
        if component not in registered:
            failures.append({
                "component": component,
                "method": "register",
                "error": {"code": "not_registered", "message": f"{component} add-in did not connect"},
            })
            continue
        start = time.monotonic()
        command = bridge.issue(component, "probe_capabilities", {})
        try:
            result = bridge.wait_result(command.id, timeout)
            timings[f"{component}_probe_ms"] = round((time.monotonic() - start) * 1000)
            if result.ok:
                caps = result.value.get("capabilities", {})
                capabilities[component] = caps
            else:
                failures.append({
                    "component": component,
                    "method": "probe_capabilities",
                    "error": dict(result.error or {}),
                })
        except TimeoutError:
            bridge.state.cancel(command.id)
            failures.append({
                "component": component,
                "method": "probe_capabilities",
                "error": {"code": "timeout", "message": "probe timed out"},
                "cancellation": "mapped",
            })

    commands = (
        ("writer", "smoke_docx", "smoke.docx", "docx"),
        ("presentation", "smoke_pptx", "smoke.pptx", "pptx"),
        ("spreadsheet", "smoke_xlsx", "smoke.xlsx", "xlsx"),
        ("writer", "smoke_pdf", "smoke.pdf", "pdf"),
    )

    for component, method, filename, fmt in commands:
        if component not in registered:
            continue
        out_path = output_dir / filename
        policy.require_allowed(out_path)
        params = {"outputPath": str(out_path.resolve()), "imagePath": str(image_path.resolve())}
        start = time.monotonic()
        command = bridge.issue(component, method, params)
        try:
            result = bridge.wait_result(command.id, timeout)
            timings[f"{method}_ms"] = round((time.monotonic() - start) * 1000)
            if result.ok:
                container_path = result.value.get("containerPath")
                if container_path and not out_path.is_file():
                    import shutil
                    src = Path(container_path)
                    if src.is_file():
                        shutil.copy2(src, out_path)
                        src.unlink(missing_ok=True)
                validate_artifact(out_path, fmt)
                artifacts.append({
                    "format": fmt,
                    "path": str(out_path.resolve()),
                    "size": out_path.stat().st_size,
                })
                if fmt == "pdf":
                    pdf_preset = result.value.get("preset", "")
                for key, val in result.value.get("capabilities", {}).items():
                    comp = key.split(".")[0] if "." in key else component
                    if comp in capabilities:
                        capabilities[comp][key] = val
            else:
                failures.append({
                    "component": component,
                    "method": method,
                    "error": dict(result.error or {}),
                })
        except TimeoutError:
            bridge.state.cancel(command.id)
            failures.append({
                "component": component,
                "method": method,
                "error": {"code": "timeout", "message": f"{method} timed out"},
                "cancellation": "mapped",
            })
        except ArtifactError as exc:
            failures.append({
                "component": component,
                "method": method,
                "error": {"code": "artifact_invalid", "message": str(exc)},
            })

    status = "passed" if len(artifacts) == 4 and not failures else "failed"
    return {
        "schemaVersion": 1,
        "status": status,
        "backend": "mac-wps-jsapi-probe",
        "platform": "macOS",
        "wpsVersion": "",
        "wpsjsVersion": "2.2.3",
        "pdfPreset": pdf_preset,
        "artifacts": artifacts,
        "failures": failures,
        "timingsMs": timings,
        "capabilities": capabilities,
    }


def run_phase0(
    output_dir: Path,
    node: str | None = None,
    timeout: float = 90,
) -> Path:
    from .runtime import COMPONENT_CONFIG, ProbeRuntime, read_wps_version

    output_dir.mkdir(parents=True, exist_ok=True)
    run_dir = output_dir / ".runtime"
    probe_root = Path(__file__).resolve().parent.parent.parent.parent.parent / "macos" / "wps-jsapi-probe"

    image_path = output_dir / "fixture.png"
    image_path.write_bytes(base64.b64decode(PNG_FIXTURE))

    policy = PathPolicy((output_dir,))
    origins = {"*"}

    wps_version = read_wps_version()

    with LoopbackBridge(origins) as bridge:
        runtime = ProbeRuntime(probe_root, run_dir, node=node)
        try:
            runtime.prepare_profiles(bridge.url, bridge.token)
            runtime.start_servers()
            runtime.activate_components()
            registered = set()
            for comp in ("writer", "presentation", "spreadsheet"):
                try:
                    bridge.wait_registered({comp}, min(timeout, 60))
                    registered.add(comp)
                except TimeoutError:
                    pass
            report = _execute_commands(bridge, policy, output_dir, image_path, timeout, registered)
        finally:
            runtime.restore_registration()

    report["wpsVersion"] = wps_version

    caps_path = output_dir / "capabilities.json"
    _write_json(caps_path, {
        "schemaVersion": 1,
        "platform": "macOS",
        "wpsVersion": wps_version,
        "components": report.pop("capabilities", {}),
    })

    report_path = output_dir / "phase0-report.json"
    _write_json(report_path, report)
    return report_path
