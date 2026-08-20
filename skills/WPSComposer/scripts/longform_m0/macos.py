"""Isolated macOS native runner for the long-form M0 capability gate."""

from __future__ import annotations

import copy
import shutil
import tempfile
import time
from pathlib import Path
from typing import Any, Callable, Mapping, Optional

from ..artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    publish_artifact_group,
    validate_office_package,
    validate_pdf,
)
from ..macos_probe.bridge import LoopbackBridge
from ..macos_probe.models import ProbeResult
from ..macos_probe.runtime import ProbeRuntime, read_wps_version
from . import host_checks
from .contracts import (
    PROBE_VERSION,
    PROTOCOL_VERSION,
    RESOURCE_MANIFEST_VERSION,
    REQUIRED_IDS,
    SCHEMA_VERSION,
    validate_platform_evidence,
    write_canonical_json,
)

EMPTY_MANIFEST = {
    "version": RESOURCE_MANIFEST_VERSION,
    "entries": [],
    "digest": "a6a20076da005b27c9afc3a5d5b2457798c0ac817d1abc38b2fee4398ac3f133",
}
ORIGINS = {
    "http://127.0.0.1:3889",
    "http://127.0.0.1:3890",
    "http://127.0.0.1:3891",
}
_RESULT_FIELDS = frozenset(
    {
        "probeVersion",
        "protocolVersion",
        "resourceManifestVersion",
        "platform",
        "wpsVersion",
        "capabilities",
        "failures",
        "docxPath",
        "pdfPath",
    }
)
_REMOTE_ERROR_CODES = frozenset(
    {
        "PROTOCOL_MISMATCH",
        "RESOURCE_MANIFEST_MISMATCH",
        "GENERATION_COMMAND_FAILED",
        "M0_PROBE_FAILED",
    }
)


class MacosM0Failed(RuntimeError):
    """Raised after redacted macOS M0 failure evidence has been written."""

    def __init__(self, evidence_path: Path):
        super().__init__("macOS long-form M0 probe failed")
        self.evidence_path = Path(evidence_path)


class _M0Error(RuntimeError):
    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code


def _remaining(deadline: float) -> float:
    budget = deadline - time.monotonic()
    if budget <= 0:
        raise _M0Error("COMMAND_TIMEOUT", "macOS M0 deadline expired")
    return budget


def _failure_capabilities() -> list[dict[str, Any]]:
    rows = []
    for capability_id in range(1, 16):
        if capability_id == 2:
            rows.append(
                {
                    "id": 2,
                    "status": "passed",
                    "checks": ["not-applicable-macos"],
                    "artifacts": [],
                    "metrics": {},
                }
            )
        else:
            rows.append(
                {
                    "id": capability_id,
                    "status": "not-run",
                    "checks": [],
                    "artifacts": [],
                    "metrics": {},
                }
            )
    return rows


def _failure_evidence(wps_version: str, code: str) -> dict[str, Any]:
    return {
        "schemaVersion": SCHEMA_VERSION,
        "probeVersion": PROBE_VERSION,
        "platform": "macos",
        "wpsVersion": wps_version or "unknown",
        "protocolVersion": PROTOCOL_VERSION,
        "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
        "capabilities": _failure_capabilities(),
        "artifacts": {},
        "failures": [
            {"code": code, "message": "macOS long-form M0 probe failed"}
        ],
    }


def _normalize_remote_failure(result: ProbeResult) -> _M0Error:
    error = result.error if isinstance(result.error, Mapping) else {}
    raw_code = error.get("code")
    code = raw_code if raw_code in _REMOTE_ERROR_CODES else "M0_PROBE_FAILED"
    return _M0Error(str(code), "native M0 command failed")


def _validate_result(
    value: Mapping[str, Any],
    staged_docx: Path,
    staged_pdf: Path,
    wps_version: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    if set(value) != _RESULT_FIELDS:
        raise _M0Error("PROTOCOL_ERROR", "native result fields are invalid")
    expected = {
        "probeVersion": PROBE_VERSION,
        "protocolVersion": PROTOCOL_VERSION,
        "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
        "platform": "macos",
        "wpsVersion": wps_version,
    }
    for key, expected_value in expected.items():
        if value.get(key) != expected_value:
            raise _M0Error("PROTOCOL_ERROR", f"native result {key} is invalid")
    try:
        reported_docx = Path(str(value["docxPath"])).expanduser().resolve()
        reported_pdf = Path(str(value["pdfPath"])).expanduser().resolve()
    except (OSError, RuntimeError, ValueError) as exc:
        raise _M0Error("PROTOCOL_ERROR", "native result paths are invalid") from exc
    if reported_docx != staged_docx or reported_pdf != staged_pdf:
        raise _M0Error("PROTOCOL_ERROR", "native result paths do not match staging")
    capabilities = value["capabilities"]
    failures = value["failures"]
    if not isinstance(capabilities, list) or not isinstance(failures, list):
        raise _M0Error("PROTOCOL_ERROR", "native result evidence is invalid")
    return copy.deepcopy(capabilities), copy.deepcopy(failures)


def _complete_evidence(
    capabilities: list[dict[str, Any]],
    failures: list[dict[str, Any]],
    artifacts: Mapping[str, Any],
    wps_version: str,
) -> dict[str, Any]:
    artifact_names = [artifacts["docx"]["name"], artifacts["pdf"]["name"]]
    capability_ids = set()
    for capability in capabilities:
        if not isinstance(capability, dict):
            raise _M0Error("PROTOCOL_ERROR", "native capability is invalid")
        if set(capability) != {"id", "status", "checks", "metrics"}:
            raise _M0Error("PROTOCOL_ERROR", "native capability fields are invalid")
        capability_id = capability.get("id")
        if isinstance(capability_id, bool) or not isinstance(capability_id, int):
            raise _M0Error("PROTOCOL_ERROR", "native capability id is invalid")
        capability_ids.add(capability_id)
        capability["artifacts"] = (
            list(artifact_names)
            if capability.get("status") == "passed" and capability_id >= 3
            else []
        )
        if capability_id == 5:
            metrics = capability.get("metrics")
            if not isinstance(metrics, dict):
                raise _M0Error("PROTOCOL_ERROR", "native capability metrics are invalid")
            metrics["hostPdfSnapshot"] = artifacts["pdf"]["snapshot"]
    if capability_ids != set(range(1, 16)):
        raise _M0Error("PROTOCOL_ERROR", "native capability ids are incomplete")
    raw = {
        "schemaVersion": SCHEMA_VERSION,
        "probeVersion": PROBE_VERSION,
        "platform": "macos",
        "wpsVersion": wps_version,
        "protocolVersion": PROTOCOL_VERSION,
        "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
        "capabilities": capabilities,
        "artifacts": {
            "docx": {
                "name": artifacts["docx"]["name"],
                "sha256": artifacts["docx"]["sha256"],
            },
            "pdf": {
                "name": artifacts["pdf"]["name"],
                "sha256": artifacts["pdf"]["sha256"],
            },
        },
        "failures": failures,
    }
    return raw


def _exception_code(error: BaseException, stage: str) -> str:
    if isinstance(error, _M0Error):
        return error.code
    if isinstance(error, ArtifactTransportError):
        return error.code
    if isinstance(error, ArtifactValidationError):
        return "STAGED_ARTIFACT_INVALID"
    if isinstance(error, TimeoutError):
        return "COMMAND_TIMEOUT"
    if stage == "dependency":
        return "DEPENDENCY_UNAVAILABLE"
    if isinstance(error, ValueError) and stage == "evidence":
        return "EVIDENCE_INVALID"
    return "RUNTIME_FAILURE"


def run_macos_probe(
    output_dir: Path,
    timeout: float = 600,
    *,
    bridge_factory: Callable[..., Any] = LoopbackBridge,
    runtime_factory: Callable[..., Any] = ProbeRuntime,
    version_reader: Callable[[], str] = read_wps_version,
) -> Path:
    """Run one installed-WPS macOS M0 probe and return its evidence path."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    output = host_checks.prepare_evidence_directory(output_dir)
    evidence_path = output / "platform-evidence.json"
    final_docx = output / "probe.docx"
    final_pdf = output / "probe.pdf"
    runtime_root = Path(tempfile.mkdtemp(prefix="wpscomposer-longform-m0-"))
    runtime_dir = runtime_root / "runtime"
    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    deadline = time.monotonic() + timeout
    try:
        wps_version = version_reader() or "unknown"
    except Exception:
        wps_version = "unknown"
    runtime: Optional[Any] = None
    stage = "dependency"
    try:
        host_checks.require_probe_dependencies()
        stage = "runtime"
        with bridge_factory(ORIGINS) as bridge:
            runtime = runtime_factory(
                probe_root,
                runtime_dir,
                bridge.url,
                bridge.token,
                deadline=deadline,
            )
            with runtime:
                if runtime.staging_dir is None:
                    raise _M0Error(
                        "STAGING_UNAVAILABLE", "WPS staging was not created"
                    )
                staged_docx = (runtime.staging_dir / "longform-m0.docx").resolve()
                staged_pdf = (runtime.staging_dir / "longform-m0.pdf").resolve()
                runtime.prepare_profiles()
                runtime.start_servers(deadline=deadline)
                runtime.activate_components()
                bridge.wait_registered({"writer"}, _remaining(deadline))
                command = bridge.issue(
                    "writer",
                    "probe_longform_m0",
                    {
                        "manifest": copy.deepcopy(EMPTY_MANIFEST),
                        "probeVersion": PROBE_VERSION,
                        "protocolVersion": PROTOCOL_VERSION,
                        "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
                        "stagedDocxPath": str(staged_docx),
                        "stagedPdfPath": str(staged_pdf),
                    },
                )
                try:
                    result = bridge.wait_result(command.id, _remaining(deadline))
                except TimeoutError as exc:
                    state = getattr(bridge, "state", None)
                    cancel = getattr(state, "cancel", None)
                    if callable(cancel):
                        cancel(command.id)
                    raise _M0Error("COMMAND_TIMEOUT", "native command timed out") from exc
                if not result.ok:
                    raise _normalize_remote_failure(result)
                if not isinstance(result.value, Mapping):
                    raise _M0Error("PROTOCOL_ERROR", "native result is invalid")
                capabilities, failures = _validate_result(
                    result.value, staged_docx, staged_pdf, wps_version
                )
                stage = "publication"
                publish_artifact_group(
                    (
                        (
                            staged_docx,
                            final_docx,
                            False,
                            lambda path: validate_office_package(path, "docx"),
                        ),
                        (staged_pdf, final_pdf, False, validate_pdf),
                    )
                )
        if runtime is None or not runtime.registration_restored:
            raise _M0Error(
                "REGISTRATION_RESTORE_FAILED",
                "WPS registration was not restored",
            )
        stage = "evidence"
        artifact_evidence = host_checks.validate_native_artifacts(
            final_docx, final_pdf
        )
        raw = _complete_evidence(
            capabilities, failures, artifact_evidence, wps_version
        )
        host_checks.validate_evidence_privacy(raw)
        validate_platform_evidence(raw)
        write_canonical_json(evidence_path, raw)
        return evidence_path
    except Exception as error:
        final_docx.unlink(missing_ok=True)
        final_pdf.unlink(missing_ok=True)
        code = _exception_code(error, stage)
        raw = _failure_evidence(wps_version, code)
        host_checks.validate_evidence_privacy(raw)
        validate_platform_evidence(raw)
        write_canonical_json(evidence_path, raw)
        raise MacosM0Failed(evidence_path) from error
    finally:
        if runtime is None or getattr(runtime, "registration_restored", False):
            shutil.rmtree(runtime_root, ignore_errors=True)
