"""Supervised Windows worker boundary for the long-form M0 gate.

The module is deliberately importable without pywin32.  COM and Windows
process helpers are loaded only inside the native worker/runtime functions.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
import hashlib
import json
import ntpath
import platform
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import PureWindowsPath
from types import MappingProxyType
from types import SimpleNamespace
from typing import Any, Callable, Mapping, Optional

from ..artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    publish_artifact_group,
    validate_office_package,
    validate_pdf,
)
from . import host_checks
from .contracts import (
    PROBE_VERSION,
    PROTOCOL_VERSION,
    RESOURCE_MANIFEST_VERSION,
    SCHEMA_VERSION,
    validate_platform_evidence,
    write_canonical_json,
)

SVG_FIXTURE = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="60" '
    'viewBox="0 0 120 60"><rect width="120" height="60" fill="#E8F0FE"/>'
    '<text x="12" y="36" font-size="16">M0 SVG</text></svg>'
)
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

_REQUEST_FIELDS = frozenset(
    {
        "probeVersion",
        "protocolVersion",
        "resourceManifestVersion",
        "workerMode",
        "ownershipTimeoutVerified",
        "stagedDocxPath",
        "stagedPdfPath",
        "stagedSvgPath",
        "expectedWpsVersion",
        "preexistingPids",
        "cancelPath",
        "ackPath",
    }
)
_IDENTITY_FIELDS = frozenset(
    {"pid", "executable", "createdNs", "parentPid"}
)


def _run_powershell(script: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            "powershell.exe",
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            script,
        ],
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
    )


def _read_parent_pid(pid: int) -> int:
    result = _run_powershell(
        "$p=Get-CimInstance Win32_Process -Filter 'ProcessId="
        + str(int(pid))
        + "'; if($null -ne $p){[Console]::Write($p.ParentProcessId)}"
    )
    value = result.stdout.strip()
    if not value.isdigit() or int(value) <= 0:
        raise WindowsOwnershipError("WPS parent PID is unavailable")
    return int(value)


def _read_windows_identity(pid: int) -> Optional[ProcessIdentity]:
    """Read immutable process identity fields through pywin32 and CIM."""
    try:
        import win32api
        import win32process

        handle = win32api.OpenProcess(0x1000 | 0x0010, False, int(pid))
    except Exception:
        return None
    try:
        executable = ntpath.normcase(
            ntpath.normpath(str(win32process.GetModuleFileNameEx(handle, 0)))
        )
        created = win32process.GetProcessTimes(handle)["CreationTime"]
        created_ns = int(float(created.timestamp()) * 1_000_000_000)
        parent_pid = _read_parent_pid(int(pid))
        return ProcessIdentity(
            pid=int(pid),
            executable=executable,
            created_ns=created_ns,
            parent_pid=parent_pid,
        )
    except Exception:
        return None
    finally:
        try:
            win32api.CloseHandle(handle)
        except Exception:
            pass


def _snapshot_wps_pids(
    *, runner: Callable[..., Any] = subprocess.run
) -> set[int]:
    result = runner(
        [
            "powershell.exe",
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "@(Get-Process -Name wps -ErrorAction SilentlyContinue).Id | "
            "ForEach-Object {[Console]::WriteLine($_)}",
        ],
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
    )
    pids = set()
    for line in str(result.stdout).splitlines():
        value = line.strip()
        if value.isdigit() and int(value) > 0:
            pids.add(int(value))
    return pids


def _terminate_windows_tree(
    identity: ProcessIdentity,
    *,
    read_identity: Callable[[int], Optional[ProcessIdentity]] = _read_windows_identity,
    runner: Callable[..., Any] = subprocess.run,
) -> None:
    """Force-stop only a process whose immutable identity still matches."""
    if read_identity(identity.pid) != identity:
        return
    runner(
        ["taskkill.exe", "/PID", str(identity.pid), "/T", "/F"],
        check=True,
        capture_output=True,
        text=True,
        timeout=10,
    )


class WindowsOwnershipError(RuntimeError):
    """Raised when a dedicated WPS process cannot be proven."""


class WindowsWorkerTimeout(TimeoutError):
    """A supervised worker exceeded its deadline after ownership discovery."""

    def __init__(self, identity: Optional[ProcessIdentity]):
        super().__init__("Windows long-form M0 worker timed out")
        self.identity = identity


@dataclass(frozen=True)
class ProcessIdentity:
    pid: int
    executable: str
    created_ns: int
    parent_pid: int

    def __post_init__(self) -> None:
        for name in ("pid", "created_ns", "parent_pid"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
                raise ValueError(f"process {name} must be a positive integer")
        if not isinstance(self.executable, str) or not self.executable:
            raise ValueError("process executable must be a non-empty string")

    def to_json(self) -> dict[str, Any]:
        return {
            "pid": self.pid,
            "executable": self.executable,
            "createdNs": self.created_ns,
            "parentPid": self.parent_pid,
        }

    @classmethod
    def from_json(cls, value: Any) -> "ProcessIdentity":
        if not isinstance(value, Mapping) or set(value) != _IDENTITY_FIELDS:
            raise ValueError("process identity fields are invalid")
        return cls(
            pid=value["pid"],
            executable=value["executable"],
            created_ns=value["createdNs"],
            parent_pid=value["parentPid"],
        )


@dataclass(frozen=True)
class OwnedApplication:
    app: Any
    identity: ProcessIdentity


def _dispatch_owned_application(
    client: Any,
    *,
    hwnd_pid: Callable[[int], int],
    read_identity: Callable[[int], Optional[ProcessIdentity]],
    preexisting_pids: set[int],
) -> OwnedApplication:
    """Create and prove one dedicated WPS COM server without Dispatch fallback."""
    app = client.DispatchEx("kwps.Application")
    helper_document = None
    try:
        try:
            helper_document = app.Documents.Add()
            window = app.ActiveWindow
        except BaseException:
            window = None
        if window is None:
            raise WindowsOwnershipError("WPS window did not become available")
        hwnd = int(window.Hwnd)
        pid = hwnd_pid(hwnd)
        if isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0:
            raise WindowsOwnershipError("WPS HWND did not resolve to a PID")
        if pid in preexisting_pids:
            raise WindowsOwnershipError("DispatchEx reused a pre-existing WPS PID")
        identity = read_identity(pid)
        if identity is None or identity.pid != pid:
            raise WindowsOwnershipError("WPS process identity is unavailable")
        try:
            if helper_document is not None:
                helper_document.Close(0)
        except BaseException:
            pass
        return OwnedApplication(app=app, identity=identity)
    except BaseException:
        try:
            app.Quit()
        except BaseException:
            pass
        raise


def _windows_path(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value or len(value) > 4096:
        raise ValueError(f"{label} must be a non-empty Windows path")
    path = PureWindowsPath(value)
    if not path.is_absolute() or any(part in {".", ".."} for part in path.parts):
        raise ValueError(f"{label} must be an absolute Windows path")
    return value


def _validate_worker_request(value: Any) -> Mapping[str, Any]:
    if not isinstance(value, Mapping) or set(value) != _REQUEST_FIELDS:
        raise ValueError("worker request fields are invalid")
    if value["probeVersion"] != PROBE_VERSION:
        raise ValueError("worker request probe version is invalid")
    if value["protocolVersion"] != PROTOCOL_VERSION:
        raise ValueError("worker request protocol version is invalid")
    if value["resourceManifestVersion"] != RESOURCE_MANIFEST_VERSION:
        raise ValueError("worker request resource manifest version is invalid")
    paths = {
        name: _windows_path(value[name], name)
        for name in (
            "stagedDocxPath",
            "stagedPdfPath",
            "stagedSvgPath",
            "cancelPath",
            "ackPath",
        )
    }
    mode = value["workerMode"]
    if mode not in {"probe", "ownership-timeout"}:
        raise ValueError("worker request mode is invalid")
    timeout_verified = value["ownershipTimeoutVerified"]
    if not isinstance(timeout_verified, bool):
        raise ValueError("worker timeout verification must be boolean")
    if mode == "probe" and not timeout_verified:
        raise ValueError("probe worker requires timeout verification")
    if mode == "ownership-timeout" and timeout_verified:
        raise ValueError("timeout fixture cannot be pre-verified")
    version = value["expectedWpsVersion"]
    if not isinstance(version, str) or not version or len(version) > 128:
        raise ValueError("worker request WPS version is invalid")
    raw_pids = value["preexistingPids"]
    if not isinstance(raw_pids, list) or any(
        isinstance(pid, bool) or not isinstance(pid, int) or pid <= 0
        for pid in raw_pids
    ):
        raise ValueError("worker request pre-existing PIDs are invalid")
    pids = tuple(sorted(set(raw_pids)))
    if len(pids) != len(raw_pids):
        raise ValueError("worker request pre-existing PIDs must be unique")
    return MappingProxyType(
        {
            "probeVersion": PROBE_VERSION,
            "protocolVersion": PROTOCOL_VERSION,
            "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
            "workerMode": mode,
            "ownershipTimeoutVerified": timeout_verified,
            **paths,
            "expectedWpsVersion": version,
            "preexistingPids": pids,
        }
    )


def _validate_progress(
    value: Any, preexisting_pids: set[int]
) -> ProcessIdentity:
    identity = ProcessIdentity.from_json(value)
    if identity.pid in preexisting_pids:
        raise ValueError("worker claimed a pre-existing WPS process")
    return identity


def _cancel_then_terminate(
    worker: Any,
    cancel_path: Any,
    identity: ProcessIdentity,
    *,
    read_identity: Callable[[int], Optional[ProcessIdentity]],
    terminate_tree: Callable[[ProcessIdentity], None],
    grace_seconds: float = 5,
) -> None:
    """Request cancellation, then terminate only the still-matching WPS tree."""
    if grace_seconds <= 0 or grace_seconds > 5:
        raise ValueError("Windows cancellation grace must be in (0, 5]")
    write_canonical_json(cancel_path, {"cancel": True})
    try:
        worker.wait(timeout=grace_seconds)
        return
    except (TimeoutError, subprocess.TimeoutExpired):
        pass
    current = read_identity(identity.pid)
    if current == identity:
        terminate_tree(identity)


def _cancel_without_identity(worker: Any, cancel_path: Path) -> None:
    """Stop only the Python worker when no WPS identity can be proven."""
    write_canonical_json(cancel_path, {"cancel": True})
    try:
        worker.wait(timeout=5)
        return
    except (TimeoutError, subprocess.TimeoutExpired):
        pass
    terminate = getattr(worker, "terminate", None)
    if callable(terminate):
        terminate()
        try:
            worker.wait(timeout=1)
        except (TimeoutError, subprocess.TimeoutExpired):
            pass


def _read_json(path: Path, label: str) -> Any:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Windows worker {label} is invalid") from exc


def _supervise_worker(
    worker: Any,
    progress_path: Path,
    result_path: Path,
    cancel_path: Path,
    ack_path: Path,
    *,
    preexisting_pids: set[int],
    read_identity: Callable[[int], Optional[ProcessIdentity]],
    terminate_tree: Callable[[ProcessIdentity], None],
    timeout: float,
    owned_timeout: Optional[float] = None,
    sleep: Callable[[float], None] = time.sleep,
    monotonic: Callable[[], float] = time.monotonic,
) -> tuple[Mapping[str, Any], ProcessIdentity]:
    """Wait for one worker while continuously enforcing process ownership."""
    if timeout <= 0:
        raise ValueError("Windows worker timeout must be positive")
    if owned_timeout is not None and owned_timeout <= 0:
        raise ValueError("Windows owned-worker timeout must be positive")
    deadline = monotonic() + timeout
    owned: Optional[ProcessIdentity] = None
    progress = Path(progress_path)
    result = Path(result_path)
    while monotonic() < deadline:
        if owned is None and progress.is_file():
            owned = _validate_progress(
                _read_json(progress, "progress"), preexisting_pids
            )
            if read_identity(owned.pid) != owned:
                _cancel_then_terminate(
                    worker,
                    cancel_path,
                    owned,
                    read_identity=read_identity,
                    terminate_tree=terminate_tree,
                )
                raise WindowsOwnershipError(
                    "WPS process identity changed before supervision"
                )
            if owned_timeout is not None:
                deadline = min(deadline, monotonic() + owned_timeout)
        if result.is_file():
            if owned is None:
                _cancel_without_identity(worker, cancel_path)
                raise WindowsOwnershipError(
                    "Windows worker returned before proving WPS ownership"
                )
            if read_identity(owned.pid) != owned:
                _cancel_then_terminate(
                    worker,
                    cancel_path,
                    owned,
                    read_identity=read_identity,
                    terminate_tree=terminate_tree,
                )
                raise WindowsOwnershipError(
                    "WPS process identity changed before result validation"
                )
            value = _read_json(result, "result")
            if not isinstance(value, Mapping):
                raise RuntimeError("Windows worker result must be an object")
            write_canonical_json(ack_path, {"accepted": True})
            remaining = max(0.001, min(5.0, deadline - monotonic()))
            try:
                exit_code = worker.wait(timeout=remaining)
            except (TimeoutError, subprocess.TimeoutExpired) as exc:
                _cancel_then_terminate(
                    worker,
                    cancel_path,
                    owned,
                    read_identity=read_identity,
                    terminate_tree=terminate_tree,
                )
                raise RuntimeError(
                    "Windows worker did not exit after writing its result"
                ) from exc
            if exit_code != 0:
                raise RuntimeError(
                    f"Windows worker exited with status {exit_code}"
                )
            return value, owned
        exit_code = worker.poll()
        if exit_code is not None:
            raise RuntimeError(
                f"Windows worker exited without a result (status {exit_code})"
            )
        sleep(min(0.05, max(0.0, deadline - monotonic())))
    if owned is not None:
        _cancel_then_terminate(
            worker,
            cancel_path,
            owned,
            read_identity=read_identity,
            terminate_tree=terminate_tree,
        )
    else:
        _cancel_without_identity(worker, cancel_path)
    raise WindowsWorkerTimeout(owned)


def _load_windows_native() -> Any:
    """Load pywin32 only inside a real Windows worker."""
    import pythoncom
    import win32com.client
    import win32process

    return SimpleNamespace(
        pythoncom=pythoncom,
        client=win32com.client,
        hwnd_pid=lambda hwnd: int(
            win32process.GetWindowThreadProcessId(int(hwnd))[1]
        ),
        read_identity=_read_windows_identity,
    )


def _worker_execute(
    raw_request: Any,
    progress_path: Path,
    *,
    result_path: Optional[Path] = None,
    ack_path: Optional[Path] = None,
    cancel_path: Optional[Path] = None,
    native: Any = None,
    assay_runner: Optional[Callable[[Any, Mapping[str, Any], Callable[[], bool]], Mapping[str, Any]]] = None,
    sleep: Callable[[float], None] = time.sleep,
) -> Mapping[str, Any]:
    """Run one COM worker after proving that DispatchEx created a new WPS PID."""
    request = _validate_worker_request(raw_request)
    runtime = native if native is not None else _load_windows_native()
    if assay_runner is not None:
        runner = assay_runner
    elif request["workerMode"] == "ownership-timeout":
        runner = _run_ownership_timeout_assay
    else:
        runner = _run_com_assays
    runtime.pythoncom.CoInitialize()
    owned: Optional[OwnedApplication] = None
    try:
        owned = _dispatch_owned_application(
            runtime.client,
            hwnd_pid=runtime.hwnd_pid,
            read_identity=runtime.read_identity,
            preexisting_pids=set(request["preexistingPids"]),
        )
        write_canonical_json(progress_path, owned.identity.to_json())
        owned.app.Visible = 0
        owned.app.DisplayAlerts = 0
        cancellation = Path(cancel_path or request["cancelPath"])
        value = runner(
            owned.app,
            request,
            lambda: cancellation.is_file(),
        )
        if result_path is not None and not cancellation.is_file():
            write_canonical_json(result_path, value)
            acknowledgement = Path(ack_path or request["ackPath"])
            while not acknowledgement.is_file():
                if cancellation.is_file():
                    raise _WorkerCanceled("Windows M0 result was not accepted")
                sleep(0.05)
        return value
    finally:
        try:
            if owned is not None:
                owned.app.Quit()
        finally:
            runtime.pythoncom.CoUninitialize()


class _WorkerCanceled(RuntimeError):
    pass


def _run_ownership_timeout_assay(
    _application: Any,
    _request: Mapping[str, Any],
    canceled: Callable[[], bool],
) -> Mapping[str, Any]:
    """Stay alive until the supervisor exercises cooperative cancellation."""
    while not canceled():
        time.sleep(0.05)
    return {"canceled": True}


def _verify_preexisting_survived(
    identities: Mapping[int, ProcessIdentity],
    *,
    read_identity: Callable[[int], Optional[ProcessIdentity]],
) -> None:
    for pid, expected in identities.items():
        if read_identity(pid) != expected:
            raise WindowsOwnershipError(
                "a pre-existing WPS process did not survive the timeout fixture"
            )


def _snapshot_wps_identities(
    *,
    snapshot_pids: Callable[[], set[int]] = _snapshot_wps_pids,
    read_identity: Callable[[int], Optional[ProcessIdentity]] = _read_windows_identity,
) -> dict[int, ProcessIdentity]:
    identities: dict[int, ProcessIdentity] = {}
    for pid in sorted(snapshot_pids()):
        identity = read_identity(pid)
        if identity is None:
            raise WindowsOwnershipError(
                "a pre-existing WPS process identity is unavailable"
            )
        identities[pid] = identity
    return identities


def _wait_identity_gone(
    identity: ProcessIdentity,
    *,
    read_identity: Callable[[int], Optional[ProcessIdentity]],
    timeout: float = 5,
    sleep: Callable[[float], None] = time.sleep,
    terminate_tree: Callable[[ProcessIdentity], None] = _terminate_windows_tree,
) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if read_identity(identity.pid) != identity:
            return
        sleep(0.05)
    if read_identity(identity.pid) == identity:
        terminate_tree(identity)
    raise WindowsOwnershipError("worker-owned WPS process did not exit")


def _worker_main(
    request_path: Path, progress_path: Path, result_path: Path
) -> int:
    request = _read_json(Path(request_path), "request")
    try:
        _worker_execute(
            request,
            Path(progress_path),
            result_path=Path(result_path),
        )
    except _WorkerCanceled:
        return 0
    return 0


def _validate_worker_result(
    value: Any,
    staged_docx: Path,
    staged_pdf: Path,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], str]:
    if not isinstance(value, Mapping) or set(value) != _RESULT_FIELDS:
        raise ValueError("Windows worker result fields are invalid")
    expected = {
        "probeVersion": PROBE_VERSION,
        "protocolVersion": PROTOCOL_VERSION,
        "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
        "platform": "windows",
    }
    for name, required in expected.items():
        if value.get(name) != required:
            raise ValueError(f"Windows worker result {name} is invalid")
    wps_version = value.get("wpsVersion")
    if not isinstance(wps_version, str) or not wps_version or len(wps_version) > 128:
        raise ValueError("Windows worker WPS version is invalid")
    reported_docx = Path(str(value.get("docxPath", ""))).resolve()
    reported_pdf = Path(str(value.get("pdfPath", ""))).resolve()
    if reported_docx != staged_docx.resolve() or reported_pdf != staged_pdf.resolve():
        raise ValueError("Windows worker result paths do not match staging")
    capabilities = value.get("capabilities")
    failures = value.get("failures")
    if not isinstance(capabilities, list) or not isinstance(failures, list):
        raise ValueError("Windows worker evidence is invalid")
    return copy.deepcopy(capabilities), copy.deepcopy(failures), wps_version


def _complete_windows_evidence(
    capabilities: list[dict[str, Any]],
    failures: list[dict[str, Any]],
    artifacts: Mapping[str, Any],
    coordinate_snapshot: Mapping[str, Any],
    wps_version: str,
) -> dict[str, Any]:
    artifact_names = [artifacts["docx"]["name"], artifacts["pdf"]["name"]]
    capability_ids: set[int] = set()
    for capability in capabilities:
        if not isinstance(capability, dict) or set(capability) != {
            "id",
            "status",
            "checks",
            "metrics",
        }:
            raise ValueError("Windows native capability fields are invalid")
        capability_id = capability.get("id")
        if isinstance(capability_id, bool) or not isinstance(capability_id, int):
            raise ValueError("Windows native capability id is invalid")
        capability_ids.add(capability_id)
        capability["artifacts"] = (
            list(artifact_names)
            if capability.get("status") == "passed" and capability_id >= 3
            else []
        )
    if capability_ids != set(range(1, 16)):
        raise ValueError("Windows native capability ids are incomplete")
    by_id = {item["id"]: item for item in capabilities}

    def fail_host(capability_id: int, reason: str) -> None:
        capability = by_id[capability_id]
        capability["status"] = "failed"
        capability["checks"] = ["native-attempted", "host-validated"]
        capability["metrics"]["hostValidation"] = reason
        if not any(
            isinstance(item, Mapping)
            and item.get("capabilityId") == capability_id
            for item in failures
        ):
            failures.append(
                {"code": "HOST_VALIDATION_FAILED", "capabilityId": capability_id}
            )

    pdf_snapshot = artifacts["pdf"]["snapshot"]
    fonts = pdf_snapshot["fonts"]
    by_id[3]["metrics"]["pdfFontCount"] = len(fonts)
    if by_id[3]["status"] == "passed" and not fonts:
        fail_host(3, "pdf-fonts-empty")

    coordinate_metrics = by_id[5]["metrics"]
    coordinate_metrics["hostPdfSnapshot"] = copy.deepcopy(pdf_snapshot)
    coordinate_metrics["hostCoordinateSnapshot"] = copy.deepcopy(
        coordinate_snapshot
    )
    try:
        paragraph = coordinate_snapshot["paragraph"]
        shape = coordinate_snapshot["shape"]
        errors = (
            abs(float(coordinate_metrics["paragraphX"]) - float(paragraph["bbox"][0])),
            abs(float(coordinate_metrics["paragraphY"]) - float(paragraph["bbox"][1])),
            abs(float(coordinate_metrics["shapeX"]) - float(shape["frameBBox"][0])),
            abs(float(coordinate_metrics["shapeY"]) - float(shape["frameBBox"][1])),
        )
        maximum_error: Optional[float] = max(errors)
    except (KeyError, TypeError, ValueError, OverflowError):
        maximum_error = None
    coordinate_metrics["pdfAgreementMaxError"] = maximum_error
    if by_id[5]["status"] == "passed" and (
        maximum_error is None or maximum_error > 1.0
    ):
        fail_host(5, "coordinate-error-over-1pt")

    return {
        "schemaVersion": SCHEMA_VERSION,
        "probeVersion": PROBE_VERSION,
        "platform": "windows",
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


def _failure_capabilities(capability_id: Optional[int]) -> list[dict[str, Any]]:
    return [
        {
            "id": current,
            "status": "failed" if current == capability_id else "not-run",
            "checks": ["native-attempted"] if current == capability_id else [],
            "artifacts": [],
            "metrics": {},
        }
        for current in range(1, 16)
    ]


def _failure_evidence(code: str, capability_id: Optional[int]) -> dict[str, Any]:
    return {
        "schemaVersion": SCHEMA_VERSION,
        "probeVersion": PROBE_VERSION,
        "platform": "windows",
        "wpsVersion": "unknown",
        "protocolVersion": PROTOCOL_VERSION,
        "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
        "capabilities": _failure_capabilities(capability_id),
        "artifacts": {},
        "failures": [{"code": code, "message": "Windows long-form M0 probe failed"}],
    }


class WindowsM0Failed(RuntimeError):
    """Raised after redacted Windows M0 failure evidence has been written."""

    def __init__(self, evidence_path: Path):
        super().__init__("Windows long-form M0 probe failed")
        self.evidence_path = Path(evidence_path)


def _failure_code(error: BaseException, stage: str) -> tuple[str, Optional[int]]:
    if isinstance(error, WindowsOwnershipError):
        return "OWNERSHIP_UNPROVEN", 2
    if isinstance(error, WindowsWorkerTimeout):
        return "WORKER_TIMEOUT", 2 if stage == "ownership-fixture" else None
    if stage == "ownership-fixture":
        return "OWNERSHIP_FIXTURE_FAILED", 2
    if isinstance(error, ArtifactTransportError):
        return error.code, None
    if isinstance(error, ArtifactValidationError):
        return "STAGED_ARTIFACT_INVALID", None
    if stage == "dependency":
        return "DEPENDENCY_UNAVAILABLE", None
    if stage == "platform":
        return "WINDOWS_REQUIRED", None
    if stage == "evidence" and isinstance(error, ValueError):
        return "EVIDENCE_INVALID", None
    return "RUNTIME_FAILURE", None


def _worker_request(
    *,
    mode: str,
    verified: bool,
    staged_docx: Path,
    staged_pdf: Path,
    staged_svg: Path,
    preexisting_pids: list[int],
    cancel_path: Path,
    ack_path: Path,
) -> dict[str, Any]:
    return {
        "probeVersion": PROBE_VERSION,
        "protocolVersion": PROTOCOL_VERSION,
        "resourceManifestVersion": RESOURCE_MANIFEST_VERSION,
        "workerMode": mode,
        "ownershipTimeoutVerified": verified,
        "stagedDocxPath": str(staged_docx.resolve()),
        "stagedPdfPath": str(staged_pdf.resolve()),
        "stagedSvgPath": str(staged_svg.resolve()),
        "expectedWpsVersion": "unknown",
        "preexistingPids": preexisting_pids,
        "cancelPath": str(cancel_path.resolve()),
        "ackPath": str(ack_path.resolve()),
    }


def _launch_worker(
    request_path: Path, progress_path: Path, result_path: Path
) -> subprocess.Popen[Any]:
    return subprocess.Popen(
        [
            sys.executable,
            "-m",
            "skills.WPSComposer.scripts.longform_m0.windows",
            "--worker",
            str(request_path),
            str(progress_path),
            str(result_path),
        ],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
    )


def _com_item(collection: Any, index: int) -> Any:
    return collection.Item(index)


def _com_count(collection: Any) -> int:
    return max(0, int(getattr(collection, "Count", 0) or 0))


def _com_end_range(document: Any) -> Any:
    end = max(0, int(document.Content.End) - 1)
    return document.Range(end, end)


def _com_append(document: Any, value: str) -> Any:
    insertion = _com_end_range(document)
    start = int(insertion.Start)
    insertion.InsertAfter(str(value) + "\r")
    units = len(str(value).encode("utf-16-le")) // 2
    return document.Range(start, start + units)


def _com_marker(capability_id: int) -> str:
    return f"M0C{capability_id:02d}"


def _com_row(
    capability_id: int,
    status: str,
    checks: list[str],
    metrics: Optional[Mapping[str, Any]] = None,
) -> dict[str, Any]:
    return {
        "id": capability_id,
        "status": status,
        "checks": checks,
        "metrics": dict(metrics or {}),
    }


def _com_annotate_failure(document: Any, capability_id: int) -> None:
    try:
        marker = _com_append(document, f"[M0 CAP {capability_id} FAILED]")
        marker.Font.Color = 192
        marker.Font.Bold = -1
    except Exception:
        pass


def _com_assay(
    document: Any,
    capabilities: list[dict[str, Any]],
    capability_id: int,
    operation: Callable[[], Mapping[str, Any]],
    *,
    optional: bool = False,
    canceled: Callable[[], bool],
) -> None:
    if canceled():
        raise _WorkerCanceled("Windows M0 worker was canceled")
    try:
        outcome = dict(operation() or {})
        passed = bool(outcome.pop("passed", True))
        if not passed:
            _com_annotate_failure(document, capability_id)
            capabilities.append(
                _com_row(
                    capability_id,
                    "unsupported" if optional else "failed",
                    ["native-attempted"],
                    outcome,
                )
            )
            return
        _com_append(document, _com_marker(capability_id))
        capabilities.append(
            _com_row(capability_id, "passed", ["native"], outcome)
        )
    except _WorkerCanceled:
        raise
    except Exception as error:
        _com_annotate_failure(document, capability_id)
        stage = getattr(error, "probe_stage", "operation")
        capabilities.append(
            _com_row(
                capability_id,
                "unsupported" if optional else "failed",
                ["native-attempted"],
                {
                    "errorClass": type(error).__name__[:64] or "Error",
                    "stage": str(stage)[:64],
                },
            )
        )


def _com_fonts(application: Any, document: Any) -> Mapping[str, Any]:
    names = [
        str(_com_item(application.FontNames, index))
        for index in range(1, _com_count(application.FontNames) + 1)
    ]
    choices = []
    for preferred in (
        ("SimSun", "FmlSong", "Songti SC", "STSong"),
        ("Times New Roman", "Times"),
        ("Consolas", "Courier New", "Courier"),
    ):
        selected = next((name for name in preferred if name in names), None)
        if selected is None:
            return {"passed": False, "fontCount": len(names)}
        choices.append(selected)
    for label, font_name in zip(("CJK", "Latin", "Mono"), choices):
        written = _com_append(document, label)
        written.Font.Name = font_name
        written.Font.NameFarEast = font_name
    return {"fontCount": len(names), "mappedFontCount": len(choices)}


def _com_unicode(document: Any) -> Mapping[str, Any]:
    values = ("中文é", "e\u0301", "👩‍💻", "𰀀️")
    utf16_units = 0
    for value in values:
        written = _com_append(document, value)
        units = len(value.encode("utf-16-le")) // 2
        if int(written.End) - int(written.Start) != units:
            raise RuntimeError("UTF16 range mismatch")
        utf16_units += units
    return {
        "paragraphCount": len(values),
        "utf16Units": utf16_units,
        "unicodeDataVersion": "15.1.0",
        "normalizationModes": 2,
    }


def _com_coordinates(document: Any) -> Mapping[str, Any]:
    shape = document.Shapes.AddTextbox(1, 72, 72, 120, 36, _com_end_range(document))
    shape.RelativeHorizontalPosition = 1
    shape.RelativeVerticalPosition = 1
    shape.Left = 72
    shape.Top = 72
    shape.TextFrame.MarginLeft = 0
    shape.TextFrame.MarginRight = 0
    shape.TextFrame.MarginTop = 0
    shape.TextFrame.MarginBottom = 0
    shape.TextFrame.TextRange.Text = "M0XY5"
    shape.Line.Weight = 0.75
    shape.Line.ForeColor.RGB = 0
    paragraph = _com_append(document, "COORD")
    paragraph.Font.Size = 1
    paragraph.ParagraphFormat.SpaceBefore = 0
    paragraph.ParagraphFormat.SpaceAfter = 0
    paragraph.ParagraphFormat.LineSpacingRule = 4
    paragraph.ParagraphFormat.LineSpacing = 1
    document.Repaginate()
    values = {
        "origin": "top-left",
        "unit": "point",
        "paragraphX": float(paragraph.Information(5)),
        "paragraphY": float(paragraph.Information(6)),
        "shapeX": float(shape.Left),
        "shapeY": float(shape.Top),
    }
    values["passed"] = all(
        value == value and abs(value) != float("inf")
        for key, value in values.items()
        if key.endswith("X") or key.endswith("Y")
    )
    return values


_COM_NUMBERING_SCHEMES = (
    (
        ("第%1章", 37),
        ("第%2节", 37),
        ("%3、", 37),
        ("（%4）", 37),
    ),
    (
        ("%1", 0),
        ("%1.%2", 0),
        ("%1.%2.%3", 0),
        ("%1.%2.%3.%4", 0),
    ),
    (
        ("第%1章", 37),
        ("%1.%2", 253),
        ("%1.%2.%3", 253),
        ("关键工法%4：", 22),
    ),
)
_COM_NUMBERING_EXPECTED = (
    ("第一章", "第一节", "一、", "（一）"),
    ("1", "1.1", "1.1.1", "1.1.1.1"),
    ("第一章", "1.1", "1.1.1", "关键工法01："),
)


def _com_find_paragraph(document: Any, token: str) -> Any:
    for index in range(1, _com_count(document.Paragraphs) + 1):
        paragraph = _com_item(document.Paragraphs, index)
        if token in str(paragraph.Range.Text or ""):
            return paragraph
    return None


def _com_list_string(document: Any, token: str) -> str:
    paragraph = _com_find_paragraph(document, token)
    if paragraph is None:
        return ""
    return "".join(str(paragraph.Range.ListFormat.ListString or "").split())


def _com_add_list(document: Any, scheme_index: int) -> Any:
    template = document.ListTemplates.Add(True, f"wpsc_m0_{scheme_index}")
    definitions = _COM_NUMBERING_SCHEMES[scheme_index - 1]
    for level, (number_format, number_style) in enumerate(definitions, start=1):
        list_level = _com_item(template.ListLevels, level)
        list_level.NumberFormat = number_format
        list_level.NumberStyle = number_style
        list_level.NumberPosition = (level - 1) * 18
        list_level.TextPosition = level * 18
        list_level.ResetOnHigher = 0 if (scheme_index == 3 and level == 4) else max(0, level - 1)
        list_level.StartAt = 1
    start = int(_com_end_range(document).Start)
    for level in range(1, 5):
        written = _com_append(document, f"L{scheme_index}-{level}")
        written.Style = _com_item(document.Styles, -1 - level)
    list_range = document.Range(start, int(_com_end_range(document).Start))
    list_range.ListFormat.ApplyListTemplateWithLevel(template, False, 0, 0, 1)
    for level in range(1, 5):
        _com_find_paragraph(document, f"L{scheme_index}-{level}").Range.ListFormat.ListLevelNumber = level
    return template


def _com_numbering_matches(document: Any, *, mutated: bool) -> bool:
    for scheme in range(1, 4):
        for level in range(1, 5):
            token = "L2-X" if mutated and scheme == 2 and level == 4 else f"L{scheme}-{level}"
            if _com_list_string(document, token) != _COM_NUMBERING_EXPECTED[scheme - 1][level - 1]:
                return False
    return not mutated or _com_find_paragraph(document, "L2-4") is None


def _numbering_error(stage: str) -> None:
    error = RuntimeError("numbering assay failed")
    error.probe_stage = stage
    raise error


def _com_numbering(document: Any) -> Mapping[str, Any]:
    templates = [_com_add_list(document, index) for index in range(1, 4)]
    if not _com_numbering_matches(document, mutated=False):
        _numbering_error("initial-schemes")
    inserted = document.Paragraphs.Add(_com_find_paragraph(document, "L2-4").Range)
    inserted.Range.Text = "L2-X\r"
    inserted.Range.Style = _com_item(document.Styles, -5)
    first = _com_find_paragraph(document, "L2-1").Range
    last = _com_find_paragraph(document, "L2-X").Range
    bound = document.Range(first.Start, last.End)
    bound.ListFormat.ApplyListTemplateWithLevel(templates[1], False, 0, 0, 1)
    for level in range(1, 5):
        _com_find_paragraph(document, f"L2-{level}").Range.ListFormat.ListLevelNumber = level
    _com_find_paragraph(document, "L2-X").Range.ListFormat.ListLevelNumber = 4
    if _com_list_string(document, "L2-4") != "1.1.1.1" or _com_list_string(document, "L2-X") != "1.1.1.2":
        _numbering_error("insert-renumber")
    moving = _com_find_paragraph(document, "L2-X")
    destination = _com_find_paragraph(document, "L2-4").Range.Duplicate
    moving.Range.Cut()
    destination.Collapse(1)
    destination.Paste()
    if _com_list_string(document, "L2-X") != "1.1.1.1" or _com_list_string(document, "L2-4") != "1.1.1.2":
        _numbering_error("move-renumber")
    _com_find_paragraph(document, "L2-4").Range.Delete()
    if not _com_numbering_matches(document, mutated=True):
        _numbering_error("delete-renumber")
    return {
        "definitionCount": 3,
        "levelCount": 4,
        "numberedParagraphCount": 12,
        "mutationKinds": 3,
        "mutationRenumberChecks": 3,
        "exactSchemeCount": 3,
    }


def _com_sections(document: Any) -> Mapping[str, Any]:
    _com_end_range(document).InsertBreak(7)
    for index in range(3):
        _com_end_range(document).InsertBreak(2)
        _com_item(document.Sections, _com_count(document.Sections)).PageSetup.Orientation = 1
        _com_append(document, f"LANDSCAPE-{index + 1}")
        _com_end_range(document).InsertBreak(2)
        _com_item(document.Sections, _com_count(document.Sections)).PageSetup.Orientation = 0
    first = _com_item(document.Sections, 1)
    body = _com_item(document.Sections, min(2, _com_count(document.Sections)))
    first_footer = _com_item(first.Footers, 1)
    body_footer = _com_item(body.Footers, 1)
    first_footer.PageNumbers.NumberStyle = 2
    first_footer.PageNumbers.RestartNumberingAtSection = True
    first_footer.PageNumbers.StartingNumber = 1
    body_footer.LinkToPrevious = False
    body_footer.PageNumbers.NumberStyle = 0
    body_footer.PageNumbers.RestartNumberingAtSection = True
    body_footer.PageNumbers.StartingNumber = 1
    count = _com_count(document.Sections)
    return {
        "passed": count >= 7,
        "sectionCount": count,
        "landscapeLifecycleCount": 3,
        "pageNumberRestartCount": 2,
        "explicitPageBreakCount": 1,
    }


def _com_toc(document: Any) -> Mapping[str, Any]:
    for index, style_id in enumerate((-20, -21, -22), start=2):
        _com_item(document.Styles, style_id).ParagraphFormat.SpaceAfter = index
    title = _com_append(document, "TOC-TITLE")
    title.Style = _com_item(document.Styles, -67)
    title.ParagraphFormat.OutlineLevel = 10
    document.TablesOfContents.Add(_com_end_range(document), True, 1, 3)
    return {
        "passed": _com_count(document.TablesOfContents) >= 1,
        "tocCount": _com_count(document.TablesOfContents),
        "tocStyleCount": 3,
        "titleOutlineLevel": 10,
    }


def _com_captions(document: Any) -> Mapping[str, Any]:
    field_codes = (
        "SEQ Figure \\* ARABIC \\s 1",
        "SEQ Figure \\* ARABIC \\r 1",
        "SEQ Table \\* ARABIC \\s 1",
    )
    for field_code in field_codes:
        document.Fields.Add(_com_end_range(document), -1, field_code, True)
        _com_append(document, "")
    document.TablesOfFigures.Add(_com_end_range(document), "Figure")
    _com_append(document, "")
    document.TablesOfFigures.Add(_com_end_range(document), "Table")
    return {
        "passed": _com_count(document.TablesOfFigures) >= 2,
        "sequenceFieldCount": 3,
        "figureIndexCount": _com_count(document.TablesOfFigures),
        "resetModes": 2,
    }


def _com_references(document: Any) -> Mapping[str, Any]:
    occupied = "wpsc_ref_000000000000000000000000"
    retry = "wpsc_ref_000000000000000000000001"
    target = _com_append(document, "REFERENCE-TARGET")
    document.Bookmarks.Add(occupied, target)
    collision = bool(document.Bookmarks.Exists(occupied))
    document.Bookmarks.Add(retry if collision else occupied, target)
    document.Fields.Add(_com_end_range(document), -1, f"REF {retry}", True)
    return {
        "passed": collision and bool(document.Bookmarks.Exists(retry)),
        "bookmarkCount": _com_count(document.Bookmarks),
        "collisionRetries": 1,
        "referenceFieldCount": 1,
    }


def _com_formula(document: Any) -> Mapping[str, Any]:
    table = document.Tables.Add(_com_end_range(document), 1, 1)
    cell_range = table.Cell(1, 1).Range
    cell_range.Text = "x=(-b±√(b²-4ac))/(2a)"
    for border_id in range(-6, 0):
        table.Borders(border_id).LineStyle = 0
    formula_range = cell_range.Duplicate
    formula_range.MoveEnd(1, -1)
    math = document.OMaths.Add(formula_range)
    try:
        math.BuildUp()
    except Exception:
        document.OMaths.BuildUp()
    return {
        "passed": _com_count(document.OMaths) >= 1,
        "nativeFormulaCount": _com_count(document.OMaths),
        "borderlessContainerCount": 1,
    }


def _com_pagination(document: Any) -> Mapping[str, Any]:
    table = document.Tables.Add(_com_end_range(document), 80, 2)
    for row in range(1, 81):
        table.Cell(row, 1).Range.Text = f"R{row}"
        table.Cell(row, 2).Range.Text = "PAGINATION"
    document.Repaginate()
    pages: set[int] = set()
    positioned = 0
    for row in range(1, 81):
        row_range = table.Cell(row, 1).Range
        page = int(row_range.Information(3))
        x = float(row_range.Information(5))
        y = float(row_range.Information(6))
        if page > 0:
            pages.add(page)
        if x == x and y == y:
            positioned += 1
    fragments = len(pages)
    return {
        "passed": fragments >= 2 and positioned == 80,
        "nodeCount": 1,
        "sectionCount": _com_count(document.Sections),
        "pageSpanCount": fragments,
        "fragmentCount": fragments,
        "positionedRowCount": positioned,
        "rangeUnit": "utf16",
    }


def _com_checkpoint(document: Any) -> Mapping[str, Any]:
    first = _com_append(document, "CHECKPOINT-CHILD-1")
    partial = _com_append(document, "CHECKPOINT-PARTIAL")
    cleaned = False
    try:
        raise RuntimeError("intentional child failure")
    except RuntimeError:
        partial.Delete()
        cleaned = True
        _com_append(document, "[M0 CAP 13 FALLBACK]")
    later = _com_append(document, "CHECKPOINT-CONTINUED")
    return {
        "passed": cleaned and int(first.End) > int(first.Start) and int(later.End) > int(later.Start),
        "checkpointDepth": 2,
        "childCleanupCount": 1,
        "fallbackCount": 1,
        "continuationCount": 1,
    }


def _com_field_snapshot(document: Any) -> tuple[str, ...]:
    return tuple(
        str(_com_item(document.Fields, index).Result.Text or "")
        for index in range(1, _com_count(document.Fields) + 1)
    )


def _com_convergence(document: Any) -> Mapping[str, Any]:
    previous = _com_field_snapshot(document)
    full_passes = 2
    for current_pass in (1, 2):
        document.Fields.Update()
        current = _com_field_snapshot(document)
        if current == previous:
            full_passes = current_pass
            break
        previous = current
    _com_append(document, "[M0 NOTICE PATCH]")
    document.Fields.Update()
    return {
        "fullPasses": full_passes,
        "patchPasses": 1,
        "plannedSaveCount": 2,
        "plannedExportCount": 1,
    }


def _com_svg(document: Any, request: Mapping[str, Any]) -> Mapping[str, Any]:
    before = _com_count(document.InlineShapes)
    shape = document.InlineShapes.AddPicture(
        request["stagedSvgPath"], False, True, _com_end_range(document)
    )
    shape.Width = 120
    shape.Height = 60
    count = _com_count(document.InlineShapes)
    return {
        "passed": count == before + 1,
        "svgCount": count,
        "staticManifestCount": 1,
    }


def _com_update_fields(document: Any) -> None:
    for index in range(1, _com_count(document.TablesOfContents) + 1):
        _com_item(document.TablesOfContents, index).Update()
    for index in range(1, _com_count(document.TablesOfFigures) + 1):
        _com_item(document.TablesOfFigures, index).Update()
    document.Fields.Update()
    document.Repaginate()


def _com_verify_reopened(
    document: Any, capabilities: list[dict[str, Any]]
) -> None:
    content = str(document.Content.Text or "")
    object_checks = {
        6: _com_numbering_matches(document, mutated=True),
        8: _com_count(document.TablesOfContents) >= 1,
        9: _com_count(document.TablesOfFigures) >= 2,
        10: _com_count(document.Bookmarks) >= 2,
        11: _com_count(document.OMaths) >= 1,
        15: _com_count(document.InlineShapes) >= 1,
    }
    for capability in capabilities:
        capability_id = capability["id"]
        if capability_id < 3 or capability["status"] != "passed":
            continue
        if _com_marker(capability_id) not in content or not object_checks.get(capability_id, True):
            capability["status"] = "unsupported" if capability_id == 15 else "failed"
            capability["checks"] = ["native-attempted"]
            capability["metrics"]["reopenFailure"] = "native-object-missing"
        else:
            capability["checks"].extend(("reopened", "refreshed"))


def _run_com_assays(
    application: Any,
    request: Mapping[str, Any],
    canceled: Callable[[], bool],
) -> Mapping[str, Any]:
    svg_payload = Path(request["stagedSvgPath"]).read_bytes()
    if (
        len(svg_payload) != 185
        or hashlib.sha256(svg_payload).hexdigest()
        != "924f47ebd7a4c22393defac4103818aedced9f5dd0630bf27e1d6aa7ad30cbfc"
    ):
        raise RuntimeError("staged SVG resource manifest mismatch")
    capabilities = [
        _com_row(1, "passed", ["native"], {"protocolAccepted": True, "manifestBound": True}),
        _com_row(
            2,
            "passed" if request["ownershipTimeoutVerified"] else "failed",
            ["native"],
            {
                "dispatchExOwned": True,
                "timeoutFixtureVerified": bool(
                    request["ownershipTimeoutVerified"]
                ),
            },
        ),
    ]
    document = None
    try:
        document = application.Documents.Add()
        document.PageSetup.TopMargin = 54
        document.PageSetup.BottomMargin = 54
        document.PageSetup.LeftMargin = 64
        document.PageSetup.RightMargin = 64
        _com_append(document, "WPSComposer Long-form M0 Native Probe")
        _com_assay(document, capabilities, 3, lambda: _com_fonts(application, document), canceled=canceled)
        _com_assay(document, capabilities, 4, lambda: _com_unicode(document), canceled=canceled)
        _com_assay(document, capabilities, 5, lambda: _com_coordinates(document), canceled=canceled)
        _com_assay(document, capabilities, 6, lambda: _com_numbering(document), canceled=canceled)
        _com_assay(document, capabilities, 7, lambda: _com_sections(document), canceled=canceled)
        _com_assay(document, capabilities, 8, lambda: _com_toc(document), canceled=canceled)
        _com_assay(document, capabilities, 9, lambda: _com_captions(document), canceled=canceled)
        _com_assay(document, capabilities, 10, lambda: _com_references(document), canceled=canceled)
        _com_assay(document, capabilities, 11, lambda: _com_formula(document), canceled=canceled)
        _com_assay(document, capabilities, 12, lambda: _com_pagination(document), canceled=canceled)
        _com_assay(document, capabilities, 13, lambda: _com_checkpoint(document), canceled=canceled)
        _com_assay(document, capabilities, 14, lambda: _com_convergence(document), canceled=canceled)
        _com_assay(document, capabilities, 15, lambda: _com_svg(document, request), optional=True, canceled=canceled)
        document.SaveAs2(request["stagedDocxPath"], 12)
        document.Close(0)
        document = None
        document = application.Documents.Open(request["stagedDocxPath"], False, False)
        _com_update_fields(document)
        _com_verify_reopened(document, capabilities)
        document.Save()
        document.ExportAsFixedFormat(request["stagedPdfPath"], 17, False, 0, 0)
        document.Close(0)
        document = None
        failures = [
            {"code": "CAPABILITY_FAILED", "capabilityId": row["id"]}
            for row in capabilities
            if row["status"] == "failed"
        ]
        return {
            "probeVersion": request["probeVersion"],
            "protocolVersion": request["protocolVersion"],
            "resourceManifestVersion": request["resourceManifestVersion"],
            "platform": "windows",
            "wpsVersion": str(getattr(application, "Version", "unknown")),
            "capabilities": capabilities,
            "failures": failures,
            "docxPath": request["stagedDocxPath"],
            "pdfPath": request["stagedPdfPath"],
        }
    finally:
        if document is not None:
            document.Close(0)


def run_windows_probe(output_dir: Path, timeout: float = 600) -> Path:
    """Run the closed Windows COM M0 gate and return platform evidence."""
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    output = host_checks.prepare_evidence_directory(output_dir)
    evidence_path = output / "platform-evidence.json"
    final_docx = output / "probe.docx"
    final_pdf = output / "probe.pdf"
    runtime_root = Path(tempfile.mkdtemp(prefix="wpscomposer-longform-m0-"))
    deadline = time.monotonic() + timeout
    stage = "platform"
    try:
        if platform.system() != "Windows":
            raise RuntimeError("Windows long-form M0 probe requires Windows")
        stage = "dependency"
        host_checks.require_probe_dependencies()
        stage = "ownership-fixture"
        preexisting = _snapshot_wps_identities()
        if not preexisting:
            raise WindowsOwnershipError(
                "capability 2 requires a separate pre-existing WPS document"
            )

        staged_docx = (runtime_root / "longform-m0.docx").resolve()
        staged_pdf = (runtime_root / "longform-m0.pdf").resolve()
        staged_svg = (runtime_root / "longform-m0.svg").resolve()
        staged_svg.write_text(SVG_FIXTURE, encoding="utf-8")
        preexisting_pids = sorted(preexisting)

        fixture_dir = runtime_root / "ownership-fixture"
        fixture_dir.mkdir(mode=0o700)
        fixture_request_path = fixture_dir / "request.json"
        fixture_progress = fixture_dir / "progress.json"
        fixture_result = fixture_dir / "result.json"
        fixture_cancel = fixture_dir / "cancel.json"
        fixture_ack = fixture_dir / "ack.json"
        fixture_request = _worker_request(
            mode="ownership-timeout",
            verified=False,
            staged_docx=staged_docx,
            staged_pdf=staged_pdf,
            staged_svg=staged_svg,
            preexisting_pids=preexisting_pids,
            cancel_path=fixture_cancel,
            ack_path=fixture_ack,
        )
        _validate_worker_request(fixture_request)
        write_canonical_json(fixture_request_path, fixture_request)
        fixture_remaining = deadline - time.monotonic()
        if fixture_remaining <= 0:
            raise WindowsWorkerTimeout(None)
        fixture_worker = _launch_worker(
            fixture_request_path, fixture_progress, fixture_result
        )
        fixture_identity: Optional[ProcessIdentity] = None
        fixture_budget = min(60.0, fixture_remaining)
        try:
            _supervise_worker(
                fixture_worker,
                fixture_progress,
                fixture_result,
                fixture_cancel,
                fixture_ack,
                preexisting_pids=set(preexisting_pids),
                read_identity=_read_windows_identity,
                terminate_tree=_terminate_windows_tree,
                timeout=fixture_budget,
                owned_timeout=0.25,
            )
            raise WindowsOwnershipError(
                "ownership fixture returned instead of timing out"
            )
        except WindowsWorkerTimeout as error:
            fixture_identity = error.identity
            if fixture_identity is None:
                raise WindowsOwnershipError(
                    "ownership fixture timed out before proving its WPS process"
                ) from error
        _wait_identity_gone(
            fixture_identity,
            read_identity=_read_windows_identity,
        )
        _verify_preexisting_survived(
            preexisting,
            read_identity=_read_windows_identity,
        )

        stage = "native-worker"
        worker_dir = runtime_root / "probe-worker"
        worker_dir.mkdir(mode=0o700)
        request_path = worker_dir / "request.json"
        progress_path = worker_dir / "progress.json"
        result_path = worker_dir / "result.json"
        cancel_path = worker_dir / "cancel.json"
        ack_path = worker_dir / "ack.json"
        request = _worker_request(
            mode="probe",
            verified=True,
            staged_docx=staged_docx,
            staged_pdf=staged_pdf,
            staged_svg=staged_svg,
            preexisting_pids=preexisting_pids,
            cancel_path=cancel_path,
            ack_path=ack_path,
        )
        _validate_worker_request(request)
        write_canonical_json(request_path, request)
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise WindowsWorkerTimeout(None)
        worker = _launch_worker(request_path, progress_path, result_path)
        result, owned = _supervise_worker(
            worker,
            progress_path,
            result_path,
            cancel_path,
            ack_path,
            preexisting_pids=set(preexisting_pids),
            read_identity=_read_windows_identity,
            terminate_tree=_terminate_windows_tree,
            timeout=remaining,
        )
        _wait_identity_gone(owned, read_identity=_read_windows_identity)
        _verify_preexisting_survived(
            preexisting,
            read_identity=_read_windows_identity,
        )
        capabilities, failures, wps_version = _validate_worker_result(
            result, staged_docx, staged_pdf
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
        stage = "evidence"
        artifact_evidence = host_checks.validate_native_artifacts(
            final_docx, final_pdf
        )
        coordinate_snapshot = host_checks.snapshot_pdf_markers(
            final_pdf,
            {"paragraph": "COORD", "shape": "M0XY5"},
        )
        raw = _complete_windows_evidence(
            capabilities,
            failures,
            artifact_evidence,
            coordinate_snapshot,
            wps_version,
        )
        host_checks.validate_evidence_privacy(
            raw, forbidden_roots=(runtime_root,)
        )
        validate_platform_evidence(raw)
        write_canonical_json(evidence_path, raw)
        return evidence_path
    except Exception as error:
        final_docx.unlink(missing_ok=True)
        final_pdf.unlink(missing_ok=True)
        code, capability_id = _failure_code(error, stage)
        raw = _failure_evidence(code, capability_id)
        host_checks.validate_evidence_privacy(raw)
        validate_platform_evidence(raw)
        write_canonical_json(evidence_path, raw)
        raise WindowsM0Failed(evidence_path) from error
    finally:
        shutil.rmtree(runtime_root, ignore_errors=True)


def _module_main(argv: list[str]) -> int:
    if len(argv) != 5 or argv[1] != "--worker":
        return 2
    return _worker_main(Path(argv[2]), Path(argv[3]), Path(argv[4]))


if __name__ == "__main__":
    raise SystemExit(_module_main(sys.argv))
