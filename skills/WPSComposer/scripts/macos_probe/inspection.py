"""macOS WPS JSAPI backend for inspecting existing documents.

Reads structured content (slides, shapes, text, tables) from an existing
Office file by driving the real WPS layout engine through the JSAPI bridge.
This is the read-only counterpart to ``conversion.py``: it stages the source
file, issues an inspect command to the WPS add-in, and returns the JSON
snapshot produced by the add-in — without converting to PDF.
"""

from __future__ import annotations

from pathlib import Path
import shutil
import tempfile
import time
from typing import Any, Callable, Optional

from .._dispatch import WPSUnavailable
from ..artifact_transport import (
    ArtifactTransportError,
    publish_artifact,
    validate_office_package,
)
from .bridge import LoopbackBridge
from .models import PathPolicy, ProtocolError
from .runtime import ProbeRuntime


ORIGINS = {
    "http://127.0.0.1:3889",
    "http://127.0.0.1:3890",
    "http://127.0.0.1:3891",
}
# Maps a document extension (lower-cased, dot-prefixed) to the
# (component, method) pair used to inspect it through the JSAPI bridge.
INSPECTABLE = {
    ".ppt": ("presentation", "inspect_presentation"),
    ".pptx": ("presentation", "inspect_presentation"),
    ".pptm": ("presentation", "inspect_presentation"),
    ".pps": ("presentation", "inspect_presentation"),
    ".ppsx": ("presentation", "inspect_presentation"),
    ".ppsm": ("presentation", "inspect_presentation"),
    ".doc": ("writer", "inspect_document"),
    ".docx": ("writer", "inspect_document"),
    ".docm": ("writer", "inspect_document"),
    ".xls": ("spreadsheet", "inspect_workbook"),
    ".xlsx": ("spreadsheet", "inspect_workbook"),
    ".xlsm": ("spreadsheet", "inspect_workbook"),
}

# Formats that support edit through the JSAPI bridge.
EDITABLE = {
    ".pptx": ("presentation", "edit_presentation"),
}


class InspectionError(RuntimeError):
    """Stable runtime error returned by the macOS inspection backend."""

    def __init__(self, *, code: str, source: str, component: str, message: str):
        super().__init__(message)
        self.code = code
        self.source = source
        self.component = component
        self.message = message

    def to_dict(self) -> dict:
        return {
            "code": self.code,
            "source": self.source,
            "component": self.component,
            "message": self.message,
        }


def _error(source: str, component: str, code: str, message: str) -> InspectionError:
    return InspectionError(
        code=code, source=source, component=component, message=message
    )


def _wait_for_registration(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    component: str,
    deadline: float,
) -> None:
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


def _redact_staging(message: str, staging_dir: Path) -> str:
    return str(message).replace(str(staging_dir), "<wps-staging>")


def _run_inspection(
    source: Path,
    component: str,
    method: str,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    timeout: float,
    *,
    include_text: bool = True,
    max_shapes: Optional[int] = None,
) -> dict[str, Any]:
    if runtime.staging_dir is None:
        raise _error(
            str(source), component, "STAGING_UNAVAILABLE",
            "WPS container staging session was not created",
        )
    runtime.prepare_profiles()
    runtime.start_servers()
    runtime.activate_component(component)
    deadline = time.monotonic() + timeout
    try:
        _wait_for_registration(bridge, runtime, component, deadline)
    except TimeoutError as exc:
        raise _error(
            str(source), component, "INSPECTION_COMMAND_FAILED",
            "Timed out waiting for the WPS add-in to register",
        ) from exc

    policy = PathPolicy((runtime.staging_dir,))
    staged_source = policy.require_allowed(
        runtime.staging_dir / f"source{source.suffix.lower()}"
    )
    try:
        shutil.copy2(source, staged_source)
    except OSError as exc:
        raise _error(
            str(source), component, "STAGING_SAVE_FAILED",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc

    params: dict[str, Any] = {"sourcePath": str(staged_source)}
    if not include_text:
        params["includeText"] = False
        params["includeCells"] = False
    if max_shapes is not None:
        limit = int(max_shapes)
        # Send under every format-specific key so each JS handler finds
        # the one it expects: maxShapes (PPT), maxElements (DOCX),
        # maxCells (XLSX).
        params["maxShapes"] = limit
        params["maxElements"] = limit
        params["maxCells"] = limit

    command = bridge.issue(component, method, params)
    try:
        result = bridge.wait_result(
            command.id, max(0.0, deadline - time.monotonic())
        )
    except TimeoutError as exc:
        bridge.state.cancel(command.id)
        raise _error(
            str(source), component, "INSPECTION_COMMAND_FAILED",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    if not result.ok:
        details = dict(result.error or {})
        raise _error(
            str(source), component, "INSPECTION_COMMAND_FAILED",
            _redact_staging(
                str(details.get("message") or "WPS inspection failed"),
                runtime.staging_dir,
            ),
        )
    value = dict(result.value or {})
    # Replace the staged path with the user-facing original so snapshots do
    # not leak the private WPS container path.
    if isinstance(value.get("path"), str):
        value["path"] = str(source)
    return value


def inspect_macos(
    source: str,
    *,
    include_text: bool = True,
    max_shapes: Optional[int] = None,
    enabled: Optional[bool] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 90,
) -> dict[str, Any]:
    """Inspect an existing Office file through the WPS JSAPI bridge.

    Returns the structured snapshot (slides / shapes / text / tables)
    produced by the add-in. Currently supports presentation formats only.
    """
    _ = enabled  # parity with conversion_macos signature; not yet gated
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    source_path = Path(source).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"Source file not found: {source_path}")
    suffix = source_path.suffix.lower()
    try:
        component, method = INSPECTABLE[suffix]
    except KeyError as exc:
        supported = ", ".join(sorted(INSPECTABLE))
        raise ValueError(
            f"Inspection is not yet supported for '{suffix}'. "
            f"Supported: {supported}."
        ) from exc

    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    runtime_root = Path(
        tempfile.mkdtemp(prefix="wpscomposer-macos-inspect-")
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
                    return _run_inspection(
                        source_path, component, method, bridge, runtime,
                        timeout,
                        include_text=include_text,
                        max_shapes=max_shapes,
                    )
            except InspectionError:
                raise
            except Exception as exc:
                if not getattr(runtime, "registration_restored", True):
                    recovery = runtime.runtime_dir / "registration-recovery"
                    raise _error(
                        str(source_path), component,
                        "REGISTRATION_RESTORE_FAILED",
                        f"WPS registration restore failed; recovery "
                        f"retained at {recovery}",
                    ) from exc
                raise _error(
                    str(source_path), component,
                    "INSPECTION_COMMAND_FAILED",
                    "Mac WPS inspection command failed",
                ) from exc
    finally:
        if runtime is None or getattr(runtime, "registration_restored", True):
            shutil.rmtree(runtime_root, ignore_errors=True)



def edit_macos(
    source: str,
    patches: list,
    output: Optional[str] = None,
    *,
    enabled: Optional[bool] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 120,
    atomic: bool = True,
    raise_on_error: bool = False,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Apply patches to an existing presentation through the WPS JSAPI bridge.

    Saves the result to *output* (or back to *source* when omitted) and
    returns ``{"path": ..., "patches": [...]}``.
    """
    _ = enabled
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    if not patches:
        raise ValueError("at least one patch is required")
    source_path = Path(source).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"Source file not found: {source_path}")
    suffix = source_path.suffix.lower()
    try:
        component, method = EDITABLE[suffix]
    except KeyError as exc:
        supported = ", ".join(sorted(EDITABLE))
        raise ValueError(
            f"Editing is not yet supported for '{suffix}'. "
            f"Supported: {supported}."
        ) from exc
    output_path = (
        Path(output).expanduser().resolve() if output else source_path
    )
    if output_path.suffix.lower() != ".pptx":
        raise ValueError("macOS edit output must use '.pptx'")
    in_place = output_path == source_path
    if output_path.exists() and not in_place and not overwrite:
        raise FileExistsError(f"Output already exists: {output_path}")

    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    runtime_root = Path(
        tempfile.mkdtemp(prefix="wpscomposer-macos-edit-")
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
                    return _run_edit(
                        source_path, output_path, component, method,
                        patches, bridge, runtime, timeout,
                        atomic=atomic,
                        raise_on_error=raise_on_error,
                        overwrite=overwrite or in_place,
                    )
            except InspectionError:
                raise
            except Exception as exc:
                if not getattr(runtime, "registration_restored", True):
                    recovery = runtime.runtime_dir / "registration-recovery"
                    raise _error(
                        str(source_path), component,
                        "REGISTRATION_RESTORE_FAILED",
                        f"WPS registration restore failed; recovery "
                        f"retained at {recovery}",
                    ) from exc
                raise _error(
                    str(source_path), component,
                    "EDIT_COMMAND_FAILED",
                    "Mac WPS edit command failed",
                ) from exc
    finally:
        if runtime is None or getattr(runtime, "registration_restored", True):
            shutil.rmtree(runtime_root, ignore_errors=True)


def _run_edit(
    source: Path,
    output: Path,
    component: str,
    method: str,
    patches: list,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    timeout: float,
    *,
    atomic: bool = True,
    raise_on_error: bool = False,
    overwrite: bool = False,
) -> dict[str, Any]:
    if runtime.staging_dir is None:
        raise _error(str(source), component, "STAGING_UNAVAILABLE",
                     "WPS container staging session was not created")
    runtime.prepare_profiles()
    runtime.start_servers()
    runtime.activate_component(component)
    deadline = time.monotonic() + timeout
    try:
        _wait_for_registration(bridge, runtime, component, deadline)
    except TimeoutError as exc:
        raise _error(str(source), component, "EDIT_COMMAND_FAILED",
                     "Timed out waiting for the WPS add-in to register") from exc

    policy = PathPolicy((runtime.staging_dir,))
    staged_source = policy.require_allowed(
        runtime.staging_dir / f"source{source.suffix.lower()}"
    )
    staged_output = policy.require_allowed(
        runtime.staging_dir / f"output{output.suffix.lower()}"
    )
    try:
        shutil.copy2(source, staged_source)
    except OSError as exc:
        raise _error(str(source), component, "STAGING_SAVE_FAILED",
                     _redact_staging(str(exc), runtime.staging_dir)) from exc

    command = bridge.issue(component, method, {
        "sourcePath": str(staged_source),
        "outputPath": str(staged_output),
        "patches": patches,
        "atomic": bool(atomic),
        "raiseOnError": bool(raise_on_error),
    })
    try:
        result = bridge.wait_result(
            command.id, max(0.0, deadline - time.monotonic())
        )
    except TimeoutError as exc:
        bridge.state.cancel(command.id)
        raise _error(str(source), component, "EDIT_COMMAND_FAILED",
                     _redact_staging(str(exc), runtime.staging_dir)) from exc
    if not result.ok:
        details = dict(result.error or {})
        raise _error(str(source), component, "EDIT_COMMAND_FAILED",
                     _redact_staging(
                         str(details.get("message") or "WPS edit failed"),
                         runtime.staging_dir))
    value = dict(result.value or {})
    reports = list(value.get("patches") or [])
    if len(reports) != len(patches):
        raise _error(
            str(source), component, "PROTOCOL_ERROR",
            "WPS edit returned an incomplete patch report",
        )
    for report in reports:
        if report.get("rejected"):
            report["ok"] = False
    failures = [report for report in reports if not report.get("ok")]
    if failures and atomic:
        return {"path": None, "patches": reports, "saved": False}
    if failures and raise_on_error:
        raise _error(
            str(source), component, "PATCH_REJECTED",
            f"{len(failures)} of {len(reports)} patch(es) failed",
        )

    # The bridge must report exactly the path reserved for this command. A
    # merely in-root path could be the untouched staged source or stale output.
    try:
        reported = policy.require_allowed(str(value.get("path", "")))
    except ProtocolError as exc:
        raise _error(
            str(source), component, "PROTOCOL_ERROR",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    if reported != staged_output:
        raise _error(
            str(source), component, "PROTOCOL_ERROR",
            "WPS edit reported an unexpected staged output path",
        )

    _wait_for_stable_edit_artifact(
        staged_output, deadline, source=str(source), component=component
    )
    try:
        published = publish_artifact(
            staged_output,
            output,
            overwrite=overwrite,
            validator=lambda path: validate_office_package(path, "pptx"),
        )
    except FileExistsError:
        raise
    except ArtifactTransportError as exc:
        raise _error(str(source), component, exc.code, str(exc)) from exc
    except OSError as exc:
        raise _error(
            str(source), component, "ARTIFACT_PUBLISH_FAILED", str(exc)
        ) from exc
    value["path"] = str(published)
    value["saved"] = True
    return value


def _wait_for_stable_edit_artifact(
    path: Path,
    deadline: float,
    *,
    source: str,
    component: str,
) -> None:
    """Wait for two identical file observations within the existing deadline."""
    previous = None
    while True:
        try:
            stat = path.stat()
            current = (stat.st_size, stat.st_mtime_ns)
        except FileNotFoundError:
            current = None
        if current is not None and current == previous:
            return
        previous = current
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise _error(
                source,
                component,
                "STAGED_ARTIFACT_INVALID",
                "WPS edit output did not finish landing before the deadline",
            )
        time.sleep(min(0.05, remaining))


def macos_inspection_available() -> bool:
    """True when the JSAPI bridge backend can run on this platform.

    Used by document_api to decide whether to route inspect() through the
    bridge instead of failing with WPSUnavailable on macOS.
    """
    import platform
    if platform.system() != "Darwin":
        return False
    return Path("/Applications/wpsoffice.app").is_dir()


def raise_if_com_unavailable() -> None:
    """Raise WPSUnavailable if COM dispatch is impossible (non-Windows)."""
    from .._dispatch import _require
    _require()
