"""macOS WPS JSAPI backend for staged Office-to-PDF conversion."""

from __future__ import annotations

from pathlib import Path
import shutil
import tempfile
import time
from typing import Callable, Optional

from ..artifact_transport import (
    ArtifactTransportError,
    publish_artifact,
    validate_pdf,
)
from ..conversion import (
    REMOTE_CONVERSION_ERROR_CODES,
    ConversionError,
    ConversionRequest,
    normalize_conversion_error_code,
)
from .bridge import LoopbackBridge
from .models import PathPolicy, ProtocolError
from .runtime import ProbeRuntime


ORIGINS = {
    "http://127.0.0.1:3889",
    "http://127.0.0.1:3890",
    "http://127.0.0.1:3891",
}
METHODS = {
    "writer": "convert_writer_pdf",
    "spreadsheet": "convert_workbook_pdf",
    "presentation": "convert_presentation_pdf",
}

# Enabled after two six-format WPS 12.1.26035 acceptance runs on 2026-07-18.
MACOS_CONVERSION_ENABLED = True


def _error(
    request: ConversionRequest,
    code: str,
    message: str,
) -> ConversionError:
    return ConversionError(
        code=code,
        source=str(request.source),
        component=request.component,
        backend="mac-wps-jsapi",
        message=message,
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


def _run_conversion(
    request: ConversionRequest,
    method: str,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
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
    deadline = time.monotonic() + timeout
    try:
        _wait_for_registration(bridge, runtime, request.component, deadline)
    except TimeoutError as exc:
        raise _error(
            request,
            "CONVERSION_COMMAND_FAILED",
            "Timed out waiting for the WPS add-in to register",
        ) from exc

    policy = PathPolicy((runtime.staging_dir,))
    staged_source = policy.require_allowed(
        runtime.staging_dir / f"source{request.source.suffix.lower()}"
    )
    staged_output = policy.require_allowed(
        runtime.staging_dir / "converted.pdf"
    )
    try:
        shutil.copy2(request.source, staged_source)
    except OSError as exc:
        raise _error(
            request,
            "STAGING_SAVE_FAILED",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    command = bridge.issue(
        request.component,
        method,
        {
            "sourcePath": str(staged_source),
            "outputPath": str(staged_output),
        },
    )
    try:
        result = bridge.wait_result(
            command.id, max(0.0, deadline - time.monotonic())
        )
    except TimeoutError as exc:
        bridge.state.cancel(command.id)
        raise _error(
            request,
            "CONVERSION_COMMAND_FAILED",
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc
    if not result.ok:
        details = dict(result.error or {})
        raise _error(
            request,
            normalize_conversion_error_code(
                details.get("code"), allowed=REMOTE_CONVERSION_ERROR_CODES
            ),
            _redact_staging(
                str(details.get("message") or "WPS conversion failed"),
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
    if reported != staged_output:
        raise _error(
            request,
            "PROTOCOL_ERROR",
            "WPS returned an unexpected staged output path",
        )
    try:
        return publish_artifact(
            staged_output,
            request.output,
            overwrite=request.overwrite,
            validator=validate_pdf,
        )
    except ArtifactTransportError as exc:
        raise _error(
            request,
            exc.code,
            _redact_staging(str(exc), runtime.staging_dir),
        ) from exc


def convert_macos(
    request: ConversionRequest,
    *,
    enabled: Optional[bool] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    timeout: float = 90,
) -> Path:
    """Convert one staged source through its typed WPS JSAPI command."""
    gate_enabled = MACOS_CONVERSION_ENABLED if enabled is None else bool(enabled)
    if not gate_enabled:
        raise _error(
            request,
            "MACOS_GATE_NOT_PASSED",
            "Mac WPS conversion is disabled until the real acceptance gate passes",
        )
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    try:
        method = METHODS[request.component]
    except KeyError as exc:
        raise _error(
            request,
            "UNSUPPORTED_COMPONENT",
            f"Unsupported conversion component: {request.component}",
        ) from exc

    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos/wps-jsapi-probe"
    runtime_root = Path(
        tempfile.mkdtemp(prefix="wpscomposer-macos-convert-")
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
                    return _run_conversion(
                        request, method, bridge, runtime, timeout
                    )
            except ConversionError:
                raise
            except Exception as exc:
                if not getattr(runtime, "registration_restored", True):
                    recovery = runtime.runtime_dir / "registration-recovery"
                    raise _error(
                        request,
                        "REGISTRATION_RESTORE_FAILED",
                        f"WPS registration restore failed; recovery retained at {recovery}",
                    ) from exc
                raise
    finally:
        if runtime is None or getattr(runtime, "registration_restored", True):
            shutil.rmtree(runtime_root, ignore_errors=True)
