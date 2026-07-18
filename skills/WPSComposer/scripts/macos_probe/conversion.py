"""macOS WPS JSAPI backend for staged Office-to-PDF conversion."""

from __future__ import annotations

from pathlib import Path
import shutil
import tempfile
import time
from typing import Callable, Optional

from ..artifact_transport import publish_artifact, validate_pdf
from ..conversion import ConversionError, ConversionRequest
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

# This is flipped only after the real, two-run macOS acceptance gate passes.
MACOS_CONVERSION_ENABLED = False


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
    with bridge_factory(ORIGINS) as bridge:
        with tempfile.TemporaryDirectory(
            prefix="wpscomposer-macos-convert-"
        ) as directory:
            runtime = runtime_factory(
                probe_root,
                Path(directory) / "runtime",
                bridge.url,
                bridge.token,
            )
            with runtime:
                if runtime.staging_dir is None:
                    raise _error(
                        request,
                        "STAGING_UNAVAILABLE",
                        "WPS container staging session was not created",
                    )
                runtime.prepare_profiles()
                runtime.start_servers()
                runtime.activate_component(request.component)
                _wait_for_registration(
                    bridge, runtime, request.component, timeout
                )

                policy = PathPolicy((runtime.staging_dir,))
                staged_source = policy.require_allowed(
                    runtime.staging_dir / f"source{request.source.suffix.lower()}"
                )
                staged_output = policy.require_allowed(
                    runtime.staging_dir / "converted.pdf"
                )
                shutil.copy2(request.source, staged_source)
                command = bridge.issue(
                    request.component,
                    method,
                    {
                        "sourcePath": str(staged_source),
                        "outputPath": str(staged_output),
                    },
                )
                result = bridge.wait_result(command.id, timeout)
                if not result.ok:
                    details = dict(result.error or {})
                    raise _error(
                        request,
                        str(details.get("code") or "JSAPI_COMMAND_FAILED"),
                        str(details.get("message") or "WPS conversion failed"),
                    )
                try:
                    reported = policy.require_allowed(
                        str(result.value.get("path", ""))
                    )
                except ProtocolError as exc:
                    raise _error(request, "PROTOCOL_ERROR", str(exc)) from exc
                if reported != staged_output:
                    raise _error(
                        request,
                        "PROTOCOL_ERROR",
                        "WPS returned an unexpected staged output path",
                    )
                return publish_artifact(
                    staged_output,
                    request.output,
                    overwrite=request.overwrite,
                    validator=validate_pdf,
                )
