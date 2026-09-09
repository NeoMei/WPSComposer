"""Stable cross-platform Office-to-PDF conversion facade."""

from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path
import sys
from typing import Callable, Optional, Tuple
import warnings

from .artifact_transport import (
    ArtifactTransportError,
    ArtifactValidationError,
    validate_pdf,
)
from .presentation import present_artifact, validate_open_result
from .office_engines import com_engine, resolve_engine, validate_engine, validate_timeout
from .msoffice.errors import NativeWordError, NATIVE_WORD_ERROR_CODES, RECOVERY_FIELDS
from .msoffice.office_errors import NativeOfficeError, NATIVE_OFFICE_ERROR_CODES


_COMPONENT_BY_SUFFIX = {
    ".doc": "writer",
    ".docx": "writer",
    ".xls": "spreadsheet",
    ".xlsx": "spreadsheet",
    ".ppt": "presentation",
    ".pptx": "presentation",
}

STABLE_CONVERSION_ERROR_CODES = NATIVE_WORD_ERROR_CODES | NATIVE_OFFICE_ERROR_CODES | frozenset(
    {
        "ARTIFACT_PUBLISH_FAILED",
        "BACKEND_UNAVAILABLE",
        "CONVERSION_COMMAND_FAILED",
        "CONVERSION_FAILED",
        "FINAL_ARTIFACT_INVALID",
        "INTERACTIVE_INPUT_REQUIRED",
        "MACOS_GATE_NOT_PASSED",
        "NO_VISIBLE_WORKSHEETS",
        "PROTOCOL_ERROR",
        "REGISTRATION_RESTORE_FAILED",
        "STAGED_ARTIFACT_INVALID",
        "STAGING_SAVE_FAILED",
        "STAGING_UNAVAILABLE",
        "UNSUPPORTED_COMPONENT",
    }
)
REMOTE_CONVERSION_ERROR_CODES = frozenset(
    {
        "CONVERSION_COMMAND_FAILED",
        "INTERACTIVE_INPUT_REQUIRED",
        "NO_VISIBLE_WORKSHEETS",
    }
)


def normalize_conversion_error_code(
    code: object,
    *,
    allowed=STABLE_CONVERSION_ERROR_CODES,
    fallback: str = "CONVERSION_COMMAND_FAILED",
) -> str:
    value = str(code) if isinstance(code, str) else ""
    return value if value in allowed else fallback


@dataclass(frozen=True)
class ConversionRequest:
    """Validated source, destination, and routing information."""

    source: Path
    output: Path
    component: str
    overwrite: bool
    engine: str = "wps"
    timeout: float = 600


class ConversionError(RuntimeError):
    """Stable runtime error returned by every conversion backend."""

    def __init__(
        self,
        *,
        code: str,
        source: str,
        component: str,
        backend: str,
        message: str,
    ):
        super().__init__(message)
        self.code = code
        self.source = source
        self.component = component
        self.backend = backend
        self.message = message

    def to_dict(self) -> dict:
        result = {
            "code": self.code,
            "source": self.source,
            "component": self.component,
            "backend": self.backend,
            "message": self.message,
        }
        for name in RECOVERY_FIELDS:
            value = getattr(self, name, None)
            if value is not None:
                result[name] = value
        return result


Backend = Callable[[ConversionRequest], Path]


def _build_request(
    source: str,
    output: Optional[str],
    *,
    overwrite: bool,
) -> ConversionRequest:
    source_path = Path(source).expanduser().resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"Source file not found: {source_path}")
    suffix = source_path.suffix.lower()
    try:
        component = _COMPONENT_BY_SUFFIX[suffix]
    except KeyError as exc:
        supported = ", ".join(sorted(_COMPONENT_BY_SUFFIX))
        raise ValueError(
            f"Unsupported source format '{source_path.suffix}'. Use: {supported}."
        ) from exc

    output_path = (
        source_path.with_suffix(".pdf")
        if output is None
        else Path(output).expanduser().resolve()
    )
    if output_path.suffix.lower() != ".pdf":
        raise ValueError(f"Output path must end in .pdf: {output_path}")
    if output_path.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {output_path}")
    return ConversionRequest(
        source=source_path,
        output=output_path,
        component=component,
        overwrite=bool(overwrite),
    )


def _select_backend(request: ConversionRequest) -> Tuple[str, Backend]:
    if request.engine == "msoffice":
        if request.component != "writer":
            if sys.platform == "darwin":
                from .msoffice.macos_office_runtime import convert
                return "mac-" + request.component + "-applescript", lambda req: convert(req, timeout=req.timeout)
            if sys.platform == "win32":
                from .msoffice.windows_office_runtime import convert
                return "windows-" + request.component + "-com", lambda req: convert(req, timeout=req.timeout)
        if sys.platform == "win32":
            from .msoffice.windows_runtime import convert
            return "windows-word-com", lambda req: convert(req, timeout=req.timeout)
        if sys.platform == "darwin":
            from .msoffice.macos_runtime import convert
            return "mac-word-applescript", lambda req: convert(req, timeout=req.timeout)
    if sys.platform == "win32":
        from .windows_conversion import convert_windows

        return "windows-com", convert_windows
    if sys.platform == "darwin":
        from .macos_probe.conversion import convert_macos

        return "mac-wps-jsapi", convert_macos
    raise ConversionError(
        code="BACKEND_UNAVAILABLE",
        source=str(request.source),
        component=request.component,
        backend=sys.platform,
        message=f"Unsupported platform: {sys.platform}",
    )


def convert_to_pdf(
    source: str,
    output: Optional[str] = None,
    *,
    overwrite: bool = False,
    open_result: bool = False,
    engine: str = "wps",
    timeout: float = 600,
) -> str:
    """Convert to PDF using WPS or native Word (DOC/DOCX only).

    ``auto`` prefers installed WPS, then Word, with no execution-time fallback.
    Native Word applies ``timeout`` to staging, export and atomic publication.
    """
    validate_open_result(open_result)
    validate_engine(engine)
    validate_timeout(timeout)
    request = _build_request(source, output, overwrite=overwrite)
    request = replace(request, engine=resolve_engine(engine, request.component), timeout=timeout)
    backend_name, backend = _select_backend(request)
    try:
        with com_engine(request.engine):
            result = Path(backend(request)).expanduser().resolve()
    except (NativeWordError, NativeOfficeError) as exc:
        error = ConversionError(
            code=exc.code, source=str(request.source), component=request.component,
            backend=backend_name, message=exc.safe_message,
        )
        for name in RECOVERY_FIELDS:
            setattr(error, name, getattr(exc, name))
        raise error from None
    except (FileNotFoundError, FileExistsError, ValueError):
        raise
    except ConversionError as exc:
        normalized = normalize_conversion_error_code(exc.code)
        if normalized == exc.code:
            raise
        raise ConversionError(
            code=normalized,
            source=exc.source,
            component=exc.component,
            backend=exc.backend,
            message=exc.message,
        ) from exc
    except ArtifactTransportError as exc:
        raise ConversionError(
            code=normalize_conversion_error_code(exc.code),
            source=str(request.source),
            component=request.component,
            backend=backend_name,
            message=str(exc),
        ) from exc
    except Exception as exc:
        raise ConversionError(
            code="CONVERSION_FAILED",
            source=str(request.source),
            component=request.component,
            backend=backend_name,
            message=str(exc),
        ) from exc
    try:
        validate_pdf(result)
    except ArtifactValidationError as exc:
        raise ConversionError(
            code="FINAL_ARTIFACT_INVALID",
            source=str(request.source),
            component=request.component,
            backend=backend_name,
            message=str(exc),
        ) from exc
    if open_result:
        try:
            present_artifact(result)
        except Exception as exc:
            warnings.warn(
                f"Final artifact was published but could not be opened: {exc}",
                RuntimeWarning,
                stacklevel=2,
            )
    return str(result)
