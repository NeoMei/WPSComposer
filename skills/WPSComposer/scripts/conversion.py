"""Stable cross-platform Office-to-PDF conversion facade."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys
from typing import Callable, Optional, Tuple

from .artifact_transport import ArtifactTransportError, validate_pdf


_COMPONENT_BY_SUFFIX = {
    ".doc": "writer",
    ".docx": "writer",
    ".xls": "spreadsheet",
    ".xlsx": "spreadsheet",
    ".ppt": "presentation",
    ".pptx": "presentation",
}


@dataclass(frozen=True)
class ConversionRequest:
    """Validated source, destination, and routing information."""

    source: Path
    output: Path
    component: str
    overwrite: bool


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
        return {
            "code": self.code,
            "source": self.source,
            "component": self.component,
            "backend": self.backend,
            "message": self.message,
        }


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
) -> str:
    """Convert one Word, Excel, or PowerPoint file to an absolute PDF path."""
    request = _build_request(source, output, overwrite=overwrite)
    backend_name, backend = _select_backend(request)
    try:
        result = Path(backend(request)).expanduser().resolve()
        validate_pdf(result)
        return str(result)
    except (FileNotFoundError, FileExistsError, ValueError, ConversionError):
        raise
    except ArtifactTransportError as exc:
        raise ConversionError(
            code=exc.code,
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
