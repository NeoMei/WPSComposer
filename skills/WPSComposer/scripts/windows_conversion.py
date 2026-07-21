"""Windows WPS/Microsoft Office COM backend for Office-to-PDF conversion."""

from __future__ import annotations

from pathlib import Path
import tempfile
from typing import Mapping, Optional, Type

from .artifact_transport import ArtifactTransportError, publish_artifact, validate_pdf
from .conversion import ConversionError, ConversionRequest
from .sheet import SheetComposer
from .slide import SlideComposer
from .writer import WriterComposer


_DEFAULT_COMPOSERS = {
    "writer": WriterComposer,
    "spreadsheet": SheetComposer,
    "presentation": SlideComposer,
}


def convert_windows(
    request: ConversionRequest,
    *,
    composer_classes: Optional[Mapping[str, Type]] = None,
) -> Path:
    """Convert through the format-specific document-level COM API."""
    classes = _DEFAULT_COMPOSERS if composer_classes is None else composer_classes
    try:
        composer_class = classes[request.component]
    except KeyError as exc:
        raise ValueError(
            f"Unsupported conversion component: {request.component}"
        ) from exc

    with tempfile.TemporaryDirectory(prefix="wpscomposer-convert-") as directory:
        staged = Path(directory) / "converted.pdf"
        aliases = {str(staged.parent), str(staged.parent.resolve())}

        def redact(message: object) -> str:
            redacted = str(message)
            for alias in aliases:
                redacted = redacted.replace(alias, "<conversion-staging>")
            return redacted

        try:
            with composer_class.open_document(
                str(request.source), read_only=True, visible=False
            ) as composer:
                if request.component == "writer":
                    composer.doc.ExportAsFixedFormat(str(staged), 17)
                elif request.component == "spreadsheet":
                    composer.doc.ExportAsFixedFormat(0, str(staged))
                elif request.component == "presentation":
                    composer.doc.SaveAs(str(staged), 32)
                else:
                    raise ValueError(
                        f"Unsupported conversion component: {request.component}"
                    )
        except ValueError:
            raise
        except Exception as exc:
            raise ConversionError(
                code="CONVERSION_COMMAND_FAILED",
                source=str(request.source),
                component=request.component,
                backend="windows-com",
                message=redact(exc),
            ) from exc
        try:
            return publish_artifact(
                staged,
                request.output,
                overwrite=request.overwrite,
                validator=validate_pdf,
            )
        except ArtifactTransportError as exc:
            raise ConversionError(
                code=exc.code,
                source=str(request.source),
                component=request.component,
                backend="windows-com",
                message=redact(exc),
            ) from exc
