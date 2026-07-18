"""Windows WPS/Microsoft Office COM backend for Office-to-PDF conversion."""

from __future__ import annotations

from pathlib import Path
import tempfile
from typing import Mapping, Optional, Type

from .artifact_transport import publish_artifact, validate_pdf
from .conversion import ConversionRequest
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
        return publish_artifact(
            staged,
            request.output,
            overwrite=request.overwrite,
            validator=validate_pdf,
        )
