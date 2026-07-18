"""Pinned WPS JSAPI document templates for macOS generation probes."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import os
from pathlib import Path
import shutil

from ..artifact_transport import validate_office_package


class TemplateError(RuntimeError):
    """Raised when a pinned WPS template cannot be used safely."""


@dataclass(frozen=True)
class TemplateSpec:
    component: str
    filename: str
    output_name: str
    format_name: str
    sha256: str


TEMPLATES = {
    "writer": TemplateSpec(
        "writer",
        "wpsDemo.docx",
        "generated.docx",
        "docx",
        "95c2da9c75b65f7da18da345847a65ecab21512978c15e65c5b601512d52fd8e",
    ),
    "spreadsheet": TemplateSpec(
        "spreadsheet",
        "etDemo.xlsx",
        "generated.xlsx",
        "xlsx",
        "999138c4d4d22c2eb7c80e114d623da0ac310406ab71317256758754aae02dc9",
    ),
    "presentation": TemplateSpec(
        "presentation",
        "wppDemo.pptx",
        "generated.pptx",
        "pptx",
        "5bafe9e14e99c7b1f7f81e0d1e32c59eb2d52c93ccfdd4ca786b0c8616174600",
    ),
}


def template_for_component(component: str) -> TemplateSpec:
    """Return the immutable template specification for a WPS component."""
    try:
        return TEMPLATES[component]
    except KeyError as exc:
        raise TemplateError(f"Unsupported WPS component: {component}") from exc


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def clone_template(probe_root: Path, staging_dir: Path, component: str) -> Path:
    """Validate a pinned WPS template and copy it into the private session."""
    spec = template_for_component(component)
    source = Path(probe_root) / "node_modules/wpsjs/src/lib/res" / spec.filename
    try:
        digest = _sha256(source)
    except OSError as exc:
        raise TemplateError(f"Pinned {component} template digest mismatch") from exc
    if digest != spec.sha256:
        raise TemplateError(f"Pinned {component} template digest mismatch")
    validate_office_package(source, spec.format_name)
    target = Path(staging_dir) / spec.output_name
    shutil.copyfile(source, target)
    os.chmod(target, 0o600)
    return target
