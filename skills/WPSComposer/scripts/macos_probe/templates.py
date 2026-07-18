"""Pinned WPS JSAPI document templates for macOS generation probes."""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import os
from pathlib import Path
import shutil
import tempfile

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
    target = Path(staging_dir) / spec.output_name
    try:
        if _sha256(source) != spec.sha256:
            raise TemplateError(f"Pinned {component} template digest mismatch")
    except OSError as exc:
        raise TemplateError(f"Pinned {component} template digest mismatch") from exc
    temporary: Path | None = None
    target_created = False
    published = False
    try:
        descriptor, temporary_name = tempfile.mkstemp(
            dir=staging_dir,
            prefix=".wpscomposer-template-",
            suffix=".tmp",
        )
        temporary = Path(temporary_name)
        os.close(descriptor)
        os.chmod(temporary, 0o600)
        try:
            shutil.copyfile(source, temporary)
        except OSError as exc:
            raise TemplateError(
                f"Pinned {component} template digest mismatch"
            ) from exc
        if _sha256(temporary) != spec.sha256:
            raise TemplateError(f"Pinned {component} template digest mismatch")
        validate_office_package(temporary, spec.format_name)
        os.chmod(temporary, 0o600)
        os.link(temporary, target)
        target_created = True
        temporary.unlink()
        temporary = None
        published = True
        return target
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
        if target_created and not published:
            target.unlink(missing_ok=True)
