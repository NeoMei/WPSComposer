"""Pinned WPS JSAPI document templates for macOS generation probes."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import shutil
import tempfile

from ..artifact_transport import validate_office_package


class TemplateError(RuntimeError):
    """Raised when a pinned WPS template cannot be used safely."""


class AddinAssetError(RuntimeError):
    """Raised when a generated add-in asset drifted from its manifest."""


ADDIN_ASSET_MANIFEST = "asset-manifest.json"
GENERATED_ADDIN_ASSETS = ("writer-longform-v2.js",)


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


def build_addin_asset_manifest(assets: Path) -> dict[str, object]:
    """Build the canonical manifest for generated long-form add-in assets."""

    root = Path(assets)
    return {
        "version": 1,
        "assets": {
            name: _sha256(root / name)
            for name in GENERATED_ADDIN_ASSETS
        },
    }


def write_addin_asset_manifest(assets: Path) -> Path:
    """Regenerate the manifest from the canonical template asset directory."""

    root = Path(assets)
    target = root / ADDIN_ASSET_MANIFEST
    target.write_text(
        json.dumps(
            build_addin_asset_manifest(root),
            ensure_ascii=False,
            sort_keys=True,
            indent=2,
        ) + "\n",
        encoding="utf-8",
    )
    return target


def verify_addin_assets(assets: Path) -> None:
    """Reject a generated add-in asset that no longer matches its manifest."""

    root = Path(assets)
    try:
        raw = json.loads((root / ADDIN_ASSET_MANIFEST).read_text(encoding="utf-8"))
        expected = build_addin_asset_manifest(root)
    except (OSError, ValueError, TypeError, KeyError):
        raise AddinAssetError("writer-longform-v2 add-in asset is unverified") from None
    if raw != expected:
        raise AddinAssetError("writer-longform-v2 add-in asset digest mismatch") from None


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


def _main() -> int:
    parser = argparse.ArgumentParser(
        description="Regenerate WPS add-in asset manifest"
    )
    parser.add_argument("--write-addin-manifest", type=Path)
    args = parser.parse_args()
    if args.write_addin_manifest is None:
        parser.error("--write-addin-manifest is required")
    write_addin_asset_manifest(args.write_addin_manifest)
    verify_addin_assets(args.write_addin_manifest)
    return 0


if __name__ == "__main__":  # pragma: no cover - exercised as an asset command
    raise SystemExit(_main())
