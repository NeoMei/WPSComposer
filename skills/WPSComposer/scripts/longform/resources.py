"""Resource preflight for long-form documents."""

from __future__ import annotations

import base64
import hashlib
import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

from ..document_model import (
    DegradationBlock,
    ExcalidrawBlock,
    FigureBlock,
    ImageBlock,
    Section,
)

RESOURCE_NOT_FOUND = "RESOURCE_NOT_FOUND"
RESOURCE_PATH_ESCAPES_BASE = "RESOURCE_PATH_ESCAPES_BASE"
RESOURCE_ABSOLUTE_PATH_OUTSIDE = "RESOURCE_ABSOLUTE_PATH_OUTSIDE"
RESOURCE_MEDIA_TYPE_UNSUPPORTED = "RESOURCE_MEDIA_TYPE_UNSUPPORTED"
RESOURCE_READ_FAILED = "RESOURCE_READ_FAILED"
RESOURCE_DECODE_FAILED = "RESOURCE_DECODE_FAILED"

_EXTENSION_MEDIA_TYPES: dict[str, str] = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".tiff": "image/tiff",
    ".tif": "image/tiff",
    ".bmp": "image/bmp",
    ".webp": "image/webp",
    ".svg": "image/svg+xml",
}

_RASTER_MEDIA_TYPES = frozenset({
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/tiff",
    "image/bmp",
    "image/webp",
})


@dataclass(frozen=True)
class PreflightResource:
    """A resource that passed preflight checks."""

    resource_id: str
    source_path: str
    source_sha256: str
    payload_sha256: str
    byte_length: int
    media_type: str
    normalizer_id: str


@dataclass(frozen=True)
class ResourceDegradation:
    """A deterministic degradation for a resource that failed preflight."""

    node_id: Optional[str]
    code: str
    message: str
    fallback_text: str


@dataclass(frozen=True)
class ResourcePreflight:
    """Result of resource preflight."""

    resources: list[PreflightResource] = field(default_factory=list)
    degradations: list[ResourceDegradation] = field(default_factory=list)
    manifest: dict[str, Any] = field(default_factory=dict)


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _resource_id_for_path(source_path: str) -> str:
    """Deterministic opaque resource id derived from the relative source path."""
    digest = hashlib.sha256(source_path.encode("utf-8")).digest()
    return "wpsc-rsrc:" + base64.urlsafe_b64encode(digest).decode("ascii").rstrip("=")[:16]


def _media_type_for_path(path: str) -> Optional[str]:
    lower = path.lower()
    if lower.endswith(".excalidraw.md"):
        return "image/svg+xml"
    ext = os.path.splitext(lower)[1]
    return _EXTENSION_MEDIA_TYPES.get(ext)


def _is_inside_base(target: Path, base: Path) -> bool:
    try:
        target.relative_to(base)
        return True
    except ValueError:
        return False


def _resolve_resource_path(source_path: str, base_dir: str) -> tuple[Optional[Path], Optional[str]]:
    """Resolve a resource path against base_dir, returning (path, error_code).

    Returns (None, code) when the path is rejected.  Resolves symlinks so that
    a symlink pointing outside base_dir is treated the same as a traversal path.
    """
    base = Path(base_dir).resolve()
    raw = Path(source_path)

    if raw.is_absolute():
        try:
            resolved = raw.resolve()
        except (OSError, RuntimeError):
            return None, RESOURCE_PATH_ESCAPES_BASE
        if not _is_inside_base(resolved, base):
            return None, RESOURCE_ABSOLUTE_PATH_OUTSIDE
        return resolved, None

    try:
        resolved = (base / source_path).resolve()
    except (OSError, RuntimeError):
        return None, RESOURCE_PATH_ESCAPES_BASE

    if not _is_inside_base(resolved, base):
        return None, RESOURCE_PATH_ESCAPES_BASE
    return resolved, None


def _scan_nodes(nodes: list[Any]) -> list[tuple[str, Optional[str]]]:
    """Collect (path, node_id) pairs from arbitrary document nodes."""
    results: list[tuple[str, Optional[str]]] = []
    stack = list(nodes)
    while stack:
        node = stack.pop()
        if isinstance(node, Section):
            stack.extend(reversed(node.elements))
        elif isinstance(node, FigureBlock):
            stack.extend(reversed(node.images))
        elif isinstance(node, ImageBlock):
            if node.path:
                results.append((node.path, getattr(node, "node_id", None)))
        elif isinstance(node, ExcalidrawBlock):
            if node.path:
                results.append((node.path, getattr(node, "node_id", None)))
    return results


def _build_manifest(resources: list[PreflightResource]) -> dict[str, Any]:
    entries = [
        {
            "resourceId": r.resource_id,
            "sourceSha256": r.source_sha256,
            "payloadSha256": r.payload_sha256,
            "byteLength": r.byte_length,
            "mediaType": r.media_type,
            "normalizerId": r.normalizer_id,
        }
        for r in sorted(resources, key=lambda x: x.resource_id)
    ]
    envelope: dict[str, Any] = {"version": "1", "entries": entries}
    canonical = json.dumps(envelope, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    envelope["digest"] = hashlib.sha256(canonical).hexdigest()
    return envelope


def preflight_resources(nodes: list[Any], base_dir: str) -> ResourcePreflight:
    """Preflight file resources referenced by document nodes.

    Returns validated resources, deterministic degradations for rejected paths,
    and a redacted resource manifest suitable for a generation-plan envelope.
    """
    resources: list[PreflightResource] = []
    degradations: list[ResourceDegradation] = []
    seen_paths: set[str] = set()

    base = Path(base_dir).resolve()

    for source_path, node_id in _scan_nodes(nodes):
        if not source_path:
            continue

        normalized_source_path = source_path.replace(os.sep, "/")
        if normalized_source_path in seen_paths:
            continue
        seen_paths.add(normalized_source_path)

        resolved, error_code = _resolve_resource_path(source_path, str(base))
        if error_code:
            degradations.append(
                ResourceDegradation(
                    node_id=node_id,
                    code=error_code,
                    message=f"Resource path rejected: {source_path}",
                    fallback_text=f"[{error_code}] {source_path}",
                )
            )
            continue

        if resolved is None:
            continue

        media_type = _media_type_for_path(source_path)
        if media_type is None:
            degradations.append(
                ResourceDegradation(
                    node_id=node_id,
                    code=RESOURCE_MEDIA_TYPE_UNSUPPORTED,
                    message=f"Unsupported media type for resource: {source_path}",
                    fallback_text=f"[{RESOURCE_MEDIA_TYPE_UNSUPPORTED}] {source_path}",
                )
            )
            continue

        try:
            data = resolved.read_bytes()
        except OSError as exc:
            code = RESOURCE_READ_FAILED if exc.errno in {13} else RESOURCE_NOT_FOUND
            degradations.append(
                ResourceDegradation(
                    node_id=node_id,
                    code=code,
                    message=f"Could not read resource {source_path}: {exc}",
                    fallback_text=f"[{code}] {source_path}",
                )
            )
            continue

        # M1 preflight deliberately defers deep media decoding to the staging
        # and executor layers.  We record the media type and hashes here.

        source_digest = _sha256(data)
        resources.append(
            PreflightResource(
                resource_id=_resource_id_for_path(normalized_source_path),
                source_path=normalized_source_path,
                source_sha256=source_digest,
                payload_sha256=source_digest,
                byte_length=len(data),
                media_type=media_type,
                normalizer_id="none",
            )
        )

    manifest = _build_manifest(resources)
    return ResourcePreflight(
        resources=resources,
        degradations=degradations,
        manifest=manifest,
    )


__all__ = [
    "RESOURCE_ABSOLUTE_PATH_OUTSIDE",
    "RESOURCE_DECODE_FAILED",
    "RESOURCE_MEDIA_TYPE_UNSUPPORTED",
    "RESOURCE_NOT_FOUND",
    "RESOURCE_PATH_ESCAPES_BASE",
    "RESOURCE_READ_FAILED",
    "PreflightResource",
    "ResourceDegradation",
    "ResourcePreflight",
    "preflight_resources",
]
