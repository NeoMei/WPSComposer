"""Resource preflight for long-form documents."""

from __future__ import annotations

import base64
import hashlib
import io
import json
import math
import os
import re
import warnings
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional
from urllib.parse import unquote_to_bytes

from PIL import Image, ImageCms, ImageOps, UnidentifiedImageError

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
RESOURCE_TOO_LARGE = "RESOURCE_TOO_LARGE"
RESOURCE_PIXEL_LIMIT_EXCEEDED = "RESOURCE_PIXEL_LIMIT_EXCEEDED"
RESOURCE_SIDE_LIMIT_EXCEEDED = "RESOURCE_SIDE_LIMIT_EXCEEDED"
MULTIFRAME_FLATTENED = "MULTIFRAME_FLATTENED"

MAX_RESOURCE_BYTES = 50 * 1024 * 1024
MAX_IMAGE_PIXELS = 80_000_000
MAX_IMAGE_SIDE = 32_768

_FORMAT_MEDIA_TYPES = {
    "PNG": "image/png",
    "JPEG": "image/jpeg",
    "TIFF": "image/tiff",
    "BMP": "image/bmp",
    "GIF": "image/gif",
}
_SVG_FORBIDDEN_DECLARATIONS = re.compile(br"<!\s*(?:DOCTYPE|ENTITY)\b", re.IGNORECASE)
_SVG_URL = re.compile(r"url\(\s*(['\"]?)(.*?)\1\s*\)", re.IGNORECASE)
_SVG_LENGTH = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)(?:px)?\s*$", re.IGNORECASE)


@dataclass(frozen=True)
class ImageProfile:
    """Decoded image metadata used by pure layout policy."""

    pixel_width: int
    pixel_height: int
    dpi_x: Optional[float]
    dpi_y: Optional[float]
    frame_count: int
    source_format: str
    has_icc: bool


_RasterPreparation = tuple[str, str, bytes, ImageProfile, bool]


@dataclass(frozen=True)
class PreflightResource:
    """A resource that passed preflight checks."""

    resource_id: str
    source_path: str = field(repr=False, compare=False)
    source_sha256: str
    payload_sha256: str
    byte_length: int
    media_type: str
    normalizer_id: str
    image_profile: ImageProfile = field(
        default_factory=lambda: ImageProfile(0, 0, None, None, 1, "UNKNOWN", False)
    )
    payload_bytes: bytes = field(default=b"", repr=False, compare=False)


@dataclass(frozen=True)
class PreparedLongformResource:
    """Private normalized resource transport for long-form executors."""

    id: str
    media_type: str
    source_sha256: str
    payload_sha256: str
    normalizer_id: str
    payload_bytes: bytes = field(repr=False, compare=False)
    image_profile: ImageProfile


@dataclass(frozen=True)
class ResourceDegradation:
    """A deterministic degradation for a resource that failed preflight."""

    node_id: Optional[str]
    code: str
    message: str
    fallback_text: str
    source_path: str = field(repr=False, compare=False)


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


class _ResourceRejected(Exception):
    def __init__(self, code: str) -> None:
        super().__init__(code)
        self.code = code


def _finite_positive(value: Any) -> Optional[float]:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) and number > 0 else None


def _dpi_pair(info: dict[str, Any]) -> tuple[Optional[float], Optional[float]]:
    raw = info.get("dpi")
    if isinstance(raw, (tuple, list)) and len(raw) >= 2:
        return _finite_positive(raw[0]), _finite_positive(raw[1])
    value = _finite_positive(raw)
    return (value, value) if value is not None else (None, None)


def _valid_icc(raw: Any) -> Optional[bytes]:
    if not isinstance(raw, bytes) or not raw:
        return None
    try:
        ImageCms.ImageCmsProfile(io.BytesIO(raw))
    except (OSError, TypeError, ValueError):
        return None
    return raw


def _check_geometry(width: int, height: int) -> None:
    if width > MAX_IMAGE_SIDE or height > MAX_IMAGE_SIDE:
        raise _ResourceRejected(RESOURCE_SIDE_LIMIT_EXCEEDED)
    if width * height > MAX_IMAGE_PIXELS:
        raise _ResourceRejected(RESOURCE_PIXEL_LIMIT_EXCEEDED)


def _save_lossless_png(image: Image.Image, icc: Optional[bytes]) -> bytes:
    output = io.BytesIO()
    kwargs: dict[str, Any] = {"format": "PNG", "compress_level": 9}
    if icc is not None:
        kwargs["icc_profile"] = icc
    image.save(output, **kwargs)
    return output.getvalue()


def _prepare_raster_if_identified(data: bytes) -> Optional[_RasterPreparation]:
    with warnings.catch_warnings():
        warnings.simplefilter("error", Image.DecompressionBombWarning)
        try:
            with Image.open(io.BytesIO(data)) as image:
                source_format = str(image.format or "").upper()
                media_type = _FORMAT_MEDIA_TYPES.get(source_format)
                if media_type is None:
                    raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)

                frame_count = int(getattr(image, "n_frames", 1) or 1)
                info = dict(image.info)
                dpi_x, dpi_y = _dpi_pair(info)
                icc = _valid_icc(info.get("icc_profile"))
                orientation = int(image.getexif().get(274, 1) or 1)

                # Every frame/page is decoded so the accepted format is based
                # on decoded content rather than a header or filename suffix.
                for frame_index in range(frame_count):
                    image.seek(frame_index)
                    _check_geometry(*image.size)
                    image.load()

                image.seek(0)
                first = image.copy()
                transformed = orientation != 1
                if transformed:
                    first = ImageOps.exif_transpose(first)
                _check_geometry(*first.size)

                multiframe = frame_count > 1 and source_format in {"GIF", "TIFF"}
                if multiframe or transformed:
                    payload = _save_lossless_png(first, icc)
                    payload_media_type = "image/png"
                    if multiframe:
                        normalizer = (
                            "gif-first-frame-png-v1"
                            if source_format == "GIF"
                            else "tiff-first-page-png-v1"
                        )
                    else:
                        normalizer = "exif-transpose-png-v1"
                else:
                    payload = data
                    payload_media_type = media_type
                    normalizer = "none-v1"

                profile = ImageProfile(
                    pixel_width=int(first.width),
                    pixel_height=int(first.height),
                    dpi_x=dpi_x,
                    dpi_y=dpi_y,
                    frame_count=frame_count,
                    source_format=source_format,
                    has_icc=icc is not None,
                )
                return payload_media_type, normalizer, payload, profile, multiframe
        except _ResourceRejected:
            raise
        except (Image.DecompressionBombWarning, Image.DecompressionBombError):
            raise _ResourceRejected(RESOURCE_DECODE_FAILED)
        except UnidentifiedImageError:
            return None
        except (OSError, SyntaxError, ValueError, EOFError):
            raise _ResourceRejected(RESOURCE_DECODE_FAILED)


def _prepare_raster(data: bytes) -> _RasterPreparation:
    prepared = _prepare_raster_if_identified(data)
    if prepared is None:
        raise _ResourceRejected(RESOURCE_DECODE_FAILED)
    return prepared


def _decode_data_uri(uri: str) -> tuple[str, bytes]:
    if not uri.lower().startswith("data:") or "," not in uri:
        raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
    metadata, encoded = uri[5:].split(",", 1)
    parts = metadata.split(";")
    media_type = parts[0].lower() or "text/plain"
    try:
        if any(part.lower() == "base64" for part in parts[1:]):
            payload = base64.b64decode(encoded, validate=True)
        else:
            payload = unquote_to_bytes(encoded)
    except (ValueError, UnicodeError):
        raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
    return media_type, payload


def _svg_number(value: Optional[str]) -> Optional[int]:
    if value is None:
        return None
    match = _SVG_LENGTH.fullmatch(value)
    if not match:
        return None
    return max(1, int(math.ceil(float(match.group(1)))))


def _prepare_svg(data: bytes) -> tuple[str, str, bytes, ImageProfile, bool]:
    if _SVG_FORBIDDEN_DECLARATIONS.search(data):
        raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)

    count = 0
    depth = 0
    root: Optional[ET.Element] = None
    embedded_bytes = 0

    def inspect_data_uri(uri: str, *, raster_required: bool = False) -> None:
        nonlocal embedded_bytes
        media_type, payload = _decode_data_uri(uri)
        declared_raster = (
            media_type.startswith("image/") and media_type != "image/svg+xml"
        )
        try:
            prepared_raster = _prepare_raster_if_identified(payload)
        except _ResourceRejected:
            raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
        if prepared_raster is None and (raster_required or declared_raster):
            raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
        embedded_bytes += len(payload)
        if len(data) + embedded_bytes > MAX_RESOURCE_BYTES:
            raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)

    def inspect_url_references(value: str) -> None:
        for match in _SVG_URL.finditer(value):
            inspect_data_uri(match.group(2))

    try:
        for event, element in ET.iterparse(io.BytesIO(data), events=("start", "end")):
            if event == "end":
                tag = element.tag.rsplit("}", 1)[-1].lower()
                if tag == "style":
                    inspect_url_references(element.text or "")
                depth -= 1
                continue
            count += 1
            depth += 1
            if count > 100_000 or depth > 256:
                raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
            if root is None:
                root = element
            tag = element.tag.rsplit("}", 1)[-1].lower()
            if tag in {"script", "foreignobject"}:
                raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
            for raw_name, raw_value in element.attrib.items():
                name = raw_name.rsplit("}", 1)[-1].lower()
                value = raw_value.strip()
                if name in {"font-face-uri", "font-family"} and "url(" in value.lower():
                    raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
                if name in {"href", "src"}:
                    inspect_data_uri(
                        value, raster_required=tag in {"image", "feimage"}
                    )
                else:
                    inspect_url_references(value)
    except _ResourceRejected:
        raise
    except (ET.ParseError, UnicodeError, ValueError):
        raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)

    if root is None or root.tag.rsplit("}", 1)[-1].lower() != "svg":
        raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
    width = _svg_number(root.attrib.get("width"))
    height = _svg_number(root.attrib.get("height"))
    if width is None or height is None:
        view_box = root.attrib.get("viewBox") or root.attrib.get("viewbox")
        if view_box:
            try:
                values = [float(item) for item in re.split(r"[\s,]+", view_box.strip())]
                if len(values) == 4:
                    width = max(1, int(math.ceil(values[2])))
                    height = max(1, int(math.ceil(values[3])))
            except (ValueError, OverflowError):
                pass
    width = width or 1
    height = height or 1
    _check_geometry(width, height)
    return (
        "image/svg+xml",
        "svg-static-v1",
        data,
        ImageProfile(width, height, None, None, 1, "SVG", False),
        False,
    )


def _has_supported_raster_signature(data: bytes) -> bool:
    return (
        data.startswith(b"\x89PNG\r\n\x1a\n")
        or data.startswith(b"\xff\xd8\xff")
        or data.startswith(
            (
                b"GIF87a",
                b"GIF89a",
                b"BM",
                b"II*\x00",
                b"MM\x00*",
                b"II+\x00",
                b"MM\x00+",
            )
        )
    )


def _prepare_payload(data: bytes) -> tuple[str, str, bytes, ImageProfile, bool]:
    if data.lstrip().startswith(b"<"):
        return _prepare_svg(data)
    try:
        return _prepare_raster(data)
    except _ResourceRejected as rejected:
        if (
            rejected.code == RESOURCE_DECODE_FAILED
            and not _has_supported_raster_signature(data)
        ):
            raise _ResourceRejected(RESOURCE_MEDIA_TYPE_UNSUPPORTED)
        raise


def _degradation(
    node_id: Optional[str], code: str, source_path: str
) -> ResourceDegradation:
    return ResourceDegradation(
        node_id=node_id,
        code=code,
        message=f"Resource rejected: {code}",
        fallback_text=f"[{code}]",
        source_path=source_path,
    )


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
            degradations.append(_degradation(node_id, error_code, normalized_source_path))
            continue

        if resolved is None:
            continue

        try:
            with resolved.open("rb") as source:
                data = source.read(MAX_RESOURCE_BYTES + 1)
        except OSError as exc:
            code = RESOURCE_READ_FAILED if exc.errno in {13} else RESOURCE_NOT_FOUND
            degradations.append(_degradation(node_id, code, normalized_source_path))
            continue

        if len(data) > MAX_RESOURCE_BYTES:
            degradations.append(
                _degradation(node_id, RESOURCE_TOO_LARGE, normalized_source_path)
            )
            continue

        try:
            media_type, normalizer_id, payload, profile, multiframe = _prepare_payload(data)
        except _ResourceRejected as rejected:
            degradations.append(
                _degradation(node_id, rejected.code, normalized_source_path)
            )
            continue
        if len(payload) > MAX_RESOURCE_BYTES:
            degradations.append(
                _degradation(node_id, RESOURCE_TOO_LARGE, normalized_source_path)
            )
            continue

        source_digest = _sha256(data)
        resources.append(
            PreflightResource(
                resource_id=_resource_id_for_path(normalized_source_path),
                source_path=normalized_source_path,
                source_sha256=source_digest,
                payload_sha256=_sha256(payload),
                byte_length=len(payload),
                media_type=media_type,
                normalizer_id=normalizer_id,
                image_profile=profile,
                payload_bytes=payload,
            )
        )
        if multiframe:
            degradations.append(
                _degradation(node_id, MULTIFRAME_FLATTENED, normalized_source_path)
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
    "RESOURCE_TOO_LARGE",
    "RESOURCE_PIXEL_LIMIT_EXCEEDED",
    "RESOURCE_SIDE_LIMIT_EXCEEDED",
    "MULTIFRAME_FLATTENED",
    "MAX_RESOURCE_BYTES",
    "MAX_IMAGE_PIXELS",
    "MAX_IMAGE_SIDE",
    "ImageProfile",
    "PreflightResource",
    "PreparedLongformResource",
    "ResourceDegradation",
    "ResourcePreflight",
    "preflight_resources",
]
