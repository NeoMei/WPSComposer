from __future__ import annotations

import base64
import hashlib
import io
import struct
from dataclasses import replace
from pathlib import Path

import pytest
from PIL import Image

from skills.WPSComposer.scripts.document_model import ExcalidrawBlock, ImageBlock
from skills.WPSComposer.scripts.longform import resources as resource_module
from skills.WPSComposer.scripts.longform.pipeline import _build_executor_resources
from skills.WPSComposer.scripts.longform.resources import (
    MULTIFRAME_FLATTENED,
    MAX_IMAGE_PIXELS,
    MAX_IMAGE_SIDE,
    MAX_RESOURCE_BYTES,
    RESOURCE_DECODE_FAILED,
    RESOURCE_MEDIA_TYPE_UNSUPPORTED,
    RESOURCE_PIXEL_LIMIT_EXCEEDED,
    RESOURCE_SIDE_LIMIT_EXCEEDED,
    RESOURCE_TOO_LARGE,
    ImageProfile,
    PreparedLongformResource,
    PreflightResource,
    preflight_resources,
)


MEDIA = Path(__file__).parent / "fixtures" / "media"


def _preflight(tmp_path: Path, name: str, payload: bytes):
    target = tmp_path / name
    target.write_bytes(payload)
    return preflight_resources([ImageBlock(path=name, alt="fixture")], str(tmp_path))


def _minimal_bigtiff() -> bytes:
    entries = [
        (256, 3, 1, 1),
        (257, 3, 1, 1),
        (258, 3, 1, 8),
        (259, 3, 1, 1),
        (262, 3, 1, 1),
        (273, 16, 1, 232),
        (277, 3, 1, 1),
        (278, 3, 1, 1),
        (279, 16, 1, 1),
        (284, 3, 1, 1),
    ]
    payload = bytearray(b"II+\x00" + struct.pack("<HHQ", 8, 0, 16))
    payload += struct.pack("<Q", len(entries))
    for tag, field_type, count, value in entries:
        payload += struct.pack("<HHQQ", tag, field_type, count, value)
    payload += struct.pack("<Q", 0) + b"\x80"
    return bytes(payload)


def test_m3_resource_limits_are_exact() -> None:
    assert MAX_RESOURCE_BYTES == 50 * 1024 * 1024
    assert MAX_IMAGE_PIXELS == 80_000_000
    assert MAX_IMAGE_SIDE == 32_768


@pytest.mark.parametrize(
    ("fixture", "media_type", "source_format"),
    [
        ("png-wrong.jpg", "image/png", "PNG"),
        ("jpeg-wrong.png", "image/jpeg", "JPEG"),
        ("tiff-wrong.dat", "image/tiff", "TIFF"),
        ("bmp-wrong.gif", "image/bmp", "BMP"),
        ("gif-wrong.bmp", "image/gif", "GIF"),
    ],
)
def test_raster_media_type_comes_from_fully_decoded_content(
    tmp_path: Path, fixture: str, media_type: str, source_format: str
) -> None:
    result = _preflight(tmp_path, fixture, (MEDIA / fixture).read_bytes())

    assert len(result.resources) == 1
    assert result.resources[0].media_type == media_type
    assert result.resources[0].image_profile.source_format == source_format


def test_pillow_decodable_bigtiff_is_accepted_without_a_manual_prefix_allowlist(
    tmp_path: Path,
) -> None:
    payload = _minimal_bigtiff()
    with Image.open(io.BytesIO(payload)) as decoded:
        decoded.load()
        assert decoded.format == "TIFF"

    result = _preflight(tmp_path, "bigtiff.bin", payload)

    assert result.degradations == []
    assert result.resources[0].media_type == "image/tiff"
    assert result.resources[0].image_profile.source_format == "TIFF"


def test_renamed_webp_is_rejected(tmp_path: Path) -> None:
    result = _preflight(tmp_path, "looks-like.png", (MEDIA / "webp-renamed.png").read_bytes())

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_MEDIA_TYPE_UNSUPPORTED]


def test_corrupt_signature_is_rejected_after_decode(tmp_path: Path) -> None:
    result = _preflight(tmp_path, "corrupt.png", b"\x89PNG\r\n\x1a\ntruncated")

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_DECODE_FAILED]


def test_50_mib_resource_limit_is_enforced_before_decode(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(resource_module, "MAX_RESOURCE_BYTES", 31)
    result = _preflight(tmp_path, "large.png", b"x" * 32)

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_TOO_LARGE]


def test_source_is_opened_once_and_read_once_with_a_bounded_limit(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    source = tmp_path / "large.png"
    source.write_bytes(b"x" * 32)
    monkeypatch.setattr(resource_module, "MAX_RESOURCE_BYTES", 31)
    original_open = Path.open
    open_count = 0
    read_sizes: list[int] = []

    class TrackingReader:
        def __init__(self, handle: object) -> None:
            self.handle = handle

        def __enter__(self) -> "TrackingReader":
            return self

        def __exit__(self, *args: object) -> None:
            self.handle.close()

        def read(self, size: int = -1) -> bytes:
            read_sizes.append(size)
            return self.handle.read(size)

    def tracking_open(path: Path, *args: object, **kwargs: object):
        nonlocal open_count
        handle = original_open(path, *args, **kwargs)
        if path == source:
            open_count += 1
            return TrackingReader(handle)
        return handle

    monkeypatch.setattr(Path, "open", tracking_open)

    result = preflight_resources([ImageBlock(path=source.name)], str(tmp_path))

    assert [item.code for item in result.degradations] == [RESOURCE_TOO_LARGE]
    assert open_count == 1
    assert read_sizes == [32]


@pytest.mark.parametrize(
    ("limit_name", "limit", "code"),
    [
        ("MAX_IMAGE_PIXELS", 11, RESOURCE_PIXEL_LIMIT_EXCEEDED),
        ("MAX_IMAGE_SIDE", 3, RESOURCE_SIDE_LIMIT_EXCEEDED),
    ],
)
def test_image_geometry_limits_are_enforced(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    limit_name: str,
    limit: int,
    code: str,
) -> None:
    monkeypatch.setattr(resource_module, limit_name, limit)
    result = _preflight(tmp_path, "tiny.png", (MEDIA / "png-wrong.jpg").read_bytes())

    assert result.resources == []
    assert [item.code for item in result.degradations] == [code]


def test_pillow_decompression_bomb_warning_is_a_decode_failure(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr(Image, "MAX_IMAGE_PIXELS", 8)
    result = _preflight(tmp_path, "bomb.png", (MEDIA / "png-wrong.jpg").read_bytes())

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_DECODE_FAILED]


def test_single_frame_untransformed_raster_uses_original_payload() -> None:
    payload = (MEDIA / "png-wrong.jpg").read_bytes()
    result = preflight_resources(
        [ImageBlock(path="png-wrong.jpg", alt="png")], str(MEDIA)
    )
    resource = result.resources[0]

    assert resource.normalizer_id == "none-v1"
    assert resource.payload_bytes == payload
    assert resource.source_sha256 == resource.payload_sha256 == hashlib.sha256(payload).hexdigest()


def test_exif_orientation_normalizes_losslessly_and_retains_valid_icc() -> None:
    payload = (MEDIA / "oriented.jpg").read_bytes()
    result = preflight_resources(
        [ImageBlock(path="oriented.jpg", alt="oriented")], str(MEDIA)
    )
    resource = result.resources[0]

    assert resource.normalizer_id == "exif-transpose-png-v1"
    assert resource.media_type == "image/png"
    assert resource.source_sha256 == hashlib.sha256(payload).hexdigest()
    assert resource.payload_sha256 == hashlib.sha256(resource.payload_bytes).hexdigest()
    assert resource.source_sha256 != resource.payload_sha256
    assert (resource.image_profile.pixel_width, resource.image_profile.pixel_height) == (3, 2)
    assert resource.image_profile.has_icc is True
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.format == "PNG"
        assert normalized.size == (3, 2)
        assert normalized.info.get("icc_profile")


@pytest.mark.parametrize(
    ("fixture", "normalizer_id"),
    [
        ("animated.gif", "gif-first-frame-png-v1"),
        ("multipage.tiff", "tiff-first-page-png-v1"),
    ],
)
def test_multiframe_inputs_flatten_first_frame_without_path_leaks(
    fixture: str, normalizer_id: str
) -> None:
    result = preflight_resources(
        [ImageBlock(path=fixture, alt="multi")], str(MEDIA)
    )
    resource = result.resources[0]

    assert resource.normalizer_id == normalizer_id
    assert resource.media_type == "image/png"
    assert resource.image_profile.frame_count == 2
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.format == "PNG"
        assert getattr(normalized, "n_frames", 1) == 1
    assert [item.code for item in result.degradations] == [MULTIFRAME_FLATTENED]
    diagnostic = repr(result.degradations[0]) + result.degradations[0].message
    assert fixture not in diagnostic
    assert str(MEDIA) not in diagnostic


def test_preflight_private_payload_and_path_do_not_affect_repr_or_equality() -> None:
    result = preflight_resources(
        [ImageBlock(path="png-wrong.jpg", alt="png")], str(MEDIA)
    )
    resource = result.resources[0]
    altered = replace(resource, source_path="private/other.png", payload_bytes=b"other")

    assert altered == resource
    assert "png-wrong.jpg" not in repr(resource)
    assert "payload_bytes" not in repr(resource)
    assert repr(resource) == repr(altered)


def test_manifest_entry_remains_exactly_six_allowed_keys() -> None:
    result = preflight_resources(
        [ImageBlock(path="png-wrong.jpg", alt="png")], str(MEDIA)
    )

    assert set(result.manifest["entries"][0]) == {
        "resourceId",
        "sourceSha256",
        "payloadSha256",
        "byteLength",
        "mediaType",
        "normalizerId",
    }
    serialized = repr(result.manifest)
    assert "payload_bytes" not in serialized
    assert str(MEDIA) not in serialized


def test_executor_resources_use_private_normalized_bytes_without_rereading_source(
    tmp_path: Path,
) -> None:
    source = tmp_path / "oriented.jpg"
    source.write_bytes((MEDIA / "oriented.jpg").read_bytes())
    preflight = preflight_resources([ImageBlock(path=source.name)], str(tmp_path))
    source.unlink()

    (prepared,) = _build_executor_resources(str(tmp_path), preflight)

    assert isinstance(prepared, PreparedLongformResource)
    accepted = preflight.resources[0]
    assert prepared.id == accepted.resource_id
    assert prepared.payload_bytes == accepted.payload_bytes
    assert prepared.source_sha256 == accepted.source_sha256
    assert prepared.payload_sha256 == accepted.payload_sha256
    assert prepared.image_profile == accepted.image_profile
    assert "oriented.jpg" not in repr(prepared)
    assert "payload_bytes" not in repr(prepared)
    assert prepared == replace(prepared, payload_bytes=b"different")


def test_static_svg_is_retained_as_svg() -> None:
    payload = (MEDIA / "static.svg").read_bytes()
    result = preflight_resources([ImageBlock(path="static.svg", alt="svg")], str(MEDIA))
    resource = result.resources[0]

    assert result.degradations == []
    assert resource.media_type == "image/svg+xml"
    assert resource.normalizer_id == "svg-static-v1"
    assert resource.payload_bytes == payload
    assert resource.image_profile == ImageProfile(40, 20, None, None, 1, "SVG", False)


@pytest.mark.parametrize(
    "payload",
    [
        b'{"type":"excalidraw"}',
        b'<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>',
    ],
)
def test_excalidraw_markdown_must_contain_safe_static_svg(
    tmp_path: Path, payload: bytes
) -> None:
    target = tmp_path / "unsafe.excalidraw.md"
    target.write_bytes(payload)

    result = preflight_resources(
        [ExcalidrawBlock(path=target.name, alt="unsafe")], str(tmp_path)
    )

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


@pytest.mark.parametrize(
    "svg",
    [
        '<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>',
        '<svg xmlns="http://www.w3.org/2000/svg"><foreignObject/></svg>',
        '<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.invalid/x.png"/></svg>',
        '<svg xmlns="http://www.w3.org/2000/svg"><image href="local.png"/></svg>',
        '<svg xmlns="http://www.w3.org/2000/svg"><style>@font-face{src:url(font.woff)}</style></svg>',
        '<svg xmlns="http://www.w3.org/2000/svg"><style>.x{fill:url(https://example.invalid/x)}</style></svg>',
        '<!DOCTYPE svg><svg xmlns="http://www.w3.org/2000/svg"/>',
        '<!ENTITY x "boom"><svg xmlns="http://www.w3.org/2000/svg"/>',
    ],
)
def test_unsafe_svg_features_are_rejected(tmp_path: Path, svg: str) -> None:
    result = _preflight(tmp_path, "unsafe.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_MEDIA_TYPE_UNSUPPORTED]


def test_svg_element_limit_is_enforced(tmp_path: Path) -> None:
    svg = '<svg xmlns="http://www.w3.org/2000/svg">' + "<g/>" * 100_000 + "<g/></svg>"
    result = _preflight(tmp_path, "many.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_MEDIA_TYPE_UNSUPPORTED]


def test_svg_depth_limit_is_enforced(tmp_path: Path) -> None:
    svg = '<svg xmlns="http://www.w3.org/2000/svg">' + "<g>" * 256 + "</g>" * 256 + "</svg>"
    result = _preflight(tmp_path, "deep.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_MEDIA_TYPE_UNSUPPORTED]


@pytest.mark.parametrize("tag", ["image", "feImage"])
def test_embedded_svg_raster_uses_same_resource_limits(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, tag: str
) -> None:
    embedded = base64.b64encode((MEDIA / "png-wrong.jpg").read_bytes()).decode("ascii")
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20">'
        f'<{tag} href="data:image/png;base64,{embedded}"/></svg>'
    )
    monkeypatch.setattr(resource_module, "MAX_IMAGE_SIDE", 3)
    result = _preflight(tmp_path, "embedded.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [RESOURCE_MEDIA_TYPE_UNSUPPORTED]


def test_style_data_uri_raster_uses_same_resource_limits(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    embedded = base64.b64encode((MEDIA / "png-wrong.jpg").read_bytes()).decode(
        "ascii"
    )
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f"<style>.cursor{{cursor:url(data:image/png;base64,{embedded})}}</style>"
        "</svg>"
    )
    monkeypatch.setattr(resource_module, "MAX_IMAGE_SIDE", 3)

    result = _preflight(tmp_path, "style-data-uri.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_ordinary_attribute_data_uri_raster_uses_same_resource_limits(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    embedded = base64.b64encode((MEDIA / "png-wrong.jpg").read_bytes()).decode(
        "ascii"
    )
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f'<rect width="10" height="10" cursor="url(data:image/png;base64,{embedded})"/>'
        "</svg>"
    )
    monkeypatch.setattr(resource_module, "MAX_IMAGE_SIDE", 3)

    result = _preflight(tmp_path, "attribute-data-uri.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_svg_data_uri_occurrence_is_counted_once(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    payload = (MEDIA / "png-wrong.jpg").read_bytes()
    embedded = base64.b64encode(payload).decode("ascii")
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20">'
        f'<image href="data:image/png;base64,{embedded}"/>'
        "</svg>"
    ).encode("utf-8")
    monkeypatch.setattr(resource_module, "MAX_RESOURCE_BYTES", len(svg) + len(payload))

    result = _preflight(tmp_path, "exact-budget.svg", svg)

    assert len(result.resources) == 1
    assert result.degradations == []


def test_mislabeled_raster_data_uri_uses_decoded_content_for_geometry_limits(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    embedded = base64.b64encode((MEDIA / "png-wrong.jpg").read_bytes()).decode(
        "ascii"
    )
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f"<style>.cursor{{cursor:url(data:application/octet-stream;base64,{embedded})}}</style>"
        "</svg>"
    )
    monkeypatch.setattr(resource_module, "MAX_IMAGE_SIDE", 3)

    result = _preflight(tmp_path, "mislabeled-raster.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_mislabeled_webp_data_uri_remains_unsupported(tmp_path: Path) -> None:
    embedded = base64.b64encode((MEDIA / "webp-renamed.png").read_bytes()).decode(
        "ascii"
    )
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f'<rect cursor="url(data:application/octet-stream;base64,{embedded})"/>'
        "</svg>"
    )

    result = _preflight(tmp_path, "mislabeled-webp.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_genuine_nonimage_data_uri_does_not_receive_raster_geometry_checks(
    tmp_path: Path,
) -> None:
    embedded = base64.b64encode(b"not an image").decode("ascii")
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f"<style>.data{{fill:url(data:application/octet-stream;base64,{embedded})}}</style>"
        "</svg>"
    )

    result = _preflight(tmp_path, "nonimage-data.svg", svg.encode("utf-8"))

    assert len(result.resources) == 1
    assert result.degradations == []


def test_corrupt_payload_declared_as_supported_image_is_rejected(tmp_path: Path) -> None:
    embedded = base64.b64encode(b"not an image").decode("ascii")
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f'<rect cursor="url(data:image/png;base64,{embedded})"/>'
        "</svg>"
    )

    result = _preflight(tmp_path, "corrupt-declared-image.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]
