from __future__ import annotations

import base64
import hashlib
import io
import struct
from dataclasses import replace
from pathlib import Path

import pytest
from PIL import Image, ImageCms, ImageOps

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


def _image_payload(image: Image.Image, format_name: str, **kwargs: object) -> bytes:
    output = io.BytesIO()
    image.save(output, format=format_name, **kwargs)
    return output.getvalue()


def _assert_rgb_pixels_close(
    actual: Image.Image, expected: Image.Image, tolerance: int = 1
) -> None:
    """Compare decoded visual RGB pixels after a documented mode conversion."""
    actual_pixels = list(actual.convert("RGB").getdata())
    expected_pixels = list(expected.convert("RGB").getdata())
    assert len(actual_pixels) == len(expected_pixels)
    assert max(
        abs(actual_channel - expected_channel)
        for actual_pixel, expected_pixel in zip(actual_pixels, expected_pixels)
        for actual_channel, expected_channel in zip(actual_pixel, expected_pixel)
    ) <= tolerance


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
        with Image.open(io.BytesIO(payload)) as source:
            source.load()
            expected = ImageOps.exif_transpose(source)
        _assert_rgb_pixels_close(normalized, expected, tolerance=0)


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
        with Image.open(MEDIA / fixture) as source:
            source.seek(0)
            source.load()
            expected = ImageOps.exif_transpose(source.copy())
        _assert_rgb_pixels_close(normalized, expected, tolerance=0)
    assert [item.code for item in result.degradations] == [MULTIFRAME_FLATTENED]
    diagnostic = repr(result.degradations[0]) + result.degradations[0].message
    assert fixture not in diagnostic
    assert str(MEDIA) not in diagnostic


def test_oriented_cmyk_jpeg_normalizes_to_deterministic_rgb_png(
    tmp_path: Path,
) -> None:
    source_image = Image.new("CMYK", (2, 3))
    source_image.putdata(
        [
            (0, 10, 20, 0),
            (30, 0, 40, 10),
            (60, 20, 0, 5),
            (0, 80, 30, 20),
            (50, 40, 30, 10),
            (10, 20, 90, 0),
        ]
    )
    exif = Image.Exif()
    exif[274] = 6
    # A structurally valid RGB profile is incompatible with CMYK pixels. The
    # converted output must not retain stale metadata that describes RGB input.
    incompatible_icc = ImageCms.ImageCmsProfile(
        ImageCms.createProfile("sRGB")
    ).tobytes()
    payload = _image_payload(
        source_image,
        "JPEG",
        quality=100,
        subsampling=0,
        exif=exif,
        icc_profile=incompatible_icc,
    )

    result = _preflight(tmp_path, "oriented-cmyk.jpg", payload)
    resource = result.resources[0]

    assert resource.normalizer_id == "exif-transpose-png-v1"
    assert resource.media_type == "image/png"
    assert resource.image_profile.has_icc is False
    with Image.open(io.BytesIO(payload)) as decoded_source:
        decoded_source.load()
        expected = ImageOps.exif_transpose(decoded_source).convert("RGB")
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.mode == "RGB"
        assert normalized.info.get("icc_profile") is None
        _assert_rgb_pixels_close(normalized, expected, tolerance=1)


def test_multiframe_cmyk_tiff_normalizes_first_page_to_rgb_png(tmp_path: Path) -> None:
    first = Image.new("CMYK", (2, 2))
    first.putdata(
        [(0, 10, 20, 0), (40, 0, 30, 10), (20, 60, 0, 5), (5, 15, 70, 20)]
    )
    second = Image.new("CMYK", (2, 2), (90, 80, 70, 60))
    payload = _image_payload(
        first,
        "TIFF",
        save_all=True,
        append_images=[second],
        compression="raw",
    )

    result = _preflight(tmp_path, "cmyk-pages.tiff", payload)
    resource = result.resources[0]

    assert resource.normalizer_id == "tiff-first-page-png-v1"
    assert [item.code for item in result.degradations] == [MULTIFRAME_FLATTENED]
    with Image.open(io.BytesIO(payload)) as decoded_source:
        decoded_source.seek(0)
        decoded_source.load()
        expected = decoded_source.convert("RGB")
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.mode == "RGB"
        _assert_rgb_pixels_close(normalized, expected, tolerance=1)


def test_multiframe_float_tiff_normalizes_first_page_deterministically(
    tmp_path: Path,
) -> None:
    first = Image.new("F", (2, 2))
    first.putdata([0.0, 64.5, 255.0, 512.0])
    second = Image.new("F", (2, 2), 999.0)
    payload = _image_payload(
        first,
        "TIFF",
        save_all=True,
        append_images=[second],
        compression="raw",
    )

    result = _preflight(tmp_path, "numeric-float.tiff", payload)
    resource = result.resources[0]

    assert resource.normalizer_id == "tiff-first-page-png-v1"
    assert [item.code for item in result.degradations] == [MULTIFRAME_FLATTENED]
    with Image.open(io.BytesIO(payload)) as decoded_source:
        decoded_source.seek(0)
        decoded_source.load()
        expected = decoded_source.convert("RGB")
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.mode == "RGB"
        _assert_rgb_pixels_close(normalized, expected, tolerance=1)


@pytest.mark.parametrize(
    ("mode", "endian"),
    [("I;16", "<"), ("I;16L", "<"), ("I;16B", ">"), ("I", None)],
)
def test_multiframe_unsigned_16bit_tiff_preserves_exact_first_page_samples(
    tmp_path: Path, mode: str, endian: str | None
) -> None:
    samples = [0, 64, 1024, 65535]
    second_samples = [5, 1000, 30000, 60000]
    if mode == "I":
        first = Image.new("I", (2, 2))
        first.putdata(samples)
        second = Image.new("I", (2, 2))
        second.putdata(second_samples)
    else:
        assert endian is not None
        first = Image.frombytes(mode, (2, 2), struct.pack(f"{endian}4H", *samples))
        second = Image.frombytes(
            mode, (2, 2), struct.pack(f"{endian}4H", *second_samples)
        )
    payload = _image_payload(
        first,
        "TIFF",
        save_all=True,
        append_images=[second],
        compression="raw",
    )

    result = _preflight(tmp_path, f"unsigned-{mode}.tiff", payload)
    resource = result.resources[0]

    assert resource.normalizer_id == "tiff-first-page-png-v1"
    assert [item.code for item in result.degradations] == [MULTIFRAME_FLATTENED]
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.mode == "I;16"
        assert list(normalized.getdata()) == samples


def test_oriented_unsigned_16bit_tiff_preserves_exact_transposed_samples(
    tmp_path: Path,
) -> None:
    samples = [0, 64, 1024, 4096, 32768, 65535]
    source = Image.frombytes("I;16", (2, 3), struct.pack("<6H", *samples))
    exif = Image.Exif()
    exif[274] = 6
    payload = _image_payload(source, "TIFF", compression="raw", exif=exif)

    result = _preflight(tmp_path, "oriented-16bit.tiff", payload)
    resource = result.resources[0]

    assert resource.normalizer_id == "exif-transpose-png-v1"
    with Image.open(io.BytesIO(payload)) as decoded_source:
        decoded_source.load()
        expected = list(ImageOps.exif_transpose(decoded_source).getdata())
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.mode == "I;16"
        assert list(normalized.getdata()) == expected


def test_wide_signed_integer_tiff_uses_deterministic_rgb_fallback(
    tmp_path: Path,
) -> None:
    first = Image.new("I", (2, 2))
    first.putdata([-1, 0, 65535, 65536])
    second = Image.new("I", (2, 2), 1)
    payload = _image_payload(first, "TIFF", save_all=True, append_images=[second])

    result = _preflight(tmp_path, "wide-signed.tiff", payload)
    resource = result.resources[0]

    with Image.open(io.BytesIO(payload)) as decoded_source:
        decoded_source.seek(0)
        decoded_source.load()
        expected = decoded_source.convert("RGB")
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.mode == "RGB"
        _assert_rgb_pixels_close(normalized, expected, tolerance=0)


def test_grayscale_raster_drops_structurally_valid_rgb_icc(tmp_path: Path) -> None:
    rgb_icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    first = Image.new("L", (2, 2))
    first.putdata([0, 64, 128, 255])
    second = Image.new("L", (2, 2), 32)
    payload = _image_payload(
        first,
        "TIFF",
        save_all=True,
        append_images=[second],
        icc_profile=rgb_icc,
    )

    result = _preflight(tmp_path, "gray-rgb-profile.tiff", payload)
    resource = result.resources[0]

    assert resource.image_profile.has_icc is False
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        assert normalized.mode == "L"
        assert normalized.info.get("icc_profile") is None


def test_rgb_raster_retains_compatible_icc_after_normalization(tmp_path: Path) -> None:
    rgb_icc = ImageCms.ImageCmsProfile(ImageCms.createProfile("sRGB")).tobytes()
    first = Image.new("RGB", (2, 2), (20, 40, 60))
    second = Image.new("RGB", (2, 2), (80, 100, 120))
    payload = _image_payload(
        first,
        "TIFF",
        save_all=True,
        append_images=[second],
        icc_profile=rgb_icc,
    )

    result = _preflight(tmp_path, "rgb-profile.tiff", payload)
    resource = result.resources[0]

    assert resource.image_profile.has_icc is True
    with Image.open(io.BytesIO(resource.payload_bytes)) as normalized:
        normalized.load()
        retained_icc = normalized.info.get("icc_profile")
        assert retained_icc
        retained_profile = ImageCms.ImageCmsProfile(io.BytesIO(retained_icc))
        assert retained_profile.profile.xcolor_space == "RGB "


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
    "svg",
    [
        (
            r'<svg xmlns="http://www.w3.org/2000/svg"><style>'
            r'@\69mport "https://example.invalid/a.css";</style></svg>'
        ),
        (
            r'<svg xmlns="http://www.w3.org/2000/svg"><style>'
            r'.x{fill:u\72l(https://example.invalid/x)}</style></svg>'
        ),
        (
            '<svg xmlns="http://www.w3.org/2000/svg"><style>'
            '@keyframes pulse{to{opacity:0}}.x{animation:pulse 1s}'
            "</style></svg>"
        ),
        '<svg xmlns="http://www.w3.org/2000/svg"><rect style="fill:red"/></svg>',
    ],
)
def test_svg_style_elements_and_attributes_are_always_rejected(
    tmp_path: Path, svg: str
) -> None:
    result = _preflight(tmp_path, "style.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


@pytest.mark.parametrize(
    "element",
    [
        "animateColor",
        "discard",
        'ev:handler xmlns:ev="http://www.w3.org/2001/xml-events"',
        'html:iframe xmlns:html="http://www.w3.org/1999/xhtml"',
        "unknownWidget",
    ],
)
def test_svg_animation_interaction_and_unknown_elements_are_rejected(
    tmp_path: Path, element: str
) -> None:
    svg = f'<svg xmlns="http://www.w3.org/2000/svg"><{element}/></svg>'

    result = _preflight(tmp_path, "active-element.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


@pytest.mark.parametrize(
    "attribute",
    [
        'xmlns:ev="http://www.w3.org/2001/xml-events" ev:handler="#handler"',
        'xmlns:ev="http://www.w3.org/2001/xml-events" ev:event="click"',
        'xmlns:unknown="urn:active" unknown:payload="x"',
    ],
)
def test_svg_xml_events_and_unknown_attribute_namespaces_are_rejected(
    tmp_path: Path, attribute: str
) -> None:
    svg = f'<svg xmlns="http://www.w3.org/2000/svg"><rect {attribute}/></svg>'

    result = _preflight(tmp_path, "active-attribute.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_svg_non_declaration_processing_instruction_is_rejected(tmp_path: Path) -> None:
    svg = (
        '<?xml version="1.0"?><?reviewer active="yes"?>'
        '<svg xmlns="http://www.w3.org/2000/svg"><rect width="1" height="1"/></svg>'
    )

    result = _preflight(tmp_path, "processing-instruction.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_svg_entities_are_normalized_before_reference_checks(tmp_path: Path) -> None:
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg">'
        '<rect fill="url(&#x68;ttps://example.invalid/external.svg)"/>'
        "</svg>"
    )

    result = _preflight(tmp_path, "entity-reference.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_safe_static_svg_allowlist_supports_shapes_gradients_clips_and_use(
    tmp_path: Path,
) -> None:
    svg = b'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20">
  <defs>
    <linearGradient id="gradient"><stop offset="0" stop-color="#fff"/></linearGradient>
    <clipPath id="clip"><circle cx="10" cy="10" r="8"/></clipPath>
    <g id="shape"><path d="M0 0 L10 0 L10 10 Z"/></g>
  </defs>
  <rect width="40" height="20" fill="url(#gradient)" clip-path="url(#clip)"/>
  <use href="&#35;shape" x="5" y="5"/>
  <line x1="0" y1="0" x2="40" y2="20"/>
  <polygon points="0,0 2,0 1,2"/>
  <text x="1" y="15"><tspan>safe</tspan></text>
</svg>'''

    result = _preflight(tmp_path, "safe-static.svg", svg)

    assert len(result.resources) == 1
    assert result.resources[0].media_type == "image/svg+xml"
    assert result.degradations == []


@pytest.mark.parametrize(
    "attribute",
    [
        'onload="alert(1)"',
        'OnClick="alert(1)"',
        'xmlns:event="urn:test" event:OnLoad="alert(1)"',
    ],
)
def test_svg_event_handler_attributes_are_rejected(
    tmp_path: Path, attribute: str
) -> None:
    svg = f'<svg xmlns="http://www.w3.org/2000/svg" {attribute}/>'

    result = _preflight(tmp_path, "event.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


@pytest.mark.parametrize(
    "tag",
    ["animate", "animateTransform", "animateMotion", "set", "ANIMATE"],
)
def test_svg_smil_animation_elements_are_rejected(tmp_path: Path, tag: str) -> None:
    svg = f'<svg xmlns="http://www.w3.org/2000/svg"><{tag}/></svg>'

    result = _preflight(tmp_path, "smil.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


@pytest.mark.parametrize(
    "svg",
    [
        '<svg xmlns="http://www.w3.org/2000/svg"><style>@import "https://example.invalid/a.css";</style></svg>',
        "<svg xmlns='http://www.w3.org/2000/svg'><style>@IMPORT\n 'https://example.invalid/a.css';</style></svg>",
        "<svg xmlns='http://www.w3.org/2000/svg'><rect style=\"@import 'https://example.invalid/a.css';\"/></svg>",
        '<svg xmlns="http://www.w3.org/2000/svg"><style>@import url(  "https://example.invalid/a.css"  );</style></svg>',
    ],
)
def test_svg_external_css_imports_are_rejected(tmp_path: Path, svg: str) -> None:
    result = _preflight(tmp_path, "css-import.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_svg_xml_stylesheet_processing_instruction_is_rejected(tmp_path: Path) -> None:
    svg = (
        '<?xml-stylesheet type="text/css" href="https://example.invalid/a.css"?>'
        '<svg xmlns="http://www.w3.org/2000/svg"/>'
    )

    result = _preflight(tmp_path, "stylesheet.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


@pytest.mark.parametrize("reference_name", ["href", "xlink:href", "src"])
def test_svg_external_reference_variants_are_rejected(
    tmp_path: Path, reference_name: str
) -> None:
    namespace = (
        ' xmlns:xlink="http://www.w3.org/1999/xlink"'
        if ":" in reference_name
        else ""
    )
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg"{namespace}>'
        f'<image {reference_name}="  HTTPS://example.invalid/a.png  "/>'
        "</svg>"
    )

    result = _preflight(tmp_path, "external-reference.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


@pytest.mark.parametrize("declared_type", ["image/svg+xml", "application/octet-stream"])
def test_svg_nested_active_data_payloads_are_rejected(
    tmp_path: Path, declared_type: str
) -> None:
    nested = base64.b64encode(
        b'<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>'
    ).decode("ascii")
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f'<rect style="fill:url( data:{declared_type};base64,{nested} )"/>'
        "</svg>"
    )

    result = _preflight(tmp_path, "nested-active.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


def test_safe_embedded_raster_data_uri_remains_allowed(tmp_path: Path) -> None:
    embedded = base64.b64encode((MEDIA / "png-wrong.jpg").read_bytes()).decode(
        "ascii"
    )
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="4" height="3">'
        f'<image href="data:image/png;base64,{embedded}" width="4" height="3"/>'
        "</svg>"
    )

    result = _preflight(tmp_path, "safe-raster.svg", svg.encode("utf-8"))

    assert len(result.resources) == 1
    assert result.resources[0].media_type == "image/svg+xml"
    assert result.degradations == []


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


def test_nonimage_data_uri_is_rejected_by_static_svg_contract(
    tmp_path: Path,
) -> None:
    embedded = base64.b64encode(b"not an image").decode("ascii")
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="2" height="2">'
        f"<style>.data{{fill:url(data:application/octet-stream;base64,{embedded})}}</style>"
        "</svg>"
    )

    result = _preflight(tmp_path, "nonimage-data.svg", svg.encode("utf-8"))

    assert result.resources == []
    assert [item.code for item in result.degradations] == [
        RESOURCE_MEDIA_TYPE_UNSUPPORTED
    ]


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
