from __future__ import annotations

import hashlib
import os
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.document_model import (
    DegradationBlock,
    DocumentIssue,
    ExcalidrawBlock,
    FigureBlock,
    ImageBlock,
    Section,
)
from skills.WPSComposer.scripts.longform.resources import (
    RESOURCE_ABSOLUTE_PATH_OUTSIDE,
    RESOURCE_MEDIA_TYPE_UNSUPPORTED,
    RESOURCE_NOT_FOUND,
    RESOURCE_PATH_ESCAPES_BASE,
    RESOURCE_READ_FAILED,
    PreflightResource,
    ResourceDegradation,
    preflight_resources,
)


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _make_image(base_dir: Path, rel_path: str, data: bytes) -> Path:
    target = base_dir / rel_path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(data)
    return target


def test_missing_resource_returns_degradation(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    nodes = [ImageBlock(path="missing.png", alt="diagram")]

    result = preflight_resources(nodes, str(base_dir))

    assert result.resources == []
    assert len(result.degradations) == 1
    deg = result.degradations[0]
    assert isinstance(deg, ResourceDegradation)
    assert deg.code == RESOURCE_NOT_FOUND
    assert "missing.png" in deg.fallback_text
    assert isinstance(result.manifest, dict)
    assert result.manifest["entries"] == []


def test_unsupported_extension_returns_degradation(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    (base_dir / "bad.exe").write_bytes(b"MZ")
    nodes = [ImageBlock(path="bad.exe", alt="diagram")]

    result = preflight_resources(nodes, str(base_dir))

    assert result.resources == []
    assert len(result.degradations) == 1
    deg = result.degradations[0]
    assert deg.code == RESOURCE_MEDIA_TYPE_UNSUPPORTED
    assert "bad.exe" in deg.fallback_text


def test_valid_raster_produces_hash_and_resource(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    data = b"\x89PNG\r\n\x1a\nfake-png-body"
    _make_image(base_dir, "diagram.png", data)
    nodes = [ImageBlock(path="diagram.png", alt="diagram")]

    result = preflight_resources(nodes, str(base_dir))

    assert len(result.resources) == 1
    resource = result.resources[0]
    assert isinstance(resource, PreflightResource)
    assert resource.source_path == "diagram.png"
    assert resource.source_sha256 == _sha256(data)
    assert resource.payload_sha256 == _sha256(data)
    assert resource.byte_length == len(data)
    assert resource.media_type == "image/png"
    assert resource.normalizer_id == "none"
    assert result.degradations == []


def test_valid_jpeg_resource(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    data = b"\xff\xd8\xff\xe0fake-jpeg"
    _make_image(base_dir, "photo.jpg", data)
    nodes = [ImageBlock(path="photo.jpg", alt="photo")]

    result = preflight_resources(nodes, str(base_dir))

    assert len(result.resources) == 1
    assert result.resources[0].media_type == "image/jpeg"
    assert result.resources[0].source_sha256 == _sha256(data)


def test_svg_resource_manifest_binding(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    svg = '<svg xmlns="http://www.w3.org/2000/svg"><rect width="10" height="10"/></svg>'
    data = svg.encode("utf-8")
    _make_image(base_dir, "drawing.svg", data)
    nodes = [ImageBlock(path="drawing.svg", alt="drawing")]

    result = preflight_resources(nodes, str(base_dir))

    assert len(result.resources) == 1
    resource = result.resources[0]
    assert resource.media_type == "image/svg+xml"
    assert resource.source_sha256 == _sha256(data)
    assert resource.payload_sha256 == _sha256(data)
    assert result.degradations == []
    manifest = result.manifest
    assert manifest["version"] == "1"
    entry = manifest["entries"][0]
    assert entry["mediaType"] == "image/svg+xml"
    assert entry["normalizerId"] == "none"
    assert "source" not in entry
    assert "path" not in entry


def test_manifest_redacts_absolute_paths(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    data = b"\x89PNG\r\n\x1a\nsecret"
    _make_image(base_dir, "assets/secret.png", data)
    nodes = [ImageBlock(path="assets/secret.png", alt="secret")]

    result = preflight_resources(nodes, str(base_dir))

    manifest_json = str(result.manifest)
    assert str(base_dir) not in manifest_json
    assert "assets/secret.png" not in manifest_json
    assert result.resources[0].source_path == "assets/secret.png"


def test_figure_block_images_are_processed(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    a = b"\x89PNG\r\n\x1a\nA"
    b = b"\x89PNG\r\n\x1a\nB"
    _make_image(base_dir, "a.png", a)
    _make_image(base_dir, "b.png", b)
    figure = FigureBlock(
        identifier="fig:pair",
        caption="Pair",
        images=[
            ImageBlock(path="a.png", alt="first"),
            ImageBlock(path="b.png", alt="second"),
        ],
    )

    result = preflight_resources([figure], str(base_dir))

    assert len(result.resources) == 2
    hashes = {r.source_sha256 for r in result.resources}
    assert hashes == {_sha256(a), _sha256(b)}
    assert result.degradations == []


def test_path_escaping_base_dir_is_rejected(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    sibling = tmp_path / "secret.png"
    sibling.write_bytes(b"\x89PNG\r\n\x1a\n")
    nodes = [ImageBlock(path="../secret.png", alt="escape")]

    result = preflight_resources(nodes, str(base_dir))

    assert result.resources == []
    assert len(result.degradations) == 1
    assert result.degradations[0].code == RESOURCE_PATH_ESCAPES_BASE


def test_absolute_path_outside_base_dir_is_rejected(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    outside = tmp_path / "outside.png"
    outside.write_bytes(b"\x89PNG\r\n\x1a\n")
    nodes = [ImageBlock(path=str(outside), alt="outside")]

    result = preflight_resources(nodes, str(base_dir))

    assert result.resources == []
    assert len(result.degradations) == 1
    assert result.degradations[0].code == RESOURCE_ABSOLUTE_PATH_OUTSIDE


def test_symlink_escaping_base_dir_is_rejected(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    target = tmp_path / "real.png"
    target.write_bytes(b"\x89PNG\r\n\x1a\n")
    link = base_dir / "link.png"
    os.symlink(target, link)
    nodes = [ImageBlock(path="link.png", alt="link")]

    result = preflight_resources(nodes, str(base_dir))

    assert result.resources == []
    assert len(result.degradations) == 1
    assert result.degradations[0].code == RESOURCE_PATH_ESCAPES_BASE


def test_valid_symlink_inside_base_dir_is_accepted(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    target = base_dir / "real.png"
    target.write_bytes(b"\x89PNG\r\n\x1a\n")
    link = base_dir / "link.png"
    os.symlink(target, link)
    nodes = [ImageBlock(path="link.png", alt="link")]

    result = preflight_resources(nodes, str(base_dir))

    assert len(result.resources) == 1
    assert result.resources[0].media_type == "image/png"


def test_duplicate_references_yield_one_resource(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    data = b"\x89PNG\r\n\x1a\n"
    _make_image(base_dir, "x.png", data)
    nodes = [
        ImageBlock(path="x.png", alt="first"),
        ImageBlock(path="x.png", alt="second"),
    ]

    result = preflight_resources(nodes, str(base_dir))

    assert len(result.resources) == 1
    assert len(result.manifest["entries"]) == 1


def test_excalidraw_block_is_processed_as_svg_source(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    data = b'{"type":"excalidraw"}'
    _make_image(base_dir, "diag.excalidraw.md", data)
    nodes = [ExcalidrawBlock(path="diag.excalidraw.md", alt="diagram")]

    result = preflight_resources(nodes, str(base_dir))

    assert len(result.resources) == 1
    assert result.resources[0].media_type == "image/svg+xml"
    assert result.resources[0].source_sha256 == _sha256(data)
    assert result.degradations == []


def test_read_failure_is_degraded(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    target = base_dir / "unreadable.png"
    target.write_bytes(b"\x89PNG\r\n\x1a\n")
    target.chmod(0o000)
    try:
        nodes = [ImageBlock(path="unreadable.png", alt="diagram")]
        result = preflight_resources(nodes, str(base_dir))
        assert len(result.degradations) == 1
        assert result.degradations[0].code == RESOURCE_READ_FAILED
    finally:
        target.chmod(0o644)


def test_degradation_block_is_not_a_resource(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    nodes = [
        DegradationBlock(
            issue=DocumentIssue(code="TEST", message="test", placement="block"),
            fallback_text="nothing",
        )
    ]

    result = preflight_resources(nodes, str(base_dir))

    assert result.resources == []
    assert result.degradations == []


def test_empty_nodes_yield_empty_result(tmp_path: Path) -> None:
    result = preflight_resources([], str(tmp_path))
    assert result.resources == []
    assert result.degradations == []
    assert result.manifest["entries"] == []


def test_section_walks_into_elements(tmp_path: Path) -> None:
    base_dir = tmp_path / "project"
    base_dir.mkdir()
    data = b"\x89PNG\r\n\x1a\nZ"
    _make_image(base_dir, "z.png", data)
    section = Section(
        level=1,
        heading="Chapter",
        elements=[ImageBlock(path="z.png", alt="diagram")],
    )

    result = preflight_resources([section], str(base_dir))

    assert len(result.resources) == 1
    assert result.resources[0].source_sha256 == _sha256(data)
