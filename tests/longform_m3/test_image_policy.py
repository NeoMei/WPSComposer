from __future__ import annotations

import math

import pytest

from skills.WPSComposer.scripts.document_model import FigureBlock, ImageBlock
from skills.WPSComposer.scripts.longform.image_policy import (
    IMAGE_DPI_TOO_LOW,
    IMAGE_DPI_WARNING,
    IMAGE_LABEL_TOO_SMALL,
    FigureImageLayout,
    classify_image_quality,
    resolve_figure_layout,
)
from skills.WPSComposer.scripts.longform.policy import page_content_width_pt
from skills.WPSComposer.scripts.longform.resources import ImageProfile, PreflightResource


def _resource(
    resource_id: str,
    path: str,
    width: int,
    height: int,
    *,
    dpi_x: float | None = None,
    dpi_y: float | None = None,
    media_type: str = "image/png",
) -> PreflightResource:
    return PreflightResource(
        resource_id=resource_id,
        source_path=path,
        source_sha256="a" * 64,
        payload_sha256="a" * 64,
        byte_length=4,
        media_type=media_type,
        normalizer_id="none-v1" if media_type != "image/svg+xml" else "svg-static-v1",
        image_profile=ImageProfile(width, height, dpi_x, dpi_y, 1, "PNG", False),
        payload_bytes=b"test",
    )


def _figure(width: str = "auto", *, layout: str = "stack", orientation: str = "portrait") -> FigureBlock:
    return FigureBlock(
        images=[ImageBlock(path="a.png", alt="a")],
        width=width,
        layout=layout,
        orientation=orientation,
        kind="photo",
    )


def test_auto_uses_credible_embedded_dpi_and_preserves_aspect_ratio() -> None:
    resource = _resource("r1", "a.png", 1200, 600, dpi_x=300, dpi_y=300)

    (layout,) = resolve_figure_layout(_figure(), [resource], 500)

    assert layout == FigureImageLayout("r1", 288.0, 144.0, 300.0, "image/png", "none-v1")


@pytest.mark.parametrize("dpi", [None, 10.0, 5000.0, math.nan])
def test_auto_falls_back_to_96_dpi_when_embedded_dpi_is_not_credible(dpi: float | None) -> None:
    resource = _resource("r1", "a.png", 400, 200, dpi_x=dpi, dpi_y=dpi)

    (layout,) = resolve_figure_layout(_figure(), [resource], 500)

    assert layout.display_width_pt == pytest.approx(300.0)
    assert layout.display_height_pt == pytest.approx(150.0)
    assert layout.effective_dpi == pytest.approx(96.0)


@pytest.mark.parametrize(
    ("width", "expected"),
    [
        ("column", 244.0),
        ("full", 500.0),
        ("360pt", 360.0),
        ("900pt", 500.0),
    ],
)
def test_width_modes_are_capped_by_the_available_content_slot(width: str, expected: float) -> None:
    resource = _resource("r1", "a.png", 2000, 1000, dpi_x=300, dpi_y=300)

    (layout,) = resolve_figure_layout(_figure(width), [resource], 500)

    assert layout.display_width_pt == pytest.approx(expected)
    assert layout.display_height_pt == pytest.approx(expected / 2)


def test_two_column_layout_uses_fixed_12_point_gap() -> None:
    figure = FigureBlock(
        images=[ImageBlock(path="a.png"), ImageBlock(path="b.png")],
        width="full",
        layout="columns",
        columns=2,
    )
    resources = [
        _resource("r1", "a.png", 100, 100),
        _resource("r2", "b.png", 200, 100),
    ]

    layouts = resolve_figure_layout(figure, resources, 500)

    assert [item.display_width_pt for item in layouts] == [244.0, 244.0]
    assert [item.display_height_pt for item in layouts] == [244.0, 122.0]
    assert sum(item.display_width_pt for item in layouts) + 12.0 == 500.0


def test_portrait_and_explicit_landscape_use_distinct_policy_slot_widths() -> None:
    portrait = page_content_width_pt("portrait")
    landscape = page_content_width_pt("landscape")
    resource = _resource("r1", "a.png", 4000, 2000, dpi_x=300, dpi_y=300)

    portrait_layout = resolve_figure_layout(_figure("full"), [resource], portrait)[0]
    landscape_layout = resolve_figure_layout(
        _figure("full", orientation="landscape"), [resource], landscape
    )[0]

    assert landscape > portrait
    assert portrait_layout.display_width_pt == pytest.approx(portrait)
    assert landscape_layout.display_width_pt == pytest.approx(landscape)
    assert portrait_layout.display_height_pt * 2 == pytest.approx(portrait_layout.display_width_pt)
    assert landscape_layout.display_height_pt * 2 == pytest.approx(landscape_layout.display_width_pt)


def test_extreme_image_ratio_never_changes_orientation_or_available_width() -> None:
    figure = _figure("full", orientation="portrait")
    resource = _resource("r1", "a.png", 4000, 100)

    (layout,) = resolve_figure_layout(figure, [resource], 440)

    assert figure.orientation == "portrait"
    assert layout.display_width_pt == 440
    assert layout.display_height_pt == 11


@pytest.mark.parametrize("kind", ["photo", "scan"])
def test_photo_and_scan_below_150_dpi_receive_nonvisible_warning(kind: str) -> None:
    layout = FigureImageLayout("r", 300, 150, 120, "image/png", "none-v1")

    issues = classify_image_quality(layout, kind)

    assert [(item.code, item.severity, item.visible) for item in issues] == [
        (IMAGE_DPI_WARNING, "warning", False)
    ]


@pytest.mark.parametrize("kind", ["auto", "photo", "scan", "screenshot", "diagram"])
def test_any_raster_below_96_dpi_is_a_visible_block_degradation(kind: str) -> None:
    layout = FigureImageLayout("r", 300, 150, 95.9, "image/png", "none-v1")

    issues = classify_image_quality(layout, kind)

    assert [(item.code, item.severity, item.visible) for item in issues] == [
        (IMAGE_DPI_TOO_LOW, "degradation", True)
    ]


@pytest.mark.parametrize("kind", ["screenshot", "diagram"])
def test_screenshot_and_diagram_between_96_and_150_have_no_dpi_only_marker(kind: str) -> None:
    layout = FigureImageLayout("r", 300, 150, 120, "image/png", "none-v1")

    assert classify_image_quality(layout, kind) == ()


def test_label_size_check_runs_only_with_reliable_metadata() -> None:
    layout = FigureImageLayout("r", 300, 150, 200, "image/png", "none-v1")

    assert classify_image_quality(layout, "diagram", label_size_pt=7, label_metadata_reliable=False) == ()
    issues = classify_image_quality(
        layout, "diagram", label_size_pt=7, label_metadata_reliable=True
    )
    assert [(item.code, item.severity, item.visible) for item in issues] == [
        (IMAGE_LABEL_TOO_SMALL, "degradation", True)
    ]


def test_low_dpi_and_reliable_small_label_are_reported_independently() -> None:
    layout = FigureImageLayout("r", 300, 150, 95.9, "image/png", "none-v1")

    issues = classify_image_quality(
        layout, "diagram", label_size_pt=7, label_metadata_reliable=True
    )

    assert [(item.code, item.severity, item.visible) for item in issues] == [
        (IMAGE_DPI_TOO_LOW, "degradation", True),
        (IMAGE_LABEL_TOO_SMALL, "degradation", True),
    ]
