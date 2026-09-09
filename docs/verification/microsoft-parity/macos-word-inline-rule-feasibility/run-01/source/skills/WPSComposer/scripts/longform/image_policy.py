"""Pure deterministic image sizing and quality policy."""

from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Any, Optional, Sequence


IMAGE_DPI_TOO_LOW = "IMAGE_DPI_TOO_LOW"
IMAGE_DPI_WARNING = "IMAGE_DPI_WARNING"
IMAGE_LABEL_TOO_SMALL = "IMAGE_LABEL_TOO_SMALL"


@dataclass(frozen=True)
class FigureImageLayout:
    resource_id: str
    display_width_pt: float
    display_height_pt: float
    effective_dpi: float
    media_type: str
    normalizer_id: str


@dataclass(frozen=True)
class ImageQualityIssue:
    code: str
    severity: str
    visible: bool
    resource_id: str


def resolve_figure_layout(
    figure: Any,
    resources: Sequence[Any],
    content_width_pt: float,
    column_gap_pt: float = 12.0,
) -> tuple[FigureImageLayout, ...]:
    if not math.isfinite(content_width_pt) or content_width_pt <= 0:
        raise ValueError("content_width_pt must be finite and positive")
    if not math.isfinite(column_gap_pt) or column_gap_pt < 0:
        raise ValueError("column_gap_pt must be finite and non-negative")

    image_paths = [str(getattr(image, "path", "")).replace("\\", "/") for image in figure.images]
    by_path = {
        str(getattr(resource, "source_path", "")).replace("\\", "/"): resource
        for resource in resources
    }
    ordered = [by_path[path] for path in image_paths if path in by_path]
    if not ordered and len(resources) == len(image_paths):
        ordered = list(resources)

    is_columns = getattr(figure, "layout", "stack") in {"columns", "side-by-side"}
    if is_columns:
        if len(ordered) != 2:
            raise ValueError("column figure layout requires exactly two resources")
        slot_width = (content_width_pt - column_gap_pt) / 2.0
        if slot_width <= 0:
            raise ValueError("column gap leaves no image slot")
    else:
        slot_width = content_width_pt

    requested_width = str(getattr(figure, "width", "auto") or "auto").strip().lower()
    layouts = []
    for resource in ordered:
        profile = resource.image_profile
        pixel_width = float(profile.pixel_width)
        pixel_height = float(profile.pixel_height)
        if pixel_width <= 0 or pixel_height <= 0:
            raise ValueError("resource image dimensions must be positive")

        embedded_dpi = _credible_dpi(profile.dpi_x, profile.dpi_y)
        natural_dpi = embedded_dpi or 96.0
        natural_width = pixel_width * 72.0 / natural_dpi
        if requested_width == "auto":
            desired_width = natural_width
        elif requested_width == "column":
            desired_width = (content_width_pt - column_gap_pt) / 2.0
        elif requested_width == "full":
            desired_width = content_width_pt
        else:
            match = re.fullmatch(r"([0-9]+(?:\.[0-9]+)?)pt", requested_width)
            if match is None:
                raise ValueError(f"unsupported figure width: {requested_width}")
            desired_width = float(match.group(1))
        if not math.isfinite(desired_width) or desired_width <= 0:
            raise ValueError("figure width must be finite and positive")
        display_width = min(desired_width, slot_width)
        display_height = display_width * pixel_height / pixel_width
        if resource.media_type == "image/svg+xml":
            effective_dpi = 0.0
        else:
            effective_dpi = min(
                pixel_width * 72.0 / display_width,
                pixel_height * 72.0 / display_height,
            )
        layouts.append(
            FigureImageLayout(
                resource_id=resource.resource_id,
                display_width_pt=display_width,
                display_height_pt=display_height,
                effective_dpi=effective_dpi,
                media_type=resource.media_type,
                normalizer_id=resource.normalizer_id,
            )
        )
    return tuple(layouts)


def classify_image_quality(
    layout: FigureImageLayout,
    kind: str,
    *,
    label_size_pt: Optional[float] = None,
    label_metadata_reliable: bool = False,
) -> tuple[ImageQualityIssue, ...]:
    issues = []
    if layout.media_type != "image/svg+xml" and layout.effective_dpi < 96.0:
        issues.append(
            ImageQualityIssue(
                IMAGE_DPI_TOO_LOW, "degradation", True, layout.resource_id
            )
        )
    normalized_kind = (kind or "auto").lower()
    if (
        layout.media_type != "image/svg+xml"
        and layout.effective_dpi >= 96.0
        and layout.effective_dpi < 150.0
        and normalized_kind in {"auto", "photo", "scan"}
    ):
        issues.append(
            ImageQualityIssue(
                IMAGE_DPI_WARNING, "warning", False, layout.resource_id
            )
        )
    if (
        label_metadata_reliable
        and label_size_pt is not None
        and math.isfinite(label_size_pt)
        and label_size_pt < 8.0
    ):
        issues.append(
            ImageQualityIssue(
                IMAGE_LABEL_TOO_SMALL, "degradation", True, layout.resource_id
            )
        )
    return tuple(issues)


def _credible_dpi(dpi_x: Optional[float], dpi_y: Optional[float]) -> Optional[float]:
    values = []
    for raw in (dpi_x, dpi_y):
        if raw is None:
            return None
        try:
            value = float(raw)
        except (TypeError, ValueError):
            return None
        if not math.isfinite(value) or not 36.0 <= value <= 1200.0:
            return None
        values.append(value)
    if max(values) / min(values) > 1.05:
        return None
    return min(values)


__all__ = [
    "IMAGE_DPI_TOO_LOW",
    "IMAGE_DPI_WARNING",
    "IMAGE_LABEL_TOO_SMALL",
    "FigureImageLayout",
    "ImageQualityIssue",
    "classify_image_quality",
    "resolve_figure_layout",
]
