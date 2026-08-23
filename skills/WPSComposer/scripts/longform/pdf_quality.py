"""Core PDF loading primitives for the long-form M5 quality gate.

This module never returns extracted customer text.  It exposes only normalized
page geometry needed by deterministic layout checks.
"""

from __future__ import annotations

from dataclasses import dataclass
from importlib import import_module, metadata
import math
from pathlib import Path
from typing import Any, Callable, Iterable, Optional, Tuple

try:  # Kept importable so the explicit dependency gate can report all gaps.
    import pdfplumber
except ImportError:  # pragma: no cover - exercised through injected gate tests
    pdfplumber = None  # type: ignore[assignment]

try:
    from pypdf import PdfReader
except ImportError:  # pragma: no cover - exercised through injected gate tests
    PdfReader = None  # type: ignore[assignment,misc]

from ..artifact_transport import validate_pdf
from .quality import normalize_bounds


_DEPENDENCIES = (
    ("PIL", "Pillow", (10, 0)),
    ("pypdf", "pypdf", (4, 0)),
    ("pdfplumber", "pdfplumber", (0, 11)),
)


class QualityDependencyError(RuntimeError):
    """Raised before WPS starts when a core quality dependency is unusable."""

    def __init__(self, missing_modules: Iterable[str]) -> None:
        self.missing_modules = tuple(sorted(set(str(item) for item in missing_modules)))
        super().__init__("Long-form quality dependencies unavailable: " + ", ".join(self.missing_modules))


def _version_tuple(value: str) -> tuple[int, ...]:
    result = []
    for part in str(value).split("."):
        digits = "".join(character for character in part if character.isdigit())
        if not digits:
            break
        result.append(int(digits))
    return tuple(result)


def require_quality_dependencies(
    *,
    importer: Callable[[str], Any] = import_module,
    version_getter: Callable[[str], str] = metadata.version,
) -> tuple[str, ...]:
    """Validate all core analyzers before any native WPS mutation."""

    failures = []
    versions = []
    for module_name, package_name, minimum in _DEPENDENCIES:
        try:
            importer(module_name)
            version = str(version_getter(package_name))
        except Exception:
            failures.append(module_name)
            continue
        if _version_tuple(version) < minimum:
            failures.append(f"{package_name}>={'.'.join(map(str, minimum))}")
            continue
        versions.append(f"{package_name} {version}")
    if failures:
        raise QualityDependencyError(failures)
    return tuple(versions)


@dataclass(frozen=True)
class PdfPage:
    """One physical PDF page in normalized top-left point coordinates."""

    physical_page: int
    width: float
    height: float
    rotation: int
    media_box: Tuple[float, float, float, float]
    crop_box: Tuple[float, float, float, float]
    glyph_bounds: Tuple[Tuple[float, float, float, float], ...] = ()
    object_bounds: Tuple[Tuple[float, float, float, float], ...] = ()


def _box(box: Any, label: str) -> Tuple[float, float, float, float]:
    try:
        values = tuple(float(item) for item in box)
    except Exception as exc:
        raise ValueError(f"PDF {label} is invalid") from exc
    if len(values) != 4 or not all(math.isfinite(item) for item in values):
        raise ValueError(f"PDF {label} is invalid")
    x0, y0, x1, y1 = values
    if x1 <= x0 or y1 <= y0:
        raise ValueError(f"PDF {label} is invalid")
    return values  # type: ignore[return-value]


def _rotation(value: Any) -> int:
    try:
        number = float(value or 0)
    except (TypeError, ValueError) as exc:
        raise ValueError("PDF rotation is invalid") from exc
    if not math.isfinite(number) or not number.is_integer():
        raise ValueError("PDF rotation is invalid")
    result = int(number) % 360
    if result not in {0, 90, 180, 270}:
        raise ValueError("PDF rotation is invalid")
    return result


def _plumber_bounds(items: Iterable[dict[str, Any]]) -> tuple[Tuple[float, float, float, float], ...]:
    result = []
    for item in items:
        try:
            raw = (item["x0"], item["top"], item["x1"], item["bottom"])
            bounds = normalize_bounds(raw)
        except (KeyError, TypeError, ValueError):
            continue
        if bounds is not None:
            result.append(bounds)
    return tuple(result)


def load_pdf_pages(path: Path) -> tuple[PdfPage, ...]:
    """Load validated page geometry without retaining extracted PDF text."""

    require_quality_dependencies()
    target = Path(path).expanduser().resolve()
    try:
        validate_pdf(target)
        if PdfReader is None or pdfplumber is None:
            raise QualityDependencyError(("pypdf", "pdfplumber"))
        reader = PdfReader(str(target), strict=True)
        with pdfplumber.open(str(target)) as plumber_pdf:
            if len(reader.pages) != len(plumber_pdf.pages):
                raise ValueError("PDF parsers disagree on page count")
            pages = []
            for index, (page, plumber_page) in enumerate(
                zip(reader.pages, plumber_pdf.pages), start=1
            ):
                media_box = _box(page.mediabox, "MediaBox")
                crop_box = _box(page.cropbox, "CropBox")
                rotation = _rotation(page.get("/Rotate", 0))
                raw_width = crop_box[2] - crop_box[0]
                raw_height = crop_box[3] - crop_box[1]
                width, height = (
                    (raw_height, raw_width)
                    if rotation in {90, 270}
                    else (raw_width, raw_height)
                )
                objects = []
                for collection in (
                    getattr(plumber_page, "images", ()),
                    getattr(plumber_page, "rects", ()),
                    getattr(plumber_page, "curves", ()),
                ):
                    objects.extend(collection or ())
                pages.append(
                    PdfPage(
                        physical_page=index,
                        width=float(width),
                        height=float(height),
                        rotation=rotation,
                        media_box=media_box,
                        crop_box=crop_box,
                        glyph_bounds=_plumber_bounds(plumber_page.chars),
                        object_bounds=_plumber_bounds(objects),
                    )
                )
        return tuple(pages)
    except QualityDependencyError:
        raise
    except ValueError as exc:
        if str(exc) == "PDF parsers disagree on page count":
            raise
        raise ValueError("PDF quality input is invalid") from None
    except Exception:
        raise ValueError("PDF quality input is invalid") from None


__all__ = [
    "PdfPage",
    "QualityDependencyError",
    "load_pdf_pages",
    "require_quality_dependencies",
]
