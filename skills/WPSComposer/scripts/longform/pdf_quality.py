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
from .executor import PaginationMap, PaginationNode
from .quality import (
    PageRole,
    QualityConfidence,
    QualityEvidence,
    QualityFinding,
    QualityReport,
    QualitySeverity,
    normalize_bounds,
)


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


@dataclass(frozen=True)
class QualityPolicy:
    """Closed inputs for deterministic quality checks.

    Node metadata comes from the semantic/execution result, never PDF text.
    Tuples keep the policy immutable and serialization-order independent.
    """

    body_margin_points: float = 54.0
    node_kinds: Tuple[Tuple[str, str], ...] = ()
    caption_targets: Tuple[Tuple[str, str], ...] = ()
    effective_dpi: Tuple[Tuple[str, float], ...] = ()
    minimum_dpi: float = 150.0
    table_header_repeated: Tuple[Tuple[str, bool], ...] = ()
    forced_split_nodes: Tuple[str, ...] = ()
    unresolved_field_nodes: Tuple[str, ...] = ()
    header_display_units: Optional[int] = None
    toc_last_page_utilization: Optional[float] = None
    bibliography_spacing_ratio: Optional[float] = None
    final_page_utilization: Optional[float] = None
    page_number_sequence_valid: Optional[bool] = None

    def __post_init__(self) -> None:
        numeric = (
            self.body_margin_points,
            self.minimum_dpi,
        )
        if any(not math.isfinite(float(value)) or float(value) <= 0 for value in numeric):
            raise ValueError("quality policy dimensions must be positive and finite")


def _role(value: Any) -> PageRole:
    return value if isinstance(value, PageRole) else PageRole(str(value))


def _node_index(pagination_map: PaginationMap) -> dict[str, PaginationNode]:
    return {node.node_id: node for node in pagination_map.nodes if node.node_id}


def _first_fragment(node: Optional[PaginationNode]):
    if node is None or not node.fragments:
        return None
    return node.fragments[0]


def _finding(
    code: str,
    severity: QualitySeverity,
    confidence: QualityConfidence,
    message: str,
    *,
    node: Optional[PaginationNode] = None,
    page: Optional[int] = None,
    bounds: Optional[Tuple[float, float, float, float]] = None,
    evidence: Optional[dict[str, Any]] = None,
    repair_key: Optional[str] = None,
) -> QualityFinding:
    fragment = _first_fragment(node)
    if page is None and fragment is not None:
        page = fragment.page
    if bounds is None and fragment is not None:
        bounds = fragment.bounds
    return QualityFinding(
        code=code,
        severity=severity,
        confidence=confidence,
        message=message,
        node_id=None if node is None else node.node_id,
        page=page,
        bounds=bounds,
        evidence=QualityEvidence.from_mapping(evidence),
        repair_key=repair_key,
    )


def analyze_pages(
    pages: Tuple[PdfPage, ...],
    pagination_map: PaginationMap,
    page_roles: Mapping[int, PageRole],
    policy: QualityPolicy,
) -> QualityReport:
    """Run the closed role-aware quality rules on normalized page geometry."""

    findings: list[QualityFinding] = []
    nodes = _node_index(pagination_map)
    kinds = dict(policy.node_kinds)
    roles = {int(page): _role(role) for page, role in page_roles.items()}
    pages_by_number = {page.physical_page: page for page in pages}
    exempt_sparse = {
        PageRole.COVER,
        PageRole.CHAPTER_START,
        PageRole.EXPLICIT_BREAK,
    }

    # A truly empty ordinary page is deterministic. Explicit sparse roles are
    # preserved because they may be intentional semantic page boundaries.
    for page in pages:
        role = roles.get(page.physical_page, PageRole.BODY)
        if not page.glyph_bounds and not page.object_bounds and role not in exempt_sparse:
            findings.append(
                _finding(
                    "UNEXPECTED_BLANK_PAGE",
                    QualitySeverity.DEGRADED,
                    QualityConfidence.HIGH,
                    "An ordinary document page contains no visible content",
                    page=page.physical_page,
                    evidence={
                        "blankCharacterCount": 0,
                        "pageRole": role.value,
                    },
                    repair_key="remove-unexpected-blank",
                )
            )

    # Visual object fragments have authoritative WPS-to-PDF mappings. Text
    # fragments without bounds are intentionally excluded from bbox repairs.
    margin = float(policy.body_margin_points)
    for node in pagination_map.nodes:
        kind = kinds.get(node.node_id)
        if kind not in {"image", "table", "formula", "visual"}:
            continue
        for fragment in node.fragments:
            bounds = fragment.bounds
            page = pages_by_number.get(fragment.page)
            if bounds is None or page is None:
                continue
            left, top, right, bottom = bounds
            overflow = max(
                margin - left,
                margin - top,
                right - (page.width - margin),
                bottom - (page.height - margin),
                0.0,
            )
            if overflow <= 0:
                continue
            if kind == "image":
                code, repair = "IMAGE_TOO_WIDE", "fit-image"
            elif kind == "table":
                code, repair = "TABLE_TOO_WIDE", "compress-table"
            else:
                code, repair = "CONTENT_OVERFLOW", None
            findings.append(
                _finding(
                    code,
                    QualitySeverity.DEGRADED,
                    QualityConfidence.HIGH,
                    "Mapped visual content crosses the document body boundary",
                    node=node,
                    page=fragment.page,
                    bounds=bounds,
                    evidence={
                        "overflowPoints": overflow,
                        "pageRole": roles.get(fragment.page, PageRole.BODY).value,
                    },
                    repair_key=repair,
                )
            )

    # A mapped heading with less than two normal text lines below it is an
    # orphan. This never relies on extracted heading text.
    line_height = 12.0
    for node_id, kind in kinds.items():
        if kind != "heading":
            continue
        node = nodes.get(node_id)
        fragment = _first_fragment(node)
        page = pages_by_number.get(fragment.page) if fragment else None
        if fragment is None or fragment.bounds is None or page is None:
            continue
        if fragment.bounds[3] > page.height - margin - (2 * line_height):
            findings.append(
                _finding(
                    "HEADING_ORPHAN",
                    QualitySeverity.DEGRADED,
                    QualityConfidence.HIGH,
                    "A heading is left at the page end without two following lines",
                    node=node,
                    repair_key="keep-heading",
                )
            )

    for caption_id, target_id in policy.caption_targets:
        caption = nodes.get(caption_id)
        target = nodes.get(target_id)
        caption_fragment = _first_fragment(caption)
        target_fragment = _first_fragment(target)
        if (
            caption_fragment is not None
            and target_fragment is not None
            and caption_fragment.page != target_fragment.page
        ):
            findings.append(
                _finding(
                    "CAPTION_SEPARATED",
                    QualitySeverity.DEGRADED,
                    QualityConfidence.HIGH,
                    "An object and its caption are on different physical pages",
                    node=caption,
                    evidence={
                        "captionPage": caption_fragment.page,
                        "objectPage": target_fragment.page,
                    },
                    repair_key="keep-caption",
                )
            )

    for node_id, dpi in policy.effective_dpi:
        if float(dpi) >= float(policy.minimum_dpi):
            continue
        findings.append(
            _finding(
                "IMAGE_LOW_DPI",
                QualitySeverity.DEGRADED,
                QualityConfidence.HIGH,
                "An image is below the configured effective DPI threshold",
                node=nodes.get(node_id),
                evidence={"actualDpi": dpi, "minimumDpi": policy.minimum_dpi},
                repair_key="low-dpi-notice",
            )
        )

    forced = set(policy.forced_split_nodes)
    for node_id, repeated in policy.table_header_repeated:
        if repeated or node_id in forced:
            continue
        findings.append(
            _finding(
                "TABLE_HEADER_NOT_REPEATED",
                QualitySeverity.DEGRADED,
                QualityConfidence.HIGH,
                "A cross-page table does not repeat its header row",
                node=nodes.get(node_id),
            )
        )

    for node_id in policy.unresolved_field_nodes:
        findings.append(
            _finding(
                "FIELD_UNRESOLVED",
                QualitySeverity.DEGRADED,
                QualityConfidence.HIGH,
                "A required native field remains unresolved after final refresh",
                node=nodes.get(node_id),
                evidence={"fieldCount": 1},
            )
        )

    if policy.header_display_units is not None and policy.header_display_units > 32:
        findings.append(
            _finding(
                "HEADER_OVERFLOW",
                QualitySeverity.DEGRADED,
                QualityConfidence.HIGH,
                "The document header exceeds the final single-line display limit",
                page=1 if pages else None,
                evidence={"headerDisplayUnits": policy.header_display_units},
                repair_key="shorten-header",
            )
        )

    if (
        policy.toc_last_page_utilization is not None
        and policy.toc_last_page_utilization < 0.25
    ):
        toc_pages = [number for number, role in roles.items() if role is PageRole.TOC]
        findings.append(
            _finding(
                "TOC_SPARSE_OVERFLOW",
                QualitySeverity.DEGRADED,
                QualityConfidence.HIGH,
                "The table of contents spills onto a sparsely used final TOC page",
                page=max(toc_pages) if toc_pages else (1 if pages else None),
                evidence={"utilization": policy.toc_last_page_utilization},
                repair_key="compact-toc",
            )
        )

    if (
        policy.bibliography_spacing_ratio is not None
        and policy.bibliography_spacing_ratio > 1.5
    ):
        bibliography_pages = [
            number for number, role in roles.items() if role is PageRole.BIBLIOGRAPHY
        ]
        findings.append(
            _finding(
                "BIBLIOGRAPHY_SPACING",
                QualitySeverity.WARNING,
                QualityConfidence.MEDIUM,
                "Bibliography character spacing appears unusually wide",
                page=min(bibliography_pages) if bibliography_pages else (1 if pages else None),
                evidence={"spacingRatio": policy.bibliography_spacing_ratio},
            )
        )

    if policy.final_page_utilization is not None and policy.final_page_utilization < 0.18:
        final_page = pages[-1].physical_page if pages else None
        role = roles.get(final_page, PageRole.BODY) if final_page is not None else PageRole.BODY
        if role not in exempt_sparse:
            findings.append(
                _finding(
                    "LAST_PAGE_SPARSE",
                    QualitySeverity.WARNING,
                    QualityConfidence.MEDIUM,
                    "The final ordinary page has low content utilization",
                    page=final_page,
                    evidence={
                        "finalPageUtilization": policy.final_page_utilization,
                        "pageRole": role.value,
                    },
                )
            )

    if policy.page_number_sequence_valid is False:
        findings.append(
            _finding(
                "PAGE_NUMBER_SEQUENCE_INVALID",
                QualitySeverity.DEGRADED,
                QualityConfidence.HIGH,
                "Native page-number fields do not follow the expected section sequence",
                page=1 if pages else None,
                evidence={"totalPages": len(pages)},
            )
        )

    return QualityReport(findings=tuple(findings), page_count=len(pages))


def analyze_pdf(
    pdf_path: Path,
    pagination_map: PaginationMap,
    page_roles: Mapping[int, PageRole],
    policy: QualityPolicy,
) -> QualityReport:
    """Validate, normalize, and analyze one WPS-exported PDF."""

    return analyze_pages(load_pdf_pages(pdf_path), pagination_map, page_roles, policy)


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
    "QualityPolicy",
    "QualityDependencyError",
    "analyze_pages",
    "analyze_pdf",
    "load_pdf_pages",
    "require_quality_dependencies",
]
