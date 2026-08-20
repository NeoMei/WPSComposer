"""Host-side dependency, artifact, PDF, and evidence privacy checks."""

from __future__ import annotations

import hashlib
import importlib.util
import math
import os
from pathlib import Path
import re
from typing import Any, Callable, Iterable, Mapping, TypeVar

try:
    import pdfplumber
except ImportError:  # kept importable so the dependency gate can report cleanly
    pdfplumber = None  # type: ignore[assignment]
try:
    from pypdf import PdfReader
except ImportError:  # kept importable so the dependency gate can report cleanly
    PdfReader = None  # type: ignore[assignment,misc]

from ..artifact_transport import validate_office_package, validate_pdf

_DEPENDENCIES = (
    ("pypdf", "pypdf"),
    ("pdfplumber", "pdfplumber"),
    ("PIL", "Pillow"),
)
_PRIVATE_KEYS = frozenset(
    {
        "text",
        "bookmarkmap",
        "fieldhash",
        "sourcepath",
        "stagingpath",
        "commandline",
    }
)
_USER_PATH = re.compile(r"(?:^|[\s\"'])/Users/")
_WINDOWS_USER_PATH = re.compile(r"(?:^|[\s\"'])[A-Za-z]:[\\/]Users[\\/]")
_T = TypeVar("_T")


def _find_missing_dependencies() -> list[str]:
    return [
        display_name
        for module_name, display_name in _DEPENDENCIES
        if importlib.util.find_spec(module_name) is None
    ]


def require_probe_dependencies() -> tuple[str, ...]:
    """Fail before native WPS startup if an M0 analysis dependency is absent."""
    missing = _find_missing_dependencies()
    if missing:
        raise RuntimeError(
            "Long-form M0 dependencies are missing: " + ", ".join(missing)
        )
    return tuple(display_name for _, display_name in _DEPENDENCIES)


def run_after_dependency_gate(executor: Callable[[], _T]) -> _T:
    """Run an executor only after the complete analysis dependency gate."""
    require_probe_dependencies()
    return executor()


def prepare_evidence_directory(path: Path) -> Path:
    """Create or accept one empty private evidence directory."""
    requested = Path(path).expanduser()
    if requested.is_symlink():
        raise ValueError("Evidence directory must not be a symbolic link")
    target = requested.resolve()
    if target.exists() and not target.is_dir():
        raise FileExistsError(f"Evidence path is not a directory: {target}")
    if target.exists() and any(target.iterdir()):
        raise FileExistsError(f"Evidence directory is not empty: {target}")
    target.mkdir(parents=True, exist_ok=True, mode=0o700)
    os.chmod(target, 0o700)
    return target


def sha256_file(path: Path) -> str:
    """Hash one regular file without loading it fully into memory."""
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        while True:
            block = stream.read(1024 * 1024)
            if not block:
                return digest.hexdigest()
            digest.update(block)


def _resolve(value: Any) -> Any:
    getter = getattr(value, "get_object", None)
    return getter() if callable(getter) else value


def _box_values(box: Any, label: str) -> list[float]:
    values = [float(value) for value in box]
    if len(values) != 4 or any(not math.isfinite(value) for value in values):
        raise ValueError(f"PDF {label} must contain four finite point values")
    return values


def _font_names(page: Any) -> list[str]:
    resources = _resolve(page.get("/Resources"))
    if not resources:
        return []
    fonts = _resolve(resources.get("/Font"))
    if not fonts:
        return []
    names = set()
    for font_reference in fonts.values():
        font = _resolve(font_reference)
        if not font:
            continue
        base_font = font.get("/BaseFont")
        if base_font:
            names.add(str(base_font).lstrip("/"))
        descendants = _resolve(font.get("/DescendantFonts")) or []
        for descendant_reference in descendants:
            descendant = _resolve(descendant_reference)
            descendant_base = descendant.get("/BaseFont") if descendant else None
            if descendant_base:
                names.add(str(descendant_base).lstrip("/"))
    return sorted(names)


def _character_bounds(chars: Iterable[dict[str, Any]]) -> Any:
    characters = list(chars)
    if not characters:
        return None
    numeric = []
    for character in characters:
        values = tuple(
            float(character[name]) for name in ("x0", "top", "x1", "bottom")
        )
        if any(not math.isfinite(value) for value in values):
            raise ValueError("PDF character bounds must be finite")
        numeric.append(values)
    return {
        "count": len(numeric),
        "bbox": [
            min(item[0] for item in numeric),
            min(item[1] for item in numeric),
            max(item[2] for item in numeric),
            max(item[3] for item in numeric),
        ],
    }


def snapshot_pdf(path: Path) -> dict[str, Any]:
    """Return layout evidence without returning any extracted document text."""
    require_probe_dependencies()
    global pdfplumber, PdfReader
    if pdfplumber is None:
        import pdfplumber as loaded_pdfplumber

        pdfplumber = loaded_pdfplumber
    if PdfReader is None:
        from pypdf import PdfReader as loaded_pdf_reader

        PdfReader = loaded_pdf_reader
    target = Path(path).expanduser().resolve()
    validate_pdf(target)
    reader = PdfReader(str(target), strict=True)
    with pdfplumber.open(str(target)) as plumber_pdf:
        if len(reader.pages) != len(plumber_pdf.pages):
            raise ValueError("PDF parsers disagree on page count")
        pages = []
        all_fonts = set()
        for index, (page, plumber_page) in enumerate(
            zip(reader.pages, plumber_pdf.pages), start=1
        ):
            rotation_value = page.get("/Rotate", 0) or 0
            rotation_number = float(rotation_value)
            if (
                not math.isfinite(rotation_number)
                or not rotation_number.is_integer()
            ):
                raise ValueError("PDF rotation must be a finite integer")
            rotation = int(rotation_number) % 360
            if rotation not in {0, 90, 180, 270}:
                raise ValueError("PDF rotation must be a multiple of 90 degrees")
            fonts = _font_names(page)
            all_fonts.update(fonts)
            pages.append(
                {
                    "physicalPage": index,
                    "mediaBox": _box_values(page.mediabox, "MediaBox"),
                    "cropBox": _box_values(page.cropbox, "CropBox"),
                    "rotation": rotation,
                    "fonts": fonts,
                    "characterBounds": _character_bounds(plumber_page.chars),
                }
            )
    return {
        "pageCount": len(pages),
        "pages": pages,
        "fonts": sorted(all_fonts),
    }


def snapshot_pdf_markers(
    path: Path, markers: Mapping[str, str]
) -> dict[str, Any]:
    """Locate fixed probe markers without returning their extracted text."""
    require_probe_dependencies()
    if not isinstance(markers, Mapping) or any(
        not isinstance(label, str)
        or not label
        or not isinstance(marker, str)
        or not marker
        for label, marker in markers.items()
    ):
        raise ValueError("PDF marker requests must be non-empty strings")
    target = Path(path).expanduser().resolve()
    validate_pdf(target)
    found: dict[str, Any] = {label: None for label in markers}
    with pdfplumber.open(str(target)) as plumber_pdf:
        for page_number, page in enumerate(plumber_pdf.pages, start=1):
            characters = list(page.chars)
            flattened = "".join(
                str(character.get("text", "")) for character in characters
            )
            for label, marker in markers.items():
                if found[label] is not None:
                    continue
                offset = flattened.find(marker)
                if offset < 0:
                    continue
                # WPS/PDFPlumber emits one entry per Unicode scalar for the
                # fixed ASCII markers used by this probe.
                hit = characters[offset : offset + len(marker)]
                if len(hit) != len(marker):
                    continue
                values = [
                    tuple(
                        float(character[name])
                        for name in ("x0", "top", "x1", "bottom")
                    )
                    for character in hit
                ]
                if any(
                    not math.isfinite(value)
                    for coordinates in values
                    for value in coordinates
                ):
                    raise ValueError("PDF marker bounds must be finite")
                found[label] = {
                    "physicalPage": page_number,
                    "bbox": [
                        min(value[0] for value in values),
                        min(value[1] for value in values),
                        max(value[2] for value in values),
                        max(value[3] for value in values),
                    ],
                }
                marker_box = found[label]["bbox"]
                containing_frames = []
                for rectangle in getattr(page, "rects", ()):
                    frame = [
                        float(rectangle[name])
                        for name in ("x0", "top", "x1", "bottom")
                    ]
                    if (
                        frame[0] <= marker_box[0]
                        and frame[1] <= marker_box[1]
                        and frame[2] >= marker_box[2]
                        and frame[3] >= marker_box[3]
                    ):
                        containing_frames.append(frame)
                if containing_frames:
                    found[label]["frameBBox"] = min(
                        containing_frames,
                        key=lambda frame: (frame[2] - frame[0])
                        * (frame[3] - frame[1]),
                    )
    return found


def validate_native_artifacts(docx_path: Path, pdf_path: Path) -> dict[str, Any]:
    """Validate native WPS outputs and return redaction-safe structural data."""
    require_probe_dependencies()
    docx = Path(docx_path).expanduser().resolve()
    pdf = Path(pdf_path).expanduser().resolve()
    validate_office_package(docx, "docx")
    validate_pdf(pdf)
    return {
        "docx": {"name": docx.name, "sha256": sha256_file(docx)},
        "pdf": {
            "name": pdf.name,
            "sha256": sha256_file(pdf),
            "snapshot": snapshot_pdf(pdf),
        },
    }


def _normalized_key(value: str) -> str:
    return "".join(character for character in value.lower() if character.isalnum())


def validate_evidence_privacy(
    value: Any, *, forbidden_roots: tuple[Path, ...] = ()
) -> None:
    """Reject document content, private locators, and non-JSON numeric values."""
    root_strings = tuple(
        str(Path(root).expanduser().resolve()) for root in forbidden_roots
    )

    def visit(current: Any, location: str) -> None:
        if current is None or isinstance(current, bool) or isinstance(current, int):
            return
        if isinstance(current, float):
            if not math.isfinite(current):
                raise ValueError(
                    f"private evidence value at {location} must be finite"
                )
            return
        if isinstance(current, str):
            if (
                _USER_PATH.search(current)
                or _WINDOWS_USER_PATH.search(current)
                or any(root and root in current for root in root_strings)
            ):
                raise ValueError(f"private evidence path at {location}")
            return
        if isinstance(current, Mapping):
            for key, item in current.items():
                if not isinstance(key, str):
                    raise ValueError(
                        f"private evidence key at {location} must be a string"
                    )
                if _normalized_key(key) in _PRIVATE_KEYS:
                    raise ValueError(f"private evidence field at {location}.{key}")
                visit(item, f"{location}.{key}")
            return
        if isinstance(current, (list, tuple)):
            for index, item in enumerate(current):
                visit(item, f"{location}[{index}]")
            return
        raise ValueError(f"private evidence value at {location} is not JSON data")

    visit(value, "evidence")
