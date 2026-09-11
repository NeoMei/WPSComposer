"""Real macOS/Windows WPS evidence runner for the M5 quality lifecycle."""

from __future__ import annotations

import hashlib
import json
import platform
from pathlib import Path
import re
import shutil
import subprocess
import time
from typing import Any, Callable

from .artifact_transport import validate_pdf
from .longform.lifecycle import run_longform_lifecycle
from .longform.pipeline import _is_absolute_path, build_longform_generation
from .longform.platform_runtime import MacLongformAdapter, WindowsLongformAdapter
from .longform.windows_executor import _create_dedicated_composer
from .macos_probe.runtime import read_wps_version


ROOT = Path(__file__).resolve().parents[3]
FIXTURES = ROOT / "tests" / "longform_m5" / "fixtures"
FIXTURE_NAMES = (
    "academic",
    "toc_dense",
    "wide_objects",
    "degradation",
    "unicode",
    "plain_short",
)
_SAFE_NAME = re.compile(r"^[A-Za-z0-9._/-]{1,180}$")
_COUNT_KEYS = frozenset({"generation", "export", "patch", "analysis"})
_STAGE_KEYS = frozenset({"generation", "export", "patch", "analysis", "total"})


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _performance_markdown(page_count: int = 60) -> str:
    parts = [
        "---\n",
        "title: M5 native performance fixture\n",
        "author: WPSComposer\n",
        "toc: true\n",
        "title_page: true\n",
        "heading_numbering: decimal\n",
        "---\n",
    ]
    for index in range(1, page_count + 1):
        parts.extend((
            f"\n# 性能章节 {index}\n\n",
            "确定性的非客户正文用于真实 WPS 分页和 PDF 质量门性能测量。\n\n",
        ))
        if index < page_count:
            parts.append(":::page-break\n:::\n")
    return "".join(parts)


class _InstrumentedAdapter:
    def __init__(self, adapter: Any) -> None:
        self.adapter = adapter
        self.counts = {key: 0 for key in _COUNT_KEYS}
        self.durations = {key: 0.0 for key in _STAGE_KEYS if key != "total"}

    def _measure(self, key: str, operation: Callable[[], Any]) -> Any:
        self.counts[key] += 1
        started = time.monotonic()
        try:
            return operation()
        finally:
            self.durations[key] += time.monotonic() - started

    def execute(self, build, directives, deadline):
        return self._measure(
            "generation", lambda: self.adapter.execute(build, directives, deadline)
        )

    def export_pdf(self, docx, deadline):
        return self._measure(
            "export", lambda: self.adapter.export_pdf(docx, deadline)
        )

    def patch_quality_notices(self, docx, notices, deadline):
        return self._measure(
            "patch",
            lambda: self.adapter.patch_quality_notices(docx, notices, deadline),
        )

    def analyze(self, pdf, outcome, build):
        return self._measure(
            "analysis", lambda: self.adapter.analyze(pdf, outcome, build)
        )

    def validate_patch(self, *args):
        return self.adapter.validate_patch(*args)

    def publish(self, *args):
        return self.adapter.publish(*args)

    def cleanup(self, *args):
        return self.adapter.cleanup(*args)


def _render_representative_pages(pdf: Path, output: Path, pages: int) -> list[str]:
    renderer = shutil.which("pdftoppm")
    if not renderer:
        raise RuntimeError(
            "M5 visual evidence requires pdftoppm (Poppler) on PATH"
        )
    screenshots = output / "screenshots"
    screenshots.mkdir(exist_ok=True)
    selected = range(1, pages + 1) if pages <= 8 else (1, (pages + 1) // 2, pages)
    rendered = []
    for page in selected:
        target = screenshots / f"{pdf.stem}-{page}"
        subprocess.run(
            [
                renderer, "-png", "-r", "120", "-f", str(page), "-l",
                str(page), "-singlefile", str(pdf), str(target),
            ],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
        )
        rendered.append(target.with_suffix(".png"))
    if not rendered or any(path.stat().st_size < 1024 for path in rendered):
        raise RuntimeError("M5 representative screenshot rendering failed")
    return [path.relative_to(output).as_posix() for path in rendered]


def _pdf_page_count(path: Path) -> int:
    from pypdf import PdfReader

    return len(PdfReader(str(path), strict=True).pages)


def _pdf_metadata(path: Path) -> tuple[str, str]:
    from pypdf import PdfReader

    metadata = PdfReader(str(path), strict=True).metadata or {}
    return str(metadata.get("/Title") or ""), str(metadata.get("/Author") or "")


def _run_one(
    name: str,
    markdown: str,
    base_dir: Path,
    output: Path,
    timeout: float,
    adapter_factory: Callable[[Any], Any],
) -> dict[str, Any]:
    build = build_longform_generation(markdown, base_dir=str(base_dir))
    adapter = adapter_factory(build)
    measured = _InstrumentedAdapter(adapter)
    target = output / f"{name}.pdf"
    started = time.monotonic()
    try:
        outcome = run_longform_lifecycle(
            build,
            measured,
            target,
            format_name="pdf",
            timeout=timeout,
            overwrite=False,
            quality_analyzer=measured.analyze,
        )
    finally:
        adapter.close()
    total = time.monotonic() - started
    validate_pdf(target)
    pages = _pdf_page_count(target)
    title, author = _pdf_metadata(target)
    if author != "WPSComposer" or not title:
        raise RuntimeError("M5 PDF metadata privacy gate failed")
    screenshots = _render_representative_pages(target, output, pages)
    durations = {
        key: round(value, 4) for key, value in measured.durations.items()
    }
    durations["total"] = round(total, 4)
    return {
        "name": name,
        "pageCount": pages,
        "operationCount": len(build.plan.operations),
        "counts": dict(sorted(measured.counts.items())),
        "stageSeconds": dict(sorted(durations.items())),
        "issueCodes": sorted({issue.code for issue in outcome.issues}),
        "artifact": {"name": target.name, "sha256": _sha256(target)},
        "screenshots": screenshots,
    }


def validate_m5_evidence_report(report: Any) -> None:
    if not isinstance(report, dict) or set(report) != {
        "version", "environment", "fixtures", "performance",
    }:
        raise ValueError("M5 evidence root is invalid")
    if report["version"] != "M5":
        raise ValueError("M5 evidence version is invalid")
    environment = report["environment"]
    if not isinstance(environment, dict) or set(environment) != {
        "system", "machine", "wpsVersion", "protocolVersion", "semanticVersion",
    }:
        raise ValueError("M5 evidence environment is invalid")
    if (
        environment["system"] not in {"Darwin", "Windows"}
        or environment["protocolVersion"] != 2
        or not isinstance(environment["machine"], str)
        or not environment["machine"]
        or not isinstance(environment["wpsVersion"], str)
        or not environment["wpsVersion"]
    ):
        raise ValueError("M5 evidence platform contract is invalid")
    entries = report["fixtures"] + [report["performance"]]
    if len(report["fixtures"]) != 6:
        raise ValueError("M5 fixture evidence is incomplete")
    for entry in entries:
        if not isinstance(entry, dict) or set(entry) != {
            "name", "pageCount", "operationCount", "counts", "stageSeconds",
            "issueCodes", "artifact", "screenshots",
        }:
            raise ValueError("M5 evidence entry is invalid")
        if not _SAFE_NAME.fullmatch(entry["name"]):
            raise ValueError("M5 evidence name is invalid")
        if type(entry["pageCount"]) is not int or entry["pageCount"] < 1:
            raise ValueError("M5 evidence page count is invalid")
        if set(entry["counts"]) != _COUNT_KEYS or any(
            type(value) is not int or value < 0 for value in entry["counts"].values()
        ):
            raise ValueError("M5 evidence counts are invalid")
        if not 1 <= entry["counts"]["generation"] <= 2:
            raise ValueError("M5 generation cap evidence is invalid")
        if not 1 <= entry["counts"]["export"] <= 3:
            raise ValueError("M5 export cap evidence is invalid")
        if entry["counts"]["patch"] not in {0, 1}:
            raise ValueError("M5 patch cap evidence is invalid")
        if set(entry["stageSeconds"]) != _STAGE_KEYS or any(
            not isinstance(value, (int, float)) or value < 0
            for value in entry["stageSeconds"].values()
        ):
            raise ValueError("M5 stage timings are invalid")
        artifact = entry["artifact"]
        if (
            not isinstance(artifact, dict)
            or set(artifact) != {"name", "sha256"}
            or _is_absolute_path(artifact["name"])
            or not re.fullmatch(r"[0-9a-f]{64}", artifact["sha256"])
        ):
            raise ValueError("M5 artifact evidence is invalid")
        if not isinstance(entry["screenshots"], list) or not entry["screenshots"] or any(
            _is_absolute_path(name) or not name.startswith("screenshots/")
            for name in entry["screenshots"]
        ):
            raise ValueError("M5 screenshot evidence is invalid")
    performance = report["performance"]
    if not 50 <= performance["pageCount"] <= 100:
        raise ValueError("M5 native performance page count is outside 50-100")
    if performance["stageSeconds"]["total"] >= 600:
        raise ValueError("M5 native performance exceeded 600 seconds")
    serialized = json.dumps(report, ensure_ascii=False)
    if "/Users/" in serialized or "\\Users\\" in serialized:
        raise ValueError("M5 evidence contains a private path")


def _read_windows_wps_version() -> str:
    composer = _create_dedicated_composer()
    try:
        version = getattr(composer._app, "Version", None)
        if callable(version):
            version = version()
        value = str(version or "").strip()
        if not value:
            raise RuntimeError("Windows WPS version is unavailable")
        return value
    finally:
        composer.close(save_changes=False)


def run_longform_m5_evidence(output: Path, timeout: float = 300.0) -> Path:
    system = platform.system()
    if system not in {"Darwin", "Windows"}:
        raise RuntimeError("M5 real evidence requires macOS or Windows WPS")
    output = Path(output).expanduser().resolve()
    output.mkdir(parents=True, exist_ok=False)
    if system == "Darwin":
        adapter_factory: Callable[[Any], Any] = MacLongformAdapter
        wps_version = read_wps_version()
    else:
        adapter_factory = WindowsLongformAdapter
        wps_version = _read_windows_wps_version()
    fixtures = []
    for name in FIXTURE_NAMES:
        fixtures.append(
            _run_one(
                name,
                (FIXTURES / f"{name}.md").read_text(encoding="utf-8"),
                FIXTURES,
                output,
                timeout,
                adapter_factory,
            )
        )
    performance = _run_one(
        "performance-60",
        _performance_markdown(60),
        FIXTURES,
        output,
        min(600.0, max(timeout, 300.0)),
        adapter_factory,
    )
    report = {
        "version": "M5",
        "environment": {
            "system": system,
            "machine": platform.machine(),
            "wpsVersion": wps_version,
            "protocolVersion": 2,
            "semanticVersion": "longform-1",
        },
        "fixtures": fixtures,
        "performance": performance,
    }
    validate_m5_evidence_report(report)
    path = output / "evidence.json"
    path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return path


__all__ = ["run_longform_m5_evidence", "validate_m5_evidence_report"]
