from __future__ import annotations

import json
import re
import shutil
import tempfile
import time
import zipfile
from pathlib import Path
from typing import Any, Callable, Optional

from ..longform.macos_executor import MacOSLongformExecutor
from ..longform.pipeline import build_longform_generation, execute_longform_plan
from .bridge import LoopbackBridge
from .models import PathPolicy
from .runtime import ProbeRuntime, read_wps_version

ORIGINS = {
    "http://127.0.0.1:3889",
    "http://127.0.0.1:3890",
    "http://127.0.0.1:3891",
}
FIXTURES_DIR = Path(__file__).resolve().parents[4] / "tests" / "longform_m2" / "fixtures"
PDF_FIXTURE_NAME = "academic"

_SECTPR_RE = re.compile(r"<w:sectPr[^>]*>.*?</w:sectPr>", re.DOTALL)
_ROLE_RE = re.compile(
    r"WpsComposerSectionRole_[0-9]+</w:instrText>.*?<w:t[^>]*>([^<]+)</w:t>",
    re.DOTALL,
)
_PG_FMT_RE = re.compile('<w:pgNumType[^>]*?w:fmt="([^"]+)"[^>]*>')
_PG_START_RE = re.compile('<w:pgNumType[^>]*?w:start="([^"]+)"[^>]*>')
_ORIENT_RE = re.compile('<w:pgSz[^>]*?w:orient="([^"]+)"[^>]*>')
_STYLE_RE = re.compile(r"<w:style[^>]*?</w:style>", re.DOTALL)
_STYLE_NAME_RE = re.compile('<w:name[^>]*?w:val="([^"]+)"')
_TOC_RE = re.compile(r"<w:instrText[^>]*>\\s*TOC\\s*</w:instrText>")
_PAGE_RE = re.compile(r"<w:instrText[^>]*>\\s*PAGE\\s*</w:instrText>")
_BORDER_RE = re.compile('<w:bottom[^>]*w:val="[^"]*"')
_ILVL_RE = re.compile(r"<w:ilvl")


def _write_json(path: Path, payload: object) -> None:
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _allocate_runtime_dir() -> tuple[Path, Path]:
    root = Path(tempfile.mkdtemp(prefix="wpscomposer-longform-m2-"))
    return root, root / "runtime"


def _copy_logs(runtime: ProbeRuntime, output_dir: Path) -> dict[str, str]:
    copied: dict[str, str] = {}
    existing = {name: path for name, path in runtime.logs.items() if path.is_file()}
    if not existing:
        return copied
    logs_dir = output_dir / "logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    for component, source in existing.items():
        target = logs_dir / source.name
        shutil.copy2(source, target)
        copied[component] = str(target.resolve())
    return copied


def _wait_for_writer_registration(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    timeout: float,
) -> None:
    """Retry writer activation until it registers with the bridge."""
    expected = {"writer"}
    deadline = time.monotonic() + timeout
    for attempt in range(4):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            bridge.wait_registered(expected, 0)
            return
        try:
            bridge.wait_registered(expected, min(10, remaining))
            return
        except TimeoutError:
            if attempt == 3:
                raise
            runtime.activate_component("writer")


def _sectpr_roles(document_xml: str) -> dict[str, list[dict[str, str]]]:
    roles: dict[str, list[dict[str, str]]] = {}
    for match in _SECTPR_RE.finditer(document_xml):
        sect = match.group(0)
        role_match = _ROLE_RE.search(sect)
        role = role_match.group(1) if role_match else "unknown"
        pg_fmt_match = _PG_FMT_RE.search(sect)
        pg_start_match = _PG_START_RE.search(sect)
        orient_match = _ORIENT_RE.search(sect)
        roles.setdefault(role, []).append(
            {
                "fmt": pg_fmt_match.group(1) if pg_fmt_match else "none",
                "start": pg_start_match.group(1) if pg_start_match else "",
                "orient": orient_match.group(1) if orient_match else "portrait",
                "hasFooterRef": str("w:footerReference" in sect),
                "hasHeaderRef": str("w:headerReference" in sect),
            }
        )
    return roles


def inspect_docx(path: Path) -> dict[str, Any]:
    """Read structural evidence from a generated DOCX package."""
    with zipfile.ZipFile(path) as package:
        document_xml = package.read("word/document.xml").decode("utf-8")
        roles = _sectpr_roles(document_xml)
        hf_parts = {
            name: package.read(name).decode("utf-8")
            for name in package.namelist()
            if name.startswith(("word/header", "word/footer"))
        }
        styles_xml: dict[str, str] = {}
        try:
            styles = package.read("word/styles.xml").decode("utf-8")
        except KeyError:
            styles = ""
        for match in _STYLE_RE.finditer(styles):
            fragment = match.group(0)
            name_match = _STYLE_NAME_RE.search(fragment)
            if name_match:
                styles_xml[name_match.group(1)] = fragment
        try:
            numbering_xml = package.read("word/numbering.xml").decode("utf-8")
        except KeyError:
            numbering_xml = ""
        try:
            content_types = package.read("[Content_Types].xml").decode("utf-8")
        except KeyError:
            content_types = ""

    header_xml = " ".join(
        xml for name, xml in hf_parts.items() if name.startswith("word/header")
    )
    footer_xml = " ".join(
        xml for name, xml in hf_parts.items() if name.startswith("word/footer")
    )
    toc_field_found = bool(_TOC_RE.search(footer_xml + document_xml))
    page_field_found = bool(_PAGE_RE.search(footer_xml))
    centered_header = '<w:jc w:val="center"' in header_xml
    header_border = bool(_BORDER_RE.search(header_xml))

    return {
        "sectPrCount": document_xml.count("<w:sectPr"),
        "roles": roles,
        "headerPartCount": sum(
            1 for name in hf_parts if name.startswith("word/header")
        ),
        "footerPartCount": sum(
            1 for name in hf_parts if name.startswith("word/footer")
        ),
        "centeredHeader": centered_header,
        "headerBottomBorder": header_border,
        "pageFieldInFooter": page_field_found,
        "tocFieldFound": toc_field_found,
        "numberingXmlPresent": bool(numbering_xml),
        "numberingLevels": bool(_ILVL_RE.search(numbering_xml)),
        "tocStyles": {
            name: name in styles_xml
            for name in ("TOC 1", "TOC 2", "TOC 3")
        },
        "contentTypes": bool(content_types),
    }


def _fixture_result(
    fixture_path: Path,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    output_dir: Path,
    timeout: float,
) -> dict[str, Any]:
    started = time.monotonic()
    name = fixture_path.stem
    artifact = output_dir / f"{name}.docx"
    try:
        markdown = fixture_path.read_text(encoding="utf-8")
        build = build_longform_generation(
            markdown, base_dir=str(fixture_path.parent)
        )
        executor = MacOSLongformExecutor(
            bridge=bridge,
            staging_dir=str(runtime.staging_dir),
        )
        outcome = execute_longform_plan(
            build,
            executor,
            deadline=time.monotonic() + timeout,
        )
        staged = Path(outcome.staged_artifact)
        shutil.copy2(staged, artifact)
        inspection = inspect_docx(artifact)
        return {
            "fixture": name,
            "status": "passed" if not outcome.issues else "issues",
            "artifact": str(artifact.resolve()),
            "size": artifact.stat().st_size,
            "appliedOperations": getattr(outcome, "applied_operations", None),
            "issueCodes": [issue.to_dict() for issue in outcome.issues],
            "paginationMap": outcome.pagination_map.to_dict(),
            "inspection": inspection,
            "timingMs": round((time.monotonic() - started) * 1000, 3),
        }
    except Exception as exc:
        return {
            "fixture": name,
            "status": "failed",
            "artifact": str(artifact.resolve()) if artifact.exists() else None,
            "error": {
                "type": type(exc).__name__,
                "message": str(exc),
            },
            "timingMs": round((time.monotonic() - started) * 1000, 3),
        }


def _convert_one_to_pdf(
    docx_path: Path,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    output_dir: Path,
    timeout: float,
) -> dict[str, Any]:
    """Convert a single DOCX to PDF using the already-running WPS writer add-in."""
    started = time.monotonic()
    output_pdf = output_dir / f"{docx_path.stem}.pdf"
    try:
        if runtime.staging_dir is None:
            raise RuntimeError("WPS staging session is not available")
        policy = PathPolicy((runtime.staging_dir,))
        staged_source = policy.require_allowed(
            runtime.staging_dir / f"pdf-source-{docx_path.stem}.docx"
        )
        staged_output = policy.require_allowed(
            runtime.staging_dir / f"pdf-output-{docx_path.stem}.pdf"
        )
        shutil.copy2(docx_path, staged_source)
        command = bridge.issue(
            "writer",
            "convert_writer_pdf",
            {
                "sourcePath": str(staged_source),
                "outputPath": str(staged_output),
            },
        )
        result = bridge.wait_result(command.id, timeout=timeout)
        if not result.ok:
            error = result.error or {}
            raise RuntimeError(
                f"PDF conversion failed: {error.get('code')}: {error.get('message')}"
            )
        reported = policy.require_allowed(str(result.value.get("path", "")))
        if reported != staged_output:
            raise RuntimeError(
                f"WPS reported unexpected PDF path: {reported}"
            )
        deadline = time.monotonic() + min(timeout, 30.0)
        while not staged_output.is_file():
            if time.monotonic() > deadline:
                raise TimeoutError("PDF output did not appear")
            time.sleep(0.1)
        shutil.copy2(staged_output, output_pdf)
        return {
            "fixture": docx_path.stem,
            "status": "passed",
            "pdf": str(output_pdf.resolve()),
            "size": output_pdf.stat().st_size,
            "timingMs": round((time.monotonic() - started) * 1000, 3),
        }
    except Exception as exc:
        return {
            "fixture": docx_path.stem,
            "status": "failed",
            "pdf": str(output_pdf.resolve()) if output_pdf.exists() else None,
            "error": {
                "type": type(exc).__name__,
                "message": str(exc),
            },
            "timingMs": round((time.monotonic() - started) * 1000, 3),
        }


def run_longform_m2_evidence(
    output_dir: Path,
    timeout: float = 300.0,
    *,
    node: Optional[str] = None,
    fixtures_dir: Optional[Path] = None,
    bridge_factory: Callable = LoopbackBridge,
    runtime_factory: Callable = ProbeRuntime,
    pdf_converter: Optional[Callable[[Path], Path]] = None,
) -> Path:
    """Generate real macOS WPS evidence for all M2 long-form fixtures.

    Args:
        output_dir: Directory where artifacts and ``evidence-report.json``
            are written.
        timeout: Overall per-fixture command timeout in seconds.
        node: Optional Node.js executable override.
        fixtures_dir: Directory containing the ``.md`` fixtures. Defaults to
            ``tests/longform_m2/fixtures``.
        bridge_factory: LoopbackBridge factory for tests.
        runtime_factory: ProbeRuntime factory for tests.
        pdf_converter: Optional callable that receives a DOCX path and returns
            a PDF path. When omitted, PDF conversion uses the live WPS bridge.

    Returns:
        Path to the generated ``evidence-report.json``.
    """
    if timeout <= 0:
        raise ValueError("timeout must be positive")
    started = time.monotonic()
    output = output_dir.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=True)
    report_path = output / "evidence-report.json"

    fixtures_root = (fixtures_dir or FIXTURES_DIR).expanduser().resolve()
    fixture_paths = sorted(
        path for path in fixtures_root.glob("*.md") if path.is_file()
    )

    runtime_root, runtime_dir = _allocate_runtime_dir()
    repository_root = Path(__file__).resolve().parents[4]
    probe_root = repository_root / "macos" / "wps-jsapi-probe"

    report: dict[str, Any] = {
        "schemaVersion": 1,
        "status": "failed",
        "backend": "mac-wps-jsapi-longform-m2",
        "platform": "macOS",
        "wpsVersion": "unknown",
        "wpsjsVersion": "2.2.3",
        "fixturesDir": str(fixtures_root),
        "fixtureCount": len(fixture_paths),
        "fixtures": [],
        "pdfConversion": None,
        "logPaths": {},
        "realWpsQuirks": [],
        "timingsMs": {},
    }

    runtime: Optional[ProbeRuntime] = None
    try:
        report["wpsVersion"] = read_wps_version()
        with bridge_factory(ORIGINS) as bridge:
            runtime = runtime_factory(
                probe_root,
                runtime_dir,
                bridge.url,
                bridge.token,
                node_override=node,
            )
            with runtime:
                if runtime.staging_dir is None:
                    raise RuntimeError("WPS staging session was not created")
                runtime.prepare_profiles()
                runtime.start_servers()
                runtime.activate_component("writer")
                _wait_for_writer_registration(bridge, runtime, timeout=60.0)

                for fixture_path in fixture_paths:
                    result = _fixture_result(
                        fixture_path,
                        bridge,
                        runtime,
                        output,
                        timeout,
                    )
                    report["fixtures"].append(result)

                successful = [
                    Path(r["artifact"])
                    for r in report["fixtures"]
                    if r.get("status") in {"passed", "issues"}
                    and r.get("artifact")
                ]
                pdf_source = next(
                    (p for p in successful if p.stem == PDF_FIXTURE_NAME),
                    successful[0] if successful else None,
                )
                if pdf_source is not None:
                    if pdf_converter is not None:
                        pdf_result = {
                            "fixture": pdf_source.stem,
                            "status": "passed",
                            "pdf": str(pdf_converter(pdf_source).resolve()),
                            "size": Path(str(pdf_converter(pdf_source))).stat().st_size,
                            "timingMs": 0,
                        }
                    else:
                        pdf_result = _convert_one_to_pdf(
                            pdf_source,
                            bridge,
                            runtime,
                            output,
                            timeout,
                        )
                    report["pdfConversion"] = pdf_result

    except Exception as exc:
        report["status"] = "failed"
        report["error"] = {"type": type(exc).__name__, "message": str(exc)}
    finally:
        if runtime is not None:
            report["logPaths"] = _copy_logs(runtime, output)
            if runtime.registration_restored:
                shutil.rmtree(runtime_root, ignore_errors=True)
            else:
                report["recoveryDirectory"] = str(
                    getattr(runtime, "recovery_dir", runtime_dir)
                )
        else:
            shutil.rmtree(runtime_root, ignore_errors=True)

    passed = sum(
        1 for r in report["fixtures"]
        if r.get("status") in {"passed", "issues"}
    )
    report["status"] = (
        "passed"
        if passed == len(fixture_paths) and fixture_paths
        else "failed"
        if not passed
        else "partial"
    )
    report["timingsMs"]["total"] = round(
        (time.monotonic() - started) * 1000, 3
    )
    _write_json(report_path, report)
    return report_path


if __name__ == "__main__":
    import sys

    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "build/longform-m2/macos-native-20260822-1"
    )
    run_longform_m2_evidence(out)
