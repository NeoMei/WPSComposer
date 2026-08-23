from __future__ import annotations

import hashlib
import json
import re
import shutil
import subprocess
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Optional, Sequence

from .longform.macos_executor import MacOSLongformExecutor
from .longform.native_math import NativeMathConversionError, convert_restricted_latex
from .longform.pipeline import (
    _is_absolute_path,
    build_longform_generation,
    execute_longform_plan,
)
from .longform.privacy import redact_private_text
from .macos_probe.bridge import LoopbackBridge
from .macos_probe.longform_evidence import (
    ORIGINS,
    _convert_one_to_pdf,
    _wait_for_writer_registration,
)
from .macos_probe.models import ProtocolError
from .macos_probe.runtime import ProbeRuntime, read_wps_version


ROOT = Path(__file__).resolve().parents[3]
FIXTURES = ROOT / "tests" / "longform_m4" / "fixtures"
PROBE_ROOT = ROOT / "macos" / "wps-jsapi-probe"
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_ALLOWED_ROOT = frozenset({
    "version", "wpsVersion", "artifacts", "metrics", "codes", "rounds", "screenshots",
})
_METRIC_KEYS = frozenset({
    "formulaCount", "editableFormulaCount", "equationSequenceCount", "referenceCount",
    "degradedFormulaCount", "formulaFallbackImageCount", "formulaNoticeCount",
    "citationCount", "bibliographyCount", "hangingBibliographyCount",
    "inlineNoticeCount", "blockNoticeCount", "documentNoticeCount", "pageCount",
    "citationOrderValid", "bibliographyOrderValid", "formulaOrderValid",
    "noticePlacementValid",
    "centeredFormulaCount", "numberAlignedFormulaCount",
})
_PRIVATE_KEYS = frozenset({
    "path", "paths", "payload", "source", "fieldresults", "bookmarknames",
    "locator", "resourceid", "resources", "traceback", "exception",
})
_NOTICE_CODES = (
    "REFERENCE_UNRESOLVED",
    "FORMULA_MALFORMED",
    "FORMULA_FALLBACK_IMAGE_UNAVAILABLE",
    "BIBLIOGRAPHY_ENTRY_MALFORMED",
    "CONFIG_VALUE_INVALID",
    "EQUATION_INSERT_FAILED",
    "BIBLIOGRAPHY_INSERT_FAILED",
    "CROSS_REFERENCE_FAILED",
)
_METRIC_LABELS = frozenset({"initial", "moved", "degradation", "runtimeRecovery"})
_SAFE_RELATIVE_PART = re.compile(r"[A-Za-z0-9._-]{1,128}")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _fixture_build(fixtures: Path, name: str):
    source = fixtures / f"{name}.md"
    return build_longform_generation(source.read_text(encoding="utf-8"), str(fixtures))


def _formula_snapshot(operation: Mapping[str, Any]) -> dict[str, Any]:
    args = operation["args"]
    content = args["content"]
    result: dict[str, Any] = {
        "id": operation["nodeId"],
        "numbering": args["numbering"],
    }
    if "nativeMath" in content:
        result["nativeMath"] = content["nativeMath"]
    else:
        result["plannedDegradation"] = {
            key: content["plannedDegradation"][key]
            for key in ("code", "placement", "fallbackKind")
        }
    fallback = args.get("fallbackResource")
    if fallback:
        result["fallback"] = (
            "resource" if "fallbackResourceId" in fallback
            else fallback["fallbackResourcePlannedDegradation"]["code"]
        )
    return result


def _notice_placements(operations: Sequence[Mapping[str, Any]]) -> list[str]:
    placements: list[str] = []
    for operation in operations:
        args = operation.get("args", {})
        if operation["op"] == "writer.reserve_document_quality_anchor":
            placements.extend(item["placement"] for item in args.get("notices", []))
        elif operation["op"] in {
            "writer.add_degradation_notice", "writer.add_inline_degradation",
            "writer.add_document_quality_notice",
        }:
            placements.append(args.get("placement", "document"))
        elif operation["op"] == "writer.add_cross_reference":
            placements.extend(
                "inline" for run in args.get("runs", []) if run.get("type") == "degradation"
            )
        elif operation["op"] == "writer.add_equation":
            content = args.get("content", {})
            if "plannedDegradation" in content:
                placements.append(content["plannedDegradation"]["placement"])
            fallback = args.get("fallbackResource", {}).get("fallbackResourcePlannedDegradation")
            if fallback:
                placements.append(fallback["placement"])
    return sorted(placements, key={"block": 0, "document": 1, "inline": 2}.get)


def _recovery_snapshot(operations: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    recovered: list[dict[str, Any]] = []
    seen: set[str] = set()
    for operation in operations:
        policy = operation.get("failurePolicy") or {}
        if policy.get("mode") != "degrade" or operation["op"] in seen:
            continue
        if operation["op"] not in {
            "writer.add_equation", "writer.add_cross_reference", "writer.add_bibliography",
        }:
            continue
        seen.add(operation["op"])
        recovered.append({
            "op": operation["op"],
            "recoverableCodes": policy["recoverableCodes"],
            "fallback": policy["fallback"],
        })
    order = {
        "writer.add_equation": 0,
        "writer.add_cross_reference": 1,
        "writer.add_bibliography": 2,
    }
    return sorted(recovered, key=lambda item: order[item["op"]])


def _fixture_snapshot(fixtures: Path, name: str) -> dict[str, Any]:
    build = _fixture_build(fixtures, name)
    operations = build.plan.to_dict()["operations"]
    citations: list[dict[str, Any]] = []
    bibliography: list[dict[str, Any]] = []
    for operation in operations:
        if operation["op"] == "writer.add_cross_reference":
            citations.extend(
                {"targetId": run["targetId"], "number": run["number"]}
                for run in operation["args"].get("runs", [])
                if run.get("type") == "citation"
            )
        elif operation["op"] == "writer.add_bibliography" and operation["args"].get("schemaVersion") == 1:
            bibliography.extend(
                {"id": item["id"], "number": item["number"], "cited": item["cited"]}
                for item in operation["args"]["entries"]
            )
    return {
        "formulas": [
            _formula_snapshot(operation)
            for operation in operations if operation["op"] == "writer.add_equation"
        ],
        "citations": citations,
        "bibliography": bibliography,
        "issueCodes": sorted({issue.code for issue in build.issues}),
        "noticePlacements": _notice_placements(operations),
        "recovery": _recovery_snapshot(operations),
    }


def build_acceptance_snapshot(fixtures: Path = FIXTURES) -> dict[str, Any]:
    failures: list[str] = []
    for source in (r"\input{/Users/alice/private.tex}", r"\unknown{C:\private\x}"):
        try:
            convert_restricted_latex(source)
        except NativeMathConversionError as error:
            failures.append(error.code)
    return {
        "version": "M4-offline-1",
        "fixtures": {
            name: _fixture_snapshot(fixtures, name)
            for name in ("acceptance_chapter", "acceptance_global", "degradation")
        },
        "privacyFailureCodes": failures,
    }


def _walk_private(value: Any, location: str = "evidence") -> None:
    if isinstance(value, Mapping):
        for key, child in value.items():
            normalized = re.sub(r"[^a-z0-9]", "", str(key).casefold())
            if normalized in _PRIVATE_KEYS:
                raise ValueError(f"private field at {location}.{key}")
            if normalized in {"sha256", "name", "screenshots"}:
                continue
            _walk_private(child, f"{location}.{key}")
    elif isinstance(value, (list, tuple)):
        for index, child in enumerate(value):
            _walk_private(child, f"{location}[{index}]")
    elif isinstance(value, str) and redact_private_text(value) != value:
        raise ValueError(f"private value at {location}")


def validate_m4_evidence_report(value: Any) -> None:
    _walk_private(value)
    if not isinstance(value, dict) or set(value) != _ALLOWED_ROOT or value["version"] != "M4":
        raise ValueError("M4 evidence root shape is invalid")
    if not isinstance(value["wpsVersion"], str) or not value["wpsVersion"]:
        raise ValueError("M4 WPS version is invalid")
    if not isinstance(value["artifacts"], list) or not value["artifacts"]:
        raise ValueError("M4 artifacts are incomplete")
    artifact_names: list[str] = []
    for artifact in value["artifacts"]:
        if not isinstance(artifact, dict) or set(artifact) != {"name", "sha256"}:
            raise ValueError("M4 artifact shape is invalid")
        name = Path(artifact["name"])
        if _is_absolute_path(artifact["name"]) or ".." in name.parts or len(name.parts) > 2:
            raise ValueError("M4 artifact name is not relative")
        if not name.parts or any(not _SAFE_RELATIVE_PART.fullmatch(part) for part in name.parts):
            raise ValueError("M4 artifact name is not privacy safe")
        if name.suffix.lower() not in {".docx", ".pdf", ".png"}:
            raise ValueError("M4 artifact suffix is invalid")
        if not re.fullmatch(r"[0-9a-f]{64}", artifact["sha256"]):
            raise ValueError("M4 artifact digest is invalid")
        artifact_names.append(str(name))
    if len(set(artifact_names)) != len(artifact_names):
        raise ValueError("M4 artifact names are duplicated")
    if not isinstance(value["metrics"], dict) or set(value["metrics"]) != _METRIC_LABELS:
        raise ValueError("M4 metrics are incomplete")
    for metrics in value["metrics"].values():
        if not isinstance(metrics, dict) or set(metrics) != _METRIC_KEYS:
            raise ValueError("M4 metric shape is invalid")
        if any(type(item) is not int or item < 0 for item in metrics.values()):
            raise ValueError("M4 metric value is invalid")
    if not isinstance(value["codes"], list) or any(
        not isinstance(code, str) or not re.fullmatch(r"[A-Z][A-Z0-9_]{2,63}", code)
        for code in value["codes"]
    ):
        raise ValueError("M4 issue codes are invalid")
    if value["codes"] != sorted(set(value["codes"])):
        raise ValueError("M4 issue codes are not canonical")
    if not isinstance(value["rounds"], list) or len(value["rounds"]) != 4 or any(
        type(rounds) is not int or rounds < 1 or rounds > 4 for rounds in value["rounds"]
    ):
        raise ValueError("M4 refresh rounds are invalid")
    if not isinstance(value["screenshots"], list) or not value["screenshots"]:
        raise ValueError("M4 screenshots are incomplete")
    if len(set(value["screenshots"])) != len(value["screenshots"]):
        raise ValueError("M4 screenshot names are duplicated")
    for raw_name in value["screenshots"]:
        name = Path(raw_name)
        if (
            not isinstance(raw_name, str) or _is_absolute_path(raw_name) or ".." in name.parts
            or len(name.parts) != 2 or name.parts[0] != "screenshots"
            or name.suffix.lower() != ".png"
            or any(not _SAFE_RELATIVE_PART.fullmatch(part) for part in name.parts)
        ):
            raise ValueError("M4 screenshot name is invalid")


def inspect_m4_docx(path: Path) -> tuple[dict[str, int], dict[str, Any]]:
    with zipfile.ZipFile(path) as package:
        xml = package.read("word/document.xml")
        app_xml = package.read("docProps/app.xml") if "docProps/app.xml" in package.namelist() else b""
    root = ET.fromstring(xml)
    paragraphs = list(root.iter(_W + "p"))
    paragraph_texts = [
        "".join(node.text or "" for node in paragraph.iter(_W + "t"))
        for paragraph in paragraphs
    ]
    bibliography = [
        paragraph for paragraph, text in zip(paragraphs, paragraph_texts)
        if re.match(r"^\[\d+\]\s", text)
    ]
    hanging = 0
    for paragraph in bibliography:
        indent = next(iter(paragraph.iter(_W + "ind")), None)
        spacing = next(iter(paragraph.iter(_W + "spacing")), None)
        if (
            indent is not None and int(indent.get(_W + "hanging", "0")) > 0
            and int(indent.get(_W + "left", "0")) > 0
            and spacing is not None and int(spacing.get(_W + "after", "0")) > 0
        ):
            hanging += 1
    full_text = "\n".join(paragraph_texts)
    page_match = re.search(rb"<Pages>(\d+)</Pages>", app_xml)
    fields: list[tuple[str, str]] = []
    stack: list[dict[str, Any]] = []

    def collect_fields(element: ET.Element, *, in_math: bool = False) -> None:
        now_in_math = in_math or element.tag in {_M + "oMath", _M + "oMathPara"}
        if now_in_math:
            return
        if element.tag == _W + "fldSimple":
            instruction = element.get(_W + "instr", "").strip()
            result = "".join(node.text or "" for node in element.iter(_W + "t")).strip()
            fields.append((instruction, result))
            return
        if element.tag == _W + "fldChar":
            kind = element.get(_W + "fldCharType")
            if kind == "begin":
                stack.append({"instruction": [], "result": [], "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                field = stack.pop()
                fields.append(("".join(field["instruction"]).strip(), "".join(field["result"]).strip()))
        elif element.tag == _W + "instrText" and stack:
            stack[-1]["instruction"].append(element.text or "")
        elif element.tag == _W + "t" and stack:
            for field in stack:
                if field["separated"]:
                    field["result"].append(element.text or "")
        for child in element:
            collect_fields(child, in_math=now_in_math)

    collect_fields(root)
    math_nodes = list(root.iter(_M + "oMath"))
    math_run_property_leaks = list(root.iter(_W + "oMath"))
    equation_paragraphs = [
        paragraph for paragraph in paragraphs
        if b"SEQ WPSC_EQ" in ET.tostring(paragraph, encoding="utf-8")
    ]
    formula_paragraphs = [
        paragraph for paragraph in paragraphs
        if any(True for _ in paragraph.iter(_M + "oMath"))
    ]
    math_field_leaks = []
    for math in math_nodes:
        raw = ET.tostring(math, encoding="utf-8")
        literal = "".join(node.text or "" for node in math.iter())
        if (
            any(True for _ in math.iter(_W + "tab"))
            or any(True for _ in math.iter(_W + "fldChar"))
            or any(True for _ in math.iter(_W + "instrText"))
            or b"SEQ WPSC_EQ" in raw
            or b"STYLEREF" in raw
            or "SEQ WPSC_EQ" in literal
            or "STYLEREF" in literal
        ):
            math_field_leaks.append(math)
    formula_math_counts = [sum(1 for _ in paragraph.iter(_M + "oMath")) for paragraph in formula_paragraphs]
    structural_tags = ("sSup", "sSubSup", "f", "rad", "nary", "d", "m")
    formula_structural_tag_counts = {
        tag: sum(1 for math in math_nodes for _ in math.iter(_M + tag))
        for tag in structural_tags
    }
    acceptance_roles = set(equation_paragraphs + formula_paragraphs + bibliography)
    acceptance_roles.update(
        paragraph for paragraph, text in zip(paragraphs, paragraph_texts)
        if text.startswith("Formula references remain native:")
    )
    def has_notice_style(element: ET.Element) -> bool:
        return (
            any(color.get(_W + "val") == "9C0006" for color in element.iter(_W + "color"))
            or any(shading.get(_W + "fill") == "FCE8E6" for shading in element.iter(_W + "shd"))
            or any(italic.get(_W + "val", "1") not in {"0", "false", "off"} for italic in element.iter(_W + "i"))
        )

    unexpected_notice_style_paragraphs = []
    for paragraph in acceptance_roles:
        paragraph_properties = paragraph.find(_W + "pPr")
        leaked = paragraph_properties is not None and has_notice_style(paragraph_properties)
        if not leaked:
            for run in paragraph.iter(_W + "r"):
                if not has_notice_style(run):
                    continue
                run_text = "".join(node.text or "" for node in run.iter(_W + "t"))
                if not any(code in run_text for code in _NOTICE_CODES):
                    leaked = True
                    break
        if leaked:
            unexpected_notice_style_paragraphs.append(paragraph)
    centered_formulas = sum(
        any(
            tab.get(_W + "val") == "center"
            for tabs in paragraph.iter(_W + "tabs")
            for tab in tabs.iter(_W + "tab")
        )
        for paragraph in equation_paragraphs
    )
    aligned_numbers = sum(
        any(
            tab.get(_W + "val") == "right"
            for tabs in paragraph.iter(_W + "tabs")
            for tab in tabs.iter(_W + "tab")
        )
        and b"SEQ WPSC_EQ" in ET.tostring(paragraph, encoding="utf-8")
        for paragraph in equation_paragraphs
    )
    equation_paragraph_records = []
    for paragraph_index, paragraph in enumerate(paragraphs):
        raw = ET.tostring(paragraph, encoding="utf-8")
        if b"SEQ WPSC_EQ" not in raw:
            continue
        equation_paragraph_records.append({
            "paragraphIndex": paragraph_index,
            "text": paragraph_texts[paragraph_index],
            "bookmarks": [
                node.get(_W + "name", "")
                for node in paragraph.iter(_W + "bookmarkStart")
                if node.get(_W + "name", "").startswith("wpsc_eq_")
            ],
            "imageCount": sum(1 for _ in paragraph.iter(_W + "drawing"))
            + sum(1 for _ in paragraph.iter(_W + "pict")),
            "mathCount": sum(1 for _ in paragraph.iter(_M + "oMath")),
        })
    formula_notice_paragraph_indices = [
        index for index, text in enumerate(paragraph_texts)
        if "EQUATION_INSERT_FAILED" in text
    ]
    degraded_formula_count = sum(
        record["mathCount"] == 0 for record in equation_paragraph_records
    )
    formula_fallback_image_count = sum(
        record["imageCount"] for record in equation_paragraph_records
    )
    metrics = {
        "formulaCount": len(math_nodes),
        "editableFormulaCount": len(math_nodes),
        "degradedFormulaCount": degraded_formula_count,
        "formulaFallbackImageCount": formula_fallback_image_count,
        "formulaNoticeCount": full_text.count("EQUATION_INSERT_FAILED"),
        "equationSequenceCount": sum("SEQ WPSC_EQ" in code for code, _ in fields),
        "referenceCount": sum(code.startswith("REF wpsc_") for code, _ in fields),
        "citationCount": sum(text.count("[1]") + text.count("[2]") for text in paragraph_texts if not re.match(r"^\[\d+\]\s", text)),
        "bibliographyCount": len(bibliography),
        "hangingBibliographyCount": hanging,
        "inlineNoticeCount": full_text.count("REFERENCE_UNRESOLVED"),
        "blockNoticeCount": sum(full_text.count(code) for code in _NOTICE_CODES if code not in {"REFERENCE_UNRESOLVED", "CONFIG_VALUE_INVALID"}),
        "documentNoticeCount": full_text.count("CONFIG_VALUE_INVALID"),
        "pageCount": int(page_match.group(1)) if page_match else 0,
        "citationOrderValid": 0,
        "bibliographyOrderValid": 0,
        "formulaOrderValid": 0,
        "noticePlacementValid": 0,
        "centeredFormulaCount": centered_formulas,
        "numberAlignedFormulaCount": aligned_numbers,
    }
    private = {
        "paragraphTexts": paragraph_texts,
        "fields": fields,
        "bookmarkOrder": [
            element.get(_W + "name", "")
            for element in root.iter(_W + "bookmarkStart")
            if element.get(_W + "name", "").startswith("wpsc_eq_")
        ],
        "formulaStructureValid": (
            not math_field_leaks
            and not math_run_property_leaks
            and all(count == 1 for count in formula_math_counts)
        ),
        "mathFieldLeakCount": len(math_field_leaks),
        "mathRunPropertyLeakCount": len(math_run_property_leaks),
        "formulaParagraphMathCounts": formula_math_counts,
        "formulaStructuralTagCounts": formula_structural_tag_counts,
        "unexpectedNoticeStyleParagraphCount": len(unexpected_notice_style_paragraphs),
        "equationParagraphRecords": equation_paragraph_records,
        "formulaNoticeParagraphIndices": formula_notice_paragraph_indices,
    }
    return metrics, private


def _copy_outcome(build: Any, bridge: LoopbackBridge, runtime: ProbeRuntime, output: Path, timeout: float):
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(runtime.staging_dir))
    outcome = execute_longform_plan(build, executor, deadline=time.monotonic() + timeout)
    shutil.copy2(outcome.staged_artifact, output)
    return outcome


def _mutate_equation(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    source: Path,
    target: Path,
    mutation: Mapping[str, Any],
    timeout: float,
) -> Mapping[str, Any]:
    if runtime.staging_dir is None:
        raise RuntimeError("WPS staging session is unavailable")
    staged_source = runtime.staging_dir / "m4-mutation-source.docx"
    staged_target = runtime.staging_dir / "m4-mutation-output.docx"
    shutil.copy2(source, staged_source)
    command = bridge.issue("writer", "mutate_longform_document", {
        "sourcePath": str(staged_source), "outputPath": str(staged_target),
        "mutations": [dict(mutation)], "resources": {},
    })
    result = bridge.wait_result(command.id, timeout=timeout)
    if not result.ok or not staged_target.is_file():
        raise RuntimeError("M4 native mutation failed")
    shutil.copy2(staged_target, target)
    return dict(result.value)


def _render_pdf(pdf: Path, output: Path) -> list[str]:
    target = output / "screenshots"
    target.mkdir(exist_ok=True)
    prefix = target / pdf.stem
    subprocess.run(
        ["pdftoppm", "-png", "-r", "144", str(pdf), str(prefix)],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
    )
    screenshots = sorted(target.glob(pdf.stem + "-*.png"))
    if not screenshots or any(path.stat().st_size < 1024 for path in screenshots):
        raise AssertionError("M4 PDF screenshots are empty")
    return [str(path.relative_to(output)) for path in screenshots]


def _assert_pdf(pdf: Path) -> None:
    import pdfplumber

    with pdfplumber.open(pdf) as document:
        if not document.pages:
            raise AssertionError("M4 PDF is empty")
        for page in document.pages:
            text = page.extract_text() or ""
            if not text.strip() and not page.images:
                raise AssertionError("M4 PDF contains a blank page")
            for image in page.images:
                if (
                    float(image["x0"]) < -0.5 or float(image["top"]) < -0.5
                    or float(image["x1"]) > float(page.width) + 0.5
                    or float(image["bottom"]) > float(page.height) + 0.5
                ):
                    raise AssertionError("M4 PDF contains clipped content")
                for word in page.extract_words():
                    value = str(word.get("text", ""))
                    is_notice = bool(re.match(
                        r"^\[(?:FORMULA|EQUATION|REFERENCE|BIBLIOGRAPHY|CONFIG|CROSS)_",
                        value,
                    ))
                    if (
                        is_notice
                        and float(word["x0"]) < float(image["x1"])
                        and float(word["x1"]) > float(image["x0"])
                        and float(word["top"]) < float(image["bottom"])
                        and float(word["bottom"]) > float(image["top"])
                    ):
                        raise AssertionError("M4 degradation notice overlaps an image")
        text = "\n".join(page.extract_text() or "" for page in document.pages)
        if "[object Object]" in text or "undefined" in text or "Traceback" in text:
            raise AssertionError("M4 PDF contains a diagnostic placeholder")


def _assert_formula_pdf(
    pdf: Path,
    mode: str,
    expectations: Sequence[Mapping[str, Any]],
) -> None:
    import pdfplumber

    with pdfplumber.open(pdf) as document:
        text = "\n".join(page.extract_text() or "" for page in document.pages)
        image_count = sum(len(page.images) for page in document.pages)
    if mode == "native":
        if "EQUATION_INSERT_FAILED" in text:
            raise AssertionError("M4 native formula PDF contains a degradation notice")
        return
    if mode != "degraded":
        raise AssertionError("M4 formula PDF mode is invalid")
    if text.count("EQUATION_INSERT_FAILED") != len(expectations):
        raise AssertionError("M4 degraded formula PDF notices are incomplete")
    expected_images = sum(bool(item["hasImage"]) for item in expectations)
    if image_count != expected_images:
        raise AssertionError("M4 degraded formula PDF image usage is not exact")
    compact_text = re.sub(r"\s+", "", text)
    for expectation in expectations:
        if expectation["hasImage"]:
            continue
        compact_source = re.sub(r"\s+", "", str(expectation["fallbackText"]))
        if compact_source not in compact_text:
            raise AssertionError("M4 degraded formula PDF source fallback is incomplete")


def _assert_formula_rendering_mode(
    metrics: Mapping[str, int],
    private: Mapping[str, Any],
    expectations: Sequence[Mapping[str, Any]],
) -> str:
    expected_count = len(expectations)
    native_count = metrics["formulaCount"]
    degraded_count = metrics["degradedFormulaCount"]
    if native_count == expected_count and degraded_count == 0:
        if not private["formulaStructureValid"] or private["mathFieldLeakCount"]:
            raise AssertionError("M4 equation fields leaked into an editable math subtree")
        required_structures = {
            "sSup": 1, "sSubSup": 1, "f": 2, "rad": 2,
            "nary": 3, "d": 2, "m": 2,
        }
        if expected_count == 12 and any(
            private["formulaStructuralTagCounts"].get(tag, 0) < count
            for tag, count in required_structures.items()
        ):
            raise AssertionError("M4 accepted formula families were not built into native structures")
        return "native"
    if native_count != 0 or degraded_count != expected_count:
        raise AssertionError("M4 formula rendering mixed native and degraded capability states")
    if private["mathRunPropertyLeakCount"] or private["mathFieldLeakCount"]:
        raise AssertionError("M4 degraded formulas retained pseudo-native math state")
    if metrics["equationSequenceCount"] != expected_count:
        raise AssertionError("M4 degraded formula numbering fields are incomplete")
    if metrics["formulaNoticeCount"] != expected_count:
        raise AssertionError("M4 formulas do not each have one explicit local degradation notice")
    expected_images = sum(bool(item["hasImage"]) for item in expectations)
    if metrics["formulaFallbackImageCount"] != expected_images:
        raise AssertionError("M4 formula fallback images were not used exactly when explicitly provided")
    records = {
        bookmark: record
        for record in private["equationParagraphRecords"]
        for bookmark in record["bookmarks"]
    }
    notice_indices = set(private["formulaNoticeParagraphIndices"])
    for expectation in expectations:
        bookmark = expectation["bookmarkName"]
        record = records.get(bookmark)
        if record is None or record["mathCount"] != 0:
            raise AssertionError("M4 degraded formula lost its numbered source location")
        if expectation["hasImage"]:
            if record["imageCount"] != 1 or record["paragraphIndex"] + 1 not in notice_indices:
                raise AssertionError("M4 explicit formula image lacks one adjacent visible notice")
        else:
            expected_source = "[EQUATION_INSERT_FAILED: " + expectation["fallbackText"] + "]"
            if record["imageCount"] != 0 or expected_source not in record["text"]:
                raise AssertionError("M4 source formula fallback is incomplete or displaced")
    return "degraded"


def _assert_acceptance_content(
    metrics: Mapping[str, int],
    private: Mapping[str, Any],
    *,
    formula_expectations: Sequence[Mapping[str, Any]],
    expected_first_bookmark: str,
    expected_reference_results: Sequence[str],
) -> str:
    mode = _assert_formula_rendering_mode(metrics, private, formula_expectations)
    if private["unexpectedNoticeStyleParagraphCount"]:
        raise AssertionError("M4 content outside notice spans inherited degradation styling")
    text = "\n".join(private["paragraphTexts"])
    if not all(item in text for item in ("Beta cited first.", "Alpha uncited-order declaration.", "Gamma remains uncited.")):
        raise AssertionError("M4 bibliography content is incomplete")
    if not (
        text.find("Beta cited first.") < text.find("Alpha uncited-order declaration.")
        < text.find("Gamma remains uncited.")
    ):
        raise AssertionError("M4 bibliography order is not cited-first")
    expected_citation_paragraph = (
        "Inline continuity begins [1], repeats [1], cites [2], and keeps missing "
        "[REFERENCE_UNRESOLVED 引用目标未解析] in this paragraph."
    )
    if expected_citation_paragraph not in private["paragraphTexts"]:
        raise AssertionError("M4 citation paragraph continuity failed")
    sequence_results = [result for code, result in private["fields"] if "SEQ WPSC_EQ" in code]
    if len(sequence_results) != 12:
        raise AssertionError("M4 equation numbering fields are incomplete")
    if private["bookmarkOrder"][0] != expected_first_bookmark:
        raise AssertionError("M4 equation move did not preserve stable identity at the new position")
    reference_results = [result for code, result in private["fields"] if code.startswith("REF wpsc_eq_")]
    if reference_results != list(expected_reference_results):
        raise AssertionError("M4 equation references did not refresh after movement")
    expected_reference_paragraph = (
        "Formula references remain native: (" + expected_reference_results[0]
        + ") then (" + expected_reference_results[1] + ")."
    )
    if expected_reference_paragraph not in private["paragraphTexts"]:
        raise AssertionError("M4 equation references moved away from their inline source location")
    return mode


def _move_last_equation_before_first(build: Any) -> dict[str, Any]:
    equations = [op for op in build.plan.to_dict()["operations"] if op["op"] == "writer.add_equation"]
    return {
        "type": "move", "kind": "equation",
        "bookmarkName": equations[-1]["args"]["bookmarkName"],
        "ownerNodeId": equations[-1]["nodeId"],
        "beforeBookmarkName": equations[0]["args"]["bookmarkName"],
        "beforeKind": "equation", "beforeOwnerNodeId": equations[0]["nodeId"],
    }


def _injected_source_recovery_operation() -> dict[str, Any]:
    return {
        "op": "writer.add_equation",
        "nodeId": "eq:runtime-source-recovery",
        "args": {
            "renderMode": "native-m4",
            "content": {"nativeMath": {
                "syntax": "wps-linear-v1",
                "linearText": "",
                "sourceHash": "0" * 64,
            }},
            "numbering": {
                "mode": "global", "sequenceId": "WPSC_EQ",
                "chapterStyleLevel": None, "resetLevel": None,
                "prefix": "(", "suffix": ")",
            },
            "bookmarkName": "wpsc_eq_1234567890abcdef12345678",
            "fallbackText": "x+y",
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["EQUATION_INSERT_FAILED"],
            "fallback": "explicit-image-then-source-notice",
        },
    }


def _fatal_engine_probe(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    build: Any,
    output: Path,
    timeout: float,
) -> str:
    if runtime.staging_dir is None:
        raise RuntimeError("WPS staging session is unavailable")
    target = runtime.staging_dir / "m4-engine-fatal.docx"
    runtime._terminate_owned_wps(time.monotonic() + 10.0)
    try:
        command = bridge.issue("writer", "generate_longform_document", {
            "plan": build.plan.to_dict(),
            "outputPath": str(target), "resources": {},
        })
    except ProtocolError:
        command = None
    try:
        result = (
            None if command is None
            else bridge.wait_result(command.id, timeout=min(timeout, 10.0))
        )
    except TimeoutError:
        result = None
    if result is not None or target.exists() or (output / target.name).exists():
        raise AssertionError("engine failure returned a public artifact")
    return "ENGINE_LOST"


def run_longform_m4_evidence(output_dir: Path, timeout: float = 300.0) -> Path:
    output = output_dir.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=False)
    runtime_root = Path(tempfile.mkdtemp(prefix="wpscomposer-longform-m4-"))
    runtime: Optional[ProbeRuntime] = None
    artifacts: list[Path] = []
    screenshots: list[str] = []
    metrics: dict[str, dict[str, int]] = {}
    codes: set[str] = set()
    rounds: list[int] = []
    try:
        with LoopbackBridge(ORIGINS) as bridge:
            runtime = ProbeRuntime(PROBE_ROOT, runtime_root / "runtime", bridge.url, bridge.token)
            with runtime:
                runtime.prepare_profiles()
                runtime.start_servers()
                runtime.activate_component("writer", isolated=True)
                _wait_for_writer_registration(bridge, runtime, timeout=60.0)

                chapter = _fixture_build(FIXTURES, "acceptance_chapter")
                initial = output / "acceptance-initial.docx"
                moved = output / "acceptance-moved.docx"
                initial_outcome = _copy_outcome(chapter, bridge, runtime, initial, timeout)
                codes.update(issue.code for issue in initial_outcome.issues)
                metrics["initial"], initial_private = inspect_m4_docx(initial)
                equations = [
                    op for op in chapter.plan.to_dict()["operations"]
                    if op["op"] == "writer.add_equation"
                ]
                formula_expectations = [
                    {
                        "bookmarkName": operation["args"]["bookmarkName"],
                        "fallbackText": operation["args"]["fallbackText"],
                        "hasImage": "fallbackResourceId" in operation["args"].get("fallbackResource", {}),
                    }
                    for operation in equations
                ]
                initial_mode = _assert_acceptance_content(
                    metrics["initial"],
                    initial_private,
                    formula_expectations=formula_expectations,
                    expected_first_bookmark=equations[0]["args"]["bookmarkName"],
                    expected_reference_results=("1-1", "1-9"),
                )
                metrics["initial"]["citationOrderValid"] = 1
                metrics["initial"]["bibliographyOrderValid"] = 1
                metrics["initial"]["formulaOrderValid"] = 1
                if initial_mode == "native" and metrics["initial"]["editableFormulaCount"] != 12:
                    raise AssertionError("not every accepted formula remained editable")
                if initial_mode == "degraded" and "EQUATION_INSERT_FAILED" not in codes:
                    raise AssertionError("M4 honest formula degradation issue is missing")

                mutation = _mutate_equation(
                    bridge, runtime, initial, moved, _move_last_equation_before_first(chapter), timeout
                )
                rounds.extend(int(item) for item in mutation.get("refreshRounds", []))
                codes.update(str(item) for item in mutation.get("issueCodes", []))
                metrics["moved"], moved_private = inspect_m4_docx(moved)
                moved_expectations = [formula_expectations[-1], *formula_expectations[:-1]]
                moved_mode = _assert_acceptance_content(
                    metrics["moved"],
                    moved_private,
                    formula_expectations=moved_expectations,
                    expected_first_bookmark=equations[-1]["args"]["bookmarkName"],
                    expected_reference_results=("1-2", "1-10"),
                )
                metrics["moved"]["citationOrderValid"] = 1
                metrics["moved"]["bibliographyOrderValid"] = 1
                metrics["moved"]["formulaOrderValid"] = 1
                if moved_mode != initial_mode:
                    raise AssertionError("M4 formula capability mode changed after save and movement")
                if moved_mode == "native" and metrics["moved"]["editableFormulaCount"] != 12:
                    raise AssertionError("moved formula was no longer editable")
                artifacts.extend((initial, moved))

                degradation = _fixture_build(FIXTURES, "degradation")
                degraded = output / "degradation.docx"
                recovered = output / "degradation-runtime-recovered.docx"
                degraded_outcome = _copy_outcome(degradation, bridge, runtime, degraded, timeout)
                codes.update(issue.code for issue in degraded_outcome.issues)
                metrics["degradation"], degradation_private = inspect_m4_docx(degraded)
                visible = "\n".join(degradation_private["paragraphTexts"])
                for code in (
                    "REFERENCE_UNRESOLVED", "FORMULA_MALFORMED",
                    "FORMULA_FALLBACK_IMAGE_UNAVAILABLE", "BIBLIOGRAPHY_ENTRY_MALFORMED",
                    "CONFIG_VALUE_INVALID",
                ):
                    if code not in visible:
                        raise AssertionError("M4 degradation notice is not visible at its declared location")
                if "before [REFERENCE_UNRESOLVED" not in visible or "malformed bibliography declaration" not in visible:
                    raise AssertionError("M4 recoverable content moved away from its source location")
                metrics["degradation"]["noticePlacementValid"] = 1
                if visible.find("CONFIG_VALUE_INVALID") > visible.find("Controlled degradation"):
                    raise AssertionError("M4 document notice is not retained at the quality anchor")
                runtime_recovery = _mutate_equation(
                    bridge, runtime, degraded, recovered,
                    {"type": "insert", "operations": [_injected_source_recovery_operation()]},
                    timeout,
                )
                rounds.extend(int(item) for item in runtime_recovery.get("refreshRounds", []))
                codes.update(str(item) for item in runtime_recovery.get("issueCodes", []))
                metrics["runtimeRecovery"], recovery_private = inspect_m4_docx(recovered)
                recovery_visible = "\n".join(recovery_private["paragraphTexts"])
                if "EQUATION_INSERT_FAILED" not in recovery_visible or "x+y" not in recovery_visible:
                    raise AssertionError("M4 named formula recovery is not visibly anchored")
                metrics["runtimeRecovery"]["noticePlacementValid"] = 1
                artifacts.extend((degraded, recovered))

                for source in (moved, recovered):
                    conversion = _convert_one_to_pdf(source, bridge, runtime, output, timeout)
                    if conversion.get("status") != "passed":
                        raise RuntimeError("M4 PDF export failed")
                    pdf = Path(conversion["pdf"])
                    _assert_pdf(pdf)
                    if source == moved:
                        _assert_formula_pdf(pdf, moved_mode, moved_expectations)
                    artifacts.append(pdf)
                    screenshots.extend(_render_pdf(pdf, output))

                engine_build = _fixture_build(FIXTURES, "acceptance_global")
                codes.add(_fatal_engine_probe(bridge, runtime, engine_build, output, timeout))
    finally:
        restored = runtime is None or runtime.registration_restored
        if restored:
            shutil.rmtree(runtime_root, ignore_errors=True)

    for screenshot in screenshots:
        artifacts.append(output / screenshot)
    if not rounds:
        raise AssertionError("M4 field refresh evidence is incomplete")
    report = {
        "version": "M4",
        "wpsVersion": read_wps_version(),
        "artifacts": [
            {"name": str(path.relative_to(output)), "sha256": _sha256(path)}
            for path in artifacts
        ],
        "metrics": metrics,
        "codes": sorted(codes),
        "rounds": rounds,
        "screenshots": screenshots,
    }
    validate_m4_evidence_report(report)
    report_path = output / "platform-evidence.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2) + "\n", encoding="utf-8")
    return report_path


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("output", nargs="?")
    parser.add_argument("--timeout", type=float, default=300.0)
    args = parser.parse_args()
    target = Path(args.output) if args.output else ROOT / "build" / "longform-m4" / (
        "macos-native-" + datetime.now().strftime("%Y%m%d-%H%M%S")
    )
    print(run_longform_m4_evidence(target, args.timeout))
