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
from pathlib import Path
from typing import Any, Mapping, Optional

from .longform.macos_executor import MacOSLongformExecutor
from .longform.pipeline import (
    _is_absolute_path,
    build_longform_generation,
    execute_longform_plan,
)
from .longform_m0.host_checks import validate_evidence_privacy
from .macos_probe.bridge import LoopbackBridge
from .macos_probe.longform_evidence import (
    ORIGINS,
    _convert_one_to_pdf,
    _wait_for_writer_registration,
)
from .macos_probe.runtime import ProbeRuntime, read_wps_version


ROOT = Path(__file__).resolve().parents[3]
FIXTURES = ROOT / "tests" / "longform_m3" / "fixtures"
PROBE_ROOT = ROOT / "macos" / "wps-jsapi-probe"
_FIELD_RE = re.compile(
    r"<w:fldChar[^>]+w:fldCharType=\"begin\"[^>]*/?>"
    r"(?P<body>.*?)<w:fldChar[^>]+w:fldCharType=\"end\"[^>]*/?>",
    re.DOTALL,
)
_INSTR_RE = re.compile(r"<w:instrText[^>]*>(.*?)</w:instrText>", re.DOTALL)
_TEXT_RE = re.compile(r"<w:t[^>]*>(.*?)</w:t>", re.DOTALL)
_BOOKMARK_RE = re.compile(r'<w:bookmarkStart[^>]+w:id="([^"]+)"[^>]+w:name="([^"]+)"[^>]*/?>')
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_FIELD_ERROR_MARKERS = (
    "错误",
    "未找到",
    "未定义书签",
    "关联章节样式",
    "未自动编号",
    "HEADING_PREFIX_AMBIGUOUS",
    "Error",
    "undefined bookmark",
)
_COUNT_KEYS = {
    "fieldCount", "figureSequenceCount", "tableSequenceCount",
    "equationSequenceCount", "referenceCount", "indexCount",
    "bookmarkCount", "imageCount", "tableCount", "mergeCount",
    "columnContainerCount", "columnGeometryCount",
    "borderlessColumnContainerCount",
}
_M3_PRIVATE_KEYS = {
    "bookmarknames", "fieldresults", "payload", "resourcehash", "paths",
}


def _reject_private_evidence_fields(value: Any, location: str = "evidence") -> None:
    if isinstance(value, Mapping):
        for key, item in value.items():
            normalized = re.sub(r"[^a-z0-9]", "", str(key).casefold())
            if normalized in _M3_PRIVATE_KEYS:
                raise ValueError(f"M3 evidence private field at {location}.{key}")
            _reject_private_evidence_fields(item, f"{location}.{key}")
    elif isinstance(value, (list, tuple)):
        for index, item in enumerate(value):
            _reject_private_evidence_fields(item, f"{location}[{index}]")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _document_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        return package.read("word/document.xml").decode("utf-8")


def inspect_native_docx(path: Path) -> dict[str, Any]:
    """Return internal structural evidence; callers must not serialize text."""
    xml = _document_xml(path)
    root = ET.fromstring(xml)
    fields: list[tuple[str, str]] = []
    stack: list[dict[str, Any]] = []
    for element in root.iter():
        if element.tag == _W + "fldChar":
            kind = element.get(_W + "fldCharType")
            if kind == "begin":
                stack.append({"instruction": [], "result": [], "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                field = stack.pop()
                fields.append((
                    "".join(field["instruction"]).strip(),
                    "".join(field["result"]).strip(),
                ))
        elif element.tag == _W + "instrText" and stack:
            stack[-1]["instruction"].append(element.text or "")
        elif element.tag == _W + "t" and stack:
            for field in stack:
                if field["separated"]:
                    field["result"].append(element.text or "")
    paragraph_texts = [
        "".join(node.text or "" for node in paragraph.iter(_W + "t")).strip()
        for paragraph in root.iter(_W + "p")
    ]
    paragraph_texts = [text for text in paragraph_texts if text]
    column_containers = []
    for table in root.iter(_W + "tbl"):
        image_nodes = list(table.iter(_W + "drawing")) + list(table.iter(_W + "pict"))
        grid_columns = [
            int(column.get(_W + "w", "0"))
            for column in table.iter(_W + "gridCol")
        ]
        if len(image_nodes) < 2:
            continue
        border_values = {
            border.tag.rsplit("}", 1)[-1]: border.get(_W + "val", "")
            for borders in table.iter(_W + "tblBorders")
            for border in borders
        }
        column_containers.append({
            "grid": grid_columns,
            "borderless": all(
                border_values.get(name) in {"none", "nil"}
                for name in ("top", "left", "bottom", "right", "insideH", "insideV")
            ),
        })
    bookmark_names = [name for _, name in _BOOKMARK_RE.findall(xml)]
    return {
        "fieldCount": len(fields),
        "figureSequenceCount": sum("SEQ WPSC_FIG" in code for code, _ in fields),
        "tableSequenceCount": sum("SEQ WPSC_TAB" in code for code, _ in fields),
        "equationSequenceCount": sum("SEQ WPSC_EQ" in code for code, _ in fields),
        "referenceCount": sum(code.startswith("REF wpsc_") for code, _ in fields),
        "indexCount": sum("TOC" in code and "WPSC_" in code for code, _ in fields),
        "bookmarkCount": len([name for name in bookmark_names if name.startswith("wpsc_")]),
        "imageCount": xml.count("<w:drawing") + xml.count("<w:pict"),
        "tableCount": xml.count("<w:tbl>"),
        "mergeCount": xml.count("<w:vMerge") + xml.count("<w:gridSpan"),
        "columnContainerCount": len(column_containers),
        "columnGeometryCount": sum(
            len(item["grid"]) == 3 and item["grid"][1] == 240
            for item in column_containers
        ),
        "borderlessColumnContainerCount": sum(
            item["borderless"] for item in column_containers
        ),
        "fieldResults": fields,
        "bookmarkNames": bookmark_names,
        "paragraphTexts": paragraph_texts,
        "xml": xml,
    }


def _assert_native_document(path: Path, *, mode: str, mutated: bool) -> dict[str, int]:
    inspection = inspect_native_docx(path)
    expected = (
        {"figureSequenceCount": 2, "tableSequenceCount": 3, "equationSequenceCount": 1}
        if mode == "chapter"
        else {"figureSequenceCount": 2, "tableSequenceCount": 3, "equationSequenceCount": 1}
    )
    for key, count in expected.items():
        if inspection[key] != count:
            raise AssertionError(f"{mode} {key} expected {count}, got {inspection[key]}")
    if inspection["indexCount"] != 2:
        raise AssertionError(f"{mode} native figure/table indexes were not retained")
    if inspection["referenceCount"] < 3:
        raise AssertionError(f"{mode} native references were not retained")
    visible_text = "\n".join(inspection["paragraphTexts"])
    if any(
        marker in instruction or marker in result or marker in visible_text
        for instruction, result in inspection["fieldResults"]
        for marker in _FIELD_ERROR_MARKERS
    ):
        raise AssertionError(f"{mode} contains an unresolved native field result")
    expected_sequence_results = {
        "WPSC_FIG": ["1", "2"],
        "WPSC_TAB": ["1", "2", "3"],
        "WPSC_EQ": ["1"],
    }
    for sequence_id, expected_results in expected_sequence_results.items():
        actual_results = [
            result
            for instruction, result in inspection["fieldResults"]
            if instruction.startswith("SEQ " + sequence_id)
        ]
        if actual_results != expected_results:
            raise AssertionError(f"{mode} {sequence_id} results are not exact and gap-free")
    if inspection["bookmarkCount"] < sum(expected.values()):
        raise AssertionError(f"{mode} native target bookmarks were not retained")
    if mutated and inspection["mergeCount"] < 1 and mode == "chapter":
        raise AssertionError("chapter table merge was not retained after mutation")
    expected_columns = 1 if mode == "chapter" else 2
    if not all(
        inspection[key] == expected_columns
        for key in (
            "columnContainerCount",
            "columnGeometryCount",
            "borderlessColumnContainerCount",
        )
    ):
        raise AssertionError(f"{mode} two-column geometry or borderless container was not retained")
    required_labels = (
        ("图 1-1", "表 1-1", "(1-1)")
        if mode == "chapter"
        else ("图 1", "表 1", "(1)")
    )
    if any(label not in visible_text for label in required_labels):
        raise AssertionError(f"{mode} exact visible labels were not retained")
    if mutated:
        expected_reference = (
            "See 图 1-2, 表 1-2, and (1-1)."
            if mode == "chapter"
            else "Global references: 图 2, 表 2, (1)."
        )
        compact_text = re.sub(r"\s+", " ", visible_text)
        if expected_reference not in compact_text:
            raise AssertionError(f"{mode} exact native REF result was not retained")
        if mode == "chapter" and not all(
            text in visible_text
            for text in ("Unnumbered appendix note", "表 1-3 Inserted chapter table")
        ):
            raise AssertionError("unnumbered H1 reset the active chapter sequence")
    index_orders = {
        "chapter": {
            False: {
                "WPSC_FIG": ["Chapter diagram", "Two-column comparison"],
                "WPSC_TAB": ["Three-line grouped data", "Grid data", "Disposable chapter grid"],
            },
            True: {
                "WPSC_FIG": ["Two-column comparison", "Chapter diagram"],
                "WPSC_TAB": ["Grid data", "Three-line grouped data", "Inserted chapter table"],
            },
        },
        "global": {
            False: {
                "WPSC_FIG": ["Raster normalizers", "Static native formats"],
                "WPSC_TAB": ["Global grid", "Disposable global grid", "Second disposable global grid"],
            },
            True: {
                "WPSC_FIG": ["Static native formats", "Raster normalizers"],
                "WPSC_TAB": ["Disposable global grid", "Global grid", "Inserted global table"],
            },
        },
    }
    for sequence_id, captions in index_orders[mode][mutated].items():
        index_results = [
            result
            for instruction, result in inspection["fieldResults"]
            if instruction.startswith("TOC") and sequence_id in instruction
        ]
        if len(index_results) != 1:
            raise AssertionError(f"{mode} {sequence_id} index result is missing")
        positions = [index_results[0].find(caption) for caption in captions]
        if any(position < 0 for position in positions) or positions != sorted(positions):
            raise AssertionError(f"{mode} {sequence_id} index order is not exact")
    xml = inspection["xml"]
    for instruction, result in inspection["fieldResults"]:
        if instruction.startswith("TOC") and any(
            title in result for title in ("图目录", "表目录")
        ):
            raise AssertionError("caption index title leaked into TOC results")
    body = re.search(r"<w:body>(.*)</w:body>", xml, re.DOTALL)
    if body and not re.search(r"(?:</w:p>|</w:tbl>)\s*<w:sectPr", body.group(1)):
        raise AssertionError("document ended with an empty trailing section")
    return {key: int(inspection[key]) for key in (
        "fieldCount", "figureSequenceCount", "tableSequenceCount",
        "equationSequenceCount", "referenceCount", "indexCount",
        "bookmarkCount", "imageCount", "tableCount", "mergeCount",
        "columnContainerCount", "columnGeometryCount", "borderlessColumnContainerCount",
    )}


def _assert_stable_bookmarks(
    initial: Mapping[str, Any],
    mutated: Mapping[str, Any],
    *,
    deleted: str,
    inserted: str,
) -> None:
    initial_names = {
        str(name) for name in initial["bookmarkNames"] if str(name).startswith("wpsc_")
    }
    mutated_names = {
        str(name) for name in mutated["bookmarkNames"] if str(name).startswith("wpsc_")
    }
    expected = (initial_names - {deleted}) | {inserted}
    if mutated_names != expected:
        raise AssertionError("native bookmark identity changed after mutation")


def _fixture_build(name: str):
    source = FIXTURES / f"{name}.md"
    return build_longform_generation(
        source.read_text(encoding="utf-8"), base_dir=str(FIXTURES)
    )


def _native_ops(build, name: str) -> list[dict[str, Any]]:
    return [
        operation
        for operation in build.plan.to_dict()["operations"]
        if operation["op"] == name
    ]


def _inserted_table(source: Mapping[str, Any], label: str) -> dict[str, Any]:
    operation = json.loads(json.dumps(source, ensure_ascii=False))
    operation["nodeId"] = f"tab:inserted-{label}"
    operation["args"]["caption"] = f"Inserted {label} table"
    operation["args"]["bookmarkName"] = "wpsc_tab_" + hashlib.sha256(
        f"m3:{label}:inserted-table".encode("ascii")
    ).hexdigest()[:24]
    operation["args"]["headers"] = ["Mutation", "Result"]
    operation["args"]["rows"] = [["insert", "passed"]]
    operation["args"]["alignments"] = ["left", "center"]
    operation["args"]["merges"] = []
    operation["args"]["plannedDegradation"] = []
    return operation


def _inserted_heading(label: str) -> dict[str, Any]:
    return {
        "op": "writer.add_heading",
        "nodeId": f"heading:inserted-{label}",
        "args": {
            "text": "Unnumbered appendix note",
            "level": 1,
            "numbering": False,
            "numberingScheme": "decimal",
            "sequenceTransparent": True,
            "keepTogether": -1,
            "keepWithNext": -1,
            "spaceBefore": 12.0,
            "spaceAfter": 6.0,
        },
    }


def _mutations(build, label: str) -> list[dict[str, Any]]:
    figures = _native_ops(build, "writer.add_captioned_figure")
    tables = _native_ops(build, "writer.add_semantic_table")
    if len(figures) < 2 or not tables:
        raise AssertionError("mutation fixture lacks native targets")
    delete_table = tables[-1]
    return [
        {
            "type": "move",
            "kind": "figure",
            "bookmarkName": figures[1]["args"]["bookmarkName"],
            "ownerNodeId": figures[1]["nodeId"],
            "beforeBookmarkName": figures[0]["args"]["bookmarkName"],
            "beforeKind": "figure",
            "beforeOwnerNodeId": figures[0]["nodeId"],
        },
        {
            "type": "move",
            "kind": "table",
            "bookmarkName": tables[1]["args"]["bookmarkName"],
            "ownerNodeId": tables[1]["nodeId"],
            "beforeBookmarkName": tables[0]["args"]["bookmarkName"],
            "beforeKind": "table",
            "beforeOwnerNodeId": tables[0]["nodeId"],
        },
        {
            "type": "insert",
            "operations": [
                _inserted_heading(label),
                _inserted_table(tables[0], label),
            ],
        },
        {
            "type": "delete",
            "kind": "table",
            "bookmarkName": delete_table["args"]["bookmarkName"],
            "ownerNodeId": delete_table["nodeId"],
        },
    ]


def _generate(
    build,
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    output: Path,
    timeout: float,
) -> Any:
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(runtime.staging_dir))
    outcome = execute_longform_plan(
        build, executor, deadline=time.monotonic() + timeout
    )
    if outcome.issues:
        raise AssertionError("native generation used a degradation or emitted a runtime issue")
    shutil.copy2(outcome.staged_artifact, output)
    return outcome


def _mutate(
    bridge: LoopbackBridge,
    runtime: ProbeRuntime,
    source: Path,
    target: Path,
    mutations: list[dict[str, Any]],
    timeout: float,
) -> dict[str, Any]:
    if runtime.staging_dir is None:
        raise RuntimeError("WPS staging session is unavailable")
    staged_source = runtime.staging_dir / f"m3-{source.stem}-source.docx"
    staged_target = runtime.staging_dir / f"m3-{target.stem}-output.docx"
    shutil.copy2(source, staged_source)
    command = bridge.issue(
        "writer",
        "mutate_longform_document",
        {
            "sourcePath": str(staged_source),
            "outputPath": str(staged_target),
            "mutations": mutations,
            "resources": {},
        },
    )
    result = bridge.wait_result(command.id, timeout=timeout)
    if not result.ok:
        code = (result.error or {}).get("code", "MUTATION_FAILED")
        raise RuntimeError(f"native mutation failed: {code}")
    value = dict(result.value)
    if value.get("mutationKinds") != ["move", "insert", "delete"]:
        raise AssertionError("native mutation sequence was incomplete")
    if value.get("issueCodes"):
        raise AssertionError("native mutation produced field or layout issues")
    if not staged_target.is_file():
        raise RuntimeError("native mutation output was not created")
    shutil.copy2(staged_target, target)
    return value


def _render_pdf(pdf: Path, output: Path) -> list[str]:
    screenshot_dir = output / "screenshots"
    screenshot_dir.mkdir(exist_ok=True)
    prefix = screenshot_dir / pdf.stem
    subprocess.run(
        ["pdftoppm", "-png", "-r", "120", str(pdf), str(prefix)],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
        text=True,
    )
    screenshots = sorted(screenshot_dir.glob(f"{pdf.stem}-*.png"))
    if not screenshots or any(path.stat().st_size < 1024 for path in screenshots):
        raise AssertionError("PDF representative rendering was empty")
    return [str(path.relative_to(output)) for path in screenshots]


def _validate_pdf_pages(pages: list[dict[str, Any]], *, mode: str) -> None:
    if not pages:
        raise AssertionError("PDF has no pages")
    full_text = "\n".join(str(page.get("text") or "") for page in pages)
    if any(marker in full_text for marker in _FIELD_ERROR_MARKERS):
        raise AssertionError(f"{mode} PDF contains a visible field or diagnostic error")
    for page in pages:
        images = page.get("images") or []
        if not str(page.get("text") or "").strip() and not images:
            raise AssertionError(f"{mode} PDF contains a blank page")
        for image in images:
            if (
                float(image["x0"]) < -0.5
                or float(image["top"]) < -0.5
                or float(image["x1"]) > float(page["width"]) + 0.5
                or float(image["bottom"]) > float(page["height"]) + 0.5
            ):
                raise AssertionError(f"{mode} PDF contains a clipped image")
    if mode == "chapter":
        cover = re.sub(r"\s+", "", str(pages[0].get("text") or ""))
        if "M3chapterevidence" not in cover or cover.isdigit():
            raise AssertionError("chapter cover page is empty or page-number-only")
        for word in pages[0].get("words") or []:
            if (
                re.fullmatch(r"[ivxlcdm]+|\d+", str(word.get("text") or ""), re.IGNORECASE)
                and float(word.get("top", 0)) > float(pages[0]["height"]) * 0.85
            ):
                raise AssertionError("chapter cover page exposes a page number")

    horizontal_pairs = 0
    for page in pages:
        images = sorted(page.get("images") or [], key=lambda item: (float(item["top"]), float(item["x0"])))
        used: set[int] = set()
        for left_index, left in enumerate(images):
            if left_index in used:
                continue
            for right_index in range(left_index + 1, len(images)):
                if right_index in used:
                    continue
                right = images[right_index]
                if abs(float(left["top"]) - float(right["top"])) > 12.0:
                    continue
                first, second = sorted((left, right), key=lambda item: float(item["x0"]))
                gap = float(second["x0"]) - float(first["x1"])
                if 8.0 <= gap <= 16.0:
                    horizontal_pairs += 1
                    used.update((left_index, right_index))
                    break
    expected_pairs = 1 if mode == "chapter" else 2
    if horizontal_pairs < expected_pairs:
        raise AssertionError(f"{mode} PDF did not retain side-by-side 12pt figure geometry")

    compact_pages = [re.sub(r"\s+", "", str(page.get("text") or "")) for page in pages]
    required_cohesion = (
        (("Two-columncomparison", "图1-1"), ("Three-linegroupeddata", "GroupMetricValue"))
        if mode == "chapter"
        else (("Staticnativeformats", "图1"), ("Globalgrid", "FormatNative"))
    )
    for first, second in required_cohesion:
        if not any(first in text and second in text for text in compact_pages):
            raise AssertionError(f"{mode} PDF caption/object page cohesion failed")


def _assert_pdf_layout(pdf: Path, *, mode: str) -> None:
    import pdfplumber

    pages: list[dict[str, Any]] = []
    with pdfplumber.open(pdf) as document:
        for page in document.pages:
            pages.append({
                "width": float(page.width),
                "height": float(page.height),
                "text": page.extract_text() or "",
                "words": page.extract_words(),
                "images": [
                    {
                        "x0": float(item["x0"]),
                        "x1": float(item["x1"]),
                        "top": float(item["top"]),
                        "bottom": float(item["bottom"]),
                    }
                    for item in page.images
                ],
            })
    _validate_pdf_pages(pages, mode=mode)


def validate_m3_evidence_report(value: Any) -> None:
    _reject_private_evidence_fields(value)
    if not isinstance(value, dict) or set(value) != {
        "status", "wpsVersion", "capabilities", "artifacts", "counts",
        "refreshRounds", "mutationKinds", "screenshots",
    }:
        raise ValueError("M3 evidence report has an invalid root shape")
    if value["status"] != "passed" or not isinstance(value["wpsVersion"], str):
        raise ValueError("M3 evidence report is not passing")
    if set(value["capabilities"]) != {"chapter", "global", "mutation", "pdf", "visual"}:
        raise ValueError("M3 evidence capabilities are incomplete")
    if any(status != "passed" for status in value["capabilities"].values()):
        raise ValueError("M3 evidence capability failed")
    for artifact in value["artifacts"]:
        if set(artifact) != {"name", "sha256", "kind"}:
            raise ValueError("M3 evidence artifact shape is invalid")
        name = Path(artifact["name"])
        if _is_absolute_path(artifact["name"]) or ".." in name.parts:
            raise ValueError("M3 evidence artifact name is not relative")
        if not re.fullmatch(r"[0-9a-f]{64}", artifact["sha256"]):
            raise ValueError("M3 evidence artifact digest is invalid")
        if artifact["kind"] not in {"docx", "pdf", "png"}:
            raise ValueError("M3 evidence artifact kind is invalid")
    if set(value["counts"]) != {
        "chapterInitial", "chapterMutated", "globalInitial", "globalMutated",
    }:
        raise ValueError("M3 evidence counts are incomplete")
    for count in value["counts"].values():
        if not isinstance(count, dict) or set(count) != _COUNT_KEYS:
            raise ValueError("M3 evidence count shape is invalid")
        if any(type(item) is not int or item < 0 for item in count.values()):
            raise ValueError("M3 evidence count value is invalid")
    if (
        not isinstance(value["refreshRounds"], list)
        or len(value["refreshRounds"]) != 10
        or any(type(item) is not int or item < 1 or item > 3 for item in value["refreshRounds"])
    ):
        raise ValueError("M3 evidence refresh rounds are invalid")
    if value["mutationKinds"] != ["move", "insert", "delete"]:
        raise ValueError("M3 evidence mutation sequence is incomplete")
    if not isinstance(value["screenshots"], list) or not value["screenshots"]:
        raise ValueError("M3 evidence screenshots are incomplete")
    if len(set(value["screenshots"])) != len(value["screenshots"]):
        raise ValueError("M3 evidence screenshot names are duplicated")
    for raw_name in value["screenshots"]:
        if not isinstance(raw_name, str):
            raise ValueError("M3 evidence screenshot name is invalid")
        name = Path(raw_name)
        if (
            _is_absolute_path(raw_name) or ".." in name.parts or len(name.parts) != 2
            or name.parts[0] != "screenshots" or name.suffix.lower() != ".png"
        ):
            raise ValueError("M3 evidence screenshot name is not relative")
    validate_evidence_privacy(value)


def run_longform_m3_evidence(output_dir: Path, timeout: float = 300.0) -> Path:
    output = output_dir.expanduser().resolve()
    output.mkdir(parents=True, exist_ok=False)
    runtime_root = Path(tempfile.mkdtemp(prefix="wpscomposer-longform-m3-"))
    runtime: Optional[ProbeRuntime] = None
    artifacts: list[Path] = []
    counts: dict[str, Mapping[str, int]] = {}
    refresh_rounds: list[int] = []
    screenshots: list[str] = []
    try:
        with LoopbackBridge(ORIGINS) as bridge:
            runtime = ProbeRuntime(
                PROBE_ROOT, runtime_root / "runtime", bridge.url, bridge.token
            )
            with runtime:
                runtime.prepare_profiles()
                runtime.start_servers()
                # An isolated instance guarantees the newly installed add-in
                # registration is loaded without disturbing an already-open
                # user WPS process. ProbeRuntime records and cleans only the
                # newly observed process identity.
                runtime.activate_component("writer", isolated=True)
                _wait_for_writer_registration(bridge, runtime, timeout=60.0)

                builds = {
                    "chapter": _fixture_build("chapter_native"),
                    "global": _fixture_build("global_media"),
                }
                for label, build in builds.items():
                    initial = output / f"{label}-initial.docx"
                    mutated = output / f"{label}-mutated.docx"
                    _generate(build, bridge, runtime, initial, timeout)
                    initial_inspection = inspect_native_docx(initial)
                    counts[f"{label}Initial"] = _assert_native_document(
                        initial, mode=label, mutated=False
                    )
                    mutations = _mutations(build, label)
                    mutation = _mutate(
                        bridge, runtime, initial, mutated,
                        mutations, timeout
                    )
                    refresh_rounds.extend(int(item) for item in mutation["refreshRounds"])
                    mutated_inspection = inspect_native_docx(mutated)
                    counts[f"{label}Mutated"] = _assert_native_document(
                        mutated, mode=label, mutated=True
                    )
                    _assert_stable_bookmarks(
                        initial_inspection,
                        mutated_inspection,
                        deleted=mutations[3]["bookmarkName"],
                        inserted=mutations[2]["operations"][-1]["args"]["bookmarkName"],
                    )
                    artifacts.extend((initial, mutated))

                    conversion = _convert_one_to_pdf(
                        mutated, bridge, runtime, output, timeout
                    )
                    if conversion.get("status") != "passed":
                        raise RuntimeError(f"{label} PDF conversion failed")
                    pdf = Path(conversion["pdf"])
                    _assert_pdf_layout(pdf, mode=label)
                    artifacts.append(pdf)
                    screenshots.extend(_render_pdf(pdf, output))
    finally:
        restored = runtime is None or runtime.registration_restored
        if restored:
            shutil.rmtree(runtime_root, ignore_errors=True)

    for screenshot in screenshots:
        artifacts.append(output / screenshot)
    report = {
        "status": "passed",
        "wpsVersion": read_wps_version(),
        "capabilities": {
            "chapter": "passed",
            "global": "passed",
            "mutation": "passed",
            "pdf": "passed",
            "visual": "passed",
        },
        "artifacts": [
            {
                "name": str(path.relative_to(output)),
                "sha256": _sha256(path),
                "kind": path.suffix.lstrip(".").lower(),
            }
            for path in artifacts
        ],
        "counts": counts,
        "refreshRounds": refresh_rounds,
        "mutationKinds": ["move", "insert", "delete"],
        "screenshots": screenshots,
    }
    validate_m3_evidence_report(report)
    report_path = output / "platform-evidence.json"
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, sort_keys=True, indent=2) + "\n",
        encoding="utf-8",
    )
    return report_path


if __name__ == "__main__":
    import argparse
    from datetime import datetime

    parser = argparse.ArgumentParser()
    parser.add_argument("output", nargs="?")
    parser.add_argument("--timeout", type=float, default=300.0)
    args = parser.parse_args()
    target = Path(args.output) if args.output else Path(
        "build/longform-m3/macos-native-" + datetime.now().strftime("%Y%m%d-%H%M%S")
    )
    print(run_longform_m3_evidence(target, timeout=args.timeout))
