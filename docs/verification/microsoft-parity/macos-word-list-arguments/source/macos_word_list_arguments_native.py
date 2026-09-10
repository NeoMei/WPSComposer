"""Opt-in native Mac Word list-argument acceptance.

This file never contacts Word on import. Native execution requires --execute,
an existing source-freeze manifest, and a new output directory.
"""
from __future__ import annotations

import argparse
from dataclasses import replace
import hashlib
import json
import math
from pathlib import Path
import shutil
import sys
import time
import traceback
from zipfile import ZipFile
import xml.etree.ElementTree as ET


ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation
from skills.WPSComposer.scripts.generation_plan import GenerationOperation
from skills.WPSComposer.scripts.msoffice.macos_runtime import MacWordAdapter
from skills.WPSComposer.scripts.msoffice.macos_script import compile_plan
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession


OWNED_SOURCES = (
    Path(__file__).resolve(),
    ROOT / "skills/WPSComposer/scripts/msoffice/macos_script.py",
    ROOT / "skills/WPSComposer/scripts/longform/windows_executor.py",
    ROOT / "tests/msoffice/test_word_list_argument_parity.py",
)
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def build_case():
    build = build_longform_generation(
        "# LIST ARGUMENT NATIVE ACCEPTANCE\n\n"
        "BASE BODY\n\n- seed\n\nBODY AFTER DEFAULT\n\nBODY AFTER ORDERED"
    )
    operations = []
    replaced = False
    for operation in build.plan.operations:
        if operation.op == "writer.add_paragraph" and operation.args.get("text") in {
            "BODY AFTER DEFAULT", "BODY AFTER ORDERED"
        }:
            continue
        if operation.op == "writer.add_paragraph" and operation.args.get("text") == "BASE BODY":
            operations.append(replace(operation, args={
                "text": "BASE BODY", "style": "Body Text", "align": 2,
                "spans": [{"text": "BASE BODY", "italic": True, "strikethrough": True}],
            }))
            continue
        if operation.op == "writer.add_list":
            replaced = True
            cases = (
                {"items": ["DEFAULT ITEM"], "ordered": False},
                {"items": ["CUSTOM ITEM"], "ordered": False, "glyph": "→", "indent": 30},
                {"items": ["EMPTY GLYPH ITEM"], "ordered": False, "glyph": "", "indent": 27.5},
                {"items": ["ORDERED ITEM"], "ordered": True, "glyph": 'ignored"\\unsafe', "indent": 33},
            )
            operations.append(replace(operation, args=cases[0], node_id="list-case:1"))
            operations.append(GenerationOperation(
                op="writer.add_paragraph", args={"text": "BODY AFTER DEFAULT", "style": "Body Text"},
                node_id="body-after-default",
            ))
            for index, args in enumerate(cases[1:], 2):
                operations.append(replace(operation, args=args, node_id=f"list-case:{index}"))
            operations.append(GenerationOperation(
                op="writer.add_paragraph", args={"text": "BODY AFTER ORDERED", "style": "Body Text"},
                node_id="body-after-ordered",
            ))
            continue
        operations.append(operation)
    if not replaced:
        raise AssertionError("seed plan did not contain writer.add_list")
    return replace(build, plan=replace(build.plan, operations=tuple(operations)))


def compile_case(target: Path):
    return compile_plan(build_case().plan, {}, target, timeout=120)


def verify_source_freeze(path: Path) -> dict[str, str]:
    data = json.loads(path.resolve(strict=True).read_text(encoding="utf-8"))
    expected = {source.relative_to(ROOT).as_posix(): sha256(source) for source in OWNED_SOURCES}
    if data.get("schema") != 1 or data.get("source_hashes") != expected:
        raise RuntimeError("source freeze does not match current implementation")
    return expected


def _paragraphs(path: Path):
    with ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        styles_root = ET.fromstring(package.read("word/styles.xml"))
    names = {}
    style_props = {}
    for style in styles_root.findall(W + "style"):
        name = style.find(W + "name")
        if name is not None:
            names[style.get(W + "styleId")] = name.get(W + "val")
        based_on = style.find(W + "basedOn")
        ppr = style.find(W + "pPr")
        rpr = style.find(W + "rPr")
        ind = None if ppr is None else ppr.find(W + "ind")
        fonts = None if rpr is None else rpr.find(W + "rFonts")
        size = None if rpr is None else rpr.find(W + "sz")
        style_props[style.get(W + "styleId")] = {
            "basedOn": None if based_on is None else based_on.get(W + "val"),
            "firstLine": None if ind is None else ind.get(W + "firstLine"),
            "eastAsia": None if fonts is None else fonts.get(W + "eastAsia"),
            "ascii": None if fonts is None else fonts.get(W + "ascii"),
            "size": None if size is None else size.get(W + "val"),
        }

    def inherited(style_id, key):
        seen = set()
        while style_id and style_id not in seen:
            seen.add(style_id)
            props = style_props.get(style_id, {})
            if props.get(key) is not None:
                return props[key]
            style_id = props.get("basedOn")
        return None

    def text_content(paragraph):
        parts = []
        for run in paragraph.findall(".//" + W + "r"):
            for node in run.iter():
                if node.tag == W + "t":
                    parts.append(node.text or "")
                elif node.tag == W + "tab":
                    parts.append("\t")
        return "".join(parts)
    rows = []
    for paragraph in root.findall("./" + W + "body/" + W + "p"):
        ppr = paragraph.find(W + "pPr")
        ind = None if ppr is None else ppr.find(W + "ind")
        tabs = [] if ppr is None else ppr.findall(W + "tabs/" + W + "tab")
        style = None if ppr is None else ppr.find(W + "pStyle")
        style_id = None if style is None else style.get(W + "val")
        run_props = paragraph.find(W + "r/" + W + "rPr")
        if run_props is None and ppr is not None:
            run_props = ppr.find(W + "rPr")
        run_fonts = None if run_props is None else run_props.find(W + "rFonts")
        run_size = None if run_props is None else run_props.find(W + "sz")
        italic = None if run_props is None else run_props.find(W + "i")
        strike = None if run_props is None else run_props.find(W + "strike")
        align = None if ppr is None else ppr.find(W + "jc")
        text = text_content(paragraph)
        rows.append(
            {
                "text": text,
                "style": names.get(style_id, style_id),
                "left": None if ind is None else ind.get(W + "left"),
                "firstLine": (None if ind is None else ind.get(W + "firstLine")) or inherited(style_id, "firstLine"),
                "hanging": None if ind is None else ind.get(W + "hanging"),
                "tabs": [tab.get(W + "pos") for tab in tabs],
                "eastAsia": (None if run_fonts is None else run_fonts.get(W + "eastAsia")) or inherited(style_id, "eastAsia"),
                "ascii": (None if run_fonts is None else run_fonts.get(W + "ascii")) or inherited(style_id, "ascii"),
                "size": (None if run_size is None else run_size.get(W + "val")) or inherited(style_id, "size"),
                "italic": italic is not None and italic.get(W + "val", "1") not in ("0", "false"),
                "strike": strike is not None and strike.get(W + "val", "1") not in ("0", "false"),
                "align": None if align is None else align.get(W + "val"),
            }
        )
    sect = root.find("./" + W + "body/" + W + "sectPr")
    size = None if sect is None else sect.find(W + "pgSz")
    margins = None if sect is None else sect.find(W + "pgMar")
    page = {
        "width": None if size is None else size.get(W + "w"),
        "height": None if size is None else size.get(W + "h"),
        "top": None if margins is None else margins.get(W + "top"),
        "bottom": None if margins is None else margins.get(W + "bottom"),
        "left": None if margins is None else margins.get(W + "left"),
        "right": None if margins is None else margins.get(W + "right"),
    }
    return rows, page


def inspect_docx(path: Path) -> dict[str, bool]:
    rows, page = _paragraphs(path)
    by_text = {row["text"]: row for row in rows}
    cases = {
        "•\tDEFAULT ITEM": 24,
        "→\tCUSTOM ITEM": 30,
        "\tEMPTY GLYPH ITEM": 27.5,
        "1.\tORDERED ITEM": 33,
    }
    checks = {}
    checks["exact_literal_prefixes"] = all(text in by_text for text in cases)
    for text, indent in cases.items():
        row = by_text.get(text, {})
        twips = str(round(indent * 20))
        checks["indent_" + hashlib.sha256(text.encode()).hexdigest()[:8]] = (
            row.get("style") == "List Paragraph"
            and row.get("left") == twips
            and row.get("hanging") == twips
            and twips in row.get("tabs", [])
            and row.get("eastAsia") == "仿宋"
            and row.get("ascii") == "Times New Roman"
            and row.get("size") == "24"
            and not row.get("italic") and not row.get("strike")
            and row.get("align") in (None, "left")
        )
    for text in ("BASE BODY", "BODY AFTER DEFAULT", "BODY AFTER ORDERED"):
        row = by_text.get(text, {})
        checks["body_reset_" + text.lower().replace(" ", "_")] = (
            row.get("style") == "Body Text"
            and row.get("firstLine") == "480"
            and row.get("left") in (None, "0")
            and row.get("hanging") in (None, "0")
        )
    checks["page_setup_preserved"] = (
        page["width"] is not None and abs(int(page["width"]) - 11906) <= 2
        and page["height"] is not None and abs(int(page["height"]) - 16838) <= 2
        and page["top"] == page["bottom"] == "1440"
        and page["left"] is not None and abs(int(page["left"]) - 1701) <= 2
        and page["right"] is not None and abs(int(page["right"]) - 1417) <= 2
    )
    return checks


def inspect_pdf(path: Path) -> dict[str, bool]:
    import fitz
    with fitz.open(path) as document:
        text = "\n".join(page.get_text() for page in document)
        pages = len(document)
    return {
        "pdf_is_readable": pages >= 1,
        "pdf_list_and_body_text": all(token in text for token in (
            "DEFAULT ITEM", "CUSTOM ITEM", "EMPTY GLYPH ITEM", "ORDERED ITEM",
            "BODY AFTER DEFAULT", "BODY AFTER ORDERED",
        )),
    }


def _readback(session: MacWordSession):
    rows = session._execute(
        [
            "set nativeRows to {}",
            "set expectedBodyStyleName to name local of (Word style (style body text) of boundDoc) as text",
            "set expectedListStyleName to name local of (Word style (style list paragraph) of boundDoc) as text",
            "set end of nativeRows to {\"builtin-styles\",expectedBodyStyleName,expectedListStyleName}",
            "repeat with paragraphIndex from 1 to count paragraphs of boundDoc",
            "set nativeParagraph to text object of paragraph paragraphIndex of boundDoc",
            "set nativeFormat to paragraph format of nativeParagraph",
            "set nativeStyleName to name local of style of nativeParagraph as text",
            "set end of nativeRows to {paragraphIndex as integer,content of nativeParagraph as text,nativeStyleName,paragraph format left indent of nativeFormat,first line indent of nativeFormat}",
            "end repeat",
        ]
    )
    return rows


def validate_readback(rows) -> bool:
    if not isinstance(rows, list):
        return False
    if (not rows or not isinstance(rows[0], list) or len(rows[0]) != 3
            or rows[0][0] != "builtin-styles"
            or not all(isinstance(value, str) and value for value in rows[0][1:])):
        return False
    body_style, list_style = rows[0][1:]
    paragraphs = rows[1:]
    if any(not isinstance(row, list) or len(row) != 5 for row in paragraphs):
        return False
    by_text = {row[1].removesuffix("\r"): row for row in paragraphs if isinstance(row[1], str)}
    for text, indent in {
        "•\tDEFAULT ITEM": 24,
        "→\tCUSTOM ITEM": 30,
        "\tEMPTY GLYPH ITEM": 27.5,
        "1.\tORDERED ITEM": 33,
    }.items():
        row = by_text.get(text)
        if (row is None or row[2] != list_style
                or abs(float(row[3]) - indent) > 0.01
                or abs(float(row[4]) + indent) > 0.01):
            return False
    for text in ("BODY AFTER DEFAULT", "BODY AFTER ORDERED"):
        row = by_text.get(text)
        if (row is None or row[2] != body_style
                or abs(float(row[3])) > 0.01 or abs(float(row[4]) - 24) > 0.01):
            return False
    return True


def run(output: Path, freeze: Path, timeout: float, *, adapter_factory=MacWordAdapter,
        session_type=MacWordSession, inventory_reader=None, pdf_inspector=inspect_pdf) -> dict:
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError("timeout must be finite and positive")
    source_hashes = verify_source_freeze(freeze)
    output = output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    report = {"status": "FAIL", "checks": {}, "source_hashes": source_hashes}
    if inventory_reader is None:
        from fixtures.microsoft_parity.macos_word_recovery import inventory as inventory_reader
    report["inventory_before"] = inventory_reader(output, "inventory-before")
    adapter = adapter_factory(build_case())
    deadline = time.monotonic() + timeout
    try:
        outcome = adapter.execute(adapter.build, (), deadline)
        staged = Path(outcome.staged_artifact)
        adapter.publish(staged, output / "list-arguments.docx", False, deadline)
        pdf = adapter.export_pdf(staged, deadline)
        adapter.publish(pdf, output / "list-arguments.pdf", False, deadline)
        if adapter.staging_root:
            shutil.copytree(adapter.staging_root, output / "generation-runtime")
        report["checks"].update(inspect_docx(output / "list-arguments.docx"))
        report["checks"].update(pdf_inspector(output / "list-arguments.pdf"))
        report["issues"] = [issue.to_dict() for issue in outcome.issues]
        report["checks"]["no_native_issues"] = not report["issues"]
    except BaseException as error:
        report["error"] = {"type": type(error).__name__, "message": str(error)}
        (output / "failure.txt").write_text(traceback.format_exc(), encoding="utf-8")
        if adapter.staging_root:
            shutil.copytree(adapter.staging_root, output / "failed-runtime", dirs_exist_ok=True)
    finally:
        try:
            adapter.close()
        except BaseException as error:
            report["cleanup_error"] = {"type": type(error).__name__, "message": str(error)}
            report.setdefault("error", report["cleanup_error"])

    if "error" not in report:
        before = sha256(output / "list-arguments.docx")
        reopened = None
        try:
            with session_type.open_document(output / "list-arguments.docx", read_only=True, visible=False) as reopened:
                report["reopen_rows"] = _readback(reopened)
                report["checks"]["native_readback_list_and_body_formats"] = validate_readback(report["reopen_rows"])
                if reopened.staging_root:
                    shutil.copytree(reopened.staging_root, output / "reopen-runtime")
            report["checks"]["read_only_reopen_exact_close"] = reopened._closed
            report["checks"]["read_only_reopen_preserves_bytes"] = sha256(output / "list-arguments.docx") == before
            report["checks"].update({"reopen_" + key: value for key, value in inspect_docx(output / "list-arguments.docx").items()})
        except BaseException as error:
            report["error"] = {"type": type(error).__name__, "message": str(error)}
            (output / "failure.txt").write_text(traceback.format_exc(), encoding="utf-8")
            if reopened and reopened.staging_root:
                shutil.copytree(reopened.staging_root, output / "failed-reopen-runtime", dirs_exist_ok=True)
    try:
        report["inventory_after"] = inventory_reader(output, "inventory-after")
        report["checks"]["document_inventory_preserved"] = report["inventory_after"] == report["inventory_before"]
    except BaseException as error:
        report["inventory_error"] = {"type": type(error).__name__, "message": str(error)}
        report.setdefault("error", report["inventory_error"])
    try:
        report["checks"]["source_freeze_still_matches"] = verify_source_freeze(freeze) == source_hashes
    except BaseException as error:
        report["source_freeze_error"] = {"type": type(error).__name__, "message": str(error)}
        report.setdefault("error", report["source_freeze_error"])
    report["artifact_hashes"] = {
        path.name: sha256(path) for path in output.iterdir() if path.is_file() and path.suffix in (".docx", ".pdf")
    }
    report["status"] = "PASS" if "error" not in report and report["checks"] and all(report["checks"].values()) else "FAIL"
    (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--source-freeze", type=Path)
    parser.add_argument("--timeout", type=float, default=240)
    args = parser.parse_args(argv)
    if not args.execute:
        print("Refusing native Word execution without --execute", file=sys.stderr)
        return 2
    if args.source_freeze is None:
        parser.error("--source-freeze is required with --execute")
    report = run(args.output, args.source_freeze, args.timeout)
    print(json.dumps({"status": report["status"], "error": report.get("error")}, ensure_ascii=False))
    return 0 if report["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
