"""Native acceptance for the first five direct macOS Word object methods."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from uuid import uuid4
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))


def _digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _image(path):
    image = Image.new("RGB", (320, 180), "#E8F1FA")
    draw = ImageDraw.Draw(image)
    draw.rectangle((10, 10, 310, 170), outline="#4472C4", width=8)
    draw.ellipse((110, 40, 210, 140), fill="#F4B183")
    image.save(path, "PNG")


def _inventory():
    return [
        "set nativeRows to {}", "repeat with di from 1 to count documents",
        "set d to document di",
        "if (posix full name of d as text) is not (posix full name of boundDoc as text) then set end of nativeRows to {name of d as text, posix full name of d as text, saved of d, content of text object of d as text}",
        "end repeat",
    ]


def _readback():
    return [
        'set nativeRows to {{"counts", count tables of boundDoc, count inline pictures of boundDoc, count shapes of boundDoc}}',
        'set end of nativeRows to {"table", number of rows of table 1 of boundDoc, number of columns of table 1 of boundDoc, heading format of row 1 of table 1 of boundDoc, allow break across pages of row 2 of table 1 of boundDoc}',
        'set c to get cell from table (table 1 of boundDoc) row 1 column 1',
        'set end of nativeRows to {"cell", content of text object of c as text, first line indent of paragraph format of text object of c, paragraph format left indent of paragraph format of text object of c, paragraph format right indent of paragraph format of text object of c, vertical alignment of c is cell align vertical center}',
        'set end of nativeRows to {"inline1", width of inline picture 1 of boundDoc, height of inline picture 1 of boundDoc, alternative text of inline picture 1 of boundDoc as text}',
        'set end of nativeRows to {"inline2", width of inline picture 2 of boundDoc, height of inline picture 2 of boundDoc}',
        'set end of nativeRows to {"inline3", width of inline picture 3 of boundDoc, height of inline picture 3 of boundDoc}',
        'set end of nativeRows to {"inline4", width of inline picture 4 of boundDoc, height of inline picture 4 of boundDoc}',
        'set blockRange to text object of inline picture 5 of boundDoc',
        'set end of nativeRows to {"block", alignment of paragraph format of blockRange is align paragraph center, keep with next of paragraph format of blockRange, space after of paragraph format of blockRange}',
        'set floatingPicture to shape 1 of boundDoc',
        'set end of nativeRows to {"floating", shape type of floatingPicture is shape type picture, width of floatingPicture, height of floatingPicture, wrap type of wrap format of floatingPicture is wrap square}',
        'set textBoxShape to shape 2 of boundDoc',
        'set end of nativeRows to {"textbox", shape type of textBoxShape is shape type text box, content of text range of text frame of textBoxShape as text, left position of textBoxShape, top of textBoxShape, width of textBoxShape, height of textBoxShape, wrap type of wrap format of textBoxShape is wrap square, font size of font object of text range of text frame of textBoxShape, bold of font object of text range of text frame of textBoxShape}',
    ]


def _verify_native(rows):
    by_name = {row[0]: row[1:] for row in rows}
    close = lambda a, b, tolerance=1.0: abs(float(a) - b) <= tolerance
    return {
        "reopen_counts": by_name["counts"] == [2, 5, 2],
        "reopen_table": by_name["table"] == [3, 3, True, False],
        "reopen_cell_utf16_format": str(by_name["cell"][0]).startswith("Header 中文😀") and by_name["cell"][1:4] == [0, 0, 0] and by_name["cell"][4] is True,
        "reopen_inline_width": close(by_name["inline1"][0], 72) and close(by_name["inline1"][1], 40.5) and by_name["inline1"][2] == "Inline 中文😀 alt",
        "reopen_inline_height": close(by_name["inline2"][0], 88.8889) and close(by_name["inline2"][1], 50),
        "reopen_inline_bounded_aspect": close(by_name["inline3"][0], 90) and close(by_name["inline3"][1], 50.625),
        "reopen_inline_max": 69.5 <= by_name["inline4"][0] <= 70.5 and 39.0 <= by_name["inline4"][1] <= 40.0,
        "reopen_block_format": by_name["block"] == [True, True, 0],
        "reopen_floating": by_name["floating"][0] is True and close(by_name["floating"][1], 96) and close(by_name["floating"][2], 54) and by_name["floating"][3] is True,
        "reopen_textbox": by_name["textbox"][0] is True and str(by_name["textbox"][1]).startswith("Textbox 中文😀") and all(close(a, b) for a, b in zip(by_name["textbox"][2:6], (72, 360, 220, 54))) and by_name["textbox"][6] is True and close(by_name["textbox"][7], 13, .2) and by_name["textbox"][8] is True,
    }


def _verify_artifacts(output):
    w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    with ZipFile(output / "business-objects.docx") as package:
        root = ET.fromstring(package.read("word/document.xml"))
        rels = ET.fromstring(package.read("word/_rels/document.xml.rels"))
        text = "".join(root.itertext())
        paragraphs = ["".join(paragraph.itertext()) for paragraph in root.findall(".//" + w + "body/" + w + "p")]
        image_rels = [rel for rel in rels if rel.get("Type", "").endswith("/image")]
        checks = {
            "ooxml_tables_merges": len(root.findall(".//" + w + "tbl")) == 2,
            "ooxml_inline_floating_textbox": len(root.findall(".//" + wp + "inline")) >= 5 and len(root.findall(".//" + wp + "anchor")) >= 2 and bool(root.findall(".//" + w + "txbxContent")),
            "ooxml_embedded_only": bool(image_rels) and not any(rel.get("TargetMode") == "External" for rel in image_rels),
            "ooxml_following_paragraph": "Following 中文😀 paragraph" in paragraphs,
        }
    from fixtures.microsoft_parity.word_object_artifact_verifier import (
        verify_picture_artifacts,
        verify_second_table_merges,
    )
    strict = verify_picture_artifacts(
        output / "business-objects.docx", output / "business-objects.pdf", output / "source.png",
        expected_extents=[(72, 40.5), (88.8889, 50), (90, 50.625), (70, 39.375), (96, 54), (84, 47.25)],
        expected_picture_anchor=(0, 0, "wrapSquare", "page", "page"),
        expected_textbox=(72, 360, 220, 54, "wrapSquare"),
    )
    checks["ooxml_tables_merges"] = checks["ooxml_tables_merges"] and verify_second_table_merges(
        root, "Horizontal merge", "Vertical merge", ("Outside", "D", "G")
    )
    checks["ooxml_embedded_only"] = checks["ooxml_embedded_only"] and strict["relationships"]
    checks["ooxml_inline_floating_textbox"] = checks["ooxml_inline_floating_textbox"] and strict["picture_geometry"] and strict["textbox_geometry"]
    import fitz
    with fitz.open(output / "business-objects.pdf") as pdf:
        text = "\n".join(page.get_text() for page in pdf)
        image_count = sum(len(page.get_images(full=True)) for page in pdf)
    checks.update({
        "pdf_visible_objects": all(value in text for value in ("Header", "Horizontal merge", "Following", "Textbox")),
        "pdf_images": image_count >= 2 and strict["pdf_placements"],
    })
    return checks


def run(output):
    from skills.WPSComposer import create_document
    from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession, apple_string
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    source = output / "source.png"
    _image(source)
    report = {"schema": 1, "status": "RUNNING", "engine": "msoffice", "checks": {}, "targets": {},
              "method_to_checks": {
                  "add_table": ["reopen_table", "reopen_cell_utf16_format", "ooxml_tables_merges"],
                  "add_merged_table": ["ooxml_tables_merges"],
                  "add_image": ["reopen_inline_width", "reopen_inline_height", "reopen_inline_bounded_aspect", "reopen_inline_max", "reopen_floating", "ooxml_embedded_only"],
                  "add_image_block": ["reopen_block_format", "ooxml_following_paragraph"],
                  "add_floating_textbox": ["reopen_textbox", "pdf_visible_objects"],
              },
              "limitations": ["Floating image alt text remains a preflight-rejected unsupported branch; see Probe A run-02."],
              "source_hashes": {"fixture": _digest(__file__), "macos_word_session": _digest(ROOT / "skills/WPSComposer/scripts/msoffice/macos_word_session.py"), "image": _digest(source)}}
    session = None
    cleanup = None
    sentinel = None
    sentinel_token = "WPSC-OBJECT-SENTINEL-" + uuid4().hex
    try:
        with create_document("writer", engine="msoffice", visible=False) as session:
            sentinel_rows = session._execute([
                "set sentinelDoc to make new document",
                f"set content of text object of sentinelDoc to {json.dumps(sentinel_token)}",
                'set nativeRows to {{"sentinel", name of sentinelDoc as text, posix full name of sentinelDoc as text, saved of sentinelDoc, content of text object of sentinelDoc as text, id of active window of sentinelDoc}}',
            ])
            if len(sentinel_rows) != 1 or sentinel_rows[0][0] != "sentinel":
                raise RuntimeError("Synthetic sentinel identity unavailable")
            sentinel = sentinel_rows[0]
            report["sentinel_before"] = sentinel
            report["sentinel_token_sha256"] = hashlib.sha256(sentinel_token.encode()).hexdigest()
            before = session._execute(_inventory())
            report["targets"]["table"] = session.add_table(3, 3, [["Header 中文😀", "H2", "H3"], ["Left", "Center", "Right"], ["Band", "Wrap " * 12, "Tail"]], col_widths=[90, 150, 105], alignments=["left", "center", "right"])
            report["targets"]["merged"] = session.add_merged_table([["Horizontal merge", "", "Outside"], ["D", "E", "Vertical merge"], ["G", "H", ""]], [(1, 1, 1, 2), (2, 3, 3, 3)])
            report["targets"]["inline_width"] = session.add_image(source, width=72, alt="Inline 中文😀 alt")
            report["targets"]["inline_height"] = session.add_image(source, height=50)
            report["targets"]["inline_bounded"] = session.add_image(source, width=100, height=100, max_width=90, max_height=60)
            report["targets"]["inline_max"] = session.add_image(source, max_width=70, max_height=40)
            report["targets"]["floating"] = session.add_image(source, width=96, height=54, inline=False, wrap=0)
            report["targets"]["block"] = session.add_image_block(source, width=84)
            session.add_paragraph("Following 中文😀 paragraph")
            report["targets"]["textbox"] = session.add_floating_textbox("Textbox 中文😀", 72, 360, 220, 54, wrap=0, fill_color="#FFF0B4", font_size=13, bold=True)
            session.save_docx(output / "business-objects.docx")
            session.export_pdf(output / "business-objects.pdf")
            after = session._execute(_inventory())
            report["checks"]["unrelated_documents_unchanged"] = before == after
            report["checks"]["nonvacuous_unsaved_sentinel"] = len(before) >= 1 and sentinel[1:5] in before
            shutil.copytree(session.staging_root, output / "native-runtime")
        report["checks"]["exact_owned_close"] = session._closed
        session = None
        source_hash = _digest(output / "business-objects.docx")
        with MacWordSession.open_document(output / "business-objects.docx", visible=False) as reopened:
            rows = reopened._execute(_readback())
            report["native_reopen_rows"] = rows
            report["checks"].update(_verify_native(rows))
            shutil.copytree(reopened.staging_root, output / "native-reopen-runtime")
        report["checks"]["reopen_preserves_docx_hash"] = _digest(output / "business-objects.docx") == source_hash
        report["checks"].update(_verify_artifacts(output))
        report["hashes"] = {name: _digest(output / name) for name in ("source.png", "business-objects.docx", "business-objects.pdf")}
        report["status"] = "PASS" if all(report["checks"].values()) else "PARTIAL"
    except BaseException as error:
        report["status"] = "FAIL"
        report["error"] = {"type": type(error).__name__, "message": str(error)}
        (output / "failure.txt").write_text(traceback.format_exc(), encoding="utf-8")
        if session and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root, output / "failed-runtime", dirs_exist_ok=True)
    finally:
        if sentinel is not None:
            try:
                cleanup = MacWordSession()
                cleanup._prepare()
                rows = cleanup._execute([
                    f'set sentinelDoc to document {apple_string(sentinel[1])}',
                    f'if (posix full name of sentinelDoc as text) is not {apple_string(sentinel[2])} then error "WPSC_SENTINEL_PATH_CHANGED"',
                    f'if (content of text object of sentinelDoc as text) is not {apple_string(sentinel_token + chr(13))} then error "WPSC_SENTINEL_TEXT_CHANGED"',
                    'if saved of sentinelDoc then error "WPSC_SENTINEL_SAVED"',
                    'close sentinelDoc saving no',
                    'set nativeRows to {{"sentinel_closed"}}',
                ], bind=False)
                report["checks"]["unsaved_sentinel_preserved_and_closed"] = rows == [["sentinel_closed"]]
                cleanup.close()
            except BaseException as cleanup_error:
                report["cleanup_error"] = {"type": type(cleanup_error).__name__, "message": str(cleanup_error)}
                if cleanup and cleanup.staging_root and cleanup.staging_root.exists():
                    shutil.copytree(cleanup.staging_root, output / "sentinel-failed-runtime", dirs_exist_ok=True)
                if cleanup and cleanup.lock:
                    cleanup.lock.close()
        if report.get("error") or report.get("cleanup_error"):
            report["status"] = "FAIL"
        else:
            report["status"] = "PASS" if report["checks"] and all(report["checks"].values()) else "PARTIAL"
    (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, required=True)
    result = run(parser.parse_args().output_dir)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
