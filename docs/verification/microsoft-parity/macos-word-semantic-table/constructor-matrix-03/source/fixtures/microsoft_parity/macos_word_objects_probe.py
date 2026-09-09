"""Guarded Probe A for Microsoft Word table, image, and text-box objects.

This fixture uses only the existing ``MacWordSession`` lifecycle, lock,
private-image staging and exact bound-document execution path.  It never runs
on import, never addresses an active document, and never quits Word.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import traceback
from zipfile import ZipFile
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

from skills.WPSComposer.scripts.artifact_transport import copy_file_before_deadline
from skills.WPSComposer.scripts.msoffice.macos_script import apple_string
from skills.WPSComposer.scripts.msoffice.macos_word_session import MacWordSession


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
W = "{" + WORD_NS + "}"


def _digest(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def _append_range(prefix: str) -> list[str]:
    return [
        "set insertionPoint to (end of content of text object of boundDoc) - 1",
        "if insertionPoint > 0 then",
        "set precedingRange to create range boundDoc start (insertionPoint - 1) end insertionPoint",
        "if (content of precedingRange as text) is not return then",
        "set boundaryRange to create range boundDoc start insertionPoint end insertionPoint",
        "set content of boundaryRange to return",
        "set insertionPoint to insertionPoint + 1",
        "end if",
        "end if",
        f"set {prefix}Range to create range boundDoc start insertionPoint end insertionPoint",
    ]


def _table_cell_commands(table: str, row: int, column: int, text: str, alignment: str) -> list[str]:
    return [
        f"set ownCell to get cell from table {table} row {row} column {column}",
        f"set content of text object of ownCell to {apple_string(text)}",
        "set first line indent of paragraph format of text object of ownCell to 0",
        "set character unit first line indent of paragraph format of text object of ownCell to 0",
        "set paragraph format left indent of paragraph format of text object of ownCell to 0",
        "set paragraph format right indent of paragraph format of text object of ownCell to 0",
        "set space before of paragraph format of text object of ownCell to 0",
        "set space after of paragraph format of text object of ownCell to 0",
        "set line spacing rule of paragraph format of text object of ownCell to line space single",
        f"set alignment of paragraph format of text object of ownCell to align paragraph {alignment}",
        "set vertical alignment of ownCell to cell align vertical center",
    ]


def build_probe_commands(private_image: Path) -> list[str]:
    """Compile the one-batch Probe A mutation and immediate native readback."""
    image = apple_string(str(Path(private_image).resolve()))
    first = [
        ["Header 中文😀", "Header ASCII", "Header Three"],
        ["左对齐中文😀", "Centered", "Right aligned"],
        ["Banded row", "Wrap text " * 8, "Tail"],
    ]
    second = [
        ["Horizontal merge", "", "M1 outside"],
        ["M2 outside", "M2 center", "Vertical merge"],
        ["M3 outside", "M3 center", ""],
    ]
    alignments = ["left", "center", "right"]
    lines = _append_range("table1")
    lines += [
        "set table1 to make new table at boundDoc with properties {text object:table1Range, number of rows:3, number of columns:3}",
        "set allow auto fit of table1 to false",
        "set allow page breaks of table1 to false",
        "set heading format of row 1 of table1 to true",
        "set left padding of table1 to 4",
        "set right padding of table1 to 4",
        "set top padding of table1 to 2",
        "set bottom padding of table1 to 2",
    ]
    for row, values in enumerate(first, 1):
        lines.append(f"set allow break across pages of row {row} of table1 to false")
        for column, text in enumerate(values, 1):
            lines += _table_cell_commands(table="table1", row=row, column=column,
                                          text=text, alignment=alignments[column - 1])
            lines.append(f"set width of ownCell to {{90, 150, 105}}'s item {column}")
            if row == 1:
                lines += [
                    "set background pattern color of shading of text object of ownCell to {17476, 29041, 50372}",
                    "set color of font object of text object of ownCell to {65535, 65535, 65535}",
                    "set bold of font object of text object of ownCell to true",
                ]
            elif row == 3:
                lines.append("set background pattern color of shading of text object of ownCell to {61680, 63222, 64764}")
    for which in ("top", "bottom", "left", "right", "horizontal", "vertical"):
        lines += [
            f"set ownBorder to get border table1 which border border {which}",
            "set line style of ownBorder to line style single",
            "set line width of ownBorder to line width75 point",
            "set color of ownBorder to {53456, 53456, 53456}",
        ]

    lines += _append_range("table2")
    lines += [
        "set table2 to make new table at boundDoc with properties {text object:table2Range, number of rows:3, number of columns:3}",
        "set allow auto fit of table2 to false",
    ]
    for row, values in enumerate(second, 1):
        lines.append(f"set allow break across pages of row {row} of table2 to false")
        for column, text in enumerate(values, 1):
            lines += _table_cell_commands(table="table2", row=row, column=column,
                                          text=text, alignment=alignments[column - 1])
            lines.append("set width of ownCell to 110")
    # Bottom-right first matches the planned all-or-nothing merge compiler order.
    lines += [
        "set mergeStart to get cell from table table2 row 2 column 3",
        "set mergeEnd to get cell from table table2 row 3 column 3",
        "merge cell mergeStart with mergeEnd",
        "set mergeStart to get cell from table table2 row 1 column 1",
        "set mergeEnd to get cell from table table2 row 1 column 2",
        "merge cell mergeStart with mergeEnd",
    ]

    lines += _append_range("inline1")
    lines += [
        f"set inline1 to make new inline picture at inline1Range with properties {{file name:{image}, link to file:false, save with document:true}}",
        "set inline1 to inline picture (count inline pictures of boundDoc) of boundDoc",
        "set lock aspect ratio of inline1 to true",
        "set width of inline1 to 72",
        f"set alternative text of inline1 to {apple_string('Inline 中文😀 alt')}",
    ]
    lines += _append_range("inline2")
    lines += [
        f"set inline2 to make new inline picture at inline2Range with properties {{file name:{image}, link to file:false, save with document:true}}",
        "set inline2 to inline picture (count inline pictures of boundDoc) of boundDoc",
        "set lock aspect ratio of inline2 to false",
        "set width of inline2 to 90",
        "set height of inline2 to 45",
        f"set alternative text of inline2 to {apple_string('Bounded two dimension alt')}",
    ]
    lines += _append_range("floating")
    lines += [
        f"set floatingPicture to make new picture at boundDoc with properties {{file name:{image}, link to file:false, save with document:true, anchor:floatingRange}}",
        "set floatingPicture to shape (count shapes of boundDoc) of boundDoc",
        "set lock aspect ratio of floatingPicture to false",
        "set left position of floatingPicture to 300",
        "set top of floatingPicture to 120",
        "set width of floatingPicture to 96",
        "set height of floatingPicture to 54",
        "set wrap type of wrap format of floatingPicture to wrap square",
    ]
    lines += _append_range("textbox")
    lines += [
        "set textBox1 to make new text box at boundDoc with properties {anchor:textboxRange, left position:72, top:360, width:220, height:54}",
        "set textBox1 to shape (count shapes of boundDoc) of boundDoc",
        "set relative horizontal position of textBox1 to relative horizontal position page",
        "set relative vertical position of textBox1 to relative vertical position page",
        "set left position of textBox1 to 72",
        "set top of textBox1 to 360",
        f"set content of text range of text frame of textBox1 to {apple_string('Textbox 中文😀')}",
        "set visible of fill format of textBox1 to true",
        "set fore color of fill format of textBox1 to {65535, 61680, 46260}",
        "set font size of font object of text range of text frame of textBox1 to 13",
        "set bold of font object of text range of text frame of textBox1 to true",
        "set wrap type of wrap format of textBox1 to wrap top bottom",
    ]
    lines += _append_range("block")
    lines += [
        f"set blockPicture to make new inline picture at blockRange with properties {{file name:{image}, link to file:false, save with document:true}}",
        "set blockPicture to inline picture (count inline pictures of boundDoc) of boundDoc",
        "set lock aspect ratio of blockPicture to true",
        "set width of blockPicture to 84",
        f"set alternative text of blockPicture to {apple_string('Block image alt')}",
        "set blockTextRange to text object of blockPicture",
        "set alignment of paragraph format of blockTextRange to align paragraph center",
        "set keep with next of paragraph format of blockTextRange to true",
        "set space after of paragraph format of blockTextRange to 0",
        "set insertionPoint to (end of content of text object of boundDoc) - 1",
        "set followingRange to create range boundDoc start insertionPoint end insertionPoint",
        f"set content of followingRange to return & {apple_string('Following 中文😀 paragraph')} & return",
        "set followingIndex to count paragraphs of boundDoc",
    ]
    lines += build_readback_commands()
    return lines


def build_readback_commands() -> list[str]:
    """Read fields required by Probe A directly from the bound Word document."""
    return [
        'set end of nativeRows to {"probe", 0, "word_version", version as text}',
        'set end of nativeRows to {"probe", 0, "table_count", count tables of boundDoc}',
        'set end of nativeRows to {"probe", 0, "inline_count", count inline pictures of boundDoc}',
        'set end of nativeRows to {"probe", 0, "shape_count", count shapes of boundDoc}',
        'set firstTable to table 1 of boundDoc',
        'set end of nativeRows to {"table1", 1, "rows", number of rows of firstTable}',
        'set end of nativeRows to {"table1", 1, "columns", number of columns of firstTable}',
        'set end of nativeRows to {"table1", 1, "auto_fit", allow auto fit of firstTable}',
        'set end of nativeRows to {"table1", 1, "repeat_header", heading format of row 1 of firstTable}',
        'set end of nativeRows to {"table1", 1, "row_split", allow break across pages of row 2 of firstTable}',
        'set firstCell to get cell from table firstTable row 1 column 1',
        'set end of nativeRows to {"table1", 1, "first_text", content of text object of firstCell as text}',
        'set end of nativeRows to {"table1", 1, "first_indent", first line indent of paragraph format of text object of firstCell}',
        'set end of nativeRows to {"table1", 1, "left_indent", paragraph format left indent of paragraph format of text object of firstCell}',
        'set end of nativeRows to {"table1", 1, "right_indent", paragraph format right indent of paragraph format of text object of firstCell}',
        'set end of nativeRows to {"table1", 1, "space_before", space before of paragraph format of text object of firstCell}',
        'set end of nativeRows to {"table1", 1, "space_after", space after of paragraph format of text object of firstCell}',
        'set end of nativeRows to {"table1", 1, "vertical_center", (vertical alignment of firstCell is cell align vertical center)}',
        'set end of nativeRows to {"table1", 1, "widths", {width of cell 1 of row 1 of firstTable, width of cell 2 of row 1 of firstTable, width of cell 3 of row 1 of firstTable}}',
        'set end of nativeRows to {"table1", 1, "alignments", {my enumIndex(alignment of paragraph format of text object of cell 1 of row 2 of firstTable, {align paragraph left, align paragraph center, align paragraph right}), my enumIndex(alignment of paragraph format of text object of cell 2 of row 2 of firstTable, {align paragraph left, align paragraph center, align paragraph right}), my enumIndex(alignment of paragraph format of text object of cell 3 of row 2 of firstTable, {align paragraph left, align paragraph center, align paragraph right})}}',
        'set firstTopBorder to get border firstTable which border border top',
        'set end of nativeRows to {"table1", 1, "top_border", (line style of firstTopBorder is line style single)}',
        'set end of nativeRows to {"table1", 1, "header_shade", background pattern color of shading of text object of firstCell}',
        'set end of nativeRows to {"table1", 1, "header_font", color of font object of text object of firstCell}',
        'set end of nativeRows to {"table2", 2, "rows", number of rows of table 2 of boundDoc}',
        'set end of nativeRows to {"table2", 2, "columns", number of columns of table 2 of boundDoc}',
        'set end of nativeRows to {"inline1", 1, "width", width of inline picture 1 of boundDoc}',
        'set end of nativeRows to {"inline1", 1, "height", height of inline picture 1 of boundDoc}',
        'set end of nativeRows to {"inline1", 1, "alt", alternative text of inline picture 1 of boundDoc as text}',
        'set end of nativeRows to {"inline2", 2, "width", width of inline picture 2 of boundDoc}',
        'set end of nativeRows to {"inline2", 2, "height", height of inline picture 2 of boundDoc}',
        'set end of nativeRows to {"inline2", 2, "alt", alternative text of inline picture 2 of boundDoc as text}',
        'set floatingPicture to shape 1 of boundDoc',
        'set end of nativeRows to {"floating", 1, "type", my enumIndex(shape type of floatingPicture, {shape type picture, shape type text box})}',
        'set end of nativeRows to {"floating", 1, "geometry", {left position of floatingPicture, top of floatingPicture, width of floatingPicture, height of floatingPicture}}',
        'set end of nativeRows to {"floating", 1, "wrap", my enumIndex(wrap type of wrap format of floatingPicture, {wrap square, wrap top bottom, wrap inline})}',
        'set end of nativeRows to {"floating", 1, "alt", alternative text of floatingPicture as text}',
        'set textBox1 to shape 2 of boundDoc',
        'set end of nativeRows to {"textbox", 2, "type", my enumIndex(shape type of textBox1, {shape type picture, shape type text box})}',
        'set end of nativeRows to {"textbox", 2, "text", content of text range of text frame of textBox1 as text}',
        'set end of nativeRows to {"textbox", 2, "geometry", {left position of textBox1, top of textBox1, width of textBox1, height of textBox1}}',
        'set end of nativeRows to {"textbox", 2, "wrap", my enumIndex(wrap type of wrap format of textBox1, {wrap square, wrap top bottom, wrap inline})}',
        'set end of nativeRows to {"textbox", 2, "fill_visible", visible of fill format of textBox1}',
        'set end of nativeRows to {"textbox", 2, "fill_color", fore color of fill format of textBox1}',
        'set end of nativeRows to {"textbox", 2, "font_size", font size of font object of text range of text frame of textBox1}',
        'set end of nativeRows to {"textbox", 2, "bold", bold of font object of text range of text frame of textBox1}',
        'set blockPicture to inline picture 3 of boundDoc',
        'set blockTextRange to text object of blockPicture',
        'set end of nativeRows to {"block", 3, "center", (alignment of paragraph format of blockTextRange is align paragraph center)}',
        'set end of nativeRows to {"block", 3, "keep_with_next", keep with next of paragraph format of blockTextRange}',
        'set end of nativeRows to {"block", 3, "space_after", space after of paragraph format of blockTextRange}',
        'set end of nativeRows to {"block", 3, "width", width of blockPicture}',
        'set end of nativeRows to {"block", 3, "height", height of blockPicture}',
        'set followingTextReadback to missing value',
        'repeat with paragraphIndex from 1 to (count paragraphs of boundDoc)',
        'set candidateText to content of text object of paragraph paragraphIndex of boundDoc as text',
        f'if candidateText contains {apple_string("Following 中文😀 paragraph")} then set followingTextReadback to candidateText',
        'end repeat',
        'set end of nativeRows to {"block", 3, "following_text", followingTextReadback as text}',
    ]


def _inventory_commands() -> list[str]:
    return [
        "set nativeRows to {}",
        "repeat with documentIndex from 1 to (count documents)",
        "set inventoryDoc to document documentIndex",
        "if (posix full name of inventoryDoc as text) is not (posix full name of boundDoc as text) then",
        'set end of nativeRows to {name of inventoryDoc as text, posix full name of inventoryDoc as text, saved of inventoryDoc, content of text object of inventoryDoc as text}',
        "end if",
        "end repeat",
    ]


def _rows_by_key(rows: list[list[object]]) -> dict[tuple[str, str], object]:
    return {(str(row[0]), str(row[2])): row[3] for row in rows}


def verify_native_rows(rows: list[list[object]]) -> dict[str, bool]:
    values = _rows_by_key(rows)
    close = lambda actual, expected, tolerance=0.8: abs(float(actual) - expected) <= tolerance
    close_vector = lambda actual, expected, tolerance=0.8: isinstance(actual, (list, tuple)) and len(actual) == len(expected) and all(close(a, b, tolerance) for a, b in zip(actual, expected))
    return {
        "native_counts": values[("probe", "table_count")] == 2 and values[("probe", "inline_count")] == 3 and values[("probe", "shape_count")] == 2,
        "native_table_dimensions": values[("table1", "rows")] == 3 and values[("table1", "columns")] == 3,
        "native_table_header_and_split": values[("table1", "repeat_header")] is True and values[("table1", "row_split")] is False,
        "native_table_indents_spacing": all(close(values[("table1", key)], 0, 0.05) for key in ("first_indent", "left_indent", "right_indent", "space_before", "space_after")),
        "native_table_alignments": values[("table1", "alignments")] == [0, 1, 2],
        "native_table_widths": close_vector(values[("table1", "widths")], (90, 150, 105), 1.5),
        "native_table_autofit_and_header_colors": values[("table1", "auto_fit")] is False and values[("table1", "header_shade")] == [17476, 29041, 50372] and values[("table1", "header_font")] == [65535, 65535, 65535],
        "native_table_vertical_center": values[("table1", "vertical_center")] is True,
        "native_table_border": values[("table1", "top_border")] is True,
        "native_table_utf16_text": str(values[("table1", "first_text")]).startswith("Header 中文😀"),
        "native_merged_table_present": values[("table2", "rows")] >= 2 and values[("table2", "columns")] >= 2,
        "native_inline_one_dimension": close(values[("inline1", "width")], 72) and close(values[("inline1", "height")], 40.5) and values[("inline1", "alt")] == "Inline 中文😀 alt",
        "native_inline_two_dimensions": close(values[("inline2", "width")], 90) and close(values[("inline2", "height")], 45) and values[("inline2", "alt")] == "Bounded two dimension alt",
        "native_floating_image": values[("floating", "type")] == 0 and values[("floating", "wrap")] == 0,
        "native_floating_geometry": close_vector(values[("floating", "geometry")], (300, 120, 96, 54), 1.0),
        "native_textbox": values[("textbox", "type")] == 1 and str(values[("textbox", "text")]).startswith("Textbox 中文😀") and values[("textbox", "wrap")] == 1,
        "native_textbox_geometry": close_vector(values[("textbox", "geometry")], (72, 360, 220, 54), 1.0),
        "native_textbox_font_fill": close(values[("textbox", "font_size")], 13, 0.2) and values[("textbox", "bold")] is True and values[("textbox", "fill_visible")] is True and values[("textbox", "fill_color")] == [65535, 61680, 46260],
        "native_block_paragraph": values[("block", "center")] is True and values[("block", "keep_with_next")] is True and close(values[("block", "space_after")], 0, 0.05) and "Following 中文😀 paragraph" in str(values[("block", "following_text")]),
    }


def verify_artifacts(output: Path) -> dict[str, bool]:
    """Inspect Word-created DOCX/PDF without rewriting either artifact."""
    output = Path(output)
    docx = output / "word-objects.docx"
    pdf = output / "word-objects.pdf"
    with ZipFile(docx) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        relationships = ET.fromstring(package.read("word/_rels/document.xml.rels"))
        tables = document.findall(".//" + W + "tbl")
        horizontal = document.findall(".//" + W + "gridSpan")
        vertical = document.findall(".//" + W + "vMerge")
        media = [name for name in package.namelist() if name.startswith("word/media/")]
        image_relationships = [
            rel for rel in relationships
            if rel.get("Type", "").endswith("/image")
        ]
        external_images = [rel for rel in image_relationships if rel.get("TargetMode") == "External"]
        text = "".join(document.itertext())
        anchors = document.findall(".//{" + DRAWING_NS + "}anchor")
        inlines = document.findall(".//{" + DRAWING_NS + "}inline")
        alt_values = [node.get("descr", "") for node in document.iter() if node.tag.endswith("}docPr")]
        textboxes = document.findall(".//" + W + "txbxContent")
        header_rows = document.findall(".//" + W + "tblHeader")
        cant_split = document.findall(".//" + W + "cantSplit")
        centered_cells = [node for node in document.findall(".//" + W + "vAlign") if node.get(W + "val") == "center"]
        paragraph_indents = document.findall(".//" + W + "tc/" + W + "p/" + W + "pPr/" + W + "ind")
        paragraph_spacing = document.findall(".//" + W + "tc/" + W + "p/" + W + "pPr/" + W + "spacing")
    from fixtures.microsoft_parity.word_object_artifact_verifier import (
        verify_picture_artifacts,
        verify_second_table_merges,
    )
    strict = verify_picture_artifacts(
        docx, pdf, output / "probe-source.png",
        expected_extents=[(72, 40.5), (90, 45), (96, 54), (84, 47.25)],
        expected_picture_anchor=(300, 120, "wrapSquare", "column", "paragraph"),
        expected_textbox=(72, 360, 220, 54, "wrapTopAndBottom"),
    )
    exact_merges = verify_second_table_merges(
        document, "Horizontal merge", "Vertical merge",
        ("M1 outside", "M2 outside", "M3 outside"),
    )
    import fitz
    with fitz.open(pdf) as pages:
        pdf_text = "\n".join(page.get_text() for page in pages)
        pdf_images = sum(len(page.get_images(full=True)) for page in pages)
    return {
        "ooxml_two_native_tables": len(tables) == 2,
        "ooxml_horizontal_and_vertical_merges": exact_merges,
        "ooxml_merge_outside_text": all(value in text for value in ("M1 outside", "M2 outside", "M3 outside")),
        "ooxml_embedded_media": len(media) >= 1 and len(image_relationships) >= 1 and not external_images and strict["relationships"],
        "ooxml_inline_and_floating_drawings": len(inlines) >= 3 and len(anchors) >= 2 and strict["picture_geometry"],
        "ooxml_alt_text": all(value in alt_values for value in ("Inline 中文😀 alt", "Bounded two dimension alt", "Block image alt")),
        "ooxml_native_textbox": bool(textboxes) and "Textbox 中文😀" in text and strict["textbox_geometry"],
        "ooxml_table_repeat_and_split": bool(header_rows) and len(cant_split) >= 6,
        # Direct zero values may be omitted by Word when it serializes the same
        # effective zero. Native readback above is the exact proof; OOXML must
        # merely contain no contradictory nonzero direct value.
        "ooxml_table_cell_format": len(centered_cells) >= 9 and all(all(node.get(W + key) in (None, "0") for key in ("firstLine", "left", "right")) for node in paragraph_indents) and all(node.get(W + "before") in (None, "0") and node.get(W + "after") in (None, "0") for node in paragraph_spacing),
        "pdf_visible_table_text": all(value in pdf_text for value in ("Header", "Centered", "Right aligned", "Horizontal merge", "M3 center")),
        "pdf_visible_textbox_and_following_text": "Textbox" in pdf_text and "Following" in pdf_text,
        "pdf_contains_images": pdf_images >= 2 and strict["pdf_placements"],
    }


def _write_probe_png(path: Path) -> None:
    from PIL import Image, ImageDraw
    image = Image.new("RGB", (320, 180), "#E8F1FA")
    draw = ImageDraw.Draw(image)
    draw.rectangle((10, 10, 310, 170), outline="#4472C4", width=8)
    draw.ellipse((110, 40, 210, 140), fill="#F4B183")
    image.save(path, format="PNG")


def run(output: Path) -> dict[str, object]:
    output = Path(output).resolve()
    output.mkdir(parents=True, exist_ok=False)
    source_image = output / "probe-source.png"
    _write_probe_png(source_image)
    report: dict[str, object] = {
        "schema": 1,
        "probe": "A",
        "status": "RUNNING",
        "checks": {},
        "limitations": [
            "Microsoft Word 16.112.3 rejects the alternative text setter on a native floating picture; run-02 retains the exact -10006 failure. Floating-image alt therefore remains unverified and must be rejected by any bounded production implementation.",
        ],
        "source_hashes": {
            "fixture": _digest(Path(__file__)),
            "macos_word_session": _digest(ROOT / "skills/WPSComposer/scripts/msoffice/macos_word_session.py"),
            "probe_image": _digest(source_image),
        },
    }
    session = None
    try:
        with MacWordSession.new_document(visible=False) as session:
            before = session._execute(_inventory_commands())
            private_image = session.staging_root / "probe-image.png"
            copy_file_before_deadline(source_image, private_image, deadline=session.publication_deadline)
            if _digest(private_image) != _digest(source_image):
                raise AssertionError("Private image staging digest mismatch")
            rows = session._execute(build_probe_commands(private_image))
            report["native_rows"] = rows
            report["checks"].update(verify_native_rows(rows))
            session.save_docx(output / "word-objects.docx")
            session.export_pdf(output / "word-objects.pdf")
            after = session._execute(_inventory_commands())
            report["inventory_before"] = before
            report["inventory_after"] = after
            report["checks"]["unrelated_documents_unchanged"] = before == after
            shutil.copytree(session.staging_root, output / "native-runtime")
        report["checks"]["exact_owned_close"] = session._closed
        session = None

        docx_hash = _digest(output / "word-objects.docx")
        with MacWordSession.open_document(output / "word-objects.docx", visible=False) as reopened:
            reopened_rows = reopened._execute(build_readback_commands())
            report["native_reopen_rows"] = reopened_rows
            report["checks"]["native_reopen_readback"] = all(verify_native_rows(reopened_rows).values())
            shutil.copytree(reopened.staging_root, output / "native-reopen-runtime")
        report["checks"]["native_reopen_preserves_docx_hash"] = _digest(output / "word-objects.docx") == docx_hash
        report["checks"].update(verify_artifacts(output))
        report["hashes"] = {
            name: _digest(output / name)
            for name in ("probe-source.png", "word-objects.docx", "word-objects.pdf")
        }
        report["status"] = "PASS" if all(report["checks"].values()) else "PARTIAL"
    except BaseException as error:
        report["status"] = "FAIL"
        report["error"] = {"type": type(error).__name__, "message": str(error)}
        (output / "failure.txt").write_text(traceback.format_exc(), encoding="utf-8")
        if session and session.staging_root and session.staging_root.exists():
            shutil.copytree(session.staging_root, output / "failed-runtime", dirs_exist_ok=True)
    (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    result = run(args.output_dir)
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result["status"] == "PASS" else 1


if __name__ == "__main__":
    raise SystemExit(main())
