"""Pure DOCX/PDF evidence checks shared by native Word object fixtures."""
from __future__ import annotations

import hashlib
import io
import posixpath
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from PIL import Image


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
EMU_PER_POINT = 12700.0


def _close_vectors(actual, expected, tolerance=0.8):
    if len(actual) != len(expected):
        return False
    remaining = [tuple(float(value) for value in vector) for vector in actual]
    for expected_vector in expected:
        for index, actual_vector in enumerate(remaining):
            if len(actual_vector) == len(expected_vector) and all(
                abs(value - float(target)) <= tolerance
                for value, target in zip(actual_vector, expected_vector)
            ):
                remaining.pop(index)
                break
        else:
            return False
    return not remaining


def _extent(node):
    value = node.find("./" + WP + "extent")
    if value is None:
        return None
    try:
        return float(value.get("cx")) / EMU_PER_POINT, float(value.get("cy")) / EMU_PER_POINT
    except (TypeError, ValueError):
        return None


def _position(node):
    horizontal = node.find("./" + WP + "positionH")
    vertical = node.find("./" + WP + "positionV")
    if horizontal is None or vertical is None:
        return None
    x = horizontal.find("./" + WP + "posOffset")
    y = vertical.find("./" + WP + "posOffset")
    try:
        return (
            horizontal.get("relativeFrom"), vertical.get("relativeFrom"),
            float(x.text) / EMU_PER_POINT, float(y.text) / EMU_PER_POINT,
        )
    except (AttributeError, TypeError, ValueError):
        return None


def _wrap_name(node):
    for child in node:
        local = child.tag.rsplit("}", 1)[-1]
        if local.startswith("wrap"):
            return local
    return None


def _pixel_identity(payload):
    with Image.open(io.BytesIO(payload)) as image:
        pixels = image.convert("RGBA")
        return pixels.size, hashlib.sha256(pixels.tobytes()).hexdigest()


def verify_second_table_merges(document, horizontal_text, vertical_text, outside_texts):
    """Bind the planned horizontal/vertical topology to the second table."""
    tables = document.findall(".//" + W + "tbl")
    if len(tables) < 2:
        return False
    rows = tables[1].findall("./" + W + "tr")
    if len(rows) != 3:
        return False
    cells = [row.findall("./" + W + "tc") for row in rows]
    if [len(value) for value in cells] != [2, 3, 3]:
        return False
    span = cells[0][0].find("./" + W + "tcPr/" + W + "gridSpan")
    start = cells[1][2].find("./" + W + "tcPr/" + W + "vMerge")
    continuation = cells[2][2].find("./" + W + "tcPr/" + W + "vMerge")
    text = lambda cell: "".join(cell.itertext())
    return (
        len(tables[1].findall(".//" + W + "gridSpan")) == 1
        and len(tables[1].findall(".//" + W + "vMerge")) == 2
        and
        span is not None and span.get(W + "val") == "2"
        and start is not None and start.get(W + "val") == "restart"
        and continuation is not None and continuation.get(W + "val") in (None, "continue")
        and text(cells[0][0]) == horizontal_text
        and text(cells[1][2]) == vertical_text
        and text(cells[2][2]) == ""
        and all(value in "".join(tables[1].itertext()) for value in outside_texts)
    )


def verify_picture_artifacts(
    docx: Path,
    pdf: Path,
    source: Path,
    *,
    expected_extents,
    expected_picture_anchor,
    expected_textbox,
):
    """Return strict relationship, geometry, and rendered-placement checks."""
    source = Path(source)
    source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
    source_pixels = _pixel_identity(source.read_bytes())
    with ZipFile(docx) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        relationships = ET.fromstring(package.read("word/_rels/document.xml.rels"))
        names = set(package.namelist())
        rels = {rel.get("Id"): rel for rel in relationships}
        picture_nodes = []
        textbox_nodes = []
        for node in document.findall(".//" + WP + "inline") + document.findall(".//" + WP + "anchor"):
            if node.find(".//" + W + "txbxContent") is not None:
                textbox_nodes.append(node)
            elif node.find(".//" + A + "blip") is not None:
                picture_nodes.append(node)

        resolved = []
        relationship_ok = len(picture_nodes) == len(expected_extents)
        for node in picture_nodes:
            blips = node.findall(".//" + A + "blip")
            rel = rels.get(blips[0].get(R + "embed")) if len(blips) == 1 else None
            target = rel.get("Target") if rel is not None else None
            if (
                rel is None
                or not rel.get("Type", "").endswith("/image")
                or rel.get("TargetMode") == "External"
                or not target
                or target.startswith("/")
            ):
                relationship_ok = False
                continue
            member = posixpath.normpath(posixpath.join("word", target.replace("\\", "/")))
            if member.startswith("../") or member not in names:
                relationship_ok = False
                continue
            payload = package.read(member)
            if hashlib.sha256(payload).hexdigest() != source_hash:
                relationship_ok = False
                continue
            resolved.append((node, member))

        extents = [_extent(node) for node, _member in resolved]
        geometry_ok = all(value is not None and min(value) > 0 for value in extents)
        geometry_ok = geometry_ok and _close_vectors(extents, expected_extents)
        picture_anchors = [node for node, _member in resolved if node.tag == WP + "anchor"]
        geometry_ok = geometry_ok and len(picture_anchors) == 1
        if picture_anchors:
            position = _position(picture_anchors[0])
            expected_x, expected_y, expected_wrap, expected_horizontal, expected_vertical = expected_picture_anchor
            geometry_ok = geometry_ok and position is not None and position[:2] == (expected_horizontal, expected_vertical)
            geometry_ok = geometry_ok and _close_vectors([position[2:]], [(expected_x, expected_y)])
            geometry_ok = geometry_ok and _wrap_name(picture_anchors[0]) == expected_wrap

        textbox_ok = len(textbox_nodes) == 1
        if textbox_nodes:
            textbox = textbox_nodes[0]
            extent = _extent(textbox)
            position = _position(textbox)
            left, top, width, height, wrap = expected_textbox
            textbox_ok = (
                textbox_ok and extent is not None and position is not None
                and position[:2] == ("page", "page")
                and _close_vectors([extent], [(width, height)])
                and _close_vectors([position[2:]], [(left, top)])
                and _wrap_name(textbox) == wrap
            )

    import fitz
    placements = []
    with fitz.open(pdf) as pages:
        for page in pages:
            seen = set()
            for raw in page.get_images(full=True):
                xref = raw[0]
                if xref in seen:
                    continue
                seen.add(xref)
                extracted = pages.extract_image(xref).get("image")
                if extracted and _pixel_identity(extracted) == source_pixels:
                    placements.extend((rect.width, rect.height) for rect in page.get_image_rects(xref))
    pdf_ok = all(min(vector) > 0 for vector in placements) and _close_vectors(placements, expected_extents)
    return {
        "relationships": relationship_ok and len(resolved) == len(expected_extents),
        "picture_geometry": geometry_ok,
        "textbox_geometry": textbox_ok,
        "pdf_placements": pdf_ok,
    }
