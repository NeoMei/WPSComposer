"""Generate deterministic source-bound DOCX fixtures for Word style probes.

The generator uses only the Python standard library and explicit OOXML.  It
does not invoke Word, WPS, LibreOffice, AppleScript, or product code.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import hashlib
from io import BytesIO
import json
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
FIXED_CORE_TIME = "2026-09-10T00:00:00Z"
NUM_ID = 4242

for prefix, uri in (
    ("w", W),
    ("r", R),
    ("cp", CP),
    ("dc", DC),
    ("dcterms", DCTERMS),
    ("xsi", XSI),
    ("ep", EP),
    ("vt", VT),
):
    ET.register_namespace(prefix, uri)


SELECTED_STYLE_CONTRACT = {
    "Heading2": {
        "role": "applied-modified-builtin",
        "name": "heading 2",
        "style_type": "paragraph",
        "custom": False,
        "auto_redefine": True,
    },
    "Heading3": {
        "role": "unused-modified-builtin",
        "name": "heading 3",
        "style_type": "paragraph",
        "custom": False,
        "auto_redefine": True,
    },
    "WPSC_Custom_Paragraph": {
        "role": "applied-custom-paragraph",
        "name": "WPSC Custom Paragraph",
        "style_type": "paragraph",
        "custom": True,
        "auto_redefine": False,
    },
    "WPSC_Custom_Character": {
        "role": "applied-custom-character",
        "name": "WPSC Custom Character",
        "style_type": "character",
        "custom": True,
        "auto_redefine": False,
    },
    "WPSC_Numbered_Style": {
        "role": "applied-numbering-linked-custom",
        "name": "WPSC Numbered Style",
        "style_type": "paragraph",
        "custom": True,
        "auto_redefine": False,
    },
    "WPSC_Unused_Custom": {
        "role": "unused-custom-paragraph",
        "name": "WPSC Unused Custom",
        "style_type": "paragraph",
        "custom": True,
        "auto_redefine": False,
    },
}


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def attr(local: str, value: object) -> dict[str, str]:
    return {qn(local): str(value)}


def child(parent: ET.Element, local: str, **attrs: object) -> ET.Element:
    return ET.SubElement(parent, qn(local), {qn(key): str(value) for key, value in attrs.items()})


def xml_bytes(root: ET.Element) -> bytes:
    body = ET.tostring(root, encoding="utf-8", short_empty_elements=True)
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _content_types(*, include_numbering: bool) -> bytes:
    root = ET.Element(f"{{{CONTENT_TYPES}}}Types")
    for extension, content_type in (
        ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
        ("xml", "application/xml"),
    ):
        ET.SubElement(root, f"{{{CONTENT_TYPES}}}Default", Extension=extension, ContentType=content_type)
    overrides = [
        ("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
        ("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    ]
    if include_numbering:
        overrides.append((
            "/word/numbering.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        ))
    for part, content_type in overrides:
        ET.SubElement(root, f"{{{CONTENT_TYPES}}}Override", PartName=part, ContentType=content_type)
    return xml_bytes(root)


def _root_relationships() -> bytes:
    root = ET.Element(f"{{{PKG_REL}}}Relationships")
    for rid, target, relationship_type in (
        ("rId1", "word/document.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"),
        ("rId2", "docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"),
        ("rId3", "docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
    ):
        ET.SubElement(root, f"{{{PKG_REL}}}Relationship", Id=rid, Target=target, Type=relationship_type)
    return xml_bytes(root)


def _document_relationships(*, include_numbering: bool) -> bytes:
    root = ET.Element(f"{{{PKG_REL}}}Relationships")
    relationships = [
        ("rId1", "styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"),
    ]
    if include_numbering:
        relationships.append((
            "rId2",
            "numbering.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        ))
    for rid, target, relationship_type in relationships:
        ET.SubElement(root, f"{{{PKG_REL}}}Relationship", Id=rid, Target=target, Type=relationship_type)
    return xml_bytes(root)


def _core_properties(title: str) -> bytes:
    root = ET.Element(f"{{{CP}}}coreProperties")
    ET.SubElement(root, f"{{{DC}}}title").text = title
    ET.SubElement(root, f"{{{DC}}}creator").text = "WPSComposer fixture generator"
    ET.SubElement(root, f"{{{CP}}}lastModifiedBy").text = "WPSComposer fixture generator"
    ET.SubElement(root, f"{{{CP}}}revision").text = "1"
    created = ET.SubElement(root, f"{{{DCTERMS}}}created", {f"{{{XSI}}}type": "dcterms:W3CDTF"})
    created.text = FIXED_CORE_TIME
    modified = ET.SubElement(root, f"{{{DCTERMS}}}modified", {f"{{{XSI}}}type": "dcterms:W3CDTF"})
    modified.text = FIXED_CORE_TIME
    return xml_bytes(root)


def _app_properties() -> bytes:
    root = ET.Element(f"{{{EP}}}Properties")
    ET.SubElement(root, f"{{{EP}}}Application").text = "WPSComposer deterministic OOXML fixture"
    ET.SubElement(root, f"{{{EP}}}AppVersion").text = "1.0"
    return xml_bytes(root)


def _add_run(paragraph: ET.Element, text: str, style_id: str | None = None) -> ET.Element:
    run = child(paragraph, "r")
    if style_id:
        rpr = child(run, "rPr")
        child(rpr, "rStyle", val=style_id)
    text_node = child(run, "t")
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = text
    return run


def _paragraph(parent: ET.Element, text: str, style_id: str = "Normal", *, section_break: bool = False) -> ET.Element:
    paragraph = child(parent, "p")
    ppr = child(paragraph, "pPr")
    child(ppr, "pStyle", val=style_id)
    if section_break:
        ppr.append(_section_properties())
    if text:
        _add_run(paragraph, text)
    return paragraph


def _section_properties() -> ET.Element:
    sect = ET.Element(qn("sectPr"))
    child(sect, "pgSz", w=12240, h=15840)
    child(sect, "pgMar", top=1440, right=1440, bottom=1440, left=1440, header=720, footer=720, gutter=0)
    child(sect, "cols", space=720)
    child(sect, "docGrid", linePitch=360)
    return sect


def _add_field(paragraph: ET.Element) -> None:
    begin = child(paragraph, "r")
    child(begin, "fldChar", fldCharType="begin")
    instruction = child(paragraph, "r")
    instr = child(instruction, "instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " PAGE "
    separate = child(paragraph, "r")
    child(separate, "fldChar", fldCharType="separate")
    _add_run(paragraph, "7", "WPSC_Custom_Character")
    end = child(paragraph, "r")
    child(end, "fldChar", fldCharType="end")


def _add_table(parent: ET.Element, text: str, style_id: str) -> None:
    table = child(parent, "tbl")
    tbl_pr = child(table, "tblPr")
    child(tbl_pr, "tblW", w=8640, type="dxa")
    grid = child(table, "tblGrid")
    child(grid, "gridCol", w=8640)
    row = child(table, "tr")
    cell = child(row, "tc")
    tc_pr = child(cell, "tcPr")
    child(tc_pr, "tcW", w=8640, type="dxa")
    _paragraph(cell, text, style_id)


def _document_xml(kind: str) -> bytes:
    document = ET.Element(qn("document"))
    body = child(document, "body")
    if kind == "empty":
        _paragraph(body, "", "Normal")
    elif kind == "nonempty":
        _paragraph(body, "NONEMPTY NORMAL", "Normal")
        _paragraph(body, "NONEMPTY APPLIED BUILTIN", "Heading2")
        custom = _paragraph(body, "NONEMPTY CUSTOM PARAGRAPH ", "WPSC_Custom_Paragraph")
        _add_run(custom, "NONEMPTY CUSTOM CHARACTER", "WPSC_Custom_Character")
        _paragraph(body, "NONEMPTY NUMBERED", "WPSC_Numbered_Style")
    elif kind == "multisection":
        first = _paragraph(body, "MULTI SECTION 1 PARAGRAPH ", "WPSC_Custom_Paragraph")
        _add_run(first, "MULTI SECTION 1 CHARACTER", "WPSC_Custom_Character")
        _paragraph(body, "MULTI SECTION 1 END", "Normal", section_break=True)
        _add_table(body, "MULTI SECTION 2 TABLE CELL", "WPSC_Custom_Paragraph")
        _paragraph(body, "MULTI SECTION 2 NUMBERED", "WPSC_Numbered_Style")
        _paragraph(body, "MULTI SECTION 2 END", "Normal", section_break=True)
        _paragraph(body, "MULTI SECTION 3 BUILTIN", "Heading2")
        field_paragraph = _paragraph(body, "MULTI SECTION 3 FIELD ", "WPSC_Custom_Paragraph")
        _add_field(field_paragraph)
    else:
        raise ValueError(f"unknown fixture kind: {kind}")
    body.append(_section_properties())
    return xml_bytes(document)


def _add_style(
    root: ET.Element,
    *,
    style_id: str,
    style_type: str,
    name: str,
    default: bool = False,
    custom: bool = False,
    based_on: str | None = None,
    next_style: str | None = None,
    link: str | None = None,
    auto_redefine: bool = False,
    qformat: bool = False,
    paragraph_values: dict[str, object] | None = None,
    run_values: dict[str, object] | None = None,
    num_id: int | None = None,
) -> ET.Element:
    attributes = {qn("type"): style_type, qn("styleId"): style_id}
    if default:
        attributes[qn("default")] = "1"
    if custom:
        attributes[qn("customStyle")] = "1"
    style = ET.SubElement(root, qn("style"), attributes)
    child(style, "name", val=name)
    if based_on:
        child(style, "basedOn", val=based_on)
    if next_style:
        child(style, "next", val=next_style)
    if link:
        child(style, "link", val=link)
    if auto_redefine:
        child(style, "autoRedefine")
    if qformat:
        child(style, "qFormat")
    if paragraph_values or num_id is not None:
        ppr = child(style, "pPr")
        if num_id is not None:
            num_pr = child(ppr, "numPr")
            child(num_pr, "ilvl", val=0)
            child(num_pr, "numId", val=num_id)
        for key, value in (paragraph_values or {}).items():
            if key == "spacing":
                child(ppr, key, after=value)
            else:
                child(ppr, key, val=value)
    if run_values:
        rpr = child(style, "rPr")
        for key, value in run_values.items():
            if value is True:
                child(rpr, key)
            else:
                child(rpr, key, val=value)
    return style


def _styles_xml(selected: bool) -> bytes:
    root = ET.Element(qn("styles"))
    defaults = child(root, "docDefaults")
    rpr_default = child(defaults, "rPrDefault")
    rpr = child(rpr_default, "rPr")
    fonts = child(rpr, "rFonts", ascii="Arial", hAnsi="Arial", eastAsia="Arial")
    fonts.set(qn("cs"), "Arial")
    child(rpr, "sz", val=22)
    child(rpr, "szCs", val=22)
    ppr_default = child(defaults, "pPrDefault")
    ppr = child(ppr_default, "pPr")
    child(ppr, "spacing", after=120, line=276, lineRule="auto")

    _add_style(root, style_id="Normal", style_type="paragraph", name="Normal", default=True, qformat=True)
    _add_style(root, style_id="DefaultParagraphFont", style_type="character", name="Default Paragraph Font", default=True)
    _add_style(root, style_id="TableNormal", style_type="table", name="Normal Table", default=True)
    if selected:
        _add_style(
            root,
            style_id="Heading2",
            style_type="paragraph",
            name="heading 2",
            based_on="Normal",
            next_style="Normal",
            auto_redefine=True,
            qformat=True,
            paragraph_values={"keepNext": 1, "spacing": 120, "outlineLvl": 1},
            run_values={"b": True, "color": "1F4E79", "sz": 28},
        )
        _add_style(
            root,
            style_id="Heading3",
            style_type="paragraph",
            name="heading 3",
            based_on="Normal",
            next_style="Normal",
            auto_redefine=True,
            qformat=True,
            paragraph_values={"keepNext": 1, "outlineLvl": 2},
            run_values={"i": True, "color": "7030A0", "sz": 24},
        )
        _add_style(
            root,
            style_id="WPSC_Custom_Paragraph",
            style_type="paragraph",
            name="WPSC Custom Paragraph",
            custom=True,
            based_on="Normal",
            link="WPSC_Custom_Character",
            qformat=True,
            paragraph_values={"spacing": 180},
            run_values={"color": "365F91"},
        )
        _add_style(
            root,
            style_id="WPSC_Custom_Character",
            style_type="character",
            name="WPSC Custom Character",
            custom=True,
            based_on="DefaultParagraphFont",
            link="WPSC_Custom_Paragraph",
            run_values={"b": True, "color": "C00000"},
        )
        _add_style(
            root,
            style_id="WPSC_Numbered_Style",
            style_type="paragraph",
            name="WPSC Numbered Style",
            custom=True,
            based_on="Normal",
            qformat=True,
            num_id=NUM_ID,
        )
        _add_style(
            root,
            style_id="WPSC_Unused_Custom",
            style_type="paragraph",
            name="WPSC Unused Custom",
            custom=True,
            based_on="Normal",
            run_values={"i": True, "color": "008000"},
        )
    return xml_bytes(root)


def _numbering_xml(selected: bool) -> bytes:
    root = ET.Element(qn("numbering"))
    if selected:
        abstract = child(root, "abstractNum", abstractNumId=NUM_ID)
        child(abstract, "multiLevelType", val="singleLevel")
        level = child(abstract, "lvl", ilvl=0)
        child(level, "start", val=1)
        child(level, "numFmt", val="decimal")
        child(level, "lvlText", val="%1.")
        child(level, "suff", val="tab")
        ppr = child(level, "pPr")
        tabs = child(ppr, "tabs")
        child(tabs, "tab", val="num", pos=720)
        child(ppr, "ind", left=720, hanging=360)
        numbering = child(root, "num", numId=NUM_ID)
        child(numbering, "abstractNumId", val=NUM_ID)
    return xml_bytes(root)


def _package(kind: str) -> bytes:
    selected = kind != "empty"
    parts = {
        "[Content_Types].xml": _content_types(include_numbering=selected),
        "_rels/.rels": _root_relationships(),
        "docProps/app.xml": _app_properties(),
        "docProps/core.xml": _core_properties(f"WPSComposer {kind} style matrix"),
        "word/_rels/document.xml.rels": _document_relationships(include_numbering=selected),
        "word/document.xml": _document_xml(kind),
        "word/styles.xml": _styles_xml(selected),
    }
    if selected:
        parts["word/numbering.xml"] = _numbering_xml(selected)
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as package:
        for name in sorted(parts):
            info = zipfile.ZipInfo(name, FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            info.extra = b""
            info.comment = b""
            package.writestr(info, parts[name], compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    return buffer.getvalue()


def _relationship_rows(xml: bytes) -> list[dict[str, str]]:
    root = ET.fromstring(xml)
    return [
        {"id": row.attrib["Id"], "target": row.attrib["Target"], "type": row.attrib["Type"]}
        for row in root
    ]


def _style_rows(styles_xml: bytes) -> dict[str, dict[str, object]]:
    root = ET.fromstring(styles_xml)
    rows: dict[str, dict[str, object]] = {}
    for style in root.findall(qn("style")):
        style_id = style.attrib[qn("styleId")]
        name = style.find(qn("name"))
        based_on = style.find(qn("basedOn"))
        link = style.find(qn("link"))
        num_id = style.find(f"{qn('pPr')}/{qn('numPr')}/{qn('numId')}")
        rows[style_id] = {
            "style_id": style_id,
            "style_type": style.attrib[qn("type")],
            "name": name.attrib[qn("val")] if name is not None else None,
            "custom": style.attrib.get(qn("customStyle")) == "1",
            "default": style.attrib.get(qn("default")) == "1",
            "auto_redefine": style.find(qn("autoRedefine")) is not None,
            "based_on": based_on.attrib[qn("val")] if based_on is not None else None,
            "link": link.attrib[qn("val")] if link is not None else None,
            "num_id": int(num_id.attrib[qn("val")]) if num_id is not None else None,
        }
    return rows


def _numbering_rows(numbering_xml: bytes) -> tuple[list[dict[str, object]], dict[int, int]]:
    root = ET.fromstring(numbering_xml)
    abstracts: dict[int, dict[str, object]] = {}
    for abstract in root.findall(qn("abstractNum")):
        abstract_id = int(abstract.attrib[qn("abstractNumId")])
        level = abstract.find(qn("lvl"))
        if level is None:
            continue
        num_format = level.find(qn("numFmt"))
        level_text = level.find(qn("lvlText"))
        abstracts[abstract_id] = {
            "abstract_num_id": abstract_id,
            "level": int(level.attrib[qn("ilvl")]),
            "num_format": num_format.attrib[qn("val")] if num_format is not None else None,
            "level_text": level_text.attrib[qn("val")] if level_text is not None else None,
        }
    num_to_abstract: dict[int, int] = {}
    rows: list[dict[str, object]] = []
    for num in root.findall(qn("num")):
        num_id = int(num.attrib[qn("numId")])
        abstract_ref = num.find(qn("abstractNumId"))
        if abstract_ref is None:
            continue
        abstract_id = int(abstract_ref.attrib[qn("val")])
        num_to_abstract[num_id] = abstract_id
        rows.append({"num_id": num_id, **abstracts[abstract_id]})
    return sorted(rows, key=lambda row: int(row["num_id"])), num_to_abstract


def _paragraph_style_locations(
    paragraph: ET.Element,
    *,
    path: str,
    section: int,
) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    paragraph_style = paragraph.find(f"{qn('pPr')}/{qn('pStyle')}")
    paragraph_text = "".join(node.text or "" for node in paragraph.iter(qn("t")))
    if paragraph_style is not None:
        rows.append({
            "kind": "paragraph",
            "path": path,
            "section": section,
            "style_id": paragraph_style.attrib[qn("val")],
            "text": paragraph_text,
        })
    field_state = "outside"
    for run_index, run in enumerate(paragraph.findall(qn("r")), 1):
        field_char = run.find(qn("fldChar"))
        if field_char is not None:
            field_type = field_char.attrib.get(qn("fldCharType"))
            if field_type == "begin":
                field_state = "instruction"
            elif field_type == "separate":
                field_state = "result"
            elif field_type == "end":
                field_state = "outside"
        run_style = run.find(f"{qn('rPr')}/{qn('rStyle')}")
        if run_style is not None:
            row = {
                "kind": "run",
                "path": f"{path}/r[{run_index}]",
                "section": section,
                "style_id": run_style.attrib[qn("val")],
                "text": "".join(node.text or "" for node in run.findall(qn("t"))),
            }
            if field_state != "outside":
                row["field_role"] = field_state
            rows.append(row)
    return rows


def _applied_locations(document_xml: bytes) -> tuple[int, list[dict[str, object]]]:
    root = ET.fromstring(document_xml)
    body = root.find(qn("body"))
    if body is None:
        raise AssertionError("document body missing")
    rows: list[dict[str, object]] = []
    section = 1
    paragraph_index = 0
    table_index = 0
    for element in body:
        if element.tag == qn("p"):
            paragraph_index += 1
            rows.extend(_paragraph_style_locations(element, path=f"body/p[{paragraph_index}]", section=section))
            if element.find(f"{qn('pPr')}/{qn('sectPr')}") is not None:
                section += 1
        elif element.tag == qn("tbl"):
            table_index += 1
            for row_index, table_row in enumerate(element.findall(qn("tr")), 1):
                for cell_index, cell in enumerate(table_row.findall(qn("tc")), 1):
                    for cell_paragraph_index, paragraph in enumerate(cell.findall(qn("p")), 1):
                        rows.extend(_paragraph_style_locations(
                            paragraph,
                            path=f"body/tbl[{table_index}]/tr[{row_index}]/tc[{cell_index}]/p[{cell_paragraph_index}]",
                            section=section,
                        ))
    final_section = body.find(qn("sectPr"))
    if final_section is None:
        raise AssertionError("final section properties missing")
    return section, rows


def _inspect_package(path: Path, kind: str) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        infos = package.infolist()
        names = [info.filename for info in infos]
        if names != sorted(names):
            raise AssertionError("ZIP members are not sorted")
        if any(info.date_time != FIXED_ZIP_TIME or info.extra or info.comment for info in infos):
            raise AssertionError("ZIP metadata is not canonical")
        parts = {name: package.read(name) for name in names}
    expected_parts = [
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/app.xml",
        "docProps/core.xml",
        "word/_rels/document.xml.rels",
        "word/document.xml",
        "word/styles.xml",
    ]
    if kind != "empty":
        expected_parts.append("word/numbering.xml")
        expected_parts.sort()
    if names != expected_parts:
        raise AssertionError("unexpected package members")
    styles = _style_rows(parts["word/styles.xml"])
    numbering, num_to_abstract = _numbering_rows(parts["word/numbering.xml"]) if kind != "empty" else ([], {})
    section_count, applied = _applied_locations(parts["word/document.xml"])
    custom_ids = sorted(style_id for style_id, row in styles.items() if row["custom"])
    selected_rows: list[dict[str, object]] = []
    if kind != "empty":
        for style_id, contract in SELECTED_STYLE_CONTRACT.items():
            style = styles.get(style_id)
            if style is None:
                raise AssertionError(f"selected style missing: {style_id}")
            locations = [row for row in applied if row["style_id"] == style_id]
            selected = {**style, "role": contract["role"], "applied_locations": locations}
            if style["num_id"] is not None:
                abstract_id = num_to_abstract.get(int(style["num_id"]))
                level = next(row["level"] for row in numbering if row["num_id"] == style["num_id"])
                selected["numbering"] = {
                    "num_id": style["num_id"],
                    "abstract_num_id": abstract_id,
                    "level": level,
                }
            selected_rows.append(selected)
    result = {
        "file": path.name,
        "sha256": sha256_file(path),
        "package_entries": names,
        "relationships": {
            "package": _relationship_rows(parts["_rels/.rels"]),
            "document": _relationship_rows(parts["word/_rels/document.xml.rels"]),
        },
        "section_count": section_count,
        "custom_style_ids": custom_ids,
        "numbering": numbering,
        "selected_styles": selected_rows,
        "all_applied_style_locations": applied,
    }
    _assert_fixture(kind, result)
    return result


def _assert_fixture(kind: str, result: dict[str, object]) -> None:
    if kind == "empty":
        if result["section_count"] != 1 or result["custom_style_ids"] or result["numbering"]:
            raise AssertionError("empty fixture is not minimal")
        return
    selected = {row["style_id"]: row for row in result["selected_styles"]}
    if set(selected) != set(SELECTED_STYLE_CONTRACT):
        raise AssertionError("selected style coverage mismatch")
    for style_id, expected in SELECTED_STYLE_CONTRACT.items():
        for key in ("name", "style_type", "custom", "auto_redefine"):
            if selected[style_id][key] != expected[key]:
                raise AssertionError(f"{kind} {style_id} {key} mismatch")
    if selected["Heading3"]["applied_locations"]:
        raise AssertionError("modified unused built-in style became applied")
    if selected["WPSC_Unused_Custom"]["applied_locations"]:
        raise AssertionError("unused custom style became applied")
    for style_id in ("Heading2", "WPSC_Custom_Paragraph", "WPSC_Custom_Character", "WPSC_Numbered_Style"):
        if not selected[style_id]["applied_locations"]:
            raise AssertionError(f"required style is not applied: {style_id}")
    if result["numbering"] != [{
        "num_id": NUM_ID,
        "abstract_num_id": NUM_ID,
        "level": 0,
        "num_format": "decimal",
        "level_text": "%1.",
    }]:
        raise AssertionError("numbering relationship mismatch")
    if kind == "nonempty" and result["section_count"] != 1:
        raise AssertionError("nonempty fixture must have one section")
    if kind == "multisection":
        if result["section_count"] != 3:
            raise AssertionError("multisection fixture must have exactly three sections")
        locations = [row for style in selected.values() for row in style["applied_locations"]]
        if not {row["section"] for row in locations}.issuperset({1, 2, 3}):
            raise AssertionError("selected styles do not cover all sections")
        if not any("/tbl[1]/" in row["path"] for row in locations):
            raise AssertionError("table-cell style application missing")
        if not any(row.get("field_role") == "result" for row in locations):
            raise AssertionError("field-result character style application missing")


def generate(out_dir: Path) -> dict[str, object]:
    out_dir.mkdir(parents=True, exist_ok=True)
    fixtures: dict[str, dict[str, object]] = {}
    for kind in ("empty", "nonempty", "multisection"):
        path = out_dir / f"{kind}.docx"
        path.write_bytes(_package(kind))
        fixtures[kind] = _inspect_package(path, kind)
    preflight = {
        "schema": "wpscomposer-word-style-matrix-v1",
        "generated_utc": FIXED_CORE_TIME,
        "purpose": "source-bound native style-classifier matrix; no Office acceptance claim",
        "fixtures": fixtures,
    }
    preflight_path = out_dir / "preflight.json"
    preflight_path.write_text(json.dumps(preflight, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    generator_path = Path(__file__).resolve()
    hashes = {
        "schema": "wpscomposer-word-style-matrix-hashes-v1",
        "generator": {"file": generator_path.name, "sha256": sha256_file(generator_path)},
        "outputs": {
            name: sha256_file(out_dir / name)
            for name in ("empty.docx", "nonempty.docx", "multisection.docx", "preflight.json")
        },
    }
    (out_dir / "HASHES.json").write_text(
        json.dumps(hashes, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return preflight


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    result = generate(args.out_dir)
    summary = {
        name: {
            "sha256": row["sha256"],
            "sections": row["section_count"],
            "custom_styles": len(row["custom_style_ids"]),
            "numbering_definitions": len(row["numbering"]),
        }
        for name, row in result["fixtures"].items()
    }
    print(json.dumps(summary, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
