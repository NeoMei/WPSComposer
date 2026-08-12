"""Native heading numbering for generated DOCX files.

WpsComposer renders heading numbers (第一章 / 第一节 / 一、 / （一）) as
plain text baked into the heading text (see heading_numbering.py). Word/WPS
native multi-level numbering instead recomputes automatically whenever
headings are added, removed, or reordered.

``apply_native_numbering()`` rewrites a generated .docx in place:

1. strips the plain-text number prefix from Heading 1-4 paragraphs
2. injects a Chinese 4-level numbering definition (word/numbering.xml)
3. binds the Heading 2-4 *styles* to levels 1-3 (style-level ``numPr``,
   so new headings created with those styles are numbered automatically)
4. binds Heading 1 paragraphs that carried a 第N章 prefix to level 0
   per-paragraph, so TOC titles like "目  录" stay unnumbered

Idempotent: a document that already carries ``word/numbering.xml`` is
left untouched. Pure stdlib (xml.etree) — no lxml dependency.
"""

from __future__ import annotations

import io
import os
import re
import zipfile
import xml.etree.ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NUM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
NUM_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"

_NUM_ID = "100"
_ABSTRACT_ID = "100"

# Plain-text prefixes produced by heading_numbering.chinese scheme.
_PREFIX_PATTERNS = {
    "Heading 1": re.compile(r"^第[一二三四五六七八九十百]+章\s*"),
    "Heading 2": re.compile(r"^第[一二三四五六七八九十百]+节\s*"),
    "Heading 3": re.compile(r"^[一二三四五六七八九十]+、\s*"),
    "Heading 4": re.compile(r"^（[一二三四五六七八九十]+）\s*"),
}

# Style-level binding: Heading N style -> numbering level.
_STYLE_LEVELS = {"Heading 2": "1", "Heading 3": "2", "Heading 4": "3"}

_NUMBERING_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<w:numbering xmlns:w="{W}">'
    f'<w:abstractNum w:abstractNumId="{_ABSTRACT_ID}">'
    '<w:multiLevelType w:val="multilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
    '<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="第%1章"/>'
    '<w:lvlJc w:val="left"/></w:lvl>'
    '<w:lvl w:ilvl="1"><w:start w:val="1"/>'
    '<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="第%2节"/>'
    '<w:lvlJc w:val="left"/></w:lvl>'
    '<w:lvl w:ilvl="2"><w:start w:val="1"/>'
    '<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="%3、"/>'
    '<w:lvlJc w:val="left"/></w:lvl>'
    '<w:lvl w:ilvl="3"><w:start w:val="1"/>'
    '<w:numFmt w:val="chineseCounting"/><w:lvlText w:val="（%4）"/>'
    '<w:lvlJc w:val="left"/></w:lvl>'
    "</w:abstractNum>"
    f'<w:num w:numId="{_NUM_ID}"><w:abstractNumId w:val="{_ABSTRACT_ID}"/></w:num>'
    "</w:numbering>\n"
)


def _parse(data: bytes) -> ET.Element:
    """Parse XML bytes, keeping the original namespace prefixes."""
    namespaces = {}
    for _event, (prefix, uri) in ET.iterparse(
        io.BytesIO(data), events=("start-ns",)
    ):
        if prefix:  # skip the default-namespace registration (not writable)
            namespaces[prefix] = uri
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    return ET.fromstring(data)


def _insert_numpr(pPr: ET.Element, ilvl: str) -> None:
    """Insert a numPr element into pPr at a schema-valid position."""
    numPr = ET.Element(f"{{{W}}}numPr")
    il = ET.SubElement(numPr, f"{{{W}}}ilvl")
    il.set(f"{{{W}}}val", ilvl)
    ni = ET.SubElement(numPr, f"{{{W}}}numId")
    ni.set(f"{{{W}}}val", _NUM_ID)
    children = list(pPr)
    idx = 0
    pstyle = pPr.find(f"{{{W}}}pStyle")
    if pstyle is not None:
        idx = children.index(pstyle) + 1
    else:
        keep = pPr.find(f"{{{W}}}keepNext")
        if keep is not None:
            idx = children.index(keep) + 1
    pPr.insert(idx, numPr)


def _apply(root: ET.Element, name_of: dict) -> None:
    """Strip numbering prefixes from Heading 1-4 paragraphs in place."""
    for p in root.iter(f"{{{W}}}p"):
        pPr = p.find(f"{{{W}}}pPr")
        if pPr is None:
            continue
        pstyle = pPr.find(f"{{{W}}}pStyle")
        if pstyle is None:
            continue
        texts = p.findall(f".//{{{W}}}t")
        if not texts:
            continue
        heading_name = name_of.get(pstyle.get(f"{{{W}}}val"))
        pat = _PREFIX_PATTERNS.get(heading_name)
        if pat is None:
            continue
        full = "".join(t.text or "" for t in texts)
        new = pat.sub("", full)
        if new == full:
            continue
        texts[0].text = new
        for t in texts[1:]:
            t.text = ""
        if heading_name == "Heading 1":
            _insert_numpr(pPr, "0")


def apply_native_numbering(docx_path: str) -> bool:
    """Rewrite a .docx in place so heading numbers are native Word/WPS
    multi-level numbering instead of plain text.

    Returns True if the document was rewritten, False if it was skipped
    (no numbering.xml part, or it already carries one).
    """
    path = os.fspath(docx_path)
    with zipfile.ZipFile(path) as zin:
        names = set(zin.namelist())
        if "word/numbering.xml" in names:
            return False  # already native-numbered — idempotent
        document_xml = zin.read("word/document.xml")
        styles_xml = zin.read("word/styles.xml")
        ct_xml = zin.read("[Content_Types].xml")
        rels_path = "word/_rels/document.xml.rels"
        rels_xml = zin.read(rels_path) if rels_path in names else None
        others = {
            name: zin.read(name)
            for name in names
            if name
            not in {
                "word/document.xml",
                "word/styles.xml",
                "[Content_Types].xml",
                "word/_rels/document.xml.rels",
                "word/numbering.xml",
            }
        }

    # --- styles.xml: styleId -> style name map, then bind Heading 2-4 ---
    styles_root = _parse(styles_xml)
    name_of = {}
    for style in styles_root.iter(f"{{{W}}}style"):
        sid = style.get(f"{{{W}}}styleId")
        name_el = style.find(f"{{{W}}}name")
        if sid is not None and name_el is not None:
            name_of[sid] = name_el.get(f"{{{W}}}val")
    for style in styles_root.iter(f"{{{W}}}style"):
        sid = style.get(f"{{{W}}}styleId")
        level = _STYLE_LEVELS.get(name_of.get(sid, ""))
        if level is None:
            continue
        pPr = style.find(f"{{{W}}}pPr")
        if pPr is None:
            pPr = ET.Element(f"{{{W}}}pPr")
            style.append(pPr)
        if pPr.find(f"{{{W}}}numPr") is None:
            _insert_numpr(pPr, level)

    # --- document.xml: strip plain-text number prefixes ---
    doc_root = _parse(document_xml)
    _apply(doc_root, name_of)

    # --- [Content_Types].xml: register numbering part ---
    ct_root = _parse(ct_xml)
    if not any(
        ov.get("PartName") == "/word/numbering.xml"
        for ov in ct_root.iter(f"{{{CT}}}Override")
    ):
        ov = ET.SubElement(ct_root, f"{{{CT}}}Override")
        ov.set("PartName", "/word/numbering.xml")
        ov.set("ContentType", NUM_CT)

    # --- document.xml.rels: link the numbering part ---
    rel_root = _parse(rels_xml) if rels_xml is not None else None
    if rel_root is not None:
        if not any(
            r.get("Target") == "numbering.xml"
            for r in rel_root.iter(f"{{{REL}}}Relationship")
        ):
            rel = ET.SubElement(rel_root, f"{{{REL}}}Relationship")
            rel.set("Id", "rIdNativeNumbering")
            rel.set("Type", NUM_REL)
            rel.set("Target", "numbering.xml")

    def _dump(root: ET.Element) -> bytes:
        return ET.tostring(root, encoding="UTF-8", xml_declaration=True)

    # --- rewrite the package atomically ---
    tmp = f"{path}.numbering.tmp"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for name in names:
                if name == "word/document.xml":
                    zout.writestr(name, _dump(doc_root))
                elif name == "word/styles.xml":
                    zout.writestr(name, _dump(styles_root))
                elif name == "[Content_Types].xml":
                    zout.writestr(name, _dump(ct_root))
                elif name == rels_path:
                    zout.writestr(name, _dump(rel_root))
                else:
                    zout.writestr(name, others[name])
            zout.writestr("word/numbering.xml", _NUMBERING_XML)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return True
