"""Convert visible heading prefixes into native Word/WPS numbering.

The post-pass preserves unrelated list definitions already present in a DOCX,
adds one WPSComposer-owned multi-level definition, and links Heading 1-4 to it.
It supports both the formal Chinese hierarchy and the common bid-document
hybrid hierarchy (第一章 / 1.1 / 1.1.1 / 关键工法01). Re-running the pass is
idempotent.
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
_NSID = "53575043"  # WPSComposer-owned definition marker.

_CHINESE_PATTERNS = {
    "Heading 1": re.compile(r"^第[一二三四五六七八九十百]+章\s*"),
    "Heading 2": re.compile(r"^第[一二三四五六七八九十百]+节\s*"),
    "Heading 3": re.compile(r"^[一二三四五六七八九十]+、\s*"),
    "Heading 4": re.compile(r"^（[一二三四五六七八九十]+）\s*"),
}
_HYBRID_PATTERNS = {
    "Heading 1": _CHINESE_PATTERNS["Heading 1"],
    "Heading 2": re.compile(r"^\d+\.\d+\s+"),
    "Heading 3": re.compile(r"^\d+\.\d+\.\d+\s+"),
    "Heading 4": re.compile(r"^关键工法\d{1,3}[：:]\s*"),
}
_STYLE_LEVELS = {
    "Heading 1": "0",
    "Heading 2": "1",
    "Heading 3": "2",
    "Heading 4": "3",
}

_NUMBERING_CHILD_ORDER = (
    "numPicBullet",
    "abstractNum",
    "num",
    "numIdMacAtCleanup",
)
_PPR_CHILD_ORDER = (
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
)


def _parse(data: bytes) -> ET.Element:
    namespaces = {}
    for _event, (prefix, uri) in ET.iterparse(
        io.BytesIO(data), events=("start-ns",)
    ):
        if prefix and not re.fullmatch(r"ns\d+", prefix):
            namespaces[prefix] = uri
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)
    ET.register_namespace("w", W)
    return ET.fromstring(data)


def _dump(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True)


def _attr(name: str) -> str:
    return f"{{{W}}}{name}"


def _full_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.iter(f"{{{W}}}t"))


def _strip_prefix(paragraph: ET.Element, pattern: re.Pattern) -> bool:
    """Remove only the matched prefix while preserving remaining runs."""
    match = pattern.match(_full_text(paragraph))
    if match is None:
        return False
    remaining = match.end()
    for node in paragraph.iter(f"{{{W}}}t"):
        text = node.text or ""
        if remaining >= len(text):
            node.text = ""
            remaining -= len(text)
        elif remaining:
            node.text = text[remaining:]
            remaining = 0
    return True


def _ensure_ppr(element: ET.Element) -> ET.Element:
    ppr = element.find(f"{{{W}}}pPr")
    if ppr is None:
        ppr = ET.Element(f"{{{W}}}pPr")
        if element.tag == f"{{{W}}}style":
            following = {
                f"{{{W}}}rPr",
                f"{{{W}}}tblPr",
                f"{{{W}}}trPr",
                f"{{{W}}}tcPr",
                f"{{{W}}}tblStylePr",
            }
            index = next(
                (
                    index
                    for index, child in enumerate(element)
                    if child.tag in following
                ),
                len(element),
            )
            element.insert(index, ppr)
        else:
            element.insert(0, ppr)
    return ppr


def _local_name(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def _insert_in_schema_order(
    parent: ET.Element, child: ET.Element, order: tuple[str, ...]
) -> int:
    """Insert *child* before the first known particle that must follow it."""
    rank = {name: index for index, name in enumerate(order)}
    child_rank = rank[_local_name(child)]
    for index, existing in enumerate(parent):
        existing_rank = rank.get(_local_name(existing))
        if existing_rank is not None and existing_rank > child_rank:
            parent.insert(index, child)
            return index
    parent.append(child)
    return len(parent) - 1


def _set_numpr(ppr: ET.Element, ilvl: str, num_id: str) -> bool:
    current = ppr.find(f"{{{W}}}numPr")
    values_match = False
    original_index = None
    if current is not None:
        level = current.find(f"{{{W}}}ilvl")
        number = current.find(f"{{{W}}}numId")
        values_match = (
            level is not None
            and level.get(_attr("val")) == ilvl
            and number is not None
            and number.get(_attr("val")) == num_id
        )
        original_index = list(ppr).index(current)
        ppr.remove(current)
    if values_match:
        numpr = current
    else:
        numpr = ET.Element(f"{{{W}}}numPr")
        level = ET.SubElement(numpr, f"{{{W}}}ilvl")
        level.set(_attr("val"), ilvl)
        number = ET.SubElement(numpr, f"{{{W}}}numId")
        number.set(_attr("val"), num_id)
    index = _insert_in_schema_order(ppr, numpr, _PPR_CHILD_ORDER)
    return not values_match or original_index != index


def _next_id(root: ET.Element, tag: str, attribute: str, preferred: str) -> str:
    used = {
        int(value)
        for node in root.findall(f"{{{W}}}{tag}")
        if (value := node.get(_attr(attribute))) is not None and value.isdigit()
    }
    candidate = int(preferred)
    while candidate in used:
        candidate += 1
    return str(candidate)


def _add_level(
    abstract: ET.Element,
    level: int,
    number_format: str,
    text: str,
    *,
    legal: bool = False,
    never_restart: bool = False,
) -> None:
    node = ET.SubElement(abstract, f"{{{W}}}lvl")
    node.set(_attr("ilvl"), str(level))
    start = ET.SubElement(node, f"{{{W}}}start")
    start.set(_attr("val"), "1")
    fmt = ET.SubElement(node, f"{{{W}}}numFmt")
    fmt.set(_attr("val"), number_format)
    if never_restart:
        restart = ET.SubElement(node, f"{{{W}}}lvlRestart")
        restart.set(_attr("val"), "0")
    if legal:
        ET.SubElement(node, f"{{{W}}}isLgl")
    label = ET.SubElement(node, f"{{{W}}}lvlText")
    label.set(_attr("val"), text)
    justify = ET.SubElement(node, f"{{{W}}}lvlJc")
    justify.set(_attr("val"), "left")


def _append_definition(
    numbering: ET.Element, abstract_id: str, num_id: str, scheme: str
) -> None:
    abstract = ET.Element(f"{{{W}}}abstractNum")
    abstract.set(_attr("abstractNumId"), abstract_id)
    nsid = ET.SubElement(abstract, f"{{{W}}}nsid")
    nsid.set(_attr("val"), _NSID)
    multilevel = ET.SubElement(abstract, f"{{{W}}}multiLevelType")
    multilevel.set(_attr("val"), "multilevel")
    if scheme == "hybrid":
        _add_level(abstract, 0, "chineseCounting", "第%1章")
        _add_level(abstract, 1, "decimal", "%1.%2", legal=True)
        _add_level(abstract, 2, "decimal", "%1.%2.%3", legal=True)
        _add_level(
            abstract,
            3,
            "decimalZero",
            "关键工法%4：",
            never_restart=True,
        )
    else:
        _add_level(abstract, 0, "chineseCounting", "第%1章")
        _add_level(abstract, 1, "chineseCounting", "第%2节")
        _add_level(abstract, 2, "chineseCounting", "%3、")
        _add_level(abstract, 3, "chineseCounting", "（%4）")
    _insert_in_schema_order(numbering, abstract, _NUMBERING_CHILD_ORDER)
    number = ET.Element(f"{{{W}}}num")
    number.set(_attr("numId"), num_id)
    reference = ET.SubElement(number, f"{{{W}}}abstractNumId")
    reference.set(_attr("val"), abstract_id)
    _insert_in_schema_order(numbering, number, _NUMBERING_CHILD_ORDER)


def _owned_ids(numbering: ET.Element) -> tuple[str, str] | None:
    abstract_id = None
    for abstract in numbering.findall(f"{{{W}}}abstractNum"):
        nsid = abstract.find(f"{{{W}}}nsid")
        if nsid is not None and nsid.get(_attr("val")) == _NSID:
            abstract_id = abstract.get(_attr("abstractNumId"))
            break
    if abstract_id is None:
        return None
    for number in numbering.findall(f"{{{W}}}num"):
        reference = number.find(f"{{{W}}}abstractNumId")
        if reference is not None and reference.get(_attr("val")) == abstract_id:
            return abstract_id, number.get(_attr("numId"))
    return None


def _definition_scheme(numbering: ET.Element, abstract_id: str) -> str:
    for abstract in numbering.findall(f"{{{W}}}abstractNum"):
        if abstract.get(_attr("abstractNumId")) != abstract_id:
            continue
        labels = {
            node.get(_attr("val"))
            for node in abstract.findall(f".//{{{W}}}lvlText")
        }
        return "hybrid" if "%1.%2" in labels else "chinese"
    return "chinese"


def _style_names(styles: ET.Element) -> dict[str, str]:
    names = {}
    for style in styles.iter(f"{{{W}}}style"):
        style_id = style.get(_attr("styleId"))
        name = style.find(f"{{{W}}}name")
        if style_id is not None and name is not None:
            raw_name = name.get(_attr("val")) or ""
            match = re.fullmatch(
                r"heading\s*([1-4])", raw_name, flags=re.IGNORECASE
            ) or re.fullmatch(
                r"heading\s*([1-4])", style_id, flags=re.IGNORECASE
            )
            names[style_id] = (
                f"Heading {match.group(1)}" if match else raw_name
            )
    return names


def _detect_scheme(document: ET.Element, names: dict[str, str]) -> str:
    for paragraph in document.iter(f"{{{W}}}p"):
        pstyle = paragraph.find(f"{{{W}}}pPr/{{{W}}}pStyle")
        if pstyle is None:
            continue
        heading = names.get(pstyle.get(_attr("val")))
        if heading in {"Heading 2", "Heading 3"}:
            if _HYBRID_PATTERNS[heading].match(_full_text(paragraph)):
                return "hybrid"
    return "chinese"


def apply_native_numbering(docx_path: str) -> bool:
    """Apply or repair native multi-level heading numbering in *docx_path*."""
    path = os.fspath(docx_path)
    with zipfile.ZipFile(path) as source:
        members = {name: source.read(name) for name in source.namelist()}

    document = _parse(members["word/document.xml"])
    styles = _parse(members["word/styles.xml"])
    content_types = _parse(members["[Content_Types].xml"])
    rels_path = "word/_rels/document.xml.rels"
    relationships = (
        _parse(members[rels_path])
        if rels_path in members
        else ET.Element(f"{{{REL}}}Relationships")
    )
    numbering = (
        _parse(members["word/numbering.xml"])
        if "word/numbering.xml" in members
        else ET.Element(f"{{{W}}}numbering")
    )
    names = _style_names(styles)
    owned = _owned_ids(numbering)
    changed = False
    if owned is None:
        abstract_id = _next_id(
            numbering, "abstractNum", "abstractNumId", _ABSTRACT_ID
        )
        num_id = _next_id(numbering, "num", "numId", _NUM_ID)
        scheme = _detect_scheme(document, names)
        _append_definition(numbering, abstract_id, num_id, scheme)
        changed = True
    else:
        abstract_id, num_id = owned
        scheme = _definition_scheme(numbering, abstract_id)

    patterns = _HYBRID_PATTERNS if scheme == "hybrid" else _CHINESE_PATTERNS

    # Bind styles so newly inserted headings inherit the list. Paragraphs
    # intentionally outside the hierarchy get a direct numId=0 override.
    for style in styles.iter(f"{{{W}}}style"):
        heading = names.get(style.get(_attr("styleId")), "")
        level = _STYLE_LEVELS.get(heading)
        if level is not None:
            changed |= _set_numpr(_ensure_ppr(style), level, num_id)

    for paragraph in document.iter(f"{{{W}}}p"):
        pstyle = paragraph.find(f"{{{W}}}pPr/{{{W}}}pStyle")
        if pstyle is None:
            continue
        heading = names.get(pstyle.get(_attr("val")))
        pattern = patterns.get(heading)
        if pattern is None:
            continue
        ppr = _ensure_ppr(paragraph)
        existing = ppr.find(f"{{{W}}}numPr/{{{W}}}numId")
        already_owned = (
            existing is not None and existing.get(_attr("val")) == num_id
        )
        if already_owned:
            continue
        if _strip_prefix(paragraph, pattern):
            level = str(int(heading[-1]) - 1)
            changed |= _set_numpr(ppr, level, num_id)
            changed = True
        else:
            changed |= _set_numpr(ppr, "0", "0")

    if not any(
        node.get("PartName") == "/word/numbering.xml"
        for node in content_types.iter(f"{{{CT}}}Override")
    ):
        node = ET.SubElement(content_types, f"{{{CT}}}Override")
        node.set("PartName", "/word/numbering.xml")
        node.set("ContentType", NUM_CT)
        changed = True
    if not any(
        node.get("Target") == "numbering.xml"
        for node in relationships.iter(f"{{{REL}}}Relationship")
    ):
        used_ids = {
            node.get("Id")
            for node in relationships.iter(f"{{{REL}}}Relationship")
        }
        relation_id = "rIdNativeNumbering"
        suffix = 2
        while relation_id in used_ids:
            relation_id = f"rIdNativeNumbering{suffix}"
            suffix += 1
        node = ET.SubElement(relationships, f"{{{REL}}}Relationship")
        node.set("Id", relation_id)
        node.set("Type", NUM_REL)
        node.set("Target", "numbering.xml")
        changed = True

    if not changed:
        return False

    members["word/document.xml"] = _dump(document)
    members["word/styles.xml"] = _dump(styles)
    members["word/numbering.xml"] = _dump(numbering)
    members["[Content_Types].xml"] = _dump(content_types)
    members[rels_path] = _dump(relationships)
    temporary = f"{path}.numbering.tmp"
    try:
        with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as output:
            for name, payload in members.items():
                output.writestr(name, payload)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)
    return True
