"""Generate deterministic OOXML fixtures for Word story-range diagnostics.

This test-asset generator uses only the Python standard library. It never
opens Word/WPS and never imports the product backend.
"""
from __future__ import annotations

import argparse
import hashlib
from io import BytesIO
import json
from pathlib import Path
import zipfile
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)
FIXED_CORE_TIME = "2026-09-10T00:00:00Z"
STORY_TYPES = ("default", "first", "even")
STORY_KINDS = ("header", "footer")

for prefix, uri in (("w", W), ("r", R), ("cp", CP), ("dc", DC),
                    ("dcterms", DCTERMS), ("xsi", XSI), ("ep", EP)):
    ET.register_namespace(prefix, uri)


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def child(parent: ET.Element, local: str, **attrs: object) -> ET.Element:
    return ET.SubElement(parent, qn(local), {qn(key): str(value) for key, value in attrs.items()})


def xml_bytes(root: ET.Element) -> bytes:
    return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ET.tostring(
        root, encoding="utf-8", short_empty_elements=True
    )


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _root_rels() -> bytes:
    root = ET.Element(f"{{{PKG_REL}}}Relationships")
    for rid, target, rel_type in (
        ("rId1", "word/document.xml", "officeDocument"),
        ("rId2", "docProps/core.xml", "metadata/core-properties"),
        ("rId3", "docProps/app.xml", "extended-properties"),
    ):
        base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
        if rel_type == "metadata/core-properties":
            base = "http://schemas.openxmlformats.org/package/2006/relationships/"
        ET.SubElement(root, f"{{{PKG_REL}}}Relationship", Id=rid, Target=target, Type=base + rel_type)
    return xml_bytes(root)


def _core(title: str) -> bytes:
    root = ET.Element(f"{{{CP}}}coreProperties")
    ET.SubElement(root, f"{{{DC}}}title").text = title
    ET.SubElement(root, f"{{{DC}}}creator").text = "WPSComposer fixture generator"
    ET.SubElement(root, f"{{{CP}}}lastModifiedBy").text = "WPSComposer fixture generator"
    ET.SubElement(root, f"{{{CP}}}revision").text = "1"
    for name in ("created", "modified"):
        node = ET.SubElement(root, f"{{{DCTERMS}}}{name}", {f"{{{XSI}}}type": "dcterms:W3CDTF"})
        node.text = FIXED_CORE_TIME
    return xml_bytes(root)


def _app() -> bytes:
    root = ET.Element(f"{{{EP}}}Properties")
    ET.SubElement(root, f"{{{EP}}}Application").text = "WPSComposer deterministic OOXML fixture"
    ET.SubElement(root, f"{{{EP}}}AppVersion").text = "1.0"
    return xml_bytes(root)


def _styles() -> bytes:
    root = ET.Element(qn("styles"))
    defaults = child(root, "docDefaults")
    child(child(defaults, "rPrDefault"), "rPr")
    child(child(defaults, "pPrDefault"), "pPr")
    style = ET.SubElement(root, qn("style"), {qn("type"): "paragraph", qn("default"): "1", qn("styleId"): "Normal"})
    child(style, "name", val="Normal")
    return xml_bytes(root)


def _settings(even_and_odd: bool) -> bytes:
    root = ET.Element(qn("settings"))
    if even_and_odd:
        child(root, "evenAndOddHeaders")
    return xml_bytes(root)


def _add_text_run(paragraph: ET.Element, text: str, *, hidden: bool = False) -> None:
    run = child(paragraph, "r")
    if hidden:
        child(child(run, "rPr"), "vanish")
    node = child(run, "t")
    node.text = text


def _story_xml(kind: str, story_type: str, marker: str) -> bytes:
    root = ET.Element(qn("hdr" if kind == "header" else "ftr"))
    paragraph = child(root, "p")
    ppr = child(paragraph, "pPr")
    child(ppr, "pStyle", val="Normal")
    ordinary = f"{marker} {kind.upper()} {story_type.upper()} "
    hidden = f"HIDDEN-{marker}-{kind}-{story_type} "
    _add_text_run(paragraph, ordinary)
    _add_text_run(paragraph, hidden, hidden=True)
    begin = child(paragraph, "r")
    child(begin, "fldChar", fldCharType="begin", fldLock="1")
    instruction = child(paragraph, "r")
    instr = child(instruction, "instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = " PAGE "
    separate = child(paragraph, "r")
    child(separate, "fldChar", fldCharType="separate")
    _add_text_run(paragraph, "7")
    end = child(paragraph, "r")
    child(end, "fldChar", fldCharType="end")
    return xml_bytes(root)


def _fixture_spec(kind: str) -> tuple[list[dict[str, object]], list[dict[str, object]], bool]:
    """Return section specs, story parts, and even/odd setting."""
    if kind == "absent":
        return [{"title_page": False, "refs": []}], [], False
    if kind == "primary":
        refs = [{"kind": k, "type": "default", "marker": "PRIMARY"} for k in STORY_KINDS]
        return [{"title_page": False, "refs": refs}], refs, False
    if kind == "inactive-first-even":
        refs = [
            {"kind": k, "type": t, "marker": "INACTIVE"}
            for k in STORY_KINDS for t in ("first", "even")
        ]
        return [{"title_page": False, "refs": refs}], refs, False
    if kind == "linked-multisection":
        first = [
            {"kind": k, "type": t, "marker": "SECTION1"}
            for k in STORY_KINDS for t in STORY_TYPES
        ]
        third = [
            {"kind": k, "type": t, "marker": "SECTION3"}
            for k in STORY_KINDS for t in STORY_TYPES
        ]
        return [
            {"title_page": True, "refs": first},
            {"title_page": True, "refs": []},
            {"title_page": True, "refs": third},
        ], first + third, True
    raise ValueError(f"unknown fixture: {kind}")


def _assign_parts(stories: list[dict[str, object]]) -> list[dict[str, object]]:
    counters = {"header": 0, "footer": 0}
    assigned: list[dict[str, object]] = []
    for story in stories:
        row = dict(story)
        story_kind = str(row["kind"])
        counters[story_kind] += 1
        row["part"] = f"{story_kind}{counters[story_kind]}.xml"
        row["rid"] = f"rIdStory{len(assigned) + 1}"
        assigned.append(row)
    return assigned


def _sect_pr(section: dict[str, object], assigned: list[dict[str, object]]) -> ET.Element:
    sect = ET.Element(qn("sectPr"))
    for raw in section["refs"]:  # type: ignore[index]
        story = next(row for row in assigned if all(row[key] == raw[key] for key in ("kind", "type", "marker")))
        ref = child(sect, f"{story['kind']}Reference", type=story["type"])
        ref.set(f"{{{R}}}id", str(story["rid"]))
    if section["title_page"]:
        child(sect, "titlePg")
    child(sect, "pgSz", w=12240, h=15840)
    child(sect, "pgMar", top=1440, right=1440, bottom=1440, left=1440, header=720, footer=720, gutter=0)
    return sect


def _document(sections: list[dict[str, object]], assigned: list[dict[str, object]]) -> bytes:
    root = ET.Element(qn("document"))
    body = child(root, "body")
    for index, section in enumerate(sections, 1):
        paragraph = child(body, "p")
        ppr = child(paragraph, "pPr")
        child(ppr, "pStyle", val="Normal")
        _add_text_run(paragraph, f"BODY SECTION {index}")
        if index < len(sections):
            ppr.append(_sect_pr(section, assigned))
        else:
            body.append(_sect_pr(section, assigned))
    return xml_bytes(root)


def _document_rels(assigned: list[dict[str, object]]) -> bytes:
    root = ET.Element(f"{{{PKG_REL}}}Relationships")
    for rid, target, rel_type in (
        ("rIdStyles", "styles.xml", "styles"),
        ("rIdSettings", "settings.xml", "settings"),
    ):
        ET.SubElement(root, f"{{{PKG_REL}}}Relationship", Id=rid, Target=target,
                      Type=f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rel_type}")
    for story in assigned:
        ET.SubElement(root, f"{{{PKG_REL}}}Relationship", Id=str(story["rid"]), Target=str(story["part"]),
                      Type=f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{story['kind']}")
    return xml_bytes(root)


def _content_types(assigned: list[dict[str, object]]) -> bytes:
    root = ET.Element(f"{{{CT}}}Types")
    ET.SubElement(root, f"{{{CT}}}Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    ET.SubElement(root, f"{{{CT}}}Default", Extension="xml", ContentType="application/xml")
    overrides = [
        ("/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"),
        ("/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"),
        ("/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"),
        ("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
        ("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml"),
    ]
    for story in assigned:
        overrides.append((
            f"/word/{story['part']}",
            f"application/vnd.openxmlformats-officedocument.wordprocessingml.{story['kind']}+xml",
        ))
    for part, content_type in overrides:
        ET.SubElement(root, f"{{{CT}}}Override", PartName=part, ContentType=content_type)
    return xml_bytes(root)


def _package(kind: str) -> bytes:
    sections, raw_stories, even_and_odd = _fixture_spec(kind)
    assigned = _assign_parts(raw_stories)
    parts = {
        "[Content_Types].xml": _content_types(assigned),
        "_rels/.rels": _root_rels(),
        "docProps/app.xml": _app(),
        "docProps/core.xml": _core(f"WPSComposer {kind} story matrix"),
        "word/_rels/document.xml.rels": _document_rels(assigned),
        "word/document.xml": _document(sections, assigned),
        "word/settings.xml": _settings(even_and_odd),
        "word/styles.xml": _styles(),
    }
    for story in assigned:
        parts[f"word/{story['part']}"] = _story_xml(str(story["kind"]), str(story["type"]), str(story["marker"]))
    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as package:
        for name in sorted(parts):
            info = zipfile.ZipInfo(name, FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            package.writestr(info, parts[name], compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    return output.getvalue()


def _field_rows(root: ET.Element) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    active: dict[str, object] | None = None
    state = "outside"
    for run in root.iter(qn("r")):
        field_char = run.find(qn("fldChar"))
        if field_char is not None:
            field_type = field_char.attrib.get(qn("fldCharType"))
            if field_type == "begin":
                active = {"type": None, "instruction": "", "result": "", "locked": field_char.attrib.get(qn("fldLock")) == "1"}
                state = "instruction"
            elif field_type == "separate":
                state = "result"
            elif field_type == "end" and active is not None:
                active["type"] = str(active["instruction"]).strip().split()[0].upper()
                rows.append(active)
                active = None
                state = "outside"
        if active is not None:
            if state == "instruction":
                active["instruction"] = str(active["instruction"]) + "".join(n.text or "" for n in run.findall(qn("instrText")))
            elif state == "result":
                active["result"] = str(active["result"]) + "".join(n.text or "" for n in run.findall(qn("t")))
    return rows


def _inspect(path: Path, kind: str) -> dict[str, object]:
    with zipfile.ZipFile(path) as package:
        infos = package.infolist()
        names = [info.filename for info in infos]
        if names != sorted(names) or any(info.date_time != FIXED_ZIP_TIME or info.extra or info.comment for info in infos):
            raise AssertionError("noncanonical ZIP")
        parts = {name: package.read(name) for name in names}
    rel_root = ET.fromstring(parts["word/_rels/document.xml.rels"])
    rels = {
        row.attrib["Id"]: {"target": row.attrib["Target"], "type": row.attrib["Type"]}
        for row in rel_root.findall(f"{{{PKG_REL}}}Relationship")
    }
    settings = ET.fromstring(parts["word/settings.xml"])
    even_and_odd = settings.find(qn("evenAndOddHeaders")) is not None
    document = ET.fromstring(parts["word/document.xml"])
    body = document.find(qn("body"))
    if body is None:
        raise AssertionError("missing body")
    sects = [node for node in body.iter(qn("sectPr"))]
    sections: list[dict[str, object]] = []
    references: list[dict[str, object]] = []
    inherited: list[dict[str, object]] = []
    last_explicit: dict[tuple[str, str], int] = {}
    for section_index, sect in enumerate(sects, 1):
        title_page = sect.find(qn("titlePg")) is not None
        sections.append({"section": section_index, "title_page": title_page})
        explicit_slots: set[tuple[str, str]] = set()
        for story_kind in STORY_KINDS:
            for ref in sect.findall(qn(f"{story_kind}Reference")):
                story_type = ref.attrib[qn("type")]
                rid = ref.attrib[f"{{{R}}}id"]
                slot = (story_kind, story_type)
                explicit_slots.add(slot)
                last_explicit[slot] = section_index
                active = story_type == "default" or (story_type == "first" and title_page) or (story_type == "even" and even_and_odd)
                references.append({
                    "section": section_index, "kind": story_kind, "type": story_type,
                    "rid": rid, "part": rels[rid]["target"], "active": active,
                    "linked_to_previous": False,
                })
        if section_index > 1:
            for slot, source_section in sorted(last_explicit.items()):
                if slot not in explicit_slots:
                    story_kind, story_type = slot
                    active = story_type == "default" or (story_type == "first" and title_page) or (story_type == "even" and even_and_odd)
                    inherited.append({
                        "section": section_index, "kind": story_kind, "type": story_type,
                        "source_section": source_section, "active": active, "linked_to_previous": True,
                    })
    story_parts = []
    story_targets = sorted({row["part"] for row in references})
    for target in story_targets:
        root = ET.fromstring(parts[f"word/{target}"])
        kind_name = "header" if root.tag == qn("hdr") else "footer"
        ref = next(row for row in references if row["part"] == target)
        ordinary_chunks: list[str] = []
        hidden_chunks: list[str] = []
        inside_field = False
        for run in root.iter(qn("r")):
            field_char = run.find(qn("fldChar"))
            if field_char is not None:
                field_type = field_char.attrib.get(qn("fldCharType"))
                if field_type == "begin":
                    inside_field = True
                elif field_type == "end":
                    inside_field = False
                continue
            if inside_field:
                continue
            text = "".join(node.text or "" for node in run.findall(qn("t")))
            if not text:
                continue
            if run.find(f"{qn('rPr')}/{qn('vanish')}") is not None:
                hidden_chunks.append(text)
            elif run.find(qn("fldChar")) is None:
                ordinary_chunks.append(text)
        story_parts.append({
            "part": target, "kind": kind_name, "type": ref["type"],
            "ordinary_text": "".join(ordinary_chunks), "hidden_text": "".join(hidden_chunks),
            "fields": _field_rows(root), "xml_sha256": hashlib.sha256(parts[f"word/{target}"]).hexdigest(),
        })
    result = {
        "file": path.name, "sha256": sha256(path), "package_entries": names,
        "section_count": len(sects), "even_and_odd_headers": even_and_odd,
        "sections": sections, "section_story_references": references,
        "inherited_story_slots": inherited, "story_parts": story_parts,
    }
    _assert_contract(kind, result)
    return result


def _assert_contract(kind: str, row: dict[str, object]) -> None:
    if kind == "absent":
        if row["section_count"] != 1 or row["story_parts"] or row["section_story_references"]:
            raise AssertionError("absent fixture contains a page story")
        return
    stories = row["story_parts"]
    if not stories or any(not item["ordinary_text"] or not item["hidden_text"] or len(item["fields"]) != 1 for item in stories):
        raise AssertionError("story payload coverage mismatch")
    if kind == "primary" and len(stories) != 2:
        raise AssertionError("primary fixture topology mismatch")
    if kind == "inactive-first-even":
        if row["even_and_odd_headers"] or row["sections"][0]["title_page"] or len(stories) != 4:
            raise AssertionError("inactive fixture activation mismatch")
        if any(item["active"] for item in row["section_story_references"]):
            raise AssertionError("inactive story unexpectedly active")
    if kind == "linked-multisection":
        if row["section_count"] != 3 or not row["even_and_odd_headers"] or len(stories) != 12:
            raise AssertionError("multisection topology mismatch")
        if len(row["inherited_story_slots"]) != 6:
            raise AssertionError("section 2 inheritance coverage mismatch")


def generate(out_dir: Path) -> dict[str, object]:
    out_dir.mkdir(parents=True, exist_ok=True)
    fixtures = {}
    for kind in ("absent", "primary", "inactive-first-even", "linked-multisection"):
        path = out_dir / f"{kind}.docx"
        path.write_bytes(_package(kind))
        fixtures[kind] = _inspect(path, kind)
    preflight = {
        "schema": "wpscomposer-word-story-matrix-v1",
        "generated_utc": FIXED_CORE_TIME,
        "purpose": "source-bound native story-range diagnostic; no Office acceptance claim",
        "fixtures": fixtures,
    }
    (out_dir / "preflight.json").write_text(json.dumps(preflight, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    hashes = {
        "schema": "wpscomposer-word-story-matrix-hashes-v1",
        "generator": {"file": Path(__file__).name, "sha256": sha256(Path(__file__).resolve())},
        "outputs": {name: sha256(out_dir / name) for name in (
            "absent.docx", "primary.docx", "inactive-first-even.docx",
            "linked-multisection.docx", "preflight.json",
        )},
    }
    (out_dir / "HASHES.json").write_text(json.dumps(hashes, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return preflight


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--out-dir", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    result = generate(args.out_dir)
    print(json.dumps({name: {"sha256": row["sha256"], "sections": row["section_count"], "stories": len(row["story_parts"])}
                      for name, row in result["fixtures"].items()}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
