"""Inspect saved native artifacts; does not launch or modify Office documents."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

import fitz
from PIL import Image, ImageDraw

from mac_word import content

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"


def validate(directory):
    directory = Path(directory)
    with zipfile.ZipFile(directory / "probe.docx") as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
        styles = ET.fromstring(archive.read("word/styles.xml"))
        numbering = ET.fromstring(archive.read("word/numbering.xml"))
        core = ET.fromstring(archive.read("docProps/core.xml"))
    def text(node):
        return "".join(t.text or "" for t in node.findall(".//w:t", NS))
    def value(node, path, default=None):
        found = node.find(path, NS)
        return found.get(W + "val", default) if found is not None else default
    table = document.find(".//w:tbl", NS)
    rows = [[text(cell) for cell in row.findall("w:tc", NS)] for row in table.findall("w:tr", NS)]
    expected = [line.split("\t") for line in content().splitlines()[10:92]]
    checks = {"table_82_rows_3_columns_exact_content": rows == expected and len(rows) == 82,
              "repeat_header_xml": table.find("w:tr/w:trPr/w:tblHeader", NS) is not None}
    paragraphs = document.findall("w:body/w:p", NS)
    body = next(p for p in paragraphs if text(p).startswith("这是中文正文"))
    indent = body.find("w:pPr/w:ind", NS)
    checks["body_font_size_indent_xml"] = (
        body.find("w:r/w:rPr/w:rFonts", NS).get(W + "eastAsia") == "仿宋"
        and value(body, "w:r/w:rPr/w:sz") == "24"
        and indent.get(W + "firstLineChars") == "200" and indent.get(W + "firstLine") == "480")
    heading_details = []
    for label, level, size in [("原生办公文档验证", 0, "32"), ("层级标题验证", 1, "30"), ("三级标题验证", 2, "30"), ("长表格验证", 0, "32")]:
        paragraph = next(p for p in paragraphs if text(p) == label)
        style_id = value(paragraph, "w:pPr/w:pStyle")
        style = next(s for s in styles if s.get(W + "styleId") == style_id)
        num_id = value(style, "w:pPr/w:numPr/w:numId")
        number = next(n for n in numbering.findall("w:num", NS) if n.get(W + "numId") == num_id)
        abstract_id = value(number, "w:abstractNumId")
        abstract = next(a for a in numbering.findall("w:abstractNum", NS) if a.get(W + "abstractNumId") == abstract_id)
        linked_level = next(l for l in abstract.findall("w:lvl", NS) if l.get(W + "ilvl") == str(level))
        ok = (value(style, "w:name") == "heading " + str(level + 1)
              and value(style, "w:pPr/w:numPr/w:ilvl", "0") == str(level)
              and value(linked_level, "w:pStyle") == style_id
              and value(linked_level, "w:lvlText") == ".".join("%" + str(i) for i in range(1, level + 2))
              and value(paragraph, "w:r/w:rPr/w:sz") == size)
        heading_details.append({"heading": label, "style_id": style_id, "num_id": num_id, "level": level, "passed": ok})
    checks["heading_styles_sizes_native_numbering_links"] = all(x["passed"] for x in heading_details)
    fields = [t.text or "" for t in document.findall(".//w:instrText", NS)]
    checks["toc_field_with_four_pagerefs"] = any('TOC \\o "1-3"' in f for f in fields) and sum("PAGEREF" in f for f in fields) == 4
    pdf = fitz.open(directory / "probe.pdf")
    pages = [p.get_text() for p in pdf]
    compact = re.sub(r"\s+", "", "".join(pages))
    checks["pdf_80_records_exactly_once"] = all(compact.count(f"原生排版验证第{i}条") == 1 for i in range(1, 81))
    checks["pdf_headers_every_page"] = all(all(h in p for h in ["编号", "检查内容", "结论"]) for p in pages)
    checks["pdf_end_marker_last_page"] = "MSOFFICE-SPIKE-END" in pages[-1]
    result = json.loads((directory / "result.json").read_text("utf-8"))
    checks["pdf_pages_match_native"] = len(pdf) == result["page_count"]
    spans = [s for p in pdf for b in p.get_text("dict")["blocks"] if "lines" in b for line in b["lines"] for s in line["spans"]]
    body_spans = [s for s in spans if "这是中文正文" in s["text"]]
    checks["pdf_body_fangsong_12pt"] = bool(body_spans) and all("FangSong" in s["font"] and abs(s["size"] - 12) < .05 for s in body_spans)
    checks["pdf_all_text_inside_pages"] = all(s["bbox"][0] >= -1 and s["bbox"][1] >= -1 and s["bbox"][2] <= p.rect.width + 1 and s["bbox"][3] <= p.rect.height + 1 for p in pdf for b in p.get_text("dict")["blocks"] if "lines" in b for l in b["lines"] for s in l["spans"])
    images = []
    for index, page in enumerate(pdf):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.25, 1.25), alpha=False)
        pix.save(directory / f"page-{index+1:02}.png")
        image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        image.thumbnail((420, 600))
        images.append(image)
    sheet = Image.new("RGB", (1260, ((len(images)+2)//3)*625), "#dddddd")
    draw = ImageDraw.Draw(sheet)
    for i, image in enumerate(images):
        x, y = (i % 3) * 420, (i // 3) * 625
        sheet.paste(image, (x, y + 25))
        draw.text((x + 10, y + 5), str(i + 1), fill="black")
    sheet.save(directory / "contact-sheet.png")
    authors = [node.text for node in core if node.tag.rsplit("}", 1)[-1] in ("creator", "lastModifiedBy")]
    report = {"checks": checks, "all_checks_passed": all(checks.values()), "page_count": len(pdf), "headings": heading_details,
              "docx_authors": authors, "pdf_metadata": pdf.metadata,
              "body_pdf_fonts": [{k:s[k] for k in ("font", "size")} for s in body_spans],
              "visual_review": "pending", "sha256": {name: hashlib.sha256((directory/name).read_bytes()).hexdigest() for name in ("probe.docx", "probe.pdf", "result.json")}}
    (directory / "artifact-validation.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"directory": str(directory), "checks": checks, "pages": len(pdf), "docx_authors": authors, "pdf_author": pdf.metadata.get("author")}, ensure_ascii=False))
    return report["all_checks_passed"]


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory")
    raise SystemExit(0 if validate(parser.parse_args().directory) else 1)
