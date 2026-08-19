"""Tests for numbering_native — plain-text heading numbers -> native numbering."""

import zipfile

import xml.etree.ElementTree as ET

from skills.WPSComposer.scripts.numbering_native import (
    apply_native_numbering,
    _NUM_ID,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"

DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="' + W + '">'
    "<w:body>"
    # H1 with plain-text 第一章 prefix (styleId 26)
    '<w:p><w:pPr><w:pStyle w:val="26"/></w:pPr>'
    "<w:r><w:t>第一章 知识库应用价值与材料清单（建议稿）</w:t></w:r></w:p>"
    # H1 TOC title — must stay unnumbered (styleId 26)
    '<w:p><w:pPr><w:pStyle w:val="26"/></w:pPr>'
    "<w:r><w:t>目  录</w:t></w:r></w:p>"
    # H2 with 第一节 prefix (styleId 27)
    '<w:p><w:pPr><w:pStyle w:val="27"/></w:pPr>'
    "<w:r><w:t>第一节 核心组织机制：部门文档柜</w:t></w:r></w:p>"
    # H3 with 一、 prefix (styleId 28)
    '<w:p><w:pPr><w:pStyle w:val="28"/></w:pPr>'
    "<w:r><w:t>一、 制度法规一键查询</w:t></w:r></w:p>"
    # H4 with （一） prefix (styleId 29)
    '<w:p><w:pPr><w:pStyle w:val="29"/></w:pPr>'
    "<w:r><w:t>（一） 市场开发中心</w:t></w:r></w:p>"
    "</w:body></w:document>"
)

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:styles xmlns:w="' + W + '">'
    '<w:style w:type="paragraph" w:styleId="10"><w:name w:val="Title"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="26">'
    '<w:name w:val="Heading 1"/><w:pPr><w:keepNext/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="27">'
    '<w:name w:val="Heading 2"/><w:pPr><w:keepNext/><w:spacing w:after="100"/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="28">'
    '<w:name w:val="Heading 3"/><w:pPr><w:keepNext/></w:pPr></w:style>'
    '<w:style w:type="paragraph" w:styleId="29">'
    '<w:name w:val="Heading 4"/><w:pPr><w:keepNext/></w:pPr></w:style>'
    "</w:styles>"
)

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="' + CT + '">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)

RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="' + REL + '">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    "</Relationships>"
)


def _make_docx(path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES.encode("utf-8"))
        zf.writestr("_rels/.rels", "".encode("utf-8"))
        zf.writestr("word/document.xml", DOCUMENT_XML.encode("utf-8"))
        zf.writestr("word/styles.xml", STYLES_XML.encode("utf-8"))
        zf.writestr("word/_rels/document.xml.rels", RELS.encode("utf-8"))


def _add_existing_numbering(path) -> None:
    """Add an unrelated list definition, as python-docx/WPS commonly does."""
    with zipfile.ZipFile(path) as source:
        members = {name: source.read(name) for name in source.namelist()}
    members["word/numbering.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:numbering xmlns:w="' + W + '">'
        '<w:abstractNum w:abstractNumId="0">'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/>'
        '<w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
        '</w:lvl></w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '</w:numbering>'
    ).encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as destination:
        for name, payload in members.items():
            destination.writestr(name, payload)


def _read(path, name):
    with zipfile.ZipFile(path) as zf:
        return zf.read(name).decode("utf-8")


def _texts(path):
    root = ET.fromstring(_read(path, "word/document.xml"))
    return [
        (p.find(f"{{{W}}}pPr/{{{W}}}pStyle").get(f"{{{W}}}val"),
         "".join(t.text or "" for t in p.iter(f"{{{W}}}t")))
        for p in root.iter(f"{{{W}}}p")
        if p.find(f"{{{W}}}pPr/{{{W}}}pStyle") is not None
    ]


def _has_numpr(element, ilvl):
    numpr = element.find(f"{{{W}}}pPr/{{{W}}}numPr")
    if numpr is None:
        return False
    il = numpr.find(f"{{{W}}}ilvl")
    ni = numpr.find(f"{{{W}}}numId")
    return (
        il is not None
        and il.get(f"{{{W}}}val") == ilvl
        and ni is not None
        and ni.get(f"{{{W}}}val") == _NUM_ID
    )


def test_strips_plain_text_prefixes(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    apply_native_numbering(docx)
    texts = _texts(docx)
    assert texts[0] == ("26", "知识库应用价值与材料清单（建议稿）")
    assert texts[1] == ("26", "目  录")  # untouched
    assert texts[2] == ("27", "核心组织机制：部门文档柜")
    assert texts[3] == ("28", "制度法规一键查询")
    assert texts[4] == ("29", "市场开发中心")


def test_binds_heading_styles_to_levels(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    apply_native_numbering(docx)
    styles = ET.fromstring(_read(docx, "word/styles.xml"))
    by_id = {
        s.get(f"{{{W}}}styleId"): s
        for s in styles.iter(f"{{{W}}}style")
    }
    assert _has_numpr(by_id["27"], "1")
    assert _has_numpr(by_id["28"], "2")
    assert _has_numpr(by_id["29"], "3")
    # Heading 1 style stays unbound (目录-style titles must not number)
    assert by_id["26"].find(f"{{{W}}}pPr/{{{W}}}numPr") is None


def test_binds_heading1_paragraph_that_had_chapter_prefix(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    apply_native_numbering(docx)
    root = ET.fromstring(_read(docx, "word/document.xml"))
    h1 = [
        p for p in root.iter(f"{{{W}}}p")
        if p.find(f"{{{W}}}pPr/{{{W}}}pStyle") is not None
        and p.find(f"{{{W}}}pPr/{{{W}}}pStyle").get(f"{{{W}}}val") == "26"
    ]
    numbered = [p for p in h1 if _has_numpr(p, "0")]
    assert len(numbered) == 1            # only 第一章 … gets level-0 binding
    assert h1[1].find(f"{{{W}}}pPr/{{{W}}}numPr") is None  # 目  录 not


def test_injects_numbering_part_and_registrations(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    apply_native_numbering(docx)
    numbering = _read(docx, "word/numbering.xml")
    for needle in ("chineseCounting", "第%1章", "第%2节", "%3、", "（%4）"):
        assert needle in numbering
    ct = _read(docx, "[Content_Types].xml")
    assert "/word/numbering.xml" in ct
    rels = _read(docx, "word/_rels/document.xml.rels")
    assert 'Target="numbering.xml"' in rels


def test_idempotent(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    assert apply_native_numbering(docx) is True
    first = _read(docx, "word/document.xml")
    assert apply_native_numbering(docx) is False  # skipped
    assert _read(docx, "word/document.xml") == first


def test_merges_native_heading_numbering_into_existing_numbering_part(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    _add_existing_numbering(docx)

    assert apply_native_numbering(docx) is True

    numbering = ET.fromstring(_read(docx, "word/numbering.xml"))
    abstract_ids = {
        node.get(f"{{{W}}}abstractNumId")
        for node in numbering.findall(f"{{{W}}}abstractNum")
    }
    assert "0" in abstract_ids  # unrelated list preserved
    assert len(abstract_ids) == 2
    assert "知识库应用价值与材料清单" in _texts(docx)[0][1]
    assert not _texts(docx)[0][1].startswith("第一章")


def test_accepts_lowercase_builtin_heading_style_names(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    with zipfile.ZipFile(docx) as source:
        members = {name: source.read(name) for name in source.namelist()}
    members["word/styles.xml"] = members["word/styles.xml"].replace(
        b'Heading 1', b'heading 1'
    ).replace(b'Heading 2', b'heading 2').replace(
        b'Heading 3', b'heading 3'
    ).replace(b'Heading 4', b'heading 4')
    with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as destination:
        for name, payload in members.items():
            destination.writestr(name, payload)

    assert apply_native_numbering(docx) is True
    texts = [text for _style, text in _texts(docx)]
    assert texts[0] == "知识库应用价值与材料清单（建议稿）"
    root = ET.fromstring(_read(docx, "word/document.xml"))
    assert sum(
        p.find(f"{{{W}}}pPr/{{{W}}}numPr") is not None
        for p in root.iter(f"{{{W}}}p")
    ) == 4


def test_preserves_hybrid_chapter_decimal_heading_appearance_with_native_links(tmp_path):
    docx = tmp_path / "doc.docx"
    _make_docx(docx)
    with zipfile.ZipFile(docx) as source:
        members = {name: source.read(name) for name in source.namelist()}
    members["word/document.xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="' + W + '"><w:body>'
        '<w:p><w:pPr><w:pStyle w:val="27"/></w:pPr>'
        '<w:r><w:t>编制说明与范围边界</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="26"/></w:pPr>'
        '<w:r><w:t>第一章 编制依据与项目理解</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="27"/></w:pPr>'
        '<w:r><w:t>1.1 编制范围与响应边界</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="28"/></w:pPr>'
        '<w:r><w:t>1.1.1 部署目标与空间组织</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="29"/></w:pPr>'
        '<w:r><w:t>关键工法01：直立采坑壁削坡与回填筑坡</w:t></w:r></w:p>'
        '</w:body></w:document>'
    ).encode("utf-8")
    with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as destination:
        for name, payload in members.items():
            destination.writestr(name, payload)

    assert apply_native_numbering(docx) is True

    texts = [text for _style, text in _texts(docx)]
    assert texts == [
        "编制说明与范围边界",
        "编制依据与项目理解",
        "编制范围与响应边界",
        "部署目标与空间组织",
        "直立采坑壁削坡与回填筑坡",
    ]
    numbering = _read(docx, "word/numbering.xml")
    assert 'w:lvlText w:val="第%1章"' in numbering
    assert 'w:lvlText w:val="%1.%2"' in numbering
    assert 'w:lvlText w:val="%1.%2.%3"' in numbering
    assert 'w:lvlText w:val="关键工法%4："' in numbering
    assert "w:isLgl" in numbering

    root = ET.fromstring(_read(docx, "word/document.xml"))
    paragraphs = list(root.iter(f"{{{W}}}p"))
    assert paragraphs[0].find(f"{{{W}}}pPr/{{{W}}}numPr") is not None
    suppressed = paragraphs[0].find(
        f"{{{W}}}pPr/{{{W}}}numPr/{{{W}}}numId"
    )
    assert suppressed.get(f"{{{W}}}val") == "0"
    for paragraph, level in zip(paragraphs[1:], ("0", "1", "2", "3")):
        ilvl = paragraph.find(
            f"{{{W}}}pPr/{{{W}}}numPr/{{{W}}}ilvl"
        )
        assert ilvl is not None
        assert ilvl.get(f"{{{W}}}val") == level
