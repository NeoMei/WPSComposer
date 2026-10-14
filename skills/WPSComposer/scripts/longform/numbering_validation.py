"""Read-only validation of the native heading-numbering contract.

This is a structural gate for a long-form DOCX.  It verifies the part of the
contract that can be proven from OOXML before publication.  Application-level
insert/delete/save/reopen behavior remains a separate native acceptance gate.
"""

from __future__ import annotations

from pathlib import Path
import re
from typing import Any, Iterable, Mapping
import xml.etree.ElementTree as ET
import zipfile

from ..artifact_transport import ArtifactValidationError


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
NS = {"w": W_NS}

_SCHEMES = {"decimal", "chinese-formal", "chinese-outline", "hybrid-bid"}
_PREFIX_RE = re.compile(r"^(?:\d+(?:\.\d+)*[.、．]?|第[零〇一二三四五六七八九十百千万两]+章|[零〇一二三四五六七八九十百千万两]+、|（[零〇一二三四五六七八九十百千万两]+）|关键工法\d{1,3}[：:])")


def _failure(code: str, detail: str) -> None:
    error = ArtifactValidationError(f"{code}: {detail}")
    error.code = code  # type: ignore[attr-defined]
    raise error


def _value(node: ET.Element | None, path: str, default: str | None = None) -> str | None:
    if node is None:
        return default
    child = node.find(path, NS)
    return child.get(W + "val", default) if child is not None else default


def _text(node: ET.Element) -> str:
    return "".join(item.text or "" for item in node.findall(".//w:t", NS))


def _style_chain(style_id: str | None, styles: Mapping[str, ET.Element]) -> list[ET.Element]:
    chain: list[ET.Element] = []
    seen: set[str] = set()
    while style_id:
        if style_id in seen:
            _failure("WPSC_HEADING_STYLE_INVALID", "heading style inheritance is invalid")
        # A minimal package (and some vendor exports) may omit the style
        # declaration while carrying the paragraph-level numPr.  Direct XML
        # still provides enough evidence; only a cyclic chain is invalid.
        if style_id not in styles:
            break
        seen.add(style_id)
        style = styles[style_id]
        chain.append(style)
        style_id = _value(style, "w:basedOn")
    return list(reversed(chain))


def _effective(paragraph: ET.Element, styles: Mapping[str, ET.Element], path: str) -> str | None:
    style_id = _value(paragraph, "w:pPr/w:pStyle")
    value: str | None = None
    for node in [*_style_chain(style_id, styles), paragraph]:
        item = node.find(path, NS)
        if item is not None and W + "val" in item.attrib:
            value = item.get(W + "val")
    return value


def _paragraph_bookmarks(document: ET.Element) -> dict[str, ET.Element]:
    result: dict[str, ET.Element] = {}
    for paragraph in document.findall(".//w:body/w:p", NS):
        for marker in paragraph.findall("w:bookmarkStart", NS):
            name = marker.get(W + "name")
            if name:
                result[name] = paragraph
    return result


def _plain_expectation(operation: Any) -> dict[str, Any]:
    if isinstance(operation, Mapping):
        op = operation.get("op")
        args = operation.get("args", {})
    else:
        op = getattr(operation, "op", None)
        args = getattr(operation, "args", {})
    if op != "writer.add_heading":
        return {}
    return {
        "bookmarkName": args.get("bookmarkName"),
        "text": args.get("text", ""),
        "level": int(args.get("level", 0)),
        "numbering": bool(args.get("numbering", False)),
        "numberingScheme": args.get("numberingScheme"),
        "sequenceTransparent": bool(args.get("sequenceTransparent", False)),
    }


def heading_expectations(plan: Any) -> list[dict[str, Any]]:
    """Extract a privacy-safe heading contract from a generation plan."""
    operations = getattr(plan, "operations", ())
    result = []
    for operation in operations:
        expectation = _plain_expectation(operation)
        if expectation:
            result.append(expectation)
    return result


def _expected_level_text(scheme: str, level: int) -> tuple[set[str], str]:
    if scheme == "decimal":
        return {"decimal", "decimalZero"}, ".".join("%" + str(index) for index in range(1, level + 1))
    if scheme == "chinese-outline":
        patterns = {1: "%1、", 2: "%2.", 3: "%2.%3", 4: "%2.%3.%4"}
        return ({"decimal", "chineseCounting", "simpChinNum"}, patterns.get(level, ""))
    if scheme == "chinese-formal":
        return ({"chineseCounting", "simpChinNum", "decimal"},
                {1: "第%1章", 2: "第%2节", 3: "%3、", 4: "（%4）"}.get(level, ""))
    if scheme == "hybrid-bid":
        return ({"chineseCounting", "simpChinNum", "decimal", "decimalZero", "legal"},
                {1: "第%1章", 2: "%1.%2", 3: "%1.%2.%3", 4: "关键工法%4："}.get(level, ""))
    _failure("WPSC_HEADING_NUMBERING_SCHEME_INVALID", "unsupported heading numbering scheme")
    raise AssertionError


def _pattern_matches(actual: str, expected: str, scheme: str) -> bool:
    if actual == expected:
        return True
    # Word/WPS commonly stores a terminal punctuation mark for a decimal
    # outline level even when the contract only names its placeholders.
    if scheme == "decimal":
        return actual.rstrip(".、．") == expected.rstrip(".、．")
    return False


def _load(path: Path) -> tuple[ET.Element, dict[str, ET.Element], ET.Element]:
    try:
        with zipfile.ZipFile(Path(path)) as package:
            document = ET.fromstring(package.read("word/document.xml"))
            styles = ET.fromstring(package.read("word/styles.xml"))
            numbering = ET.fromstring(package.read("word/numbering.xml"))
    except (OSError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
        _failure("WPSC_HEADING_NUMBERING_PACKAGE_INVALID", "DOCX numbering parts are invalid")
        raise AssertionError from exc
    return document, {item.get(W + "styleId"): item for item in styles.findall("w:style", NS) if item.get(W + "styleId")}, numbering


def validate_heading_numbering(
    path: Path,
    expectations: Iterable[Mapping[str, Any]],
) -> bool:
    """Validate heading numbering for one generated DOCX.

    The error code is attached to :class:`ArtifactValidationError` as ``code``
    so artifact transport can keep its stable failure envelope.
    """
    document, styles, numbering = _load(Path(path))
    bookmarks = _paragraph_bookmarks(document)
    instances = {item.get(W + "numId"): item for item in numbering.findall("w:num", NS)}
    abstracts = {item.get(W + "abstractNumId"): item for item in numbering.findall("w:abstractNum", NS)}
    shared_num_id: str | None = None

    for expectation in expectations:
        bookmark = expectation.get("bookmarkName")
        if not bookmark:
            _failure("WPSC_HEADING_BOOKMARK_MISSING", "numbered heading has no stable bookmark")
        paragraph = bookmarks.get(str(bookmark))
        if paragraph is None:
            _failure("WPSC_HEADING_BOOKMARK_MISSING", "heading bookmark is absent")
        expected_text = str(expectation.get("text", ""))
        if _text(paragraph).strip() != expected_text.strip():
            _failure("WPSC_HEADING_TEXT_MISMATCH", "heading bookmark resolves to unexpected text")
        level = int(expectation.get("level", 0))
        numbered = bool(expectation.get("numbering"))
        num_id = _effective(paragraph, styles, "w:pPr/w:numPr/w:numId")
        ilvl = _effective(paragraph, styles, "w:pPr/w:numPr/w:ilvl")
        if not numbered:
            if expectation.get("sequenceTransparent") and num_id not in (None, "0"):
                _failure("WPSC_UNNUMBERED_HEADING_BOUND", "transparent heading remains bound to a list")
            continue
        if num_id in (None, "0"):
            _failure("WPSC_HEADING_NUMBERING_NUMID_ZERO", "numbered heading has no effective list instance")
        if shared_num_id is None:
            shared_num_id = num_id
        elif num_id != shared_num_id:
            _failure("WPSC_HEADING_NUMBERING_SHARED_SEQUENCE", "numbered headings use multiple list instances")
        instance = instances.get(num_id)
        if instance is None:
            _failure("WPSC_HEADING_NUMBERING_INSTANCE_MISSING", "effective list instance is absent")
        abstract_id = _value(instance, "w:abstractNumId")
        abstract = abstracts.get(abstract_id)
        if abstract is None:
            _failure("WPSC_HEADING_NUMBERING_ABSTRACT_MISSING", "list abstract definition is absent")
        if ilvl is None:
            style_id = _value(paragraph, "w:pPr/w:pStyle")
            style_id = style_id or "Heading" + str(level)
            target = next((item for item in abstract.findall("w:lvl", NS) if _value(item, "w:pStyle") == style_id), None)
            ilvl = target.get(W + "ilvl") if target is not None else str(level - 1)
        if ilvl != str(level - 1):
            _failure("WPSC_HEADING_NUMBERING_LEVEL_MISMATCH", "heading list level does not match its outline level")
        native_level = next((item for item in abstract.findall("w:lvl", NS) if item.get(W + "ilvl") == ilvl), None)
        if native_level is None:
            _failure("WPSC_HEADING_NUMBERING_LEVEL_MISSING", "heading list level is absent")
        scheme = str(expectation.get("numberingScheme") or "")
        if scheme not in _SCHEMES:
            _failure("WPSC_HEADING_NUMBERING_SCHEME_INVALID", "heading numbering scheme is invalid")
        formats, expected_text = _expected_level_text(scheme, level)
        if _value(native_level, "w:numFmt") not in formats:
            _failure("WPSC_HEADING_NUMBERING_FORMAT_INVALID", "heading list format does not match its scheme")
        actual_text = _value(native_level, "w:lvlText", "") or ""
        if _value(native_level, "w:start") not in (None, "1"):
            _failure("WPSC_HEADING_NUMBERING_START_INVALID", "heading list starts at an unexpected value")
        if expected_text and not _pattern_matches(actual_text, expected_text, scheme):
            _failure("WPSC_HEADING_NUMBERING_PATTERN_INVALID", "heading list pattern does not match its scheme")
        if _PREFIX_RE.match(_text(paragraph).strip()):
            _failure("WPSC_HEADING_NUMBERING_LITERAL_PREFIX", "heading retains a literal numbering prefix")
    return True


__all__ = ["heading_expectations", "validate_heading_numbering"]
