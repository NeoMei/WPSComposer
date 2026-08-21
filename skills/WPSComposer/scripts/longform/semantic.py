
"""Long-form semantic normalization and reference resolution.

Takes a parsed StructuredDocument and produces a deterministic semantic
snapshot: normalized document, resolved configuration, reference map,
bookmark mapping, and any planned degradations.
"""

from __future__ import annotations

import copy
import dataclasses
import re
from dataclasses import dataclass
from typing import Any, Optional

from ..document_model import (
    AbstractBlock,
    CodeBlock,
    DegradationBlock,
    DocumentIssue,
    ExcalidrawBlock,
    FigureBlock,
    FormulaBlock,
    BlockQuote,
    HorizontalRule,
    ImageBlock,
    KeywordsBlock,
    ListBlock,
    MathBlock,
    PageBreakBlock,
    Paragraph,
    ReferenceListBlock,
    Section,
    SemanticTableBlock,
    Span,
    StructuredDocument,
    TableBlock,
    TaskList,
)
from .bookmark_ids import BookmarkMapResult, map_bookmarks
from .unicode_text import (
    contains_han,
    display_units,
    normalize_visible_text,
    shorten_display_units,
)

CONFIG_VALUE_INVALID = "CONFIG_VALUE_INVALID"
HEADER_SHORTENED = "HEADER_SHORTENED"
HEADING_PREFIX_AMBIGUOUS = "HEADING_PREFIX_AMBIGUOUS"
HEADING_LEVEL_GAP = "HEADING_LEVEL_GAP"
TITLE_MISSING = "TITLE_MISSING"
REFERENCE_UNRESOLVED = "REFERENCE_UNRESOLVED"
DUPLICATE_EXPLICIT_ID = "DUPLICATE_EXPLICIT_ID"
INVALID_EXPLICIT_ID = "INVALID_EXPLICIT_ID"

_VALID_BOOL = frozenset({"true", "false", "yes", "no", "1", "0"})
_VALID_CAPTION_NUMBERING = frozenset({"auto", "chapter", "global"})
_VALID_HEADING_NUMBERING = frozenset({
    "auto",
    "none",
    "chinese-formal",
    "decimal",
    "hybrid-bid",
})

_ID_RE = re.compile(r"^[A-Za-z][A-Za-z0-9._:-]{0,127}$")
_RESERVED_ID_PREFIX = "__wpsc_"

_REF_RE = re.compile(r"\{\{ref:([^}]+)\}\}")
_CITE_RE = re.compile(r"\{\{cite:([^}]+)\}\}")

_CHINESE_NUMERAL_CHARS = set(
    "零〇一二三四五六七八九十百千万两"
)


@dataclass
class LongformConfig:
    """Resolved long-form document configuration."""

    title: str = ""
    short_title: str = ""
    author: str = ""
    date: str = ""
    header: str = ""
    title_page: bool = False
    toc: bool = False
    figure_index: bool = False
    table_index: bool = False
    bibliography_include_uncited: bool = True
    caption_numbering: str = "global"
    heading_numbering: str = "none"
    layout_engine: str = "longform"

    def to_json(self) -> dict[str, Any]:
        return {
            "author": self.author,
            "bibliography_include_uncited": self.bibliography_include_uncited,
            "caption_numbering": self.caption_numbering,
            "date": self.date,
            "figure_index": self.figure_index,
            "header": self.header,
            "heading_numbering": self.heading_numbering,
            "layout_engine": self.layout_engine,
            "short_title": self.short_title,
            "table_index": self.table_index,
            "title": self.title,
            "title_page": self.title_page,
            "toc": self.toc,
        }


@dataclass
class SemanticResult:
    """Result of normalizing a long-form document."""

    document: StructuredDocument
    config: LongformConfig
    references: dict[str, dict[str, Any]]
    bookmarks: BookmarkMapResult
    issues: tuple[DocumentIssue, ...]

    def to_json(self) -> dict[str, Any]:
        return {
            "bookmarks": {
                "issues": list(self.bookmarks.issues),
                "mapping": _canonical_value(
                    dict(sorted(self.bookmarks.mapping.items()))
                ),
            },
            "config": _canonical_value(self.config.to_json()),
            "document": _canonical_value(self.document),
            "issues": [_canonical_value(i) for i in self.issues],
            "references": _canonical_value(
                dict(sorted(self.references.items()))
            ),
        }


def _canonical_value(value: Any) -> Any:
    """Recursively convert dataclasses/dicts/lists to deterministic JSON."""
    if dataclasses.is_dataclass(value):
        d = dataclasses.asdict(value)
        return {k: _canonical_value(v) for k, v in sorted(d.items())}
    if isinstance(value, (list, tuple)):
        return [_canonical_value(v) for v in value]
    if isinstance(value, dict):
        return {k: _canonical_value(v) for k, v in sorted(value.items())}
    return value


def _issue(
    code: str, message: str, placement: str = "document"
) -> DocumentIssue:
    return DocumentIssue(code=code, message=message, placement=placement)


def _normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return normalize_visible_text(text)


def _parse_bool(value: Any) -> Optional[bool]:
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in ("true", "yes", "1", "on"):
        return True
    if text in ("false", "no", "0", "off"):
        return False
    return None


def _is_auto(value: Any) -> bool:
    return str(value).strip().lower() == "auto"


def _valid_explicit_id(identifier: str) -> bool:
    if not identifier or identifier.startswith(_RESERVED_ID_PREFIX):
        return False
    return _ID_RE.fullmatch(identifier) is not None


def _is_valid_chinese_numeral(text: str) -> bool:
    if not text:
        return False
    return all(ch in _CHINESE_NUMERAL_CHARS for ch in text)


def _decimal_segment_valid(segment: str) -> bool:
    return segment.isdigit() and (segment == "0" or not segment.startswith("0"))


def _scheme_level_match(heading: str, scheme: str) -> Optional[int]:
    """Return the level (1-4) whose prefix pattern matches *heading*, if any."""
    if scheme == "chinese-formal":
        if re.match(r"^第[零〇一二三四五六七八九十百千万两]+章", heading):
            return 1
        if re.match(r"^第[零〇一二三四五六七八九十百千万两]+节", heading):
            return 2
        if re.match(r"^[零〇一二三四五六七八九十百千万两]+、", heading):
            return 3
        if re.match(r"^（[零〇一二三四五六七八九十百千万两]+）", heading):
            return 4
    elif scheme == "decimal":
        m = re.match(r"^(\d+)(?:\s+|$|[：:])", heading)
        if m and _decimal_segment_valid(m.group(1)):
            return 1
        m = re.match(r"^(\d+\.\d+)(?:\s+|$|[：:])", heading)
        if m:
            parts = m.group(1).split(".")
            if all(_decimal_segment_valid(p) for p in parts):
                return 2
        m = re.match(r"^(\d+(?:\.\d+){2})(?:\s+|$|[：:])", heading)
        if m:
            parts = m.group(1).split(".")
            if all(_decimal_segment_valid(p) for p in parts):
                return 3
        m = re.match(r"^(\d+(?:\.\d+){3})(?:\s+|$|[：:])", heading)
        if m:
            parts = m.group(1).split(".")
            if all(_decimal_segment_valid(p) for p in parts):
                return 4
    elif scheme == "hybrid-bid":
        if re.match(r"^第[零〇一二三四五六七八九十百千万两]+章", heading):
            return 1
        m = re.match(r"^(\d+\.\d+)(?:\s+|$|[：:])", heading)
        if m:
            parts = m.group(1).split(".")
            if all(_decimal_segment_valid(p) for p in parts):
                return 2
        m = re.match(r"^(\d+(?:\.\d+){2})(?:\s+|$|[：:])", heading)
        if m:
            parts = m.group(1).split(".")
            if all(_decimal_segment_valid(p) for p in parts):
                return 3
        if re.match(r"^关键工法\d{1,3}[：:]", heading):
            return 4
    return None


def _strip_prefix(heading: str, level: int, scheme: str) -> tuple[str, bool]:
    """Remove a matching scheme prefix from *heading* and report success."""
    if scheme == "chinese-formal":
        patterns = {
            1: re.compile(r"^第[零〇一二三四五六七八九十百千万两]+章\s*"),
            2: re.compile(r"^第[零〇一二三四五六七八九十百千万两]+节\s*"),
            3: re.compile(r"^[零〇一二三四五六七八九十百千万两]+、\s*"),
            4: re.compile(r"^（[零〇一二三四五六七八九十百千万两]+）\s*"),
        }
    elif scheme == "decimal":
        patterns = {
            1: re.compile(r"^\d+\s+"),
            2: re.compile(r"^\d+\.\d+\s+"),
            3: re.compile(r"^\d+(?:\.\d+){2}\s+"),
            4: re.compile(r"^\d+(?:\.\d+){3}\s+"),
        }
    elif scheme == "hybrid-bid":
        patterns = {
            1: re.compile(r"^第[零〇一二三四五六七八九十百千万两]+章\s*"),
            2: re.compile(r"^\d+\.\d+\s+"),
            3: re.compile(r"^\d+(?:\.\d+){2}\s+"),
            4: re.compile(r"^关键工法\d{1,3}[：:]\s*"),
        }
    else:
        return heading, False

    pattern = patterns.get(level)
    if pattern is None:
        return heading, False
    m = pattern.match(heading)
    if not m:
        return heading, False
    stripped = heading[m.end():].strip()
    return (stripped if stripped else ""), True


def _detect_heading_scheme(sections: list[Section]) -> str:
    """Select a heading numbering scheme from the document content."""
    headings = [s.heading for s in sections if s.level in {1, 2, 3, 4} and s.heading]
    if not headings:
        return "decimal"

    scores: dict[str, int] = {"chinese-formal": 0, "decimal": 0, "hybrid-bid": 0}
    hybrid_h4 = False
    for heading in headings:
        for scheme in scores:
            if _scheme_level_match(heading, scheme) is not None:
                scores[scheme] += 1
        if _scheme_level_match(heading, "hybrid-bid") == 4:
            hybrid_h4 = True

    if hybrid_h4:
        return "hybrid-bid"

    non_zero = {k for k, v in scores.items() if v > 0}
    if len(non_zero) == 1:
        (winner,) = non_zero
        if scores[winner] >= 2:
            return winner

    han_count = sum(1 for h in headings if contains_han(h))
    if han_count * 2 > len(headings):
        return "chinese-formal"
    return "decimal"


def _consume_title_h1(
    sections: list[Section], title: str
) -> list[Section]:
    """Consume the first H1 that equals the document title."""
    normalized_title = normalize_visible_text(title).strip() if title else ""
    if not normalized_title:
        return sections
    for idx, section in enumerate(sections):
        if section.level == 1 and normalize_visible_text(
            section.heading
        ).strip() == normalized_title:
            # Keep the section's body but remove it from the chapter stream.
            consumed = copy.deepcopy(section)
            consumed.level = 0
            consumed.heading = ""
            consumed.node_id = None
            consumed.preface = False
            sections[idx] = consumed
            return sections
    return sections


def _apply_heading_numbering(
    sections: list[Section],
    heading_numbering: str,
    title: str,
    issues: list[DocumentIssue],
) -> str:
    """Resolve heading prefixes, scheme, and level gaps in place."""
    if heading_numbering == "none":
        for section in sections:
            if section.level > 0:
                section.numbering = "none"
                section.numbering_scheme = None
        return "none"

    selected_scheme = (
        heading_numbering
        if heading_numbering != "auto"
        else _detect_heading_scheme(sections)
    )

    has_document_title = bool(title.strip())
    has_numbered_h1 = False
    # states: "absent" | "numbered" | "none_prefix" | "none_gap"
    state: dict[int, str] = {1: "absent", 2: "absent", 3: "absent", 4: "absent"}

    for section in sections:
        if section.level == 0 or section.level > 6:
            continue
        if section.level >= 5:
            section.numbering = "none"
            section.numbering_scheme = None
            continue

        if (
            has_document_title
            and not has_numbered_h1
            and section.level >= 2
        ):
            section.preface = True
            section.numbering = "none"
            section.numbering_scheme = None
            continue

        stripped, matched = _strip_prefix(
            section.heading, section.level, selected_scheme
        )

        if matched:
            numbering_candidate = selected_scheme
        else:
            has_any_prefix = _scheme_level_match(
                section.heading, selected_scheme
            ) is not None or any(
                _scheme_level_match(section.heading, s) is not None
                for s in ("chinese-formal", "decimal", "hybrid-bid")
            )
            if has_any_prefix:
                section.numbering = "none"
                section.numbering_scheme = None
                section.preface = False
                state[section.level] = "none_prefix"
                issues.append(
                    _issue(
                        HEADING_PREFIX_AMBIGUOUS,
                        f"Heading prefix does not match the selected scheme "
                        f"'{selected_scheme}' at level {section.level}.",
                    )
                )
                continue
            numbering_candidate = selected_scheme

        level = section.level
        if level == 1:
            allowed = True
        elif state[level - 1] == "numbered":
            allowed = True
        elif state[level - 1] in ("none_prefix", "none_gap"):
            section.numbering = "none"
            section.numbering_scheme = None
            section.preface = False
            state[level] = "none_prefix"
            continue
        else:
            allowed = False

        if allowed:
            section.heading = stripped
            section.numbering = numbering_candidate
            section.numbering_scheme = numbering_candidate
            section.preface = False
            state[level] = "numbered"
            for k in range(level + 1, 5):
                state[k] = "absent"
            if level == 1:
                has_numbered_h1 = True
        else:
            section.numbering = "none"
            section.numbering_scheme = None
            section.preface = False
            state[level] = "none_gap"
            issues.append(
                _issue(
                    HEADING_LEVEL_GAP,
                    f"Heading level {level} appears without its required parent.",
                )
            )

    return selected_scheme


def _derive_header(config: LongformConfig, issues: list[DocumentIssue]) -> None:
    """Resolve the running header text and shorten if necessary."""
    if config.header:
        source = config.header
    elif config.short_title:
        source = config.short_title
    elif config.title:
        source = config.title
    else:
        source = ""

    source = normalize_visible_text(source).strip()
    if display_units(source) > 64:
        source = shorten_display_units(source, max_units=64)
        issues.append(
            _issue(
                HEADER_SHORTENED,
                "Header text was shortened to 64 display units.",
            )
        )
    config.header = source


def _body_codepoint_count(sections: list[Section]) -> int:
    """Visible body code points excluding whitespace."""
    count = 0
    for section in sections:
        count += len(re.sub(r"\s", "", section.heading))
        for elem in section.elements:
            if isinstance(elem, Paragraph):
                for span in elem.spans:
                    if not span.code and not span.math:
                        count += len(re.sub(r"\s", "", span.text))
            elif isinstance(elem, ListBlock):
                for item in elem.items:
                    for span in item:
                        if not span.code and not span.math:
                            count += len(re.sub(r"\s", "", span.text))
            elif isinstance(elem, TableBlock):
                for cell in elem.headers:
                    count += len(re.sub(r"\s", "", cell))
                for row in elem.rows:
                    for cell in row:
                        count += len(re.sub(r"\s", "", cell))
            elif isinstance(elem, SemanticTableBlock):
                for cell in elem.headers:
                    count += len(re.sub(r"\s", "", cell))
                for row in elem.rows:
                    for cell in row:
                        count += len(re.sub(r"\s", "", cell))
            elif isinstance(elem, CodeBlock):
                count += len(re.sub(r"\s", "", elem.code))
            elif isinstance(elem, (ImageBlock, ExcalidrawBlock)):
                count += len(re.sub(r"\s", "", elem.alt))
            elif isinstance(elem, FigureBlock):
                for img in elem.images:
                    count += len(re.sub(r"\s", "", img.alt))
            elif isinstance(elem, (AbstractBlock, BlockQuote)):
                for para in elem.paragraphs:
                    for span in para.spans:
                        count += len(re.sub(r"\s", "", span.text))
            elif isinstance(elem, ReferenceListBlock):
                for entry in elem.entries:
                    count += len(re.sub(r"\s", "", entry))
            elif isinstance(elem, DegradationBlock):
                count += len(re.sub(r"\s", "", elem.fallback_text))
            elif isinstance(elem, (FormulaBlock, MathBlock)):
                count += len(re.sub(r"\s", "", elem.source or elem.latex))
    return count


def _build_config(
    doc: StructuredDocument, issues: list[DocumentIssue]
) -> LongformConfig:
    """Resolve explicit and automatic defaults for long-form configuration."""
    metadata = {k: _normalize_text(v) for k, v in doc.metadata.items()}

    config = LongformConfig()
    config.title = _normalize_text(doc.title)
    config.short_title = metadata.get("short_title", "")
    config.author = metadata.get("author", "")
    config.date = metadata.get("date", "")
    config.header = metadata.get("header", "")

    # title_page
    raw_title_page = doc.metadata.get("title_page", "")
    if _is_auto(raw_title_page):
        config.title_page = bool(config.title and (config.author or config.date))
    else:
        parsed = _parse_bool(raw_title_page)
        if parsed is None:
            issues.append(
                _issue(
                    CONFIG_VALUE_INVALID,
                    "Invalid value for 'title_page'; using automatic default.",
                )
            )
            config.title_page = bool(
                config.title and (config.author or config.date)
            )
        else:
            config.title_page = parsed

    if config.title_page and not config.title:
        config.title_page = False
        issues.append(_issue(TITLE_MISSING, "title_page requested but title is empty."))

    # toc
    raw_toc = doc.metadata.get("toc", "")
    if _is_auto(raw_toc):
        heading_count = sum(
            1
            for s in doc.sections
            if s.level in {1, 2, 3} and s.has_heading
        )
        config.toc = heading_count >= 3 and _body_codepoint_count(doc.sections) >= 3000
    else:
        parsed = _parse_bool(raw_toc)
        if parsed is None:
            issues.append(
                _issue(
                    CONFIG_VALUE_INVALID,
                    "Invalid value for 'toc'; using automatic default.",
                )
            )
            heading_count = sum(
                1
                for s in doc.sections
                if s.level in {1, 2, 3} and s.has_heading
            )
            config.toc = (
                heading_count >= 3 and _body_codepoint_count(doc.sections) >= 3000
            )
        else:
            config.toc = parsed

    # figure_index / table_index
    for key in ("figure_index", "table_index"):
        raw = doc.metadata.get(key, "")
        parsed = _parse_bool(raw)
        if raw and parsed is None:
            issues.append(
                _issue(
                    CONFIG_VALUE_INVALID,
                    f"Invalid boolean value for '{key}'; using false.",
                )
            )
            parsed = False
        setattr(config, key, bool(parsed))

    # bibliography_include_uncited
    raw_bib = doc.metadata.get("bibliography_include_uncited", "")
    parsed = _parse_bool(raw_bib)
    if raw_bib and parsed is None:
        issues.append(
            _issue(
                CONFIG_VALUE_INVALID,
                "Invalid boolean value for 'bibliography_include_uncited'; "
                "using true.",
            )
        )
        parsed = True
    config.bibliography_include_uncited = bool(parsed if raw_bib else True)

    # caption_numbering
    raw_caption = doc.metadata.get("caption_numbering", "")
    if _is_auto(raw_caption) or not raw_caption:
        config.caption_numbering = "auto"  # resolved later after heading numbering
    else:
        value = str(raw_caption).strip().lower()
        if value in _VALID_CAPTION_NUMBERING:
            config.caption_numbering = value
        else:
            issues.append(
                _issue(
                    CONFIG_VALUE_INVALID,
                    "Invalid value for 'caption_numbering'; using 'global'.",
                )
            )
            config.caption_numbering = "global"

    # heading_numbering
    raw_heading = doc.metadata.get("heading_numbering", "")
    if _is_auto(raw_heading) or not raw_heading:
        config.heading_numbering = "auto"
    else:
        value = str(raw_heading).strip().lower()
        if value in _VALID_HEADING_NUMBERING:
            config.heading_numbering = value
        else:
            issues.append(
                _issue(
                    CONFIG_VALUE_INVALID,
                    "Invalid value for 'heading_numbering'; using 'auto'.",
                )
            )
            config.heading_numbering = "auto"

    # layout_engine
    raw_layout = doc.metadata.get("layout_engine", "")
    if raw_layout:
        config.layout_engine = str(raw_layout).strip().lower()
    else:
        config.layout_engine = "longform"

    return config


def _collect_explicit_targets(
    doc: StructuredDocument, issues: list[DocumentIssue]
) -> tuple[dict[str, tuple[str, str]], list[str]]:
    """Collect explicit reference targets: id -> (kind, node_id)."""
    targets: dict[str, tuple[str, str]] = {}
    seen: set[str] = set()

    def register(identifier: str, kind: str, node_id: str) -> None:
        if not identifier:
            return
        if identifier.startswith(_RESERVED_ID_PREFIX) or not _valid_explicit_id(
            identifier
        ):
            issues.append(
                _issue(
                    INVALID_EXPLICIT_ID,
                    f"Invalid explicit ID '{identifier}' for {kind} block.",
                    placement="block",
                )
            )
            return
        if identifier in seen:
            issues.append(
                _issue(
                    DUPLICATE_EXPLICIT_ID,
                    f"Duplicate explicit ID '{identifier}'.",
                    placement="block",
                )
            )
            return
        seen.add(identifier)
        targets[identifier] = (kind, node_id)

    sec_counter = 0
    for section in doc.sections:
        sec_counter += 1
        if section.node_id is None:
            section.node_id = f"__wpsc_sec:{sec_counter}"
        counters: dict[str, int] = {}
        for elem in section.elements:
            def next_id(kind: str) -> str:
                counters[kind] = counters.get(kind, 0) + 1
                return f"__wpsc_{kind}:{sec_counter}:{counters[kind]}"

            if isinstance(elem, FigureBlock):
                if elem.identifier:
                    if _valid_explicit_id(elem.identifier) and elem.identifier not in seen:
                        register(elem.identifier, "fig", elem.identifier)
                        elem.node_id = elem.identifier
                    else:
                        if elem.identifier in seen:
                            issues.append(_issue(DUPLICATE_EXPLICIT_ID, f"Duplicate explicit ID '{elem.identifier}'.", placement="block"))
                        elif not _valid_explicit_id(elem.identifier):
                            register(elem.identifier, "fig", "")
                        elem.node_id = next_id("fig")
                elif elem.node_id is None:
                    elem.node_id = next_id("fig")
            elif isinstance(elem, SemanticTableBlock):
                if elem.identifier:
                    if _valid_explicit_id(elem.identifier) and elem.identifier not in seen:
                        register(elem.identifier, "tab", elem.identifier)
                        elem.node_id = elem.identifier
                    else:
                        if elem.identifier in seen:
                            issues.append(_issue(DUPLICATE_EXPLICIT_ID, f"Duplicate explicit ID '{elem.identifier}'.", placement="block"))
                        elif not _valid_explicit_id(elem.identifier):
                            register(elem.identifier, "tab", "")
                        elem.node_id = next_id("tab")
                elif elem.node_id is None:
                    elem.node_id = next_id("tab")
            elif isinstance(elem, FormulaBlock):
                if elem.identifier:
                    if _valid_explicit_id(elem.identifier) and elem.identifier not in seen:
                        register(elem.identifier, "eq", elem.identifier)
                        elem.node_id = elem.identifier
                    else:
                        if elem.identifier in seen:
                            issues.append(_issue(DUPLICATE_EXPLICIT_ID, f"Duplicate explicit ID '{elem.identifier}'.", placement="block"))
                        elif not _valid_explicit_id(elem.identifier):
                            register(elem.identifier, "eq", "")
                        elem.node_id = next_id("eq")
                elif elem.node_id is None:
                    elem.node_id = next_id("eq")
            elif isinstance(elem, ReferenceListBlock):
                if elem.identifier:
                    if _valid_explicit_id(elem.identifier) and elem.identifier not in seen:
                        register(elem.identifier, "ref", elem.identifier)
                        elem.node_id = elem.identifier
                    else:
                        if elem.identifier in seen:
                            issues.append(_issue(DUPLICATE_EXPLICIT_ID, f"Duplicate explicit ID '{elem.identifier}'.", placement="block"))
                        elif not _valid_explicit_id(elem.identifier):
                            register(elem.identifier, "ref", "")
                        elem.node_id = next_id("ref")
                elif elem.node_id is None:
                    elem.node_id = next_id("ref")
            elif isinstance(elem, PageBreakBlock) and elem.node_id is None:
                elem.node_id = next_id("pb")
            elif isinstance(elem, DegradationBlock) and elem.node_id is None:
                elem.node_id = next_id("deg")

    return targets, list(seen)


def _parse_bibliography_entries(
    entries: list[str], issues: list[DocumentIssue]
) -> list[tuple[str, str]]:
    """Parse bibliography entries into (id, text) pairs."""
    parsed: list[tuple[str, str]] = []
    current_id: Optional[str] = None
    current_text_parts: list[str] = []

    def flush() -> None:
        nonlocal current_id
        if current_id is not None:
            parsed.append((current_id, " ".join(current_text_parts).strip()))
            current_id = None
            current_text_parts.clear()

    for raw in entries:
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            continue

        yaml_id = re.match(r"^[-*]\s+id:\s*(\S+)(?:\s+text:\s*(.*))?$", stripped)
        if yaml_id:
            flush()
            current_id = yaml_id.group(1)
            if yaml_id.group(2):
                current_text_parts.append(yaml_id.group(2).strip())
            continue

        text_only = re.match(r"^text:\s*(.*)$", stripped)
        if text_only and current_id is not None:
            current_text_parts.append(text_only.group(1).strip())
            continue

        bracket = re.match(r"^\[([^\]]+)\]\s*(.*)$", stripped)
        if bracket:
            flush()
            current_id = bracket.group(1)
            current_text_parts.append(bracket.group(2).strip())
            continue

        simple = re.match(r"^(\S+?):\s+(.*)$", stripped)
        if simple:
            flush()
            current_id = simple.group(1)
            current_text_parts.append(simple.group(2).strip())
            continue

        if current_id is not None:
            current_text_parts.append(stripped)

    flush()
    return parsed


def _scan_inline_references(
    sections: list[Section], issues: list[DocumentIssue]
) -> tuple[set[str], set[str]]:
    """Scan visible text for {{ref:...}} and {{cite:...}} markers."""
    refs: set[str] = set()
    cites: set[str] = set()

    def scan_spans(spans: list[Span], collector: set[str], pattern: re.Pattern) -> None:
        for span in spans:
            if span.code or span.math:
                continue
            for m in pattern.finditer(span.text):
                collector.add(m.group(1).strip())

    for section in sections:
        for elem in section.elements:
            if isinstance(elem, Paragraph):
                scan_spans(elem.spans, refs, _REF_RE)
                scan_spans(elem.spans, cites, _CITE_RE)
            elif isinstance(elem, ListBlock):
                for item in elem.items:
                    scan_spans(item, refs, _REF_RE)
                    scan_spans(item, cites, _CITE_RE)
            elif isinstance(elem, (AbstractBlock, BlockQuote)):
                for para in elem.paragraphs:
                    scan_spans(para.spans, refs, _REF_RE)
                    scan_spans(para.spans, cites, _CITE_RE)
            elif isinstance(elem, TableBlock):
                for cell in elem.headers:
                    scan_spans([Span(text=cell)], refs, _REF_RE)
                    scan_spans([Span(text=cell)], cites, _CITE_RE)
                for row in elem.rows:
                    for cell in row:
                        scan_spans([Span(text=cell)], refs, _REF_RE)
                        scan_spans([Span(text=cell)], cites, _CITE_RE)
            elif isinstance(elem, SemanticTableBlock):
                for cell in elem.headers:
                    scan_spans([Span(text=cell)], refs, _REF_RE)
                    scan_spans([Span(text=cell)], cites, _CITE_RE)
                for row in elem.rows:
                    for cell in row:
                        scan_spans([Span(text=cell)], refs, _REF_RE)
                        scan_spans([Span(text=cell)], cites, _CITE_RE)
            elif isinstance(elem, FigureBlock):
                for img in elem.images:
                    scan_spans([Span(text=img.alt)], refs, _REF_RE)
                    scan_spans([Span(text=img.alt)], cites, _CITE_RE)
            elif isinstance(elem, DegradationBlock):
                scan_spans(
                    [Span(text=elem.fallback_text)], refs, _REF_RE
                )
                scan_spans(
                    [Span(text=elem.fallback_text)], cites, _CITE_RE
                )

    return refs, cites


def _build_references(
    doc: StructuredDocument,
    targets: dict[str, tuple[str, str]],
    explicit_ids: list[str],
    issues: list[DocumentIssue],
) -> dict[str, dict[str, Any]]:
    """Resolve cross-references and citations against collected targets."""
    references: dict[str, dict[str, Any]] = {}

    # Bibliography entries are declared inside ReferenceListBlock elements.
    bib_entries: dict[str, str] = {}
    for section in doc.sections:
        for elem in section.elements:
            if isinstance(elem, ReferenceListBlock):
                for bid, btext in _parse_bibliography_entries(
                    elem.entries, issues
                ):
                    if bid in bib_entries:
                        issues.append(
                            _issue(
                                DUPLICATE_EXPLICIT_ID,
                                f"Duplicate bibliography ID '{bid}'.",
                                placement="block",
                            )
                        )
                        continue
                    bib_entries[bid] = btext
                    if _valid_explicit_id(bid):
                        references[bid] = {
                            "cited": False,
                            "kind": "ref",
                            "node_id": bid,
                            "text": btext,
                        }
                    else:
                        issues.append(
                            _issue(
                                INVALID_EXPLICIT_ID,
                                f"Invalid bibliography ID '{bid}'.",
                                placement="block",
                            )
                        )

    # Add figure/table/formula targets to references map.
    for identifier, (kind, node_id) in targets.items():
        references[identifier] = {
            "kind": kind,
            "node_id": node_id,
        }

    requested_refs, requested_cites = _scan_inline_references(
        doc.sections, issues
    )

    for ref_id in requested_refs:
        if ref_id not in references:
            issues.append(
                _issue(
                    REFERENCE_UNRESOLVED,
                    f"Cross-reference target '{ref_id}' was not found.",
                    placement="inline",
                )
            )
        elif references[ref_id]["kind"] == "ref":
            # References are bibliography entries, not figure/table/equation.
            issues.append(
                _issue(
                    REFERENCE_UNRESOLVED,
                    f"Cross-reference target '{ref_id}' is a bibliography entry, "
                    f"not a figure/table/equation.",
                    placement="inline",
                )
            )

    for cite_id in requested_cites:
        if cite_id in references:
            references[cite_id]["cited"] = True
        else:
            issues.append(
                _issue(
                    REFERENCE_UNRESOLVED,
                    f"Citation target '{cite_id}' was not found.",
                    placement="inline",
                )
            )

    return references


def _map_bookmarks_from_targets(
    targets: dict[str, tuple[str, str]]
) -> BookmarkMapResult:
    """Map explicit IDs to deterministic internal WPS bookmark names by kind."""
    by_kind: dict[str, list[str]] = {}
    for identifier, (kind, _node_id) in targets.items():
        by_kind.setdefault(kind, []).append(identifier)

    combined_mapping: dict[str, str] = {}
    combined_issues: list[str] = []
    for kind in sorted(by_kind):
        result = map_bookmarks(by_kind[kind], kind)
        combined_mapping.update(result.mapping)
        combined_issues.extend(result.issues)

    return BookmarkMapResult(
        mapping=combined_mapping, issues=tuple(combined_issues)
    )


def _update_caption_numbering(
    config: LongformConfig, sections: list[Section]
) -> None:
    """Update auto caption numbering once heading numbering is resolved."""
    if config.caption_numbering != "auto":
        return
    has_numbered_h1 = any(
        s.level == 1 and s.numbering not in ("none", "auto")
        for s in sections
    )
    config.caption_numbering = "chapter" if has_numbered_h1 else "global"


_NORMALIZATION_ERROR = "LONGFORM_NORMALIZATION_ERROR"


def normalize_longform_document(
    doc: StructuredDocument,
) -> SemanticResult:
    """Normalize a parsed long-form document.

    Returns a SemanticResult with deterministic JSON output and no raised
    exceptions for malformed input.
    """
    try:
        doc = copy.deepcopy(doc)
        issues: list[DocumentIssue] = list(doc.issues)

        config = _build_config(doc, issues)
        config.title = _normalize_text(doc.title)

        sections = _consume_title_h1(doc.sections, config.title)
        doc.sections = sections

        selected_scheme = _apply_heading_numbering(
            doc.sections, config.heading_numbering, config.title, issues
        )
        if config.heading_numbering == "auto":
            config.heading_numbering = selected_scheme

        _derive_header(config, issues)

        targets, explicit_ids = _collect_explicit_targets(doc, issues)
        references = _build_references(doc, targets, explicit_ids, issues)
        bookmarks = _map_bookmarks_from_targets(targets)

        _update_caption_numbering(config, doc.sections)

        # Store resolved config on the document for downstream consumers.
        doc.config = _canonical_value(config.to_json())

        return SemanticResult(
            document=doc,
            config=config,
            references=references,
            bookmarks=bookmarks,
            issues=tuple(issues),
        )
    except Exception as exc:  # noqa: BLE001 - deterministic degradation
        safe_doc = copy.deepcopy(doc)
        safe_config = LongformConfig(title=_normalize_text(safe_doc.title))
        return SemanticResult(
            document=safe_doc,
            config=safe_config,
            references={},
            bookmarks=BookmarkMapResult(mapping={}, issues=()),
            issues=tuple(
                safe_doc.issues
            ) + (
                _issue(
                    _NORMALIZATION_ERROR,
                    f"Semantic normalization failed: {exc}.",
                    placement="document",
                ),
            ),
        )


__all__ = [
    "CONFIG_VALUE_INVALID",
    "DUPLICATE_EXPLICIT_ID",
    "HEADER_SHORTENED",
    "HEADING_LEVEL_GAP",
    "HEADING_PREFIX_AMBIGUOUS",
    "INVALID_EXPLICIT_ID",
    "LongformConfig",
    "REFERENCE_UNRESOLVED",
    "SemanticResult",
    "TITLE_MISSING",
    "normalize_longform_document",
]
