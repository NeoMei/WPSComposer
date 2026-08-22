
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
    BibliographyEntry,
    CodeBlock,
    CitationRun,
    DegradationBlock,
    DocumentIssue,
    ExcalidrawBlock,
    FigureBlock,
    FormulaBlock,
    BlockQuote,
    CaptionBinding,
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
    CrossReferenceRun,
    InlineDegradationRun,
    StructuredDocument,
    TableBlock,
    TableCellDegradation,
    TaskList,
)
from .native_math import NativeMathConversionError, convert_restricted_latex
from .bookmark_ids import (
    BOOKMARK_COLLISION_UNRESOLVED,
    BOOKMARK_NAME_COLLISION,
    BookmarkMapResult,
    map_bookmarks,
)
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
CAPTION_MISSING = "CAPTION_MISSING"
DUPLICATE_EXPLICIT_ID = "DUPLICATE_EXPLICIT_ID"
INVALID_EXPLICIT_ID = "INVALID_EXPLICIT_ID"
ABSTRACT_CONTENT_DEGRADED = "ABSTRACT_CONTENT_DEGRADED"
PAGE_BREAK_CONTENT_DEGRADED = "PAGE_BREAK_CONTENT_DEGRADED"
PAGE_ROLE_RESOLUTION_FAILED = "PAGE_ROLE_RESOLUTION_FAILED"
BIBLIOGRAPHY_ENTRY_MALFORMED = "BIBLIOGRAPHY_ENTRY_MALFORMED"

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
_REFERENCE_OR_CITATION_RE = re.compile(r"\{\{(ref|cite):([^}]+)\}\}")

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
    if isinstance(value, FormulaBlock):
        # The local fallback declaration is private resource input. It is
        # intentionally absent from semantic diagnostics and snapshots.
        return {
            field.name: _canonical_value(getattr(value, field.name))
            for field in sorted(dataclasses.fields(value), key=lambda item: item.name)
            if field.name not in {"fallback_image", "raw_source"}
        }
    if dataclasses.is_dataclass(value):
        return {
            field.name: _canonical_value(getattr(value, field.name))
            for field in sorted(dataclasses.fields(value), key=lambda item: item.name)
        }
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
    if _is_auto(raw_title_page) or not str(raw_title_page).strip():
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
    if _is_auto(raw_toc) or not str(raw_toc).strip():
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
    """Collect every explicit ID in semantic source order.

    The namespace is shared by headings, bibliography entries, figures,
    tables, and formulas. The first valid declaration is the only target;
    later declarations retain readable content and carry a node-local issue.
    """
    targets: dict[str, tuple[str, str]] = {}
    seen: set[str] = set()
    bibliography_declaration_counter = 0

    def register(identifier: str, kind: str, node_id: str) -> Optional[DocumentIssue]:
        if not identifier:
            return None
        if identifier.startswith(_RESERVED_ID_PREFIX) or not _valid_explicit_id(
            identifier
        ):
            issue = _issue(
                INVALID_EXPLICIT_ID,
                f"Invalid explicit ID '{identifier}' for {kind} block.",
                placement="block",
            )
            issues.append(issue)
            return issue
        if identifier in seen:
            issue = _issue(
                DUPLICATE_EXPLICIT_ID,
                f"Duplicate explicit ID '{identifier}'.",
                placement="block",
            )
            issues.append(issue)
            return issue
        seen.add(identifier)
        targets[identifier] = (kind, node_id)
        return None

    sec_counter = 0
    for section in doc.sections:
        sec_counter += 1
        explicit_section_id = (
            section.node_id
            if section.node_id and not section.node_id.startswith(_RESERVED_ID_PREFIX)
            else None
        )
        if explicit_section_id is not None:
            section.target_degradation = register(
                explicit_section_id, "sec", explicit_section_id
            )
        elif section.node_id is None:
            section.node_id = f"__wpsc_sec:{sec_counter}"
        counters: dict[str, int] = {}
        for elem in section.elements:
            def next_id(kind: str) -> str:
                counters[kind] = counters.get(kind, 0) + 1
                return f"__wpsc_{kind}:{sec_counter}:{counters[kind]}"

            if isinstance(elem, FigureBlock):
                if elem.identifier:
                    degradation = register(elem.identifier, "fig", elem.identifier)
                    elem.target_degradation = degradation
                    if degradation is None:
                        elem.node_id = elem.identifier
                    else:
                        elem.node_id = next_id("fig")
                elif elem.node_id is None:
                    elem.node_id = next_id("fig")
            elif isinstance(elem, SemanticTableBlock):
                if elem.identifier:
                    degradation = register(elem.identifier, "tab", elem.identifier)
                    elem.target_degradation = degradation
                    if degradation is None:
                        elem.node_id = elem.identifier
                    else:
                        elem.node_id = next_id("tab")
                elif elem.node_id is None:
                    elem.node_id = next_id("tab")
            elif isinstance(elem, FormulaBlock):
                if elem.identifier:
                    degradation = register(elem.identifier, "eq", elem.identifier)
                    elem.target_degradation = degradation
                    if degradation is None:
                        elem.node_id = elem.identifier
                    else:
                        elem.node_id = next_id("eq")
                elif elem.node_id is None:
                    elem.node_id = next_id("eq")
            elif isinstance(elem, ReferenceListBlock):
                if elem.identifier:
                    degradation = register(elem.identifier, "ref", elem.identifier)
                    elem.target_degradation = degradation
                    if degradation is None:
                        elem.node_id = elem.identifier
                    else:
                        elem.node_id = next_id("ref")
                elif elem.node_id is None:
                    elem.node_id = next_id("ref")
                candidates = _parse_bibliography_entries(elem.entries)
                setattr(elem, "_bibliography_candidates", candidates)
                for candidate in candidates:
                    if candidate.identifier is None:
                        degradation = DegradationBlock(
                            issue=_issue(
                                BIBLIOGRAPHY_ENTRY_MALFORMED,
                                "Bibliography entry is malformed; readable source was retained.",
                                placement="block",
                            ),
                            node_id=f"{elem.node_id}/degradation:{candidate.source_index}",
                            fallback_text=candidate.raw,
                        )
                        candidate.degradation = degradation
                        issues.append(degradation.issue)
                        continue
                    bibliography_declaration_counter += 1
                    candidate.declaration_index = bibliography_declaration_counter
                    candidate.node_id = (
                        f"{elem.node_id}/entry:{candidate.declaration_index}"
                    )
                    candidate.target_degradation = register(
                        candidate.identifier, "ref", candidate.node_id
                    )
                    if candidate.target_degradation is not None:
                        candidate.degradation = DegradationBlock(
                            issue=candidate.target_degradation,
                            node_id=(
                                f"{elem.node_id}/degradation:{candidate.source_index}"
                            ),
                            fallback_text=candidate.raw,
                        )
            elif isinstance(elem, PageBreakBlock) and elem.node_id is None:
                elem.node_id = next_id("pb")
            elif isinstance(elem, DegradationBlock) and elem.node_id is None:
                elem.node_id = next_id("deg")

    return targets, list(seen)


@dataclass
class _BibliographyCandidate:
    identifier: Optional[str]
    text: str
    raw: str
    source_index: int
    declaration_index: int = 0
    node_id: str = ""
    target_degradation: Optional[DocumentIssue] = None
    degradation: Optional[DegradationBlock] = None


def _parse_bibliography_entries(
    entries: list[str],
) -> list[_BibliographyCandidate]:
    """Parse entries without discarding any malformed visible source."""
    parsed: list[_BibliographyCandidate] = []
    declaration_index = 0
    index = 0
    while index < len(entries):
        raw = entries[index].rstrip()
        stripped = raw.strip()
        source_index = index + 1
        index += 1
        if not stripped:
            continue

        yaml_id = re.match(
            r"^[-*]\s+id:\s*(\S+)(?:\s+text:\s*(.*))?$", stripped
        )
        if yaml_id:
            identifier = normalize_visible_text(yaml_id.group(1).strip())
            text_parts = [yaml_id.group(2).strip()] if yaml_id.group(2) else []
            raw_parts = [raw]
            if index < len(entries):
                text_only = re.match(r"^\s*text:\s*(.*)$", entries[index])
                if text_only:
                    raw_parts.append(entries[index].rstrip())
                    text_parts.append(text_only.group(1).strip())
                    index += 1
            declaration_index += 1
            text = " ".join(part for part in text_parts if part).strip()
            if identifier and text:
                parsed.append(_BibliographyCandidate(
                    identifier, text, "\n".join(raw_parts), source_index,
                    declaration_index,
                ))
            else:
                parsed.append(_BibliographyCandidate(
                    None, "", "\n".join(raw_parts), source_index
                ))
            continue

        bracket = re.match(r"^\[([^\]]+)\]\s*(.*)$", stripped)
        simple = re.match(r"^(\S+?):\s+(.*)$", stripped)
        match = bracket or simple
        if match and match.group(1).strip() and match.group(2).strip():
            declaration_index += 1
            parsed.append(_BibliographyCandidate(
                normalize_visible_text(match.group(1).strip()),
                normalize_visible_text(match.group(2).strip()),
                raw,
                source_index,
                declaration_index,
            ))
        else:
            parsed.append(_BibliographyCandidate(None, "", raw, source_index))
    return parsed


def _scan_inline_references(
    doc: StructuredDocument,
) -> tuple[tuple[str, str], ...]:
    """Collect ref/cite requests in first visible semantic-source order."""
    requested: list[tuple[str, str]] = []

    def add_request(kind: str, target_id: str) -> None:
        request = (kind, normalize_visible_text(target_id.strip()))
        if request[1]:
            requested.append(request)

    def scan_spans(spans: list[Span]) -> None:
        for span in spans:
            if span.code or span.math or span.semantic_literal:
                continue
            if span.cross_reference is not None:
                add_request("ref", span.cross_reference.target_id)
            for match in _iter_visible_markers(span.text):
                add_request(match.group(1), match.group(2))

    def scan_element(elem: Any) -> None:
        if isinstance(elem, Paragraph):
            scan_spans(elem.spans)
        elif isinstance(elem, ListBlock):
            for item in elem.items:
                scan_spans(item)
        elif isinstance(elem, (AbstractBlock, BlockQuote)):
            for para in elem.paragraphs:
                scan_element(para)
        elif isinstance(elem, TableBlock):
            for cell in elem.headers:
                scan_spans([Span(text=cell)])
            for row in elem.rows:
                for cell in row:
                    scan_spans([Span(text=cell)])
        elif isinstance(elem, SemanticTableBlock):
            for cell in elem.headers:
                scan_spans([Span(text=cell)])
            for row in elem.rows:
                for cell in row:
                    scan_spans([Span(text=cell)])
        elif isinstance(elem, PageBreakBlock):
            for para in elem.content:
                scan_element(para)
        elif isinstance(elem, Section):
            for child in elem.elements:
                scan_element(child)

    if doc.abstract is not None:
        scan_element(doc.abstract)
    for section in doc.sections:
        scan_element(section)

    return tuple(requested)


_LITERAL_INLINE_RE = re.compile(
    r"`[^`]*`|\$[^$]*\$|!\[[^\]]*\]\([^)]*\)"
)


def _iter_visible_markers(text: str):
    """Yield markers outside code, inline math, and image-alt literals."""
    literal_ranges = [match.span() for match in _LITERAL_INLINE_RE.finditer(text)]
    for match in _REFERENCE_OR_CITATION_RE.finditer(text):
        if any(start <= match.start() < end for start, end in literal_ranges):
            continue
        slash_count = 0
        cursor = match.start() - 1
        while cursor >= 0 and text[cursor] == "\\":
            slash_count += 1
            cursor -= 1
        if slash_count % 2:
            continue
        yield match


def _build_references(
    doc: StructuredDocument,
    targets: dict[str, tuple[str, str]],
    bookmarks: BookmarkMapResult,
    issues: list[DocumentIssue],
    config: LongformConfig,
) -> dict[str, dict[str, Any]]:
    """Resolve cross-references and citations against collected targets."""
    references: dict[str, dict[str, Any]] = {}

    all_bibliography_candidates: list[_BibliographyCandidate] = []
    for section in doc.sections:
        for elem in section.elements:
            if isinstance(elem, ReferenceListBlock):
                all_bibliography_candidates.extend(
                    getattr(elem, "_bibliography_candidates", ())
                )
    bibliography_candidates = [
        candidate
        for candidate in all_bibliography_candidates
        if candidate.identifier is not None
    ]
    for candidate in bibliography_candidates:
        if candidate.target_degradation is not None:
            continue
        references[candidate.identifier] = {
            "cited": False,
            "declaration_index": candidate.declaration_index,
            "kind": "ref",
            "node_id": candidate.node_id,
            "number": None,
            "text": candidate.text,
        }

    # Add object and explicit-section targets from the same global namespace.
    for identifier, (kind, node_id) in targets.items():
        if kind == "ref":
            continue
        reference = {"kind": kind, "node_id": node_id}
        bookmark_name = bookmarks.mapping.get(identifier)
        if bookmark_name is not None:
            reference["bookmark_name"] = bookmark_name
        references[identifier] = reference

    next_number = 1
    reported: set[tuple[str, str]] = set()
    for request_kind, target_id in _scan_inline_references(doc):
        if request_kind == "ref":
            target = references.get(target_id)
            if target is None or target.get("kind") not in {"fig", "tab", "eq"} or not target.get("bookmark_name"):
                request_key = (request_kind, target_id)
                if request_key in reported:
                    continue
                reported.add(request_key)
                issues.append(
                    _issue(
                        REFERENCE_UNRESOLVED,
                        f"Cross-reference target '{target_id}' was not found.",
                        placement="inline",
                    )
                )
        else:
            target = references.get(target_id)
            if target is not None and target.get("kind") == "ref":
                target["cited"] = True
                if target["number"] is None:
                    target["number"] = next_number
                    next_number += 1
                continue
            request_key = (request_kind, target_id)
            if request_key not in reported:
                reported.add(request_key)
                issues.append(
                    _issue(
                        REFERENCE_UNRESOLVED,
                        f"Citation target '{target_id}' was not found.",
                        placement="inline",
                    )
                )

    if config.bibliography_include_uncited:
        for candidate in sorted(
            bibliography_candidates, key=lambda item: item.declaration_index
        ):
            target = references.get(candidate.identifier)
            if target is not None and target.get("kind") == "ref" and target["number"] is None:
                target["number"] = next_number
                next_number += 1

    ordered_entries = sorted(
        (
            BibliographyEntry(
                identifier=identifier,
                node_id=value["node_id"],
                text=value["text"],
                number=value["number"],
                declaration_index=value["declaration_index"],
                cited=bool(value["cited"]),
            )
            for identifier, value in references.items()
            if value.get("kind") == "ref" and value.get("number") is not None
        ),
        key=lambda entry: entry.number,
    )
    reference_blocks: list[ReferenceListBlock] = []
    for section in doc.sections:
        for elem in section.elements:
            if isinstance(elem, ReferenceListBlock):
                reference_blocks.append(elem)

    # Cited entries lead in citation-number order. The remaining declaration
    # stream stays source-ordered, including malformed/duplicate degradations
    # at their exact relative positions among uncited entries.
    entry_by_node = {entry.node_id: entry for entry in ordered_entries}
    cited_entries = [entry for entry in ordered_entries if entry.cited]
    declaration_tail: list[Any] = []
    for candidate in all_bibliography_candidates:
        if candidate.degradation is not None:
            declaration_tail.append(candidate.degradation)
            continue
        entry = entry_by_node.get(candidate.node_id)
        if entry is not None and not entry.cited:
            declaration_tail.append(entry)
    for block in reference_blocks:
        block.resolved_items = []
    if reference_blocks:
        reference_blocks[0].resolved_items = cited_entries + declaration_tail

    return references


def _move_bibliography_to_back_matter(doc: StructuredDocument) -> None:
    """Move non-empty bibliography blocks to one final isolated page role."""
    bibliography_blocks: list[ReferenceListBlock] = []
    retained_sections: list[Section] = []
    for section in doc.sections:
        retained_elements: list[Any] = []
        section_had_bibliography = False
        for element in section.elements:
            if isinstance(element, ReferenceListBlock):
                section_had_bibliography = True
                if element.target_degradation is not None:
                    element.resolved_items.insert(0, DegradationBlock(
                        issue=element.target_degradation,
                        node_id=f"{element.node_id}/target-degradation",
                        fallback_text="Bibliography target capability was disabled.",
                    ))
                if element.resolved_items:
                    bibliography_blocks.append(element)
            else:
                retained_elements.append(element)
        section.elements = retained_elements
        if section.elements or section.has_heading:
            # A heading that introduced only bibliography content belongs to
            # back matter and must not leave an empty body section.
            if not section.elements and section_had_bibliography:
                continue
            retained_sections.append(section)
    if bibliography_blocks:
        retained_sections.append(Section(
            level=0,
            heading="",
            elements=bibliography_blocks,
            node_id="__wpsc_sec:bibliography",
            numbering="none",
            page_role="bibliography",
        ))
    doc.sections = retained_sections


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


def _normalize_formulas(
    doc: StructuredDocument, issues: list[DocumentIssue]
) -> None:
    """Attach only converter-issued native descriptors to formula nodes."""
    for section in doc.sections:
        for element in section.elements:
            if not isinstance(element, FormulaBlock):
                continue
            try:
                element.native_math = convert_restricted_latex(element.source)
                element.content_degradation = None
            except NativeMathConversionError as error:
                issue = _issue(
                    error.code,
                    "Formula cannot use editable native math; readable source was retained.",
                    placement="block",
                )
                element.native_math = None
                element.content_degradation = issue
                issues.append(issue)


def _caption_reference_targets(
    doc: StructuredDocument,
    targets: dict[str, tuple[str, str]],
    issues: list[DocumentIssue],
) -> dict[str, tuple[str, str]]:
    """Return explicit object targets eligible for a native number bookmark."""
    eligible: dict[str, tuple[str, str]] = {}
    for section in doc.sections:
        for element in section.elements:
            kind: Optional[str] = None
            indexable = True
            if isinstance(element, FigureBlock):
                kind = "fig"
                element.caption = _normalize_text(element.caption)
                indexable = bool(element.caption)
            elif isinstance(element, SemanticTableBlock):
                kind = "tab"
                element.caption = _normalize_text(element.caption)
                indexable = bool(element.caption)
            elif isinstance(element, FormulaBlock):
                kind = "eq"
            else:
                continue

            if not indexable:
                issues.append(
                    _issue(
                        CAPTION_MISSING,
                        f"{kind} block has no non-empty caption; numbering and "
                        "cross-reference capability were disabled.",
                        placement="block",
                    )
                )
                continue

            identifier = element.identifier
            if (
                identifier
                and targets.get(identifier) == (kind, element.node_id)
            ):
                eligible[identifier] = (kind, element.node_id)
    return eligible


def _apply_caption_bindings(
    doc: StructuredDocument,
    config: LongformConfig,
    targets: dict[str, tuple[str, str]],
    bookmarks: BookmarkMapResult,
) -> None:
    """Resolve chapter/global numbering on each object in source order."""
    current_chapter_node_id: Optional[str] = None
    for section in doc.sections:
        if (
            section.level == 1
            and section.numbering not in ("none", "auto")
        ):
            current_chapter_node_id = section.node_id

        for element in section.elements:
            if isinstance(element, FigureBlock):
                kind = "fig"
                indexable = bool(element.caption)
            elif isinstance(element, SemanticTableBlock):
                kind = "tab"
                indexable = bool(element.caption)
            elif isinstance(element, FormulaBlock):
                kind = "eq"
                indexable = True
            else:
                continue

            chapter_mode = (
                config.caption_numbering in {"auto", "chapter"}
                and current_chapter_node_id is not None
            )
            mode = "chapter" if chapter_mode else "global"
            chapter_node_id = current_chapter_node_id if chapter_mode else None

            identifier = element.identifier
            bookmark_name = None
            if (
                identifier
                and indexable
                and targets.get(identifier) == (kind, element.node_id)
            ):
                bookmark_name = bookmarks.mapping.get(identifier)
            element.caption_binding = CaptionBinding(
                mode=mode,
                chapter_node_id=chapter_node_id,
                bookmark_name=bookmark_name,
                indexable=indexable,
                referenceable=bookmark_name is not None,
            )


@dataclass
class _ParagraphSlot:
    """Mutable adapter over a Paragraph or one legacy ListBlock item."""

    owner: Any
    item_index: Optional[int] = None

    @property
    def node_id(self) -> Optional[str]:
        if isinstance(self.owner, Paragraph):
            return self.owner.node_id
        return self.owner.item_node_ids[self.item_index]

    @property
    def spans(self) -> list[Span]:
        if isinstance(self.owner, Paragraph):
            return self.owner.spans
        return self.owner.items[self.item_index]

    def assign_node_id(self, node_id: str) -> None:
        if isinstance(self.owner, Paragraph):
            self.owner.node_id = node_id
        else:
            self.owner.item_node_ids[self.item_index] = node_id

    def replace_spans(self, spans: list[Span]) -> None:
        if isinstance(self.owner, Paragraph):
            self.owner.spans = spans
        else:
            self.owner.items[self.item_index] = spans


def _iter_paragraph_slots(element: Any):
    if isinstance(element, Paragraph):
        yield _ParagraphSlot(element)
    elif isinstance(element, ListBlock):
        element.item_node_ids = (
            list(element.item_node_ids[: len(element.items)])
            + [None] * max(0, len(element.items) - len(element.item_node_ids))
        )
        for item_index in range(len(element.items)):
            yield _ParagraphSlot(element, item_index)
    elif isinstance(element, (AbstractBlock, BlockQuote)):
        for paragraph in element.paragraphs:
            yield _ParagraphSlot(paragraph)
    elif isinstance(element, PageBreakBlock):
        for paragraph in element.content:
            yield _ParagraphSlot(paragraph)
    elif isinstance(element, Section):
        for child in element.elements:
            yield from _iter_paragraph_slots(child)


def _assign_slot_ids(slots, section_index: int) -> None:
    for ordinal, slot in enumerate(slots, start=1):
        if slot.node_id is None:
            slot.assign_node_id(
                f"__wpsc_para:{section_index}:{ordinal}"
            )


def _assign_paragraph_ids(doc: StructuredDocument) -> None:
    """Assign canonical IDs to abstract, paragraph, and list-item spans."""
    if doc.abstract is not None:
        _assign_slot_ids(_iter_paragraph_slots(doc.abstract), 0)
    for section_index, section in enumerate(doc.sections, start=1):
        _assign_slot_ids(_iter_paragraph_slots(section), section_index)


def _split_slot_references(
    slot: _ParagraphSlot,
    references: dict[str, dict[str, Any]],
    fallback_text: str,
) -> None:
    if slot.node_id is None:
        return
    reference_occurrence = 0
    citation_occurrence = 0
    normalized_spans: list[Span] = []
    for span in slot.spans:
        matches = list(_iter_visible_markers(span.text))
        if span.code or span.math or span.semantic_literal or not matches:
            normalized_spans.append(span)
            continue

        cursor = 0
        for match in matches:
            if match.start() > cursor:
                normalized_spans.append(
                    dataclasses.replace(
                        span,
                        text=span.text[cursor:match.start()],
                        cross_reference=None,
                        citation=None,
                        inline_degradation=None,
                    )
                )
            kind = match.group(1)
            target_id = normalize_visible_text(match.group(2).strip())
            target = references.get(target_id)
            if kind == "ref":
                reference_occurrence += 1
                resolved = (
                    target is not None
                    and target.get("kind") in {"fig", "tab", "eq"}
                    and bool(target.get("bookmark_name"))
                )
                run = CrossReferenceRun(
                    node_id=f"{slot.node_id}/ref:{reference_occurrence}",
                    target_id=target_id,
                    target_node_id=(target["node_id"] if resolved else None),
                    target_kind=(target["kind"] if resolved else None),
                    bookmark_name=(target["bookmark_name"] if resolved else None),
                    fallback_text=fallback_text,
                )
                normalized_spans.append(dataclasses.replace(
                    span,
                    text=fallback_text,
                    cross_reference=run,
                    citation=None,
                    inline_degradation=None,
                ))
            else:
                citation_occurrence += 1
                node_id = f"{slot.node_id}/cite:{citation_occurrence}"
                resolved = (
                    target is not None
                    and target.get("kind") == "ref"
                    and isinstance(target.get("number"), int)
                )
                if resolved:
                    citation_text = f"[{target['number']}]"
                    normalized_spans.append(dataclasses.replace(
                        span,
                        text=citation_text,
                        cross_reference=None,
                        citation=CitationRun(
                            node_id=node_id,
                            target_id=target_id,
                            target_node_id=target["node_id"],
                            number=target["number"],
                            fallback_text=citation_text,
                        ),
                        inline_degradation=None,
                    ))
                else:
                    citation_text = (
                        "[REFERENCE_UNRESOLVED 引用目标未解析]"
                    )
                    normalized_spans.append(dataclasses.replace(
                        span,
                        text=citation_text,
                        cross_reference=None,
                        citation=None,
                        inline_degradation=InlineDegradationRun(
                            node_id=node_id,
                            code=REFERENCE_UNRESOLVED,
                            fallback_text=citation_text,
                        ),
                    ))
            cursor = match.end()
        if cursor < len(span.text):
            normalized_spans.append(
                dataclasses.replace(
                    span,
                    text=span.text[cursor:],
                    cross_reference=None,
                    citation=None,
                    inline_degradation=None,
                )
            )
    slot.replace_spans(normalized_spans)


def _split_cross_reference_spans(
    doc: StructuredDocument, references: dict[str, dict[str, Any]]
) -> None:
    """Replace non-code/math reference markers in every paragraph container."""
    fallback_text = normalize_visible_text("引用目标未解析")
    if doc.abstract is not None:
        for slot in _iter_paragraph_slots(doc.abstract):
            _split_slot_references(slot, references, fallback_text)
    for section in doc.sections:
        for slot in _iter_paragraph_slots(section):
            _split_slot_references(slot, references, fallback_text)


def _resolve_table_citation_text(
    text: str,
    references: dict[str, dict[str, Any]],
) -> tuple[str, bool]:
    pieces: list[str] = []
    cursor = 0
    degraded = False
    for match in _iter_visible_markers(text):
        if match.group(1) != "cite":
            continue
        pieces.append(text[cursor:match.start()])
        target_id = normalize_visible_text(match.group(2).strip())
        target = references.get(target_id)
        if (
            target is not None
            and target.get("kind") == "ref"
            and isinstance(target.get("number"), int)
        ):
            pieces.append(f"[{target['number']}]")
        else:
            pieces.append("[REFERENCE_UNRESOLVED 引用目标未解析]")
            degraded = True
        cursor = match.end()
    if not pieces:
        return text, False
    pieces.append(text[cursor:])
    return "".join(pieces), degraded


def _resolve_table_citations(
    doc: StructuredDocument, references: dict[str, dict[str, Any]]
) -> None:
    fallback = "[REFERENCE_UNRESOLVED 引用目标未解析]"
    for section in doc.sections:
        for element in section.elements:
            if not isinstance(element, (TableBlock, SemanticTableBlock)):
                continue
            element.cell_degradations = []
            for column, cell in enumerate(element.headers, start=1):
                resolved, degraded = _resolve_table_citation_text(cell, references)
                element.headers[column - 1] = resolved
                if degraded:
                    element.cell_degradations.append(TableCellDegradation(
                        row=1,
                        column=column,
                        code=REFERENCE_UNRESOLVED,
                        fallback_text=fallback,
                    ))
            for row_index, row in enumerate(element.rows, start=2):
                for column, cell in enumerate(row, start=1):
                    resolved, degraded = _resolve_table_citation_text(cell, references)
                    row[column - 1] = resolved
                    if degraded:
                        element.cell_degradations.append(TableCellDegradation(
                            row=row_index,
                            column=column,
                            code=REFERENCE_UNRESOLVED,
                            fallback_text=fallback,
                        ))


def _plain_text_from_element(elem: Any) -> str:
    """Extract readable plain text from any block element."""
    if isinstance(elem, Paragraph):
        return elem.plain_text
    if isinstance(elem, ListBlock):
        return " ".join(
            "".join(s.text for s in item)
            for item in elem.items
        )
    if isinstance(elem, Section):
        result: list[Paragraph] = []
        if elem.heading:
            result.append(Paragraph.from_text(elem.heading))
        for child in elem.elements:
            result.extend(_paragraphs_from_element(child))
        return " ".join(p.plain_text for p in result)
    if isinstance(elem, (AbstractBlock, BlockQuote)):
        return " ".join(p.plain_text for p in elem.paragraphs)
    if isinstance(elem, FigureBlock):
        return " ".join(
            img.alt for img in elem.images if img.alt
        )
    if isinstance(elem, SemanticTableBlock):
        parts = [elem.caption] if elem.caption else []
        parts.extend(elem.headers)
        for row in elem.rows:
            parts.extend(row)
        return " ".join(parts)
    if isinstance(elem, FormulaBlock):
        return elem.source or ""
    if isinstance(elem, MathBlock):
        return elem.latex or ""
    if isinstance(elem, CodeBlock):
        return elem.code
    if isinstance(elem, ImageBlock):
        return elem.alt
    if isinstance(elem, ExcalidrawBlock):
        return elem.alt
    if isinstance(elem, DegradationBlock):
        return elem.fallback_text
    if isinstance(elem, ReferenceListBlock):
        return " ".join(elem.entries)
    if isinstance(elem, TaskList):
        return " ".join(text for text, _ in elem.items)
    if isinstance(elem, TableBlock):
        parts = list(elem.headers)
        for row in elem.rows:
            parts.extend(row)
        return " ".join(parts)
    if isinstance(elem, HorizontalRule):
        return ""
    if isinstance(elem, PageBreakBlock):
        return " ".join(p.plain_text for p in elem.content)
    return ""


def _paragraphs_from_element(elem: Any) -> list[Paragraph]:
    """Convert a block element into one or more readable paragraphs."""
    if isinstance(elem, Paragraph):
        return [elem]
    if isinstance(elem, ListBlock):
        return [
            Paragraph(spans=list(item))
            for item in elem.items
            if "".join(s.text for s in item).strip()
        ]
    if isinstance(elem, Section):
        result: list[Paragraph] = []
        if elem.heading:
            result.append(Paragraph.from_text(elem.heading))
        for child in elem.elements:
            result.extend(_paragraphs_from_element(child))
        return result
    plain = _plain_text_from_element(elem).strip()
    if plain:
        return [Paragraph.from_text(plain)]
    return []


def _normalize_abstract(
    abstract: AbstractBlock, issues: list[DocumentIssue]
) -> AbstractBlock:
    """Keep allowed abstract children and degrade disallowed ones in place."""
    normalized: list[Paragraph] = []
    seen_disallowed = False
    paragraph_ordinal = 0

    def iter_projection_slots(element: Any):
        if isinstance(element, (Paragraph, ListBlock)):
            yield from _iter_paragraph_slots(element)
        elif isinstance(element, Section):
            if element.heading:
                yield _ParagraphSlot(Paragraph(spans=[Span(
                    text=element.heading,
                    semantic_literal=True,
                )]))
            for child in element.elements:
                yield from iter_projection_slots(child)
        elif isinstance(element, (AbstractBlock, BlockQuote)):
            for child in element.paragraphs:
                yield from iter_projection_slots(child)
        elif isinstance(element, PageBreakBlock):
            for child in element.content:
                yield from iter_projection_slots(child)
        else:
            plain = _plain_text_from_element(element).strip()
            if plain:
                yield _ParagraphSlot(Paragraph(spans=[Span(
                    text=plain,
                    code=isinstance(element, CodeBlock),
                    math=(plain if isinstance(element, MathBlock) else ""),
                    semantic_literal=True,
                )]))

    def append_slot(slot: _ParagraphSlot) -> None:
        nonlocal paragraph_ordinal
        paragraph_ordinal += 1
        node_id = f"__wpsc_para:0:{paragraph_ordinal}"
        slot.assign_node_id(node_id)
        if isinstance(slot.owner, Paragraph):
            normalized.append(slot.owner)
        elif "".join(span.text for span in slot.spans).strip():
            normalized.append(
                Paragraph(spans=list(slot.spans), node_id=node_id)
            )

    for elem in abstract.raw_elements:
        if not isinstance(elem, (Paragraph, ListBlock)):
            seen_disallowed = True
        for slot in iter_projection_slots(elem):
            append_slot(slot)
    if seen_disallowed:
        issues.append(
            _issue(
                ABSTRACT_CONTENT_DEGRADED,
                "Abstract contains disallowed block elements; "
                "content was retained as plain text.",
                placement="block",
            )
        )
    return AbstractBlock(
        paragraphs=normalized,
        raw_elements=list(abstract.raw_elements),
    )


def _scan_page_breaks(
    sections: list[Section], issues: list[DocumentIssue]
) -> None:
    """Emit block-level issues for non-empty page-break blocks."""
    for section in sections:
        for elem in section.elements:
            if isinstance(elem, PageBreakBlock) and elem.content:
                issues.append(
                    _issue(
                        PAGE_BREAK_CONTENT_DEGRADED,
                        "Non-empty page-break block; "
                        "content was retained after the page break.",
                        placement="block",
                    )
                )


def _apply_page_role_metadata(
    doc: StructuredDocument, config: LongformConfig, issues: list[DocumentIssue]
) -> None:
    """Attach page roles to the document and each content section."""
    for section in doc.sections:
        section.outline_level = section.level
        if section.page_role != "bibliography":
            section.page_role = (
                "landscape"
                if getattr(section, "orientation", None) == "landscape"
                else "body"
            )

    try:
        from .policy import build_policy
        from .page_policy import build_page_policy

        policy = build_policy(config)
        skeleton = build_page_policy(doc, config, policy)
        doc.page_roles = [s.role for s in skeleton.sections]
    except Exception as exc:
        doc.page_roles = []
        issues.append(
            _issue(
                PAGE_ROLE_RESOLUTION_FAILED,
                f"Page role resolution failed: {exc}; using empty skeleton.",
            )
        )


def _set_title_display(doc: StructuredDocument, config: LongformConfig) -> None:
    """Preserve the document title as a non-cover title paragraph."""
    if not config.title_page and config.title:
        doc.title_display = Paragraph.from_text(config.title)


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
        _set_title_display(doc, config)

        selected_scheme = _apply_heading_numbering(
            doc.sections, config.heading_numbering, config.title, issues
        )
        if config.heading_numbering == "auto":
            config.heading_numbering = selected_scheme

        _derive_header(config, issues)

        targets, _explicit_ids = _collect_explicit_targets(doc, issues)
        _normalize_formulas(doc, issues)
        reference_targets = _caption_reference_targets(doc, targets, issues)
        bookmarks = _map_bookmarks_from_targets(reference_targets)
        for bookmark_issue in bookmarks.issues:
            semantic_issue = (
                BOOKMARK_NAME_COLLISION
                if bookmark_issue == BOOKMARK_COLLISION_UNRESOLVED
                else bookmark_issue
            )
            issues.append(
                _issue(
                    semantic_issue,
                    "A deterministic bookmark name could not be allocated; "
                    "cross-reference capability was disabled for that target.",
                    placement="block",
                )
            )
        _apply_caption_bindings(doc, config, reference_targets, bookmarks)
        if doc.abstract is not None:
            doc.abstract = _normalize_abstract(doc.abstract, issues)
        _assign_paragraph_ids(doc)
        references = _build_references(
            doc, targets, bookmarks, issues, config
        )
        _split_cross_reference_spans(doc, references)
        _resolve_table_citations(doc, references)
        _scan_page_breaks(doc.sections, issues)
        _move_bibliography_to_back_matter(doc)
        _apply_page_role_metadata(doc, config, issues)

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
    "ABSTRACT_CONTENT_DEGRADED",
    "BIBLIOGRAPHY_ENTRY_MALFORMED",
    "CAPTION_MISSING",
    "CONFIG_VALUE_INVALID",
    "DUPLICATE_EXPLICIT_ID",
    "HEADER_SHORTENED",
    "HEADING_LEVEL_GAP",
    "HEADING_PREFIX_AMBIGUOUS",
    "INVALID_EXPLICIT_ID",
    "LongformConfig",
    "PAGE_BREAK_CONTENT_DEGRADED",
    "PAGE_ROLE_RESOLUTION_FAILED",
    "REFERENCE_UNRESOLVED",
    "SemanticResult",
    "TITLE_MISSING",
    "normalize_longform_document",
]
