"""Lightweight Markdown parser — zero external dependencies.

Parses CommonMark-flavoured Markdown into a StructuredDocument.
Supports Chinese content and YAML frontmatter.
"""

from __future__ import annotations

import re
import os
from typing import List, Optional
from urllib.parse import unquote, urlparse
from urllib.request import url2pathname

from .document_model import (
    StructuredDocument, Section, Span, Paragraph,
    ListBlock, TableBlock, CodeBlock, ImageBlock, BlockQuote,
    HorizontalRule, TaskList,
)


# ---------------------------------------------------------------------------
# Inline parsing
# ---------------------------------------------------------------------------

_INLINE_CODE_RE = re.compile(r"`([^`]+)`")
_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
_ITALIC_RE = re.compile(r"\*(.+?)\*")
_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")
_IMAGE_RE = re.compile(
    r"!\[([^\]]*)\]\(\s*(?:<([^>]+)>|([^\s)]+))"
    r"(?:\s+(?:\"[^\"]*\"|'[^']*'))?\s*\)"
)

# Extended syntax
_HR_RE = re.compile(r"^(?:-{3,}|\*{3,}|_{3,})\s*$")
_TASK_RE = re.compile(r"^[-*]\s+\[([ xX])\]\s+(.+)$")
_STRIKE_RE = re.compile(r"~~(.+?)~~")
_LINK_TITLE_RE = re.compile(r'\[([^\]]+)\]\(([^)\s]+)\s+"([^"]+)"\)')
_NESTED_UL_RE = re.compile(r"^(\s+)[-*+]\s+(.+)$")
_NESTED_OL_RE = re.compile(r"^(\s+)\d+\.\s+(.+)$")


def _parse_inline(text: str) -> List[Span]:
    """Parse inline formatting: **bold**, *italic*, `code`, [link](url).

    Returns a list of Span objects preserving order.
    """
    # First, protect code spans
    codes: List[tuple] = []
    for m in _INLINE_CODE_RE.finditer(text):
        codes.append((m.start(), m.end(), m.group(1)))

    if not codes and not _BOLD_RE.search(text) and not _ITALIC_RE.search(text) and not _LINK_RE.search(text):
        return [Span(text=text)] if text else []

    # Simple approach: tokenize by priority
    spans: List[Span] = []
    remaining = text

    while remaining:
        # Find earliest match
        matches = []
        for m in _INLINE_CODE_RE.finditer(remaining):
            matches.append((m.start(), m.end(), "code", m.group(1)))
        for m in _BOLD_RE.finditer(remaining):
            matches.append((m.start(), m.end(), "bold", m.group(1)))
        for m in _ITALIC_RE.finditer(remaining):
            matches.append((m.start(), m.end(), "italic", m.group(1)))
        for m in _STRIKE_RE.finditer(remaining):
            matches.append((m.start(), m.end(), "strike", m.group(1)))
        # Link with title (tooltip) takes priority
        for m in _LINK_TITLE_RE.finditer(remaining):
            title = m.group(3) or m.group(4) or ""
            matches.append((m.start(), m.end(), "link", m.group(1), m.group(2), title))
        for m in _LINK_RE.finditer(remaining):
            matches.append((m.start(), m.end(), "link", m.group(1), m.group(2)))

        if not matches:
            if remaining:
                spans.append(Span(text=remaining))
            break

        matches.sort(key=lambda x: x[0])
        first = matches[0]

        # Text before match
        if first[0] > 0:
            spans.append(Span(text=remaining[:first[0]]))

        # The matched span
        mtype = first[2]
        if mtype == "code":
            spans.append(Span(text=first[3], code=True))
        elif mtype == "bold":
            spans.append(Span(text=first[3], bold=True))
        elif mtype == "italic":
            spans.append(Span(text=first[3], italic=True))
        elif mtype == "strike":
            spans.append(Span(text=first[3], strikethrough=True))
        elif mtype == "link":
            link_url = first[4]
            link_title = first[5] if len(first) > 5 else None
            spans.append(Span(text=first[3], link=link_url, link_title=link_title))

        remaining = remaining[first[1]:]

    # Merge adjacent plain-text spans
    merged = []
    for s in spans:
        if merged and not s.bold and not s.italic and not s.code and not s.link and not s.strikethrough \
           and not merged[-1].bold and not merged[-1].italic and not merged[-1].code and not merged[-1].link and not merged[-1].strikethrough:
            merged[-1].text += s.text
        else:
            merged.append(s)
    return merged


# ---------------------------------------------------------------------------
# Block-level parsing
# ---------------------------------------------------------------------------

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.+)$")
_UL_RE = re.compile(r"^[-*]\s+(.+)$")
_OL_RE = re.compile(r"^\d+\.\s+(.+)$")
_TABLE_SEP_RE = re.compile(r"^\|?[\s\-:]+\|[\s\-:|]+\|?$")


def _is_table_row(line: str) -> bool:
    return line.strip().startswith("|") and "|" in line[1:]


def _parse_table_row(line: str) -> List[str]:
    # Table cells are currently rendered as plain strings.  Preserve the
    # visible text while removing Markdown control markers such as ``**``
    # and backticks so they never leak into formal Writer output.
    return [
        "".join(span.text for span in _parse_inline(cell.strip()))
        for cell in line.strip().strip("|").split("|")
    ]


def _is_separator_row(line: str) -> bool:
    return bool(_TABLE_SEP_RE.match(line.strip()))


def _image_match_path(match: re.Match) -> str:
    return (match.group(2) or match.group(3) or "").strip()


def _resolve_image_path(path: str, base_dir: str = "") -> str:
    """Resolve a Markdown image target to a renderer-ready path."""
    path = unquote(path.strip())
    parsed = urlparse(path)
    if parsed.scheme in ("http", "https", "data"):
        return path
    if parsed.scheme == "file":
        return os.path.abspath(url2pathname(parsed.path))
    if os.path.isabs(path):
        return os.path.normpath(path)
    return os.path.abspath(os.path.join(base_dir or os.curdir, path))


def _extract_image_refs(text: str, base_dir: str = "") -> List[ImageBlock]:
    """Extract inline image references from text."""
    images = []
    for m in _IMAGE_RE.finditer(text):
        path = _resolve_image_path(_image_match_path(m), base_dir)
        images.append(ImageBlock(path=path, alt=m.group(1)))
    return images


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------


def parse(md_text: str, base_dir: str = "") -> StructuredDocument:
    """Parse Markdown text into a StructuredDocument.

    Args:
        md_text: Raw Markdown string.
        base_dir: Directory for resolving relative image paths.

    Returns:
        A StructuredDocument ready for rendering.
    """
    lines = md_text.split("\n")
    doc = StructuredDocument()

    # --- YAML frontmatter ---
    if lines and lines[0].strip() == "---":
        end_idx = 1
        while end_idx < len(lines) and lines[end_idx].strip() != "---":
            end_idx += 1
        if end_idx < len(lines):
            fm_lines = lines[1:end_idx]
            _parse_frontmatter(fm_lines, doc)
            lines = lines[end_idx + 1:]

    # Auto-detect title from first H1
    if not doc.title:
        for line in lines:
            m = _HEADING_RE.match(line)
            if m and m.group(1) == "#":
                doc.title = m.group(2).strip()
                break

    # --- Block-level parsing ---
    sections: List[Section] = []
    current_section: Optional[Section] = None
    i = 0

    while i < len(lines):
        line = lines[i]

        # Skip empty lines
        if not line.strip():
            i += 1
            continue

        # Horizontal rule (---, ***, ___)
        if _HR_RE.match(line.strip()):
            _ensure_section(current_section, sections)
            current_section.elements.append(HorizontalRule())
            i += 1
            continue

        # Task list (- [ ] / - [x])
        tm = _TASK_RE.match(line)
        if tm:
            items = []
            while i < len(lines):
                tm2 = _TASK_RE.match(lines[i])
                if tm2:
                    checked = tm2.group(1).lower() == "x"
                    items.append((tm2.group(2).strip(), checked))
                    i += 1
                elif not lines[i].strip():
                    i += 1
                else:
                    break
            _ensure_section(current_section, sections)
            current_section.elements.append(TaskList(items=items))
            continue

        # Heading
        hm = _HEADING_RE.match(line)
        if hm:
            level = len(hm.group(1))
            heading = hm.group(2).strip()
            current_section = Section(level=level, heading=heading)
            sections.append(current_section)
            i += 1
            continue

        # Standalone image.  Resolve relative paths while the Markdown file's
        # directory is still available; renderers should not depend on cwd.
        image_match = _IMAGE_RE.fullmatch(line.strip())
        if image_match:
            _ensure_section(current_section, sections)
            current_section.elements.append(ImageBlock(
                path=_resolve_image_path(
                    _image_match_path(image_match), base_dir
                ),
                alt=image_match.group(1).strip(),
            ))
            i += 1
            continue

        # Fenced code block
        if line.strip().startswith("```"):
            language = line.strip()[3:].strip()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code_lines.append(lines[i])
                i += 1
            i += 1  # skip closing ```
            _ensure_section(current_section, sections)
            current_section.elements.append(
                CodeBlock(code="\n".join(code_lines), language=language)
            )
            continue

        # Blockquote
        if line.startswith(">"):
            quote_lines = []
            while i < len(lines) and lines[i].startswith(">"):
                quote_lines.append(lines[i][1:].strip())
                i += 1
            _ensure_section(current_section, sections)
            paras = [Paragraph(spans=_parse_inline(ql)) for ql in quote_lines if ql]
            current_section.elements.append(BlockQuote(paragraphs=paras))
            continue

        # Table
        if _is_table_row(line):
            table_lines = []
            while i < len(lines) and _is_table_row(lines[i]):
                table_lines.append(lines[i])
                i += 1
            _parse_table_block(table_lines, current_section, sections)
            continue

        # Unordered list
        ulm = _UL_RE.match(line)
        if ulm:
            items = []
            while i < len(lines):
                ulm2 = _UL_RE.match(lines[i])
                if ulm2:
                    items.append(_parse_inline(ulm2.group(1).strip()))
                    i += 1
                elif not lines[i].strip():
                    i += 1
                else:
                    break
            _ensure_section(current_section, sections)
            current_section.elements.append(ListBlock(items=items, ordered=False))
            continue

        # Ordered list
        olm = _OL_RE.match(line)
        if olm:
            items = []
            while i < len(lines):
                olm2 = _OL_RE.match(lines[i])
                if olm2:
                    items.append(_parse_inline(olm2.group(1).strip()))
                    i += 1
                elif not lines[i].strip():
                    i += 1
                else:
                    break
            _ensure_section(current_section, sections)
            current_section.elements.append(ListBlock(items=items, ordered=True))
            continue

        # Regular paragraph
        para_lines = [line]
        i += 1
        while i < len(lines) and lines[i].strip() and not _is_block_start(lines[i]):
            para_lines.append(lines[i])
            i += 1
        para_text = " ".join(pl.strip() for pl in para_lines)
        spans = _parse_inline(para_text)
        _ensure_section(current_section, sections)
        current_section.elements.append(Paragraph(spans=spans))

    # Build document
    if not doc.title and sections and sections[0].has_heading:
        doc.title = sections[0].heading

    doc.sections = sections
    return doc


def _ensure_section(current: Optional[Section], sections: List[Section]):
    """Ensure there's a current section; create one if not."""
    if current is None:
        current = Section(level=0, heading="")
        sections.append(current)


def _detect_alignments(separator_line: str) -> list:
    cells = [c.strip() for c in separator_line.strip().strip("|").split("|")]
    alignments = []
    for cell in cells:
        left = cell.startswith(":")
        right = cell.endswith(":")
        if left and right:
            alignments.append("center")
        elif right:
            alignments.append("right")
        else:
            alignments.append("left")
    return alignments
def _is_separator_cell(cell: str) -> bool:
    return bool(re.match(r"^[\-\s:]+\+?$", cell))


_BLOCK_STARTS = {"#", "-", "*", ">", "|", "```", "---", "***", "___", "1.", "2.", "3.", "4.", "5.", "6.", "7.", "8.", "9."}


def _is_block_start(line: str) -> bool:
    s = line.strip()
    if not s:
        return False
    if _IMAGE_RE.fullmatch(s):
        return True
    for prefix in _BLOCK_STARTS:
        if s.startswith(prefix):
            return True
    return _OL_RE.match(s) is not None



def _parse_table_block(table_lines: List[str], current: Optional[Section], sections: List[Section]):
    if len(table_lines) < 2:
        return
    rows = [_parse_table_row(tl) for tl in table_lines]
    sep_idx = None
    alignments = []
    for idx, row in enumerate(rows):
        if all(_is_separator_cell(c) for c in row):
            sep_idx = idx
            alignments = _detect_alignments(table_lines[idx])
            break
    if sep_idx is not None and sep_idx > 0:
        headers = rows[sep_idx - 1]
        data = rows[sep_idx + 1:]
    else:
        headers = rows[0] if rows else []
        data = rows[1:] if len(rows) > 1 else []
    _ensure_section(current, sections)
    current.elements.append(TableBlock(headers=headers, rows=data, alignments=alignments))
def _parse_frontmatter(lines: List[str], doc: StructuredDocument):
    """Parse YAML frontmatter lines into doc.metadata."""
    for line in lines:
        line = line.strip()
        if ":" in line:
            key, _, value = line.partition(":")
            key = key.strip()
            value = value.strip().strip("'").strip('"')
            if key.lower() == "title":
                doc.title = value
            else:
                doc.metadata[key] = value


# ---------------------------------------------------------------------------
# Convenience
# ---------------------------------------------------------------------------


def parse_file(md_path: str) -> StructuredDocument:
    """Parse a Markdown file into a StructuredDocument.

    Args:
        md_path: Path to the .md file.

    Returns:
        StructuredDocument.
    """
    base_dir = os.path.dirname(os.path.abspath(md_path))
    with open(md_path, "r", encoding="utf-8") as f:
        text = f.read()
    return parse(text, base_dir=base_dir)
