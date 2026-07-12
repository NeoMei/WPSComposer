"""Structured document model — format-agnostic intermediate representation.

Sits between Markdown parsing and format-specific rendering (Writer, Slide,
Sheet, PDF).  All renderers consume this model, so the MD parser only needs
to produce one output.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any


# ---------------------------------------------------------------------------
# Inline formatting
# ---------------------------------------------------------------------------

@dataclass
class Span:
    """A formatted text span within a paragraph."""
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    strikethrough: bool = False
    link: Optional[str] = None   # URL
    link_title: Optional[str] = None  # tooltip


# ---------------------------------------------------------------------------
# Block-level elements
# ---------------------------------------------------------------------------

@dataclass
class Paragraph:
    """A paragraph with optional inline formatting spans."""
    spans: List[Span] = field(default_factory=list)
    align: int = 0  # 0=left, 1=center, 2=right

    @property
    def plain_text(self) -> str:
        return "".join(s.text for s in self.spans)

    @classmethod
    def from_text(cls, text: str) -> "Paragraph":
        return cls(spans=[Span(text=text)])


@dataclass
class ListBlock:
    """Ordered or unordered list."""
    items: List[List[Span]] = field(default_factory=list)
    ordered: bool = False


@dataclass
class TableBlock:
    """A table with headers and rows."""
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    alignments: List[str] = field(default_factory=list)  # "left"/"center"/"right" per column


@dataclass
class CodeBlock:
    """Fenced code block."""
    code: str = ""
    language: str = ""


@dataclass
class ImageBlock:
    """An embedded image."""
    path: str = ""
    alt: str = ""
    width: Optional[int] = None
    height: Optional[int] = None


@dataclass
class BlockQuote:
    """Blockquote — rendered as indented italic."""
    paragraphs: List[Paragraph] = field(default_factory=list)


# ---------------------------------------------------------------------------
# New block types for extended Markdown support
# ---------------------------------------------------------------------------

@dataclass
class HorizontalRule:
    """A horizontal rule (---, ***, ___)."""
    style: str = "single"   # "single" / "double" / "thick"


@dataclass
class TaskList:
    """A task list with checkboxes.

    items: list of (text, checked) tuples.
    """
    items: List[tuple] = field(default_factory=list)  # [(str, bool), ...]


# ---------------------------------------------------------------------------
# Section — a heading + its content
# ---------------------------------------------------------------------------

@dataclass
class Section:
    """A document section: heading level + title + block-level elements."""
    level: int                          # 1=H1, 2=H2, ..., 0=implied (no heading)
    heading: str = ""
    elements: List[Any] = field(default_factory=list)

    @property
    def has_heading(self) -> bool:
        return self.level > 0 and bool(self.heading)


# ---------------------------------------------------------------------------
# StructuredDocument — top-level container
# ---------------------------------------------------------------------------

@dataclass
class StructuredDocument:
    """Full document: metadata + ordered sections."""
    title: str = ""
    metadata: Dict[str, str] = field(default_factory=dict)
    sections: List[Section] = field(default_factory=list)

    @property
    def all_tables(self) -> List[TableBlock]:
        """Collect all tables across all sections."""
        tables = []
        for sec in self.sections:
            for elem in sec.elements:
                if isinstance(elem, TableBlock):
                    tables.append(elem)
        return tables

    @property
    def all_images(self) -> List[ImageBlock]:
        """Collect all images across all sections."""
        images = []
        for sec in self.sections:
            for elem in sec.elements:
                if isinstance(elem, ImageBlock):
                    images.append(elem)
        return images
