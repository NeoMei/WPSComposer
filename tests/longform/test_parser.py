from __future__ import annotations

import pytest
import dataclasses

from skills.WPSComposer.scripts.document_model import (
    AbstractBlock,
    DegradationBlock,
    DocumentIssue,
    FigureBlock,
    FormulaBlock,
    KeywordsBlock,
    PageBreakBlock,
    ReferenceListBlock,
    SemanticTableBlock,
    StructuredDocument,
)
from skills.WPSComposer.scripts.md_parser import parse, parse_markdown


# ---------------------------------------------------------------------------
# Frontmatter metadata -> document fields
# ---------------------------------------------------------------------------

def test_frontmatter_title_author_date_keywords_abstract_are_captured() -> None:
    md = """---
title: Longform Report
author: 张三
date: 2026-08-21
keywords:
  - longform
  - wps
abstract: "This is the abstract."
---

# Introduction

Body text.
"""
    doc = parse_markdown(md, longform=True)

    assert doc.title == "Longform Report"
    assert doc.metadata.get("author") == "张三"
    assert doc.metadata.get("date") == "2026-08-21"
    assert isinstance(doc.abstract, AbstractBlock)
    assert doc.abstract.plain_text == "This is the abstract."
    assert isinstance(doc.keywords, KeywordsBlock)
    assert doc.keywords.keywords == ["longform", "wps"]


def test_frontmatter_title_overrides_first_h1() -> None:
    doc = parse_markdown("---\ntitle: FM Title\n---\n# H1 Title\n", longform=True)
    assert doc.title == "FM Title"


def test_longform_uses_first_h1_when_no_frontmatter_title() -> None:
    doc = parse_markdown("# H1 Title\n", longform=True)
    assert doc.title == "H1 Title"


# ---------------------------------------------------------------------------
# Abstract / keywords directives
# ---------------------------------------------------------------------------

def test_abstract_directive_creates_document_abstract_block() -> None:
    md = """# Introduction

:::abstract
First abstract paragraph.

Second paragraph.
:::

Body.
"""
    doc = parse_markdown(md, longform=True)

    assert isinstance(doc.abstract, AbstractBlock)
    assert len(doc.abstract.paragraphs) == 2
    assert doc.abstract.paragraphs[0].plain_text == "First abstract paragraph."
    assert doc.abstract.paragraphs[1].plain_text == "Second paragraph."


def test_keywords_directive_creates_document_keywords_block() -> None:
    md = """# Intro

:::keywords
alpha
beta gamma
delta
:::

Body.
"""
    doc = parse_markdown(md, longform=True)

    assert isinstance(doc.keywords, KeywordsBlock)
    assert doc.keywords.keywords == ["alpha", "beta gamma", "delta"]


def test_abstract_and_keywords_directives_override_frontmatter_values() -> None:
    md = """---
abstract: Frontmatter abstract.
keywords:
  - fm-one
---

:::abstract
Directive abstract.
:::

:::keywords
kw-a
kw-b
:::
"""
    doc = parse_markdown(md, longform=True)

    assert doc.abstract.plain_text == "Directive abstract."
    assert doc.keywords.keywords == ["kw-a", "kw-b"]


def test_duplicate_abstract_or_keywords_directive_emits_degradation_block() -> None:
    md = """# Intro

:::abstract
First.
:::

:::abstract
Second.
:::

:::keywords
a
:::

:::keywords
b
:::
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    degradations = [b for b in blocks if isinstance(b, DegradationBlock)]
    assert len(degradations) >= 2
    codes = {d.issue.code for d in degradations}
    assert "LONGFORM_DUPLICATE_FRONT_BLOCK" in codes


# ---------------------------------------------------------------------------
# Explicit page breaks
# ---------------------------------------------------------------------------

def test_page_break_directive_creates_page_break_block() -> None:
    md = """# Section A

:::page-break
:::

# Section B
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    page_breaks = [b for b in blocks if isinstance(b, PageBreakBlock)]
    assert len(page_breaks) == 1


def test_empty_page_break_directive_is_allowed() -> None:
    doc = parse_markdown("# A\n\n:::page-break\n:::\n# B\n", longform=True)
    blocks = [e for s in doc.sections for e in s.elements]
    assert any(isinstance(b, PageBreakBlock) for b in blocks)


# ---------------------------------------------------------------------------
# Semantic table / figure / formula / references
# ---------------------------------------------------------------------------

def test_table_directive_with_id_and_caption_wraps_markdown_table() -> None:
    md = """# Results

:::table {#tbl:results caption="Summary"}
| Metric | Value |
| --- | --- |
| A | 1 |
| B | 2 |
:::
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    table = next(b for b in blocks if isinstance(b, SemanticTableBlock))
    assert table.identifier == "tbl:results"
    assert table.caption == "Summary"
    assert table.headers == ["Metric", "Value"]
    assert table.rows == [["A", "1"], ["B", "2"]]


def test_figure_directive_captures_image_and_caption() -> None:
    md = """# Figures

:::figure {#fig:diagram caption="System diagram"}
![Diagram](diagram.png)
:::
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    fig = next(b for b in blocks if isinstance(b, FigureBlock))
    assert fig.identifier == "fig:diagram"
    assert fig.caption == "System diagram"
    assert len(fig.images) == 1
    assert fig.images[0].alt == "Diagram"


def test_formula_directive_captures_latex_source_and_identifier() -> None:
    md = """# Math

:::formula {#eq:pythagoras}
E = mc^2
:::
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    formula = next(b for b in blocks if isinstance(b, FormulaBlock))
    assert formula.identifier == "eq:pythagoras"
    assert formula.source == "E = mc^2"


def test_references_directive_collects_entries() -> None:
    md = """# Back matter

:::references
[1] Author, Title.
[2] Another, Book.
:::
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    refs = next(b for b in blocks if isinstance(b, ReferenceListBlock))
    assert refs.entries == ["[1] Author, Title.", "[2] Another, Book."]


# ---------------------------------------------------------------------------
# Invalid / unknown directives -> degradation blocks
# ---------------------------------------------------------------------------

def test_invalid_directive_syntax_emits_degradation_block() -> None:
    md = """# Intro

:::abstract {bad value}
Body.
:::
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    degradation = next(b for b in blocks if isinstance(b, DegradationBlock))
    assert degradation.issue.code == "DIRECTIVE_SYNTAX_INVALID"


def test_unclosed_directive_emits_degradation_block_and_preserves_body() -> None:
    md = """# Intro

:::abstract
This body must remain visible.

# Still body
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    degradation = next(b for b in blocks if isinstance(b, DegradationBlock))
    assert degradation.issue.code == "DIRECTIVE_UNCLOSED"


def test_unknown_directive_emits_degradation_block() -> None:
    md = """# Intro

:::unknown
Body.
:::
"""
    doc = parse_markdown(md, longform=True)

    blocks = [e for s in doc.sections for e in s.elements]
    degradation = next(b for b in blocks if isinstance(b, DegradationBlock))
    assert degradation.issue.code == "LONGFORM_DIRECTIVE_UNKNOWN"


# ---------------------------------------------------------------------------
# Legacy compatibility: longform=False behaves exactly like before
# ---------------------------------------------------------------------------

def test_legacy_parse_ignores_longform_directives() -> None:
    md = """# Intro

:::abstract
Ignored.
:::

Text.
"""
    legacy = parse(md)
    longform = parse_markdown(md, longform=True)

    assert legacy.title == "Intro"
    assert len(legacy.sections) == 1
    assert all(
        not isinstance(e, (AbstractBlock, DegradationBlock))
        for s in legacy.sections
        for e in s.elements
    )




def test_legacy_and_longform_false_produce_identical_document() -> None:
    md = """---
title: Report
author: "Neo Mei"
date: 2026-08-21
---
# Heading 1

Paragraph with **bold** and *italic*.

## Heading 2

- Item one
- Item two

Text with inline math: $x^2 + y^2 = z^2$.

```
print("hello")
```

| A | B |
| --- | --- |
| 1 | 2 |

> A blockquote line.

![Diagram](diagram.png)

![[wikilink.png|300x200]]

See also [[Page Link]].
"""
    legacy = parse(md)
    longform_false = parse_markdown(md, longform=False)

    assert dataclasses.asdict(legacy) == dataclasses.asdict(longform_false)


def test_parse_is_exact_alias_for_longform_false() -> None:
    md = """# H1

Paragraph **bold**.
"""
    assert parse(md).title == parse_markdown(md, longform=False).title
    md = """# H1

Paragraph **bold**.
"""
    assert parse(md).title == parse_markdown(md, longform=False).title
