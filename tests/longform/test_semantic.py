from __future__ import annotations

import json
import unicodedata

import pytest

from skills.WPSComposer.scripts.document_model import (
    AbstractBlock,
    DegradationBlock,
    DocumentIssue,
    FigureBlock,
    FormulaBlock,
    KeywordsBlock,
    PageBreakBlock,
    Paragraph,
    ReferenceListBlock,
    Section,
    SemanticTableBlock,
    Span,
    StructuredDocument,
)
from skills.WPSComposer.scripts.md_parser import parse_markdown
from skills.WPSComposer.scripts.longform.semantic import (
    CONFIG_VALUE_INVALID,
    HEADER_SHORTENED,
    HEADING_LEVEL_GAP,
    HEADING_PREFIX_AMBIGUOUS,
    LongformConfig,
    REFERENCE_UNRESOLVED,
    SemanticResult,
    normalize_longform_document,
)


def _doc_from_markdown(md: str) -> StructuredDocument:
    return parse_markdown(md, longform=True)


# ---------------------------------------------------------------------------
# Title consumption
# ---------------------------------------------------------------------------

def test_frontmatter_title_is_consumed_and_first_h1_preserved_when_different() -> None:
    doc = _doc_from_markdown("---\ntitle: FM Title\n---\n# H1 Title\n\nBody.\n")
    result = normalize_longform_document(doc)

    assert result.config.title == "FM Title"
    assert result.config.title_page is False  # no author/date, auto -> False
    assert result.document.title == "FM Title"
    assert result.document.sections[0].heading == "H1 Title"
    assert result.document.sections[0].level == 1


def test_first_h1_is_consumed_as_title_when_no_frontmatter_title() -> None:
    doc = _doc_from_markdown("# H1 Title\n\nBody.\n")
    result = normalize_longform_document(doc)

    assert result.config.title == "H1 Title"
    assert result.document.title == "H1 Title"
    # The consumed H1 becomes an unheaded section so the title appears exactly once.
    assert result.document.sections[0].level == 0
    assert result.document.sections[0].heading == ""


def test_identical_frontmatter_title_and_first_h1_consumes_h1() -> None:
    doc = _doc_from_markdown('---\ntitle: "Same Title"\n---\n# Same Title\n\nBody.\n')
    result = normalize_longform_document(doc)

    assert result.config.title == "Same Title"
    # The duplicate H1 is consumed as the document title display and removed
    # from the chapter stream.
    headings = [s.heading for s in result.document.sections if s.has_heading]
    assert headings == []


# ---------------------------------------------------------------------------
# Explicit and automatic defaults
# ---------------------------------------------------------------------------

def test_explicit_config_values_override_defaults() -> None:
    md = """---
title: Report
author: Zhang
short_title: Shorty
date: 2026-08-21
title_page: true
toc: true
figure_index: true
table_index: false
bibliography_include_uncited: false
heading_numbering: decimal
caption_numbering: chapter
layout_engine: longform
---
# Intro

Text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    cfg = result.config

    assert cfg.title == "Report"
    assert cfg.short_title == "Shorty"
    assert cfg.author == "Zhang"
    assert cfg.date == "2026-08-21"
    assert cfg.title_page is True
    assert cfg.toc is True
    assert cfg.figure_index is True
    assert cfg.table_index is False
    assert cfg.bibliography_include_uncited is False
    assert cfg.heading_numbering == "decimal"
    assert cfg.caption_numbering == "chapter"
    assert cfg.layout_engine == "longform"


def test_auto_title_page_true_when_title_author_or_date_present() -> None:
    doc = _doc_from_markdown("---\ntitle: T\nauthor: A\n---\n# Intro\n")
    result = normalize_longform_document(doc)
    assert result.config.title_page is True


def test_auto_title_page_false_when_title_only() -> None:
    doc = _doc_from_markdown("---\ntitle: T\n---\n# Intro\n")
    result = normalize_longform_document(doc)
    assert result.config.title_page is False


def test_auto_toc_false_for_short_document() -> None:
    doc = _doc_from_markdown("# A\n\ntext.\n# B\n\ntext.\n# C\n\ntext.\n")
    result = normalize_longform_document(doc)
    assert result.config.toc is False


def test_auto_toc_true_for_long_document_with_three_headings() -> None:
    body = "# A\n\n" + ("x" * 3000) + "\n# B\n\ntext.\n# C\n\ntext.\n"
    doc = _doc_from_markdown(body)
    result = normalize_longform_document(doc)
    assert result.config.toc is True


def test_caption_numbering_auto_selects_chapter_after_numbered_h1() -> None:
    md = """---
title: Report
caption_numbering: auto
---
# 1 Chapter

:::figure {#fig:a caption="A figure"}
![a](a.png)
:::
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.config.caption_numbering == "chapter"


def test_caption_numbering_auto_selects_global_without_numbered_h1() -> None:
    md = """---
caption_numbering: auto
heading_numbering: none
---
# Preface

:::figure {#fig:a caption="A figure"}
![a](a.png)
:::
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.config.caption_numbering == "global"


# ---------------------------------------------------------------------------
# Heading schemes
# ---------------------------------------------------------------------------

def test_auto_detects_chinese_formal_when_majority_han() -> None:
    md = """# 引言

# 方法

# 结果
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.config.heading_numbering == "chinese-formal"


def test_auto_detects_decimal_when_majority_non_han() -> None:
    md = """# Introduction

# Method

# Result
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.config.heading_numbering == "decimal"


def test_explicit_scheme_removes_matching_prefixes() -> None:
    md = """---
title: Report
heading_numbering: decimal
---
# 1 Introduction

## 1.1 Background

### 1.1.1 Detail
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    sections = [s for s in result.document.sections if s.has_heading]
    assert sections[0].heading == "Introduction"
    assert sections[1].heading == "Background"
    assert sections[2].heading == "Detail"
    assert sections[0].numbering == "decimal"
    assert sections[1].numbering == "decimal"
    assert sections[2].numbering == "decimal"


def test_explicit_scheme_preserves_mismatched_prefixes_as_text() -> None:
    md = """---
title: Report
heading_numbering: decimal
---
# 第一章 引言

Text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    section = result.document.sections[0]
    assert section.heading == "第一章 引言"
    assert section.numbering == "none"
    codes = [i.code for i in result.issues]
    assert HEADING_PREFIX_AMBIGUOUS in codes


# ---------------------------------------------------------------------------
# Level gaps
# ---------------------------------------------------------------------------

def test_heading_level_gap_marks_child_none_and_emits_issue() -> None:
    md = """---
title: Report
heading_numbering: decimal
---
# 1 Chapter

### 1.1.1 Too deep
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    sections = [s for s in result.document.sections if s.has_heading]
    assert sections[0].level == 1
    assert sections[0].numbering == "decimal"
    assert sections[1].level == 3
    assert sections[1].numbering == "none"
    assert any(i.code == HEADING_LEVEL_GAP for i in result.issues)


def test_heading_numbering_none_skips_gap_checks() -> None:
    md = """---
heading_numbering: none
---
# Chapter

### Deep
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    sections = [s for s in result.document.sections if s.has_heading]
    assert all(s.numbering == "none" for s in sections)
    assert not any(i.code == HEADING_LEVEL_GAP for i in result.issues)


# ---------------------------------------------------------------------------
# Header shortening
# ---------------------------------------------------------------------------

def test_header_derived_from_short_title_then_title() -> None:
    doc = _doc_from_markdown("---\ntitle: Long Document Title\nshort_title: Short Header\n---\n# Intro\n")
    result = normalize_longform_document(doc)
    assert result.config.header == "Short Header"


def test_header_shortened_when_longer_than_64_display_units() -> None:
    long_header = "中" * 50  # 100 display units
    md = f"---\nheader: {long_header}\n---\n# Intro\n"
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.config.header != long_header
    from skills.WPSComposer.scripts.longform.unicode_text import display_units
    assert display_units(result.config.header) <= 64
    assert result.config.header.endswith("…")
    assert any(i.code == HEADER_SHORTENED for i in result.issues)


# ---------------------------------------------------------------------------
# CONFIG_VALUE_INVALID
# ---------------------------------------------------------------------------

def test_invalid_boolean_config_falls_back_and_emits_config_value_invalid() -> None:
    doc = _doc_from_markdown("---\ntitle_page: maybe\n---\n# Intro\n")
    result = normalize_longform_document(doc)
    assert result.config.title_page is False  # auto default
    assert any(i.code == CONFIG_VALUE_INVALID for i in result.issues)




def test_missing_title_page_and_toc_use_auto_without_config_value_invalid() -> None:
    md = """---
title: Report
---
# Intro

Body text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.config.title_page is False
    assert result.config.toc is False
    assert not any(i.code == CONFIG_VALUE_INVALID for i in result.issues)

def test_invalid_heading_numbering_enum_falls_back_to_auto() -> None:
    doc = _doc_from_markdown("---\nheading_numbering: fancy\n---\n# Intro\n")
    result = normalize_longform_document(doc)
    assert result.config.heading_numbering in {"chinese-formal", "decimal", "none"}
    assert any(i.code == CONFIG_VALUE_INVALID for i in result.issues)


# ---------------------------------------------------------------------------
# Reference collection and unresolved references
# ---------------------------------------------------------------------------

def test_collects_figure_table_formula_ids_and_maps_bookmarks() -> None:
    md = """# Body

:::figure {#fig:diagram caption="Diagram"}
![d](d.png)
:::

:::table {#tab:summary caption="Summary"}
| A |
|---|
| 1 |
:::

:::formula {#eq:euler}
e^{i\\pi} + 1 = 0
:::
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert "fig:diagram" in result.references
    assert "tab:summary" in result.references
    assert "eq:euler" in result.references
    assert result.references["fig:diagram"]["kind"] == "fig"
    assert result.references["tab:summary"]["kind"] == "tab"
    assert result.references["eq:euler"]["kind"] == "eq"
    # Bookmarks are mapped for collected explicit IDs.
    assert result.bookmarks.mapping["fig:diagram"].startswith("wpsc_fig_")
    assert result.bookmarks.mapping["tab:summary"].startswith("wpsc_tab_")
    assert result.bookmarks.mapping["eq:euler"].startswith("wpsc_eq_")


def test_duplicate_explicit_id_loses_reference_capability() -> None:
    md = """# Body

:::figure {#fig:dup caption="First"}
![a](a.png)
:::

:::figure {#fig:dup caption="Second"}
![b](b.png)
:::
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    refs = result.references["fig:dup"]
    blocks = [
        e for s in result.document.sections for e in s.elements if isinstance(e, FigureBlock)
    ]
    # Only the first figure is the reference target; the duplicate loses capability.
    assert blocks[0].node_id == refs["node_id"]
    assert blocks[1].node_id != refs["node_id"]
    assert blocks[1].node_id.startswith("__wpsc_fig")


def test_unresolved_ref_emits_issue() -> None:
    md = """# Body

See {{ref:missing}} for details.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    unresolved = [i for i in result.issues if i.code == REFERENCE_UNRESOLVED]
    assert len(unresolved) == 1


def test_bibliography_entries_are_collected_and_cited() -> None:
    md = """# Body

Claim {{cite:chen2025}}.

:::references
- id: chen2025
  text: 陈延超. 数字技术[J]. 2025.
- id: smith2024
  text: Smith. Example[J]. 2024.
:::
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert "chen2025" in result.references
    assert "smith2024" in result.references
    assert result.references["chen2025"]["kind"] == "ref"
    assert result.references["chen2025"]["cited"] is True
    assert result.references["smith2024"]["cited"] is False


# ---------------------------------------------------------------------------
# Byte-stable snapshots
# ---------------------------------------------------------------------------

def test_to_json_is_byte_stable_across_two_calls() -> None:
    doc = _doc_from_markdown("# Title\n\nParagraph.\n")
    result1 = normalize_longform_document(doc)
    result2 = normalize_longform_document(doc)
    snapshot1 = json.dumps(result1.to_json(), ensure_ascii=False, separators=(",", ":"))
    snapshot2 = json.dumps(result2.to_json(), ensure_ascii=False, separators=(",", ":"))
    assert snapshot1 == snapshot2
    assert isinstance(snapshot1.encode("utf-8"), bytes)


def test_to_json_key_order_is_deterministic() -> None:
    doc = _doc_from_markdown("---\ntitle: T\nauthor: A\ntoc: true\n---\n# H1\n")
    result = normalize_longform_document(doc)
    snapshot = json.dumps(result.to_json(), ensure_ascii=False, separators=(",", ":"))
    keys = json.loads(snapshot)["config"].keys()
    assert list(keys) == sorted(keys)


def test_to_json_reference_mapping_is_sorted_by_key() -> None:
    md = """# Body

:::figure {#fig:zebra caption="Z"}
![z](z.png)
:::

:::figure {#fig:alpha caption="A"}
![a](a.png)
:::
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    refs = result.to_json()["references"]
    assert list(refs.keys()) == sorted(refs.keys())

