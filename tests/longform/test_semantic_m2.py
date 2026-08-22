from __future__ import annotations

import pytest

from skills.WPSComposer.scripts.document_model import (
    PageBreakBlock,
    Section,
    Span,
    StructuredDocument,
)
from skills.WPSComposer.scripts.md_parser import parse_markdown
from skills.WPSComposer.scripts.longform.semantic import (
    ABSTRACT_CONTENT_DEGRADED,
    HEADING_LEVEL_GAP,
    HEADING_PREFIX_AMBIGUOUS,
    PAGE_BREAK_CONTENT_DEGRADED,
    PAGE_ROLE_RESOLUTION_FAILED,
    TITLE_MISSING,
    normalize_longform_document,
)


def _doc_from_markdown(md: str) -> StructuredDocument:
    return parse_markdown(md, longform=True)


# ---------------------------------------------------------------------------
# Front-matter extraction and title display
# ---------------------------------------------------------------------------


def test_frontmatter_title_abstract_keywords_are_extracted() -> None:
    md = """---
title: Report
author: Zhang
abstract: This is the abstract.
keywords: [AI, decisions]
toc: true
---
# Intro

Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))

    assert result.document.title == "Report"
    assert result.document.abstract is not None
    assert result.document.abstract.paragraphs[0].plain_text == "This is the abstract."
    assert result.document.keywords is not None
    assert result.document.keywords.keywords == ["AI", "decisions"]
    assert result.config.toc is True


def test_title_page_absence_preserves_non_cover_title() -> None:
    md = """---
title: Report
---
# Intro

Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))

    assert result.config.title_page is False
    assert result.document.title_display is not None
    assert result.document.title_display.plain_text == "Report"


def test_explicit_title_page_true_with_empty_title_emits_title_missing() -> None:
    md = """---
title_page: true
---
Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.config.title_page is False
    assert any(i.code == TITLE_MISSING for i in result.issues)


# ---------------------------------------------------------------------------
# Abstract content degradation
# ---------------------------------------------------------------------------


def test_abstract_with_disallowed_heading_emits_degradation() -> None:
    md = """---
title: Report
---
:::abstract
This is allowed.

# Not allowed

More text.
:::
# Intro

Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.document.abstract is not None
    assert any(i.code == ABSTRACT_CONTENT_DEGRADED for i in result.issues)
    # Readable text is preserved in abstract paragraphs.
    texts = [p.plain_text for p in result.document.abstract.paragraphs]
    assert "This is allowed." in texts
    assert "Not allowed" in " ".join(texts)


def test_abstract_with_list_is_allowed() -> None:
    md = """---
title: Report
---
:::abstract
- First item
- Second item
:::
# Intro

Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.document.abstract is not None
    assert not any(i.code == ABSTRACT_CONTENT_DEGRADED for i in result.issues)
    assert len(result.document.abstract.paragraphs) == 2


def test_abstract_does_not_duplicate_heading_children() -> None:
    md = """---
title: Report
---
:::abstract
This is allowed.

# Not allowed

More text.
:::
# Intro

Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    texts = [p.plain_text for p in result.document.abstract.paragraphs]
    assert texts == ["This is allowed.", "Not allowed", "More text."]

def test_abstract_list_preserves_inline_spans() -> None:
    md = """---
title: Report
---
:::abstract
- **bold** and *italic*
:::
# Intro
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    paras = result.document.abstract.paragraphs
    assert len(paras) == 1
    assert any(s.bold for s in paras[0].spans)
    assert any(s.italic for s in paras[0].spans)


# ---------------------------------------------------------------------------
# Page-break content degradation
# ---------------------------------------------------------------------------


def test_empty_page_break_has_no_degradation() -> None:
    md = """---
title: Report
---
# Intro

Text.

:::page-break
:::

More text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert not any(i.code == PAGE_BREAK_CONTENT_DEGRADED for i in result.issues)


def test_non_empty_page_break_emits_degradation_and_keeps_content() -> None:
    md = """---
title: Report
---
# Intro

Text.

:::page-break
Retained paragraph.
:::

More text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    issues = [i for i in result.issues if i.code == PAGE_BREAK_CONTENT_DEGRADED]
    assert len(issues) == 1
    assert issues[0].placement == "block"
    # Find the page break and verify retained content.
    pb = None
    for sec in result.document.sections:
        for elem in sec.elements:
            if isinstance(elem, PageBreakBlock):
                pb = elem
                break
    assert pb is not None
    assert len(pb.content) == 1
    assert pb.content[0].plain_text == "Retained paragraph."


def test_page_break_retains_heading_as_plain_paragraph() -> None:
    md = """---
title: Report
---
# Intro

:::page-break
### Page heading

Paragraph text.
:::

More text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    pb = None
    for sec in result.document.sections:
        for elem in sec.elements:
            if isinstance(elem, PageBreakBlock):
                pb = elem
                break
    assert pb is not None
    texts = [p.plain_text for p in pb.content]
    assert "Page heading" in texts
    assert "Paragraph text." in texts


def test_page_break_does_not_duplicate_children() -> None:
    md = """---
title: Report
---
# Intro

:::page-break
## Break heading

Break paragraph.
:::

More text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    pb = None
    for sec in result.document.sections:
        for elem in sec.elements:
            if isinstance(elem, PageBreakBlock):
                pb = elem
                break
    assert pb is not None
    assert [p.plain_text for p in pb.content] == [
        "Break heading",
        "Break paragraph.",
    ]
# ---------------------------------------------------------------------------
# Heading outline / numbering consistency
# ---------------------------------------------------------------------------


def test_section_exposes_outline_level_and_numbering_for_h1_h6() -> None:
    md = """---
title: Report
heading_numbering: decimal
---
# 1 Chapter

## 1.1 Section

##### H5

###### H6
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    sections = [s for s in result.document.sections if s.has_heading]
    assert sections[0].outline_level == 1
    assert sections[0].numbering == "decimal"
    assert sections[0].numbering_scheme == "decimal"
    assert sections[1].outline_level == 2
    assert sections[1].numbering == "decimal"
    # H5/H6 are unnumbered but still expose outline level.
    assert sections[2].outline_level == 5
    assert sections[2].numbering == "none"
    assert sections[2].numbering_scheme is None
    assert sections[3].outline_level == 6


def test_heading_level_gap_is_document_level_issue() -> None:
    md = """---
title: Report
heading_numbering: decimal
---
# 1 Chapter

### 1.1.1 Gap
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    gap_issues = [i for i in result.issues if i.code == HEADING_LEVEL_GAP]
    assert len(gap_issues) == 1
    assert gap_issues[0].placement == "document"


def test_heading_prefix_ambiguous_is_document_level_issue() -> None:
    md = """---
title: Report
heading_numbering: decimal
---
# 第一章 Mismatch

Text.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    issues = [i for i in result.issues if i.code == HEADING_PREFIX_AMBIGUOUS]
    assert len(issues) == 1
    assert issues[0].placement == "document"


# ---------------------------------------------------------------------------
# Page role metadata
# ---------------------------------------------------------------------------


def test_document_carries_page_roles_and_sections_have_page_role() -> None:
    md = """---
title: Report
author: Zhang
toc: true
---
# Intro

Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.document.page_roles is not None
    assert "front_matter" in result.document.page_roles
    assert "body" in result.document.page_roles
    for sec in result.document.sections:
        if sec.has_heading:
            assert sec.page_role == "body"
            assert sec.outline_level == sec.level


def test_page_role_fallback_emits_degradation(monkeypatch: pytest.MonkeyPatch) -> None:
    from skills.WPSComposer.scripts.longform import page_policy

    def _broken(*args, **kwargs):
        raise RuntimeError("simulated policy failure")

    monkeypatch.setattr(
        page_policy,
        "build_page_policy",
        _broken,
    )

    md = """---
title: Report
---
# Intro

Body.
"""
    result = normalize_longform_document(_doc_from_markdown(md))
    assert result.document.page_roles == []
    assert any(i.code == PAGE_ROLE_RESOLUTION_FAILED for i in result.issues)
