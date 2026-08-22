from __future__ import annotations

import json

from skills.WPSComposer.scripts.document_model import (
    BibliographyEntry,
    BlockQuote,
    CitationRun,
    DegradationBlock,
    FigureBlock,
    FormulaBlock,
    InlineDegradationRun,
    ListBlock,
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
    BIBLIOGRAPHY_ENTRY_MALFORMED,
    DUPLICATE_EXPLICIT_ID,
    REFERENCE_UNRESOLVED,
    normalize_longform_document,
)


def _normalize(markdown: str):
    return normalize_longform_document(parse_markdown(markdown, longform=True))


def _citation_runs(result):
    runs = []

    def spans(items):
        for span in items:
            if span.citation is not None:
                runs.append(span.citation)

    if result.document.abstract:
        for paragraph in result.document.abstract.paragraphs:
            spans(paragraph.spans)
    for section in result.document.sections:
        for element in section.elements:
            if isinstance(element, Paragraph):
                spans(element.spans)
            elif isinstance(element, ListBlock):
                for item in element.items:
                    spans(item)
            elif isinstance(element, BlockQuote):
                for paragraph in element.paragraphs:
                    spans(paragraph.spans)
            elif isinstance(element, PageBreakBlock):
                for paragraph in element.content:
                    spans(paragraph.spans)
    return runs


def _bibliography_entries(result):
    return [
        item
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, ReferenceListBlock)
        for item in element.resolved_items
        if isinstance(item, BibliographyEntry)
    ]


def test_first_occurrence_numbering_covers_all_supported_visible_containers() -> None:
    result = _normalize(
        """:::abstract
Abstract {{cite:a}}.
:::

Body {{cite:b}} and again {{cite:a}}.

- List {{cite:c}}.

> Quote {{cite:d}}.

:::page-break
Break content {{cite:e}}.
:::

:::table {#tab:data caption="Citations"}
| Cell |
|---|
| Table {{cite:f}} / {{cite:a}} |
:::

![literal {{cite:alt}}](alt.png)

`{{cite:code}}` ${{cite:math}}$.

:::bibliography
[f] F.
[e] E.
[d] D.
[c] C.
[b] B.
[a] A.
[alt] Alt.
[code] Code.
[math] Math.
[uncited] U.
:::
"""
    )

    assert [(run.target_id, run.number) for run in _citation_runs(result)] == [
        ("a", 1),
        ("b", 2),
        ("a", 1),
        ("c", 3),
        ("d", 4),
        ("e", 5),
    ]
    table = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, SemanticTableBlock)
    )
    assert table.rows == [["Table [6] / [1]"]]
    assert table.cell_degradations == []
    entries = _bibliography_entries(result)
    assert [(entry.identifier, entry.number, entry.cited) for entry in entries] == [
        ("a", 1, True),
        ("b", 2, True),
        ("c", 3, True),
        ("d", 4, True),
        ("e", 5, True),
        ("f", 6, True),
        ("alt", 7, False),
        ("code", 8, False),
        ("math", 9, False),
        ("uncited", 10, False),
    ]


def test_abstract_and_page_break_block_literals_are_never_citation_runs() -> None:
    result = _normalize(
        """:::abstract
![{{cite:abstract-alt}}](abstract.png)

```text
{{cite:abstract-code}}
```

$$
{{cite:abstract-math}}
$$
:::

:::page-break
![{{cite:break-alt}}](break.png)

```text
{{cite:break-code}}
```

$$
{{cite:break-math}}
$$
:::

:::bibliography
[abstract-alt] Abstract image alt.
[abstract-code] Abstract code.
[abstract-math] Abstract math.
[break-alt] Break image alt.
[break-code] Break code.
[break-math] Break math.
:::
"""
    )
    literal_ids = {
        "abstract-alt",
        "abstract-code",
        "abstract-math",
        "break-alt",
        "break-code",
        "break-math",
    }
    assert _citation_runs(result) == []
    assert all(result.references[identifier]["cited"] is False for identifier in literal_ids)
    paragraphs = list(result.document.abstract.paragraphs)
    page_break = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, PageBreakBlock)
    )
    paragraphs.extend(page_break.content)
    assert all(
        span.citation is None and span.inline_degradation is None
        for paragraph in paragraphs
        for span in paragraph.spans
    )
def test_multiple_citations_split_runs_and_unresolved_stays_same_paragraph() -> None:
    result = _normalize(
        """Before {{cite:a}}, {{cite:missing}}, and {{cite:a}} after.

:::references
[a] Entry A.
:::
"""
    )
    paragraph = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, Paragraph)
    )
    assert [span.text for span in paragraph.spans] == [
        "Before ",
        "[1]",
        ", ",
        "[REFERENCE_UNRESOLVED 引用目标未解析]",
        ", and ",
        "[1]",
        " after.",
    ]
    citations = [span.citation for span in paragraph.spans if span.citation]
    assert all(isinstance(run, CitationRun) for run in citations)
    assert [run.node_id for run in citations] == [
        f"{paragraph.node_id}/cite:1",
        f"{paragraph.node_id}/cite:3",
    ]
    degradation_span = next(span for span in paragraph.spans if span.inline_degradation)
    assert isinstance(degradation_span.inline_degradation, InlineDegradationRun)
    assert degradation_span.inline_degradation.node_id == f"{paragraph.node_id}/cite:2"
    assert degradation_span.inline_degradation.code == REFERENCE_UNRESOLVED
    assert len([issue for issue in result.issues if issue.code == REFERENCE_UNRESOLVED]) == 1


def test_unresolved_table_citation_is_static_and_cell_local() -> None:
    result = _normalize(
        """:::table {#tab:data caption="Data"}
| A | B |
|---|---|
| {{cite:known}} | before {{cite:missing}} after |
:::

:::references
[known] Known.
:::
"""
    )
    table = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, SemanticTableBlock)
    )
    assert table.rows == [["[1]", "before [REFERENCE_UNRESOLVED 引用目标未解析] after"]]
    assert len(table.cell_degradations) == 1
    degradation = table.cell_degradations[0]
    assert (degradation.row, degradation.column) == (2, 2)
    assert degradation.code == REFERENCE_UNRESOLVED
    assert degradation.fallback_text == "[REFERENCE_UNRESOLVED 引用目标未解析]"


def test_bibliography_orders_cited_first_and_can_omit_uncited() -> None:
    result = _normalize(
        """---
bibliography_include_uncited: false
---
{{cite:b}} then {{cite:a}}.

:::bibliography
[a] A.
[b] B.
[c] C.
:::
"""
    )
    entries = _bibliography_entries(result)
    assert [(entry.identifier, entry.number, entry.declaration_index, entry.cited) for entry in entries] == [
        ("b", 1, 2, True),
        ("a", 2, 1, True),
    ]
    assert result.references["c"]["cited"] is False
    assert result.references["c"]["number"] is None


def test_multiple_bibliography_blocks_share_one_global_cited_first_order() -> None:
    result = _normalize(
        """{{cite:c}} then {{cite:a}}.

:::references
[a] A.
[b] B.
:::

:::bibliography
[c] C.
[d] D.
:::
"""
    )
    entries = _bibliography_entries(result)
    assert [(entry.identifier, entry.number, entry.declaration_index) for entry in entries] == [
        ("c", 1, 3),
        ("a", 2, 1),
        ("b", 3, 2),
        ("d", 4, 4),
    ]


def test_cited_first_keeps_malformed_item_in_uncited_declaration_position() -> None:
    result = _normalize(
        """Body cites {{cite:b}} first.

:::bibliography
[a] A.
malformed source remains between a and the remaining declarations
[b] B.
[c] C.
:::
"""
    )
    refs = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, ReferenceListBlock)
    )
    assert [
        item.identifier if isinstance(item, BibliographyEntry) else item.fallback_text
        for item in refs.resolved_items
    ] == [
        "b",
        "a",
        "malformed source remains between a and the remaining declarations",
        "c",
    ]
    assert [
        item.number for item in refs.resolved_items if isinstance(item, BibliographyEntry)
    ] == [1, 2, 3]


def test_malformed_bibliography_source_remains_visible_in_original_item_order() -> None:
    result = _normalize(
        """:::bibliography
[a] A.
this malformed line must remain
[b] B.
:::
"""
    )
    refs = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, ReferenceListBlock)
    )
    assert [type(item) for item in refs.resolved_items] == [
        BibliographyEntry,
        DegradationBlock,
        BibliographyEntry,
    ]
    malformed = refs.resolved_items[1]
    assert malformed.fallback_text == "this malformed line must remain"
    assert malformed.issue.code == BIBLIOGRAPHY_ENTRY_MALFORMED
    assert malformed.issue.placement == "block"


def test_no_reference_document_has_no_bibliography_section_or_issue() -> None:
    result = _normalize("Body only.\n")
    assert not any(section.page_role == "bibliography" for section in result.document.sections)
    assert not any("BIBLIOGRAPHY" in issue.code for issue in result.issues)


def test_global_explicit_id_first_definition_wins_across_object_kinds() -> None:
    result = _normalize(
        """:::bibliography
[shared] First bibliography definition.
:::

:::figure {#shared caption="Figure remains visible"}
![f](f.png)
:::

:::table {#shared caption="Table remains visible"}
| A |
|---|
| 1 |
:::

:::equation {#shared}
x + y
:::
"""
    )
    assert result.references["shared"]["kind"] == "ref"
    objects = [
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, (FigureBlock, SemanticTableBlock, FormulaBlock))
    ]
    assert len(objects) == 3
    assert all(obj.caption_binding is not None and not obj.caption_binding.referenceable for obj in objects)
    assert all(obj.target_degradation is not None for obj in objects)
    assert len([issue for issue in result.issues if issue.code == DUPLICATE_EXPLICIT_ID]) == 3


def test_object_definition_before_bibliography_keeps_target_and_bibliography_text() -> None:
    result = _normalize(
        """:::figure {#shared caption="First target"}
![f](f.png)
:::

:::bibliography
[shared] Later bibliography text remains visible.
:::
"""
    )
    assert result.references["shared"]["kind"] == "fig"
    refs = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, ReferenceListBlock)
    )
    degradation = next(
        item for item in refs.resolved_items if isinstance(item, DegradationBlock)
    )
    assert degradation.fallback_text == "[shared] Later bibliography text remains visible."
    assert degradation.issue.code == DUPLICATE_EXPLICIT_ID


def test_explicit_section_id_participates_in_global_first_definition_wins() -> None:
    section = Section(level=1, heading="Visible", node_id="shared")
    formula = FormulaBlock(identifier="shared", source="x")
    section.elements.append(formula)
    result = normalize_longform_document(StructuredDocument(sections=[section], longform=True))

    assert result.references["shared"]["kind"] == "sec"
    assert formula.identifier == "shared"
    normalized_formula = next(
        element
        for candidate in result.document.sections
        for element in candidate.elements
        if isinstance(element, FormulaBlock)
    )
    assert normalized_formula.target_degradation is not None


def test_citation_semantic_json_is_canonical_and_byte_stable() -> None:
    markdown = "Body {{cite:b}} {{cite:a}}.\n\n:::references\n[a] A.\n[b] B.\n:::\n"
    first = _normalize(markdown).to_json()
    second = _normalize(markdown).to_json()
    assert json.dumps(first, ensure_ascii=False, separators=(",", ":")) == json.dumps(
        second, ensure_ascii=False, separators=(",", ":")
    )
    assert list(first["references"]) == sorted(first["references"])
