from __future__ import annotations

from unittest.mock import patch

from skills.WPSComposer.scripts.document_model import FigureBlock, Paragraph, SemanticTableBlock
from skills.WPSComposer.scripts.md_parser import parse_markdown
from skills.WPSComposer.scripts.longform.bookmark_ids import BOOKMARK_NAME_COLLISION
from skills.WPSComposer.scripts.longform.semantic import (
    REFERENCE_UNRESOLVED,
    normalize_longform_document,
)


def _normalize(markdown: str):
    return normalize_longform_document(parse_markdown(markdown, longform=True))


def _paragraphs(result):
    return [
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, Paragraph)
    ]


def test_multiple_references_split_into_ordered_literal_and_reference_spans() -> None:
    result = _normalize(
        """:::figure {#fig:a caption="A"}
![a](a.png)
:::

:::table {#tab:b caption="B"}
| A |
|---|
| 1 |
:::

See {{ref:fig:a}} and {{ref:tab:b}} now.
"""
    )
    paragraph = _paragraphs(result)[0]
    assert paragraph.node_id == "__wpsc_para:1:1"
    assert [span.text for span in paragraph.spans] == [
        "See ",
        "引用目标未解析",
        " and ",
        "引用目标未解析",
        " now.",
    ]
    runs = [span.cross_reference for span in paragraph.spans if span.cross_reference]
    assert [run.node_id for run in runs] == [
        "__wpsc_para:1:1/ref:1",
        "__wpsc_para:1:1/ref:2",
    ]
    assert [run.target_id for run in runs] == ["fig:a", "tab:b"]
    assert [run.target_kind for run in runs] == ["fig", "tab"]
    assert all(run.bookmark_name for run in runs)


def test_reference_markers_inside_code_and_math_spans_remain_literal() -> None:
    result = _normalize(
        """:::figure {#fig:a caption="A"}
![a](a.png)
:::

`{{ref:fig:a}}` and ${{ref:fig:a}}$ and {{ref:fig:a}}.
"""
    )
    paragraph = _paragraphs(result)[0]
    code_span = next(span for span in paragraph.spans if span.code)
    math_span = next(span for span in paragraph.spans if span.math)
    assert code_span.text == "{{ref:fig:a}}"
    assert code_span.cross_reference is None
    assert math_span.text == "{{ref:fig:a}}"
    assert math_span.cross_reference is None
    assert len([span for span in paragraph.spans if span.cross_reference]) == 1


def test_unresolved_reference_target_and_fallback_are_nfc() -> None:
    result = _normalize("See {{ref:cafe\u0301}}.\n")
    paragraph = _paragraphs(result)[0]
    run = next(span.cross_reference for span in paragraph.spans if span.cross_reference)
    assert run.target_id == "café"
    assert run.fallback_text == "引用目标未解析"
    assert run.target_node_id is None
    assert run.target_kind is None
    assert run.bookmark_name is None
    assert any(issue.code == REFERENCE_UNRESOLVED for issue in result.issues)


def test_paragraph_and_reference_occurrence_ids_are_deterministic() -> None:
    markdown = """:::figure {#fig:a caption="A"}
![a](a.png)
:::

First {{ref:fig:a}}.

Second {{ref:fig:a}} and {{ref:fig:a}}.
"""
    first = _normalize(markdown)
    second = _normalize(markdown)
    first_paragraphs = _paragraphs(first)
    second_paragraphs = _paragraphs(second)
    assert [p.node_id for p in first_paragraphs] == [p.node_id for p in second_paragraphs]
    assert [p.node_id for p in first_paragraphs] == [
        "__wpsc_para:1:1",
        "__wpsc_para:1:2",
    ]
    first_run_ids = [
        span.cross_reference.node_id
        for paragraph in first_paragraphs
        for span in paragraph.spans
        if span.cross_reference
    ]
    second_run_ids = [
        span.cross_reference.node_id
        for paragraph in second_paragraphs
        for span in paragraph.spans
        if span.cross_reference
    ]
    assert first_run_ids == second_run_ids == [
        "__wpsc_para:1:1/ref:1",
        "__wpsc_para:1:2/ref:1",
        "__wpsc_para:1:2/ref:2",
    ]


def test_bookmark_name_is_stable_when_target_moves_between_sections() -> None:
    first = _normalize(
        """# 1 First

:::figure {#fig:stable caption="Stable"}
![a](a.png)
:::

# 2 Second
"""
    )
    moved = _normalize(
        """# 1 First

# 2 Second

:::figure {#fig:stable caption="Stable"}
![a](a.png)
:::
"""
    )
    assert first.bookmarks.mapping["fig:stable"] == moved.bookmarks.mapping["fig:stable"]


def test_bookmark_collision_exhaustion_disables_target_and_inline_run() -> None:
    class _ConstantDigest:
        def hexdigest(self) -> str:
            return "a" * 64

    with patch(
        "skills.WPSComposer.scripts.longform.bookmark_ids.hashlib.sha256",
        return_value=_ConstantDigest(),
    ):
        result = _normalize(
            """:::figure {#fig:a caption="A"}
![a](a.png)
:::

:::figure {#fig:b caption="B"}
![b](b.png)
:::

{{ref:fig:a}} / {{ref:fig:b}}
"""
        )

    figures = [
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, FigureBlock)
    ]
    assert figures[0].caption_binding.referenceable is True
    assert figures[1].caption_binding.referenceable is False
    assert any(issue.code == BOOKMARK_NAME_COLLISION for issue in result.issues)
    runs = [span.cross_reference for span in _paragraphs(result)[0].spans if span.cross_reference]
    assert runs[0].bookmark_name is not None
    assert runs[1].bookmark_name is None
    assert runs[1].target_node_id is None
    assert len([issue for issue in result.issues if issue.code == REFERENCE_UNRESOLVED]) == 1


def test_empty_caption_target_is_an_unresolved_inline_reference() -> None:
    result = _normalize(
        """:::table {#tab:empty caption=""}
| A |
|---|
| 1 |
:::

See {{ref:tab:empty}}.
"""
    )
    paragraph = _paragraphs(result)[0]
    run = next(span.cross_reference for span in paragraph.spans if span.cross_reference)
    assert run.target_id == "tab:empty"
    assert run.target_node_id is None
    assert run.target_kind is None
    assert run.bookmark_name is None
    assert next(span.text for span in paragraph.spans if span.cross_reference) == "引用目标未解析"
    assert any(issue.code == REFERENCE_UNRESOLVED and issue.placement == "inline" for issue in result.issues)


def test_reference_runs_serialize_as_nested_canonical_values() -> None:
    result = _normalize(
        """:::figure {#fig:a caption="A"}
![a](a.png)
:::

See {{ref:fig:a}}.
"""
    )
    paragraph = next(
        element
        for section in result.to_json()["document"]["sections"]
        for element in section["elements"]
        if element.get("node_id", "").startswith("__wpsc_para")
    )
    run = next(span["cross_reference"] for span in paragraph["spans"] if span["cross_reference"])
    assert list(run) == sorted(run)
    assert run["target_id"] == "fig:a"
