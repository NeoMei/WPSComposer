from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import sys
from unittest.mock import patch

from skills.WPSComposer.scripts.document_model import (
    FigureBlock,
    ListBlock,
    Paragraph,
    Section,
    SemanticTableBlock,
)
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


def test_abstract_paragraphs_receive_ids_and_split_multiple_references() -> None:
    result = _normalize(
        """:::abstract
See {{ref:fig:a}}, {{ref:missing}}, and `{{ref:missing}}`.
:::

:::figure {#fig:a caption="A"}
![a](a.png)
:::
"""
    )
    abstract = result.document.abstract
    assert abstract is not None
    (paragraph,) = abstract.paragraphs
    assert paragraph.node_id == "__wpsc_para:0:1"
    runs = [span.cross_reference for span in paragraph.spans if span.cross_reference]
    assert [run.node_id for run in runs] == [
        "__wpsc_para:0:1/ref:1",
        "__wpsc_para:0:1/ref:2",
    ]
    assert runs[0].target_node_id == "fig:a"
    assert runs[0].bookmark_name is not None
    assert runs[1].target_node_id is None
    code = next(span for span in paragraph.spans if span.code)
    assert code.text == "{{ref:missing}}"
    assert code.cross_reference is None
    assert len([issue for issue in result.issues if issue.code == REFERENCE_UNRESOLVED]) == 1


def test_list_items_receive_ids_and_split_references_without_duplicate_issues() -> None:
    result = _normalize(
        """:::figure {#fig:a caption="A"}
![a](a.png)
:::

- First {{ref:fig:a}} and {{ref:missing}}.
- `{{ref:missing}}` then {{ref:fig:a}} and {{ref:fig:a}}.
"""
    )
    list_block = next(
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, ListBlock)
    )
    assert list_block.item_node_ids == [
        "__wpsc_para:1:1",
        "__wpsc_para:1:2",
    ]
    first_runs = [span.cross_reference for span in list_block.items[0] if span.cross_reference]
    second_runs = [span.cross_reference for span in list_block.items[1] if span.cross_reference]
    assert [run.node_id for run in first_runs] == [
        "__wpsc_para:1:1/ref:1",
        "__wpsc_para:1:1/ref:2",
    ]
    assert [run.node_id for run in second_runs] == [
        "__wpsc_para:1:2/ref:1",
        "__wpsc_para:1:2/ref:2",
    ]
    assert first_runs[0].target_node_id == "fig:a"
    assert first_runs[1].target_node_id is None
    assert all(run.target_node_id == "fig:a" for run in second_runs)
    code = next(span for span in list_block.items[1] if span.code)
    assert code.text == "{{ref:missing}}"
    assert code.cross_reference is None
    assert len([issue for issue in result.issues if issue.code == REFERENCE_UNRESOLVED]) == 1


def test_abstract_and_list_reference_order_is_canonical_across_runs() -> None:
    markdown = """:::abstract
Abstract {{ref:fig:a}}.
:::

:::figure {#fig:a caption="A"}
![a](a.png)
:::

- List {{ref:fig:a}}.
"""
    first = _normalize(markdown).to_json()
    second = _normalize(markdown).to_json()
    assert first == second
    assert first["document"]["abstract"]["paragraphs"][0]["node_id"] == "__wpsc_para:0:1"
    list_json = next(
        element
        for section in first["document"]["sections"]
        for element in section["elements"]
        if "item_node_ids" in element
    )
    assert list_json["item_node_ids"] == ["__wpsc_para:1:1"]


def test_issue_order_and_semantic_json_are_stable_across_hash_seeds() -> None:
    source = r'''
import json
from skills.WPSComposer.scripts.md_parser import parse_markdown
from skills.WPSComposer.scripts.longform.semantic import normalize_longform_document

markdown = """:::abstract
Abstract {{ref:abs-z}} then {{cite:abs-c}}.
:::

- List {{ref:list-b}} then {{cite:list-d}}.

Ordinary {{ref:body-a}} then {{cite:body-e}}.
"""
result = normalize_longform_document(parse_markdown(markdown, longform=True))
print(json.dumps(result.to_json(), ensure_ascii=False, separators=(",", ":")))
'''
    repo_root = Path(__file__).resolve().parents[2]
    outputs = []
    for seed in ("1", "777"):
        env = os.environ.copy()
        env["PYTHONHASHSEED"] = seed
        completed = subprocess.run(
            [sys.executable, "-c", source],
            cwd=repo_root,
            env=env,
            check=True,
            capture_output=True,
            text=True,
        )
        outputs.append(completed.stdout.encode("utf-8"))

    assert outputs[0] == outputs[1]
    payload = json.loads(outputs[0])
    messages = [
        issue["message"]
        for issue in payload["issues"]
        if issue["code"] == REFERENCE_UNRESOLVED
    ]
    assert messages == [
        "Cross-reference target 'abs-z' was not found.",
        "Citation target 'abs-c' was not found.",
        "Cross-reference target 'list-b' was not found.",
        "Citation target 'list-d' was not found.",
        "Cross-reference target 'body-a' was not found.",
        "Citation target 'body-e' was not found.",
    ]


def test_abstract_raw_list_ids_align_without_duplicate_reference_runs() -> None:
    markdown = """:::abstract
- First {{ref:fig:a}}.
- Second `{{ref:fig:a}}` and {{ref:missing}}.
:::

:::figure {#fig:a caption="A"}
![a](a.png)
:::
"""
    parsed = parse_markdown(markdown, longform=True)
    first = normalize_longform_document(parsed)
    second = normalize_longform_document(parsed)

    assert first.to_json() == second.to_json()
    abstract = first.document.abstract
    assert abstract is not None
    raw_list = next(
        element for element in abstract.raw_elements if isinstance(element, ListBlock)
    )
    assert raw_list.item_node_ids == [
        "__wpsc_para:0:1",
        "__wpsc_para:0:2",
    ]
    assert [paragraph.node_id for paragraph in abstract.paragraphs] == raw_list.item_node_ids

    raw_runs = [
        span.cross_reference
        for item in raw_list.items
        for span in item
        if span.cross_reference is not None
    ]
    normalized_runs = [
        span.cross_reference
        for paragraph in abstract.paragraphs
        for span in paragraph.spans
        if span.cross_reference is not None
    ]
    assert raw_runs == []
    assert [run.node_id for run in normalized_runs] == [
        "__wpsc_para:0:1/ref:1",
        "__wpsc_para:0:2/ref:1",
    ]
    assert normalize_longform_document(first.document).to_json() == first.to_json()


def test_nested_abstract_raw_list_ids_reuse_projected_paragraph_ids() -> None:
    markdown = """:::abstract
# Nested heading

- First {{ref:fig:a}} and {{ref:missing}}.
- Second `{{ref:missing}}` and {{ref:missing}}.
:::

:::figure {#fig:a caption="A"}
![a](a.png)
:::
"""
    parsed = parse_markdown(markdown, longform=True)
    first = normalize_longform_document(parsed)
    second = normalize_longform_document(parsed)
    assert first.to_json() == second.to_json()

    abstract = first.document.abstract
    assert abstract is not None
    raw_section = next(
        element for element in abstract.raw_elements if isinstance(element, Section)
    )
    raw_list = next(
        element for element in raw_section.elements if isinstance(element, ListBlock)
    )
    assert raw_list.item_node_ids == [
        "__wpsc_para:0:2",
        "__wpsc_para:0:3",
    ]
    assert [paragraph.node_id for paragraph in abstract.paragraphs] == [
        "__wpsc_para:0:1",
        *raw_list.item_node_ids,
    ]

    raw_runs = [
        span.cross_reference
        for item in raw_list.items
        for span in item
        if span.cross_reference is not None
    ]
    projected_runs = [
        span.cross_reference
        for paragraph in abstract.paragraphs
        for span in paragraph.spans
        if span.cross_reference is not None
    ]
    assert raw_runs == []
    assert [run.node_id for run in projected_runs] == [
        "__wpsc_para:0:2/ref:1",
        "__wpsc_para:0:2/ref:2",
        "__wpsc_para:0:3/ref:1",
    ]
    assert len([issue for issue in first.issues if issue.code == REFERENCE_UNRESOLVED]) == 1
    assert normalize_longform_document(first.document).to_json() == first.to_json()
