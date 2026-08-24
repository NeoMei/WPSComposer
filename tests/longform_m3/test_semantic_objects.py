from __future__ import annotations

from skills.WPSComposer.scripts.document_model import (
    FigureBlock,
    FormulaBlock,
    SemanticTableBlock,
    StructuredDocument,
)
from skills.WPSComposer.scripts.md_parser import parse_markdown
from skills.WPSComposer.scripts.longform.semantic import (
    CAPTION_MISSING,
    normalize_longform_document,
)


def _normalize(markdown: str):
    return normalize_longform_document(parse_markdown(markdown, longform=True))


def _semantic_blocks(result):
    return [
        element
        for section in result.document.sections
        for element in section.elements
        if isinstance(element, (FigureBlock, SemanticTableBlock, FormulaBlock))
    ]


def test_auto_caption_numbering_is_resolved_per_object() -> None:
    result = _normalize(
        """---
title: Report
heading_numbering: decimal
caption_numbering: auto
---
:::figure {#fig:before caption="Before"}
![before](before.png)
:::

# 1 Chapter

:::figure {#fig:after caption="After"}
![after](after.png)
:::
"""
    )

    before, after = _semantic_blocks(result)
    assert result.config.caption_numbering == "auto"
    assert before.caption_binding is not None
    assert before.caption_binding.mode == "global"
    assert before.caption_binding.chapter_node_id is None
    assert after.caption_binding is not None
    assert after.caption_binding.mode == "chapter"
    numbered_h1 = next(
        section
        for section in result.document.sections
        if section.level == 1 and section.numbering != "none"
    )
    assert after.caption_binding.chapter_node_id == numbered_h1.node_id


def test_explicit_global_never_switches_to_chapter() -> None:
    result = _normalize(
        """---
title: Report
heading_numbering: decimal
caption_numbering: global
---
# 1 Chapter

:::table {#tab:a caption="A"}
| A |
|---|
| 1 |
:::
"""
    )
    (table,) = _semantic_blocks(result)
    assert table.caption_binding is not None
    assert table.caption_binding.mode == "global"
    assert table.caption_binding.chapter_node_id is None


def test_explicit_chapter_falls_back_before_first_numbered_h1() -> None:
    result = _normalize(
        """---
title: Report
heading_numbering: decimal
caption_numbering: chapter
---
:::table {#tab:before caption="Before"}
| A |
|---|
| 1 |
:::

# 1 Chapter

:::table {#tab:after caption="After"}
| A |
|---|
| 2 |
:::
"""
    )
    before, after = _semantic_blocks(result)
    assert result.config.caption_numbering == "chapter"
    assert before.caption_binding.mode == "global"
    assert after.caption_binding.mode == "chapter"


def test_unnumbered_h1_does_not_reset_the_active_numbered_chapter() -> None:
    result = _normalize(
        """---
title: Report
heading_numbering: decimal
caption_numbering: auto
---
# 1 First

:::figure {#fig:first caption="First"}
![first](first.png)
:::

# 第二章 Unnumbered

:::figure {#fig:second caption="Second"}
![second](second.png)
:::
"""
    )
    first, second = _semantic_blocks(result)
    assert first.caption_binding.mode == "chapter"
    assert second.caption_binding.mode == "chapter"
    assert second.caption_binding.chapter_node_id == first.caption_binding.chapter_node_id


def test_figure_table_and_equation_have_independent_bindings() -> None:
    result = _normalize(
        """---
title: Report
heading_numbering: decimal
caption_numbering: auto
---
# 1 Chapter

:::figure {#fig:a caption="Figure"}
![a](a.png)
:::

:::table {#tab:a caption="Table"}
| A |
|---|
| 1 |
:::

:::formula {#eq:a}
x = 1
:::
"""
    )
    figure, table, equation = _semantic_blocks(result)
    bindings = [figure.caption_binding, table.caption_binding, equation.caption_binding]
    assert all(binding is not None for binding in bindings)
    assert {binding.mode for binding in bindings} == {"chapter"}
    assert len({binding.bookmark_name for binding in bindings}) == 3
    assert figure.caption_binding.bookmark_name.startswith("wpsc_fig_")
    assert table.caption_binding.bookmark_name.startswith("wpsc_tab_")
    assert equation.caption_binding.bookmark_name.startswith("wpsc_eq_")


def test_empty_figure_and_table_captions_keep_objects_but_disable_caption_contract() -> None:
    result = _normalize(
        """---
caption_numbering: auto
---
:::figure {#fig:empty caption="   "}
![kept](kept.png)
:::

:::table {#tab:empty}
| A |
|---|
| 1 |
:::
"""
    )
    figure, table = _semantic_blocks(result)

    assert len(figure.images) == 1
    assert table.rows == [["1"]]
    for block in (figure, table):
        assert block.caption_binding is not None
        assert block.caption_binding.indexable is False
        assert block.caption_binding.referenceable is False
        assert block.caption_binding.bookmark_name is None
    caption_issues = [issue for issue in result.issues if issue.code == CAPTION_MISSING]
    assert len(caption_issues) == 2
    assert all(issue.placement == "block" for issue in caption_issues)


def test_semantic_snapshot_serializes_caption_bindings_canonically() -> None:
    result = _normalize(
        """:::figure {#fig:a caption="A"}
![a](a.png)
:::
"""
    )
    document = result.to_json()["document"]
    binding = document["sections"][0]["elements"][0]["caption_binding"]
    assert list(binding) == sorted(binding)
    assert binding["mode"] == "global"


def test_contract_dataclasses_are_immutable() -> None:
    result = _normalize(
        """:::figure {#fig:a caption="A"}
![a](a.png)
:::
"""
    )
    (figure,) = _semantic_blocks(result)
    try:
        figure.caption_binding.mode = "chapter"
    except Exception as error:  # frozen dataclasses raise FrozenInstanceError
        assert type(error).__name__ == "FrozenInstanceError"
    else:  # pragma: no cover - documents the immutability contract
        raise AssertionError("CaptionBinding must be immutable")
