from __future__ import annotations

from skills.WPSComposer import parse
import pytest

from skills.WPSComposer.scripts.document_model import (
    ExcalidrawBlock,
    ImageBlock,
    Paragraph,
    TableBlock,
    TaskList,
)


def test_parse_preserves_heading_task_list_and_table():
    document = parse(
        """# 项目状态

正文 **加粗**

- [x] 已完成
- [ ] 待处理

| 任务 | 状态 |
|:---|---:|
| 基线修复 | 进行中 |
"""
    )

    assert document.title == "项目状态"
    assert len(document.sections) == 1
    elements = document.sections[0].elements
    assert [type(element) for element in elements] == [
        Paragraph,
        TaskList,
        TableBlock,
    ]
    table = elements[-1]
    assert table.headers == ["任务", "状态"]
    assert table.rows == [["基线修复", "进行中"]]
    assert table.alignments == ["left", "right"]


def test_nested_list_items_are_kept_in_parent_list():
    from skills.WPSComposer.scripts.md_parser import parse
    from skills.WPSComposer.scripts.document_model import ListBlock

    document = parse("- a\n  - b\n- c\n")
    (block,) = document.sections[0].elements
    assert isinstance(block, ListBlock)
    assert [["".join(s.text for s in item) for item in block.items]] == [["a", "b", "c"]]

    document = parse("1. one\n   1. two\n2. three\n")
    (block,) = document.sections[0].elements
    assert isinstance(block, ListBlock)
    assert block.ordered
    assert len(block.items) == 3


def test_dash_prefixed_continuation_stays_in_paragraph():
    from skills.WPSComposer.scripts.md_parser import parse
    from skills.WPSComposer.scripts.document_model import Paragraph

    document = parse("Growth was strong\n-5 percent in Q3\n")
    (block,) = document.sections[0].elements
    assert isinstance(block, Paragraph)
    assert block.plain_text == "Growth was strong -5 percent in Q3"


def test_table_starting_with_separator_row_has_no_literal_dash_headers():
    from skills.WPSComposer.scripts.md_parser import parse

    document = parse("| --- | --- |\n| a | b |\n")
    (table,) = document.sections[0].elements
    assert table.headers == []
    assert table.rows == [["a", "b"]]


def test_plain_wikilink_image_is_not_modeled_as_excalidraw(tmp_path):
    document = parse("![[photo.png|320x200]]", base_dir=str(tmp_path))

    (image,) = document.sections[0].elements
    assert isinstance(image, ImageBlock)
    assert not isinstance(image, ExcalidrawBlock)
    assert image.path == str((tmp_path / "photo.png").resolve())
    assert (image.width, image.height) == (320, 200)


def test_inline_markdown_image_is_rejected_instead_of_losing_the_image():
    with pytest.raises(ValueError, match="Inline Markdown images are not supported"):
        parse("Before ![chart](chart.png) after")


def test_image_syntax_inside_inline_code_is_literal_text():
    document = parse("Example: `![literal](not-an-image.png)`")

    (paragraph,) = document.sections[0].elements
    assert paragraph.plain_text == "Example: ![literal](not-an-image.png)"
    assert paragraph.spans[-1].code is True


def test_escaped_inline_image_syntax_is_not_rejected():
    document = parse(r"Example: \![literal](not-an-image.png)")

    (paragraph,) = document.sections[0].elements
    assert paragraph.plain_text == r"Example: \![literal](not-an-image.png)"
