from __future__ import annotations

from skills.WPSComposer import parse
from skills.WPSComposer.scripts.document_model import Paragraph, TableBlock, TaskList


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
