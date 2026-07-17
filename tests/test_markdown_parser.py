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
