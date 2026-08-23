from __future__ import annotations

import time

from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation


def _synthetic_eighty_page_markdown() -> str:
    parts = [
        "---\n",
        "title: M5 80-page performance fixture\n",
        "author: WPSComposer\n",
        "toc: true\n",
        "title_page: true\n",
        "heading_numbering: decimal\n",
        "---\n",
    ]
    for index in range(1, 81):
        parts.extend([
            f"\n# 性能章节 {index}\n\n",
            "这一页包含确定性的非客户文本，用于测量计划构建与原生分页预算。\n\n",
        ])
        if index < 80:
            parts.append(":::page-break\n:::\n")
    return "".join(parts)


def test_eighty_page_fixture_builds_within_offline_budget():
    markdown = _synthetic_eighty_page_markdown()
    started = time.monotonic()
    build = build_longform_generation(markdown, base_dir=".")
    elapsed = time.monotonic() - started
    breaks = [
        item for item in build.plan.operations if item.op == "writer.add_page_break"
    ]
    assert len(breaks) == 79
    assert len(build.plan.operations) >= 160
    assert elapsed < 5.0
