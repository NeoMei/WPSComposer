from __future__ import annotations

import json

import pytest

from skills.WPSComposer.scripts.document_model import (
    AbstractBlock,
    FigureBlock,
    ImageBlock,
    KeywordsBlock,
    Paragraph,
    Section,
    SemanticTableBlock,
    Span,
    StructuredDocument,
)
from skills.WPSComposer.scripts.generation_plan import GenerationPlan
from skills.WPSComposer.scripts.longform.plan import build_longform_plan
from skills.WPSComposer.scripts.longform.policy import LongformPolicy, build_policy
from skills.WPSComposer.scripts.longform.resources import preflight_resources
from skills.WPSComposer.scripts.longform.semantic import (
    BookmarkMapResult,
    LongformConfig,
    SemanticResult,
)


def make_config(**kwargs) -> LongformConfig:
    defaults = {
        "title": "默认标题",
        "short_title": "默认",
        "author": "作者",
        "date": "2026-08-21",
        "header": "默认页眉",
        "title_page": False,
        "toc": False,
        "figure_index": False,
        "table_index": False,
        "bibliography_include_uncited": True,
        "caption_numbering": "global",
        "heading_numbering": "none",
        "layout_engine": "longform",
    }
    defaults.update(kwargs)
    return LongformConfig(**defaults)


def make_document(title="默认标题", sections=None, abstract=None, keywords=None, title_display=None):
    return StructuredDocument(
        title=title,
        metadata={},
        sections=sections or [],
        config={},
        longform=True,
        issues=[],
        abstract=abstract,
        keywords=keywords,
        title_display=title_display,
    )


def make_semantic(
    title="默认标题",
    sections=None,
    config=None,
    references=None,
    issues=(),
    abstract=None,
    keywords=None,
    title_display=None,
) -> SemanticResult:
    doc = make_document(
        title=title,
        sections=sections or [],
        abstract=abstract,
        keywords=keywords,
        title_display=title_display,
    )
    return SemanticResult(
        document=doc,
        config=config or make_config(title=title),
        references=references or {},
        bookmarks=BookmarkMapResult(mapping={}, issues=()),
        issues=issues,
    )


def serialize_plan(plan: GenerationPlan) -> bytes:
    return json.dumps(plan.to_dict(), ensure_ascii=False, separators=(",", ":")).encode("utf-8")


def configure_sections(plan: GenerationPlan) -> list[dict]:
    return [op for op in plan.to_dict()["operations"] if op["op"] == "writer.configure_section"]


def find_op(plan: GenerationPlan, name: str) -> dict:
    for op in plan.to_dict()["operations"]:
        if op["op"] == name:
            return op
    raise AssertionError(f"operation {name} not found")


class TestSectionAwarePlanBuilder:
    """M2 Task 4: section-aware long-form plan builder."""

    def test_cover_does_not_consume_roman_page_number(self):
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", title_page=True),
            sections=[Section(level=1, heading="第一章 绪论", elements=[Paragraph(spans=[Span(text="Body.")])])],
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        sections = configure_sections(plan)
        roles = [s["args"]["role"] for s in sections]

        assert "cover" in roles
        cover = next(s for s in sections if s["args"]["role"] == "cover")
        assert cover["args"]["pageNumberFormat"] == "none"

        body = next(s for s in sections if s["args"]["role"] == "body")
        assert body["args"]["pageNumberFormat"] == "arabic"
        assert body["args"]["restartPageNumbering"] is True
        assert body["args"]["startPageNumber"] == 1
        assert not any(s["args"]["pageNumberFormat"] == "roman" for s in sections)

    def test_front_matter_roman_sequence(self):
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", toc=True),
            sections=[Section(level=1, heading="Intro", elements=[Paragraph(spans=[Span(text="Body.")])])],
            abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract text.")])]),
            keywords=KeywordsBlock(keywords=["one", "two"]),
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        sections = configure_sections(plan)
        front = next(s for s in sections if s["args"]["role"] == "front_matter")
        assert front["args"]["pageNumberFormat"] == "roman"
        assert front["args"]["restartPageNumbering"] is True
        assert front["args"]["startPageNumber"] == 1
        assert front["args"]["headerText"] == ""
        assert front["args"]["footerText"] == ""
        assert front["args"]["linkToPreviousHeader"] is False
        assert front["args"]["linkToPreviousFooter"] is False

    def test_body_restarts_arabic_numbering(self):
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author"),
            sections=[Section(level=1, heading="第一章 正文", elements=[Paragraph(spans=[Span(text="Body.")])])],
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        sections = configure_sections(plan)
        body = next(s for s in sections if s["args"]["role"] == "body")
        assert body["args"]["pageNumberFormat"] == "arabic"
        assert body["args"]["restartPageNumbering"] is True
        assert body["args"]["startPageNumber"] == 1
        assert body["args"]["headerText"] == "默认页眉"
        assert body["args"]["footerText"] == ""

    def test_no_empty_trailing_section(self):
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", title_page=True),
            sections=[],
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        sections = configure_sections(plan)
        assert all(s["args"]["role"] != "body" for s in sections)
        ops = [op["op"] for op in plan.to_dict()["operations"]]
        assert ops[-1] == "writer.finalize_fields"

    def test_toc_style_operation_includes_density_fields(self):
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", toc=True),
            sections=[Section(level=1, heading="第一章", elements=[Paragraph(spans=[Span(text="Body.")])])],
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        op = find_op(plan, "writer.configure_toc_styles")
        args = op["args"]
        assert "minFontSizePt" in args
        assert "minSpaceBeforePt" in args
        assert "minSpaceAfterPt" in args
        assert args["minFontSizePt"]["toc1"] == 10.5

    def test_heading_numbering_args_emitted(self):
        section = Section(
            level=1,
            heading="绪论",
            elements=[Paragraph(spans=[Span(text="Body.")])],
        )
        section.numbering = "chinese-formal"
        section.numbering_scheme = "chinese-formal"
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", heading_numbering="chinese-formal"),
            sections=[section],
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        headings = [op for op in plan.to_dict()["operations"] if op["op"] == "writer.add_heading"]
        assert len(headings) == 1
        args = headings[0]["args"]
        assert args["text"] == "绪论"
        assert args["numbering"] is True
        assert args["numberingScheme"] == "chinese-formal"

    def test_unnumbered_heading_does_not_emit_scheme(self):
        section = Section(
            level=1,
            heading="绪论",
            elements=[Paragraph(spans=[Span(text="Body.")])],
        )
        section.numbering = "none"
        section.numbering_scheme = None
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", heading_numbering="none"),
            sections=[section],
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        headings = [op for op in plan.to_dict()["operations"] if op["op"] == "writer.add_heading"]
        assert len(headings) == 1
        args = headings[0]["args"]
        assert args.get("numbering") is not True

    def test_deterministic_operation_order(self):
        semantic = make_semantic(
            title="Determinism",
            config=make_config(title="Determinism", author="Author", toc=True),
            sections=[
                Section(level=1, heading="第一章 A", elements=[Paragraph(spans=[Span(text="One.")])]),
                Section(level=2, heading="1.1 B", elements=[Paragraph(spans=[Span(text="Two.")])]),
            ],
            abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract.")])]),
        )
        preflight = preflight_resources([], ".")
        first = serialize_plan(build_longform_plan(semantic, preflight))
        second = serialize_plan(build_longform_plan(semantic, preflight))
        assert first == second

    def test_front_matter_grouped_before_first_body_section(self):
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", toc=True),
            sections=[Section(level=1, heading="第一章 正文", elements=[Paragraph(spans=[Span(text="Body.")])])],
            abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract.")])]),
            keywords=KeywordsBlock(keywords=["k1", "k2"]),
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        ops = plan.to_dict()["operations"]
        front_idx = next(i for i, op in enumerate(ops) if op["op"] == "writer.configure_section" and op["args"]["role"] == "front_matter")
        body_idx = next(i for i, op in enumerate(ops) if op["op"] == "writer.configure_section" and op["args"]["role"] == "body")
        assert front_idx < body_idx
        # Abstract and keywords paragraphs appear between front_matter configure_section and body configure_section.
        abstract_para_idx = next(i for i, op in enumerate(ops) if op["op"] == "writer.add_paragraph" and "Abstract" in op["args"]["text"])
        keywords_para_idx = next(i for i, op in enumerate(ops) if op["op"] == "writer.add_paragraph" and "k1" in op["args"]["text"])
        toc_idx = next(i for i, op in enumerate(ops) if op["op"] == "writer.insert_toc")
        assert front_idx < abstract_para_idx < body_idx
        assert front_idx < keywords_para_idx < body_idx
        assert front_idx < toc_idx < body_idx

    def test_each_section_has_exactly_one_configure_section(self):
        semantic = make_semantic(
            title="Report",
            config=make_config(title="Report", author="Author", title_page=True, toc=True),
            sections=[
                Section(level=1, heading="第一章 A", elements=[Paragraph(spans=[Span(text="One.")])]),
                Section(level=1, heading="第二章 B", elements=[Paragraph(spans=[Span(text="Two.")])]),
            ],
            abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract.")])]),
        )
        preflight = preflight_resources([], ".")
        plan = build_longform_plan(semantic, preflight)
        sections = configure_sections(plan)
        assert len(sections) == 3  # cover, front_matter, body
        roles = [s["args"]["role"] for s in sections]
        assert roles == ["cover", "front_matter", "body"]



class TestIndexPlacement:
    """M2 Task 4 follow-up: figure/table indexes belong in front matter when enabled."""

    def _plan_with_indexes(self, figure_index=False, table_index=False):
        elements = [Paragraph(spans=[Span(text="Body.")])]
        if figure_index:
            elements.append(
                FigureBlock(
                    identifier="fig:test",
                    node_id="fig:test",
                    caption="Test figure",
                    images=[ImageBlock(path="missing.png", alt="x")],
                    layout="stack",
                )
            )
        if table_index:
            elements.append(
                SemanticTableBlock(
                    identifier="tab:test",
                    node_id="tab:test",
                    caption="Test table",
                    headers=["H"],
                    rows=[["R"]],
                    alignments=["left"],
                )
            )
        section = Section(
            level=1,
            heading="第一章",
            elements=elements,
        )
        semantic = make_semantic(
            title="Report",
            config=make_config(
                title="Report",
                author="Author",
                toc=True,
                figure_index=figure_index,
                table_index=table_index,
            ),
            sections=[section],
            abstract=AbstractBlock(paragraphs=[Paragraph(spans=[Span(text="Abstract.")])]),
        )
        preflight = preflight_resources([], ".")
        return build_longform_plan(semantic, preflight)

    def test_figure_index_placed_in_front_matter(self):
        plan = self._plan_with_indexes(figure_index=True)
        ops = plan.to_dict()["operations"]
        front_idx = next(
            i for i, op in enumerate(ops)
            if op["op"] == "writer.configure_section" and op["args"]["role"] == "front_matter"
        )
        body_idx = next(
            i for i, op in enumerate(ops)
            if op["op"] == "writer.configure_section" and op["args"]["role"] == "body"
        )
        fig_idx = next(
            i for i, op in enumerate(ops) if op["op"] == "writer.insert_figure_index"
        )
        assert front_idx < fig_idx < body_idx

    def test_table_index_placed_in_front_matter(self):
        plan = self._plan_with_indexes(table_index=True)
        ops = plan.to_dict()["operations"]
        front_idx = next(
            i for i, op in enumerate(ops)
            if op["op"] == "writer.configure_section" and op["args"]["role"] == "front_matter"
        )
        body_idx = next(
            i for i, op in enumerate(ops)
            if op["op"] == "writer.configure_section" and op["args"]["role"] == "body"
        )
        tab_idx = next(
            i for i, op in enumerate(ops) if op["op"] == "writer.insert_table_index"
        )
        assert front_idx < tab_idx < body_idx

    def test_indexes_not_emitted_when_disabled(self):
        plan = self._plan_with_indexes(figure_index=False, table_index=False)
        ops = [op["op"] for op in plan.to_dict()["operations"]]
        assert "writer.insert_figure_index" not in ops
        assert "writer.insert_table_index" not in ops
