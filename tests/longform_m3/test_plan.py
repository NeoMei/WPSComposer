from __future__ import annotations

import json

from skills.WPSComposer.scripts.document_model import (
    CaptionBinding,
    CrossReferenceRun,
    FigureBlock,
    FormulaBlock,
    ImageBlock,
    Paragraph,
    Section,
    SemanticTableBlock,
    Span,
    StructuredDocument,
)
from skills.WPSComposer.scripts.longform.resources import (
    ImageProfile,
    PreflightResource,
    ResourcePreflight,
)
from skills.WPSComposer.scripts.generation_plan import validate_generation_plan
from skills.WPSComposer.scripts.longform.semantic import BookmarkMapResult, LongformConfig, SemanticResult
from skills.WPSComposer.scripts.longform.plan import build_longform_plan


def _semantic(elements, *, figure_index=True, table_index=True) -> SemanticResult:
    config = LongformConfig(
        title="M3", short_title="M3", author="A", date="", header="",
        title_page=False, toc=True, figure_index=figure_index, table_index=table_index,
        bibliography_include_uncited=True, caption_numbering="auto",
        heading_numbering="decimal", layout_engine="longform",
    )
    document = StructuredDocument(
        title="M3", longform=True,
        sections=[Section(level=1, heading="Chapter", node_id="sec:one", numbering="decimal", numbering_scheme="decimal", elements=elements)],
    )
    return SemanticResult(document, config, {}, BookmarkMapResult({}, ()), ())


def _preflight() -> ResourcePreflight:
    profile = ImageProfile(800, 400, 144.0, 144.0, 1, "PNG", False)
    resource = PreflightResource(
        "image-1", "fig.png", "1" * 64, "1" * 64, 16, "image/png", "none-v1", profile, b"private"
    )
    return ResourcePreflight(
        resources=(resource,), degradations=(),
        manifest={"version": "1", "entries": [], "digest": "sha256:" + "0" * 64},
    )


def _find(plan, name: str):
    return [op for op in plan.to_dict()["operations"] if op["op"] == name]


def test_plan_emits_resolved_native_objects_and_one_reference_paragraph() -> None:
    fig_mark = "wpsc_fig_" + "a" * 24
    tab_mark = "wpsc_tab_" + "b" * 24
    eq_mark = "wpsc_eq_" + "c" * 24
    ref = CrossReferenceRun("ref:1", "fig:one", "fig:one", "figure", fig_mark, "[图]")
    elements = [
        FigureBlock("fig:one", "fig:one", "Figure", [ImageBlock("fig.png", "")], "stack", "full", "portrait", "diagram", None, CaptionBinding("chapter", "sec:one", fig_mark, True, True)),
        SemanticTableBlock("tab:one", "tab:one", "Table", ["A", "B"], [["1", "2"]], ["left", "right"], "three-line", "portrait", "", True, CaptionBinding("chapter", "sec:one", tab_mark, True, True)),
        FormulaBlock("eq:one", "eq:one", "E=mc^2", None, CaptionBinding("chapter", "sec:one", eq_mark, True, True)),
        Paragraph([Span("见"), Span("", cross_reference=ref), Span("。")], node_id="para:one"),
    ]
    plan = build_longform_plan(_semantic(elements), _preflight())

    figure = _find(plan, "writer.add_captioned_figure")[0]
    assert figure["args"]["numbering"]["sequenceId"] == "WPSC_FIG"
    assert figure["args"]["bookmarkName"] == fig_mark
    assert figure["args"]["children"][0]["displayWidthPt"] > 0
    assert figure["args"]["children"][0]["normalizerId"] == "none-v1"
    table = _find(plan, "writer.add_semantic_table")[0]
    assert table["args"]["borderSpec"]["top"] == 1.5
    assert table["args"]["cellIndentPt"] == 0.0
    equation = _find(plan, "writer.add_equation")[0]
    assert equation["args"]["numbering"]["sequenceId"] == "WPSC_EQ"
    assert "nativeMath" not in equation["args"]
    reference = _find(plan, "writer.add_cross_reference")[0]
    assert [run["type"] for run in reference["args"]["runs"]] == ["text", "reference", "text"]
    assert reference["nodeId"] == "para:one"
    assert validate_generation_plan(plan.to_dict(), "writer") == plan


def test_caption_missing_omits_indexes_and_reference_targets() -> None:
    missing = FigureBlock(
        "fig:missing", "fig:missing", "", [ImageBlock("fig.png", "")],
        caption_binding=CaptionBinding("global", None, None, False, False),
    )
    plan = build_longform_plan(_semantic([missing], table_index=False), _preflight())
    assert not _find(plan, "writer.insert_figure_index")
    figure = _find(plan, "writer.add_captioned_figure")[0]
    assert "bookmarkName" not in figure["args"]
    assert figure["args"]["indexable"] is False


def test_front_matter_indexes_are_native_and_precede_body_objects() -> None:
    fig = FigureBlock(
        "fig:one", "fig:one", "Figure", [ImageBlock("fig.png", "")],
        caption_binding=CaptionBinding("global", None, "wpsc_fig_" + "a" * 24, True, True),
    )
    tab = SemanticTableBlock(
        "tab:one", "tab:one", "Table", ["A"], [["1"]], ["left"],
        caption_binding=CaptionBinding("global", None, "wpsc_tab_" + "b" * 24, True, True),
    )
    plan = build_longform_plan(_semantic([fig, tab]), _preflight()).to_dict()["operations"]
    fig_index = next(i for i, op in enumerate(plan) if op["op"] == "writer.insert_figure_index")
    body_fig = next(i for i, op in enumerate(plan) if op["op"] == "writer.add_captioned_figure")
    assert fig_index < body_fig
    assert plan[fig_index]["args"] == {"title": "图目录", "sequenceId": "WPSC_FIG", "titleStyleId": "WPSC_INDEX_TITLE"}
    assert plan[-1]["op"] == "writer.finalize_fields"


def test_operation_json_is_byte_deterministic() -> None:
    fig = FigureBlock(
        "fig:one", "fig:one", "Figure", [ImageBlock("fig.png", "")],
        caption_binding=CaptionBinding("global", None, "wpsc_fig_" + "a" * 24, True, True),
    )
    semantic = _semantic([fig], table_index=False)
    first = json.dumps(build_longform_plan(semantic, _preflight()).to_dict(), ensure_ascii=False, separators=(",", ":")).encode()
    second = json.dumps(build_longform_plan(semantic, _preflight()).to_dict(), ensure_ascii=False, separators=(",", ":")).encode()
    assert first == second


def test_unresolved_reference_is_plain_inline_fallback_not_a_ref_descriptor() -> None:
    unresolved = CrossReferenceRun("ref:1", "fig:missing", None, None, None, "引用目标未解析")
    paragraph = Paragraph(
        [Span("见"), Span("引用目标未解析", cross_reference=unresolved), Span("。")],
        node_id="para:one",
    )
    plan = build_longform_plan(
        _semantic([paragraph], figure_index=False, table_index=False), _preflight()
    )
    assert not _find(plan, "writer.add_cross_reference")
    rendered = _find(plan, "writer.add_paragraph")[0]
    assert rendered["args"]["text"] == "见引用目标未解析。"
