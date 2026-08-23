from __future__ import annotations

from skills.WPSComposer.scripts.document_model import (
    BibliographyEntry,
    BlockQuote,
    CaptionBinding,
    CitationRun,
    DocumentIssue,
    DegradationBlock,
    FormulaBlock,
    InlineDegradationRun,
    PageBreakBlock,
    Paragraph,
    ReferenceListBlock,
    Section,
    SemanticTableBlock,
    Span,
    StructuredDocument,
    TableCellDegradation,
)
from skills.WPSComposer.scripts.generation_plan import validate_generation_plan
from skills.WPSComposer.scripts.recording_composers import RecordingWriterComposer
from skills.WPSComposer.scripts.longform.native_math import convert_restricted_latex
from skills.WPSComposer.scripts.longform.plan import build_longform_plan
from skills.WPSComposer.scripts.longform.resources import (
    FORMULA_FALLBACK_IMAGE_UNAVAILABLE,
    ResourceDegradation,
    ResourcePreflight,
)
from skills.WPSComposer.scripts.longform.semantic import (
    BookmarkMapResult,
    LongformConfig,
    SemanticResult,
)


def _config() -> LongformConfig:
    return LongformConfig(
        title="M4",
        short_title="M4",
        author="",
        date="",
        header="",
        title_page=False,
        toc=False,
        figure_index=False,
        table_index=False,
        bibliography_include_uncited=True,
        caption_numbering="global",
        heading_numbering="none",
        layout_engine="longform",
    )


def _semantic(body, bibliography=(), *, issues=()) -> SemanticResult:
    sections = [
        Section(
            level=1,
            heading="Body",
            node_id="sec:body",
            elements=list(body),
        )
    ]
    if bibliography:
        sections.append(
            Section(
                level=1,
                heading="References",
                node_id="sec:bibliography",
                page_role="bibliography",
                elements=list(bibliography),
            )
        )
    document = StructuredDocument(title="M4", longform=True, sections=sections)
    return SemanticResult(
        document,
        _config(),
        {},
        BookmarkMapResult({}, ()),
        tuple(issues),
    )


def _preflight(*, bindings=None, degradations=()) -> ResourcePreflight:
    return ResourcePreflight(
        resources=[],
        degradations=list(degradations),
        manifest={
            "version": "1",
            "entries": [],
            "digest": "sha256:" + "0" * 64,
        },
        formula_bindings=dict(bindings or {}),
    )


def _ops(plan, name: str) -> list[dict]:
    return [item for item in plan.to_dict()["operations"] if item["op"] == name]


def _formula(*, valid: bool = True) -> FormulaBlock:
    issue = None if valid else DocumentIssue(
        "FORMULA_MALFORMED",
        "Formula source is malformed.",
        "block",
    )
    return FormulaBlock(
        identifier="eq:one",
        node_id="eq:one",
        source="x+y",
        raw_source="x+y",
        native_math=convert_restricted_latex("x+y") if valid else None,
        content_degradation=issue,
        caption_binding=CaptionBinding(
            "global",
            None,
            "wpsc_eq_" + "a" * 24,
            True,
            True,
        ),
    )


def _bibliography() -> ReferenceListBlock:
    return ReferenceListBlock(
        node_id="bib:block",
        resolved_items=[
            BibliographyEntry("ref:a", "ref:a", "Alpha.", 1, 2, True),
            BibliographyEntry("ref:b", "ref:b", "Beta.", 2, 1, False),
        ],
    )


def test_plan_emits_native_formula_and_orthogonal_fallback_resource() -> None:
    formula = _formula()
    plan = build_longform_plan(
        _semantic([formula]),
        _preflight(bindings={"eq:one": "wpsc-rsrc:opaque"}),
    )

    equation = _ops(plan, "writer.add_equation")[0]
    assert equation["args"] == {
        "renderMode": "native-m4",
        "content": {
            "nativeMath": {
                "syntax": "wps-linear-v1",
                "linearText": "x+y",
                "sourceHash": convert_restricted_latex("x+y").source_hash,
            }
        },
        "fallbackResource": {"fallbackResourceId": "wpsc-rsrc:opaque"},
        "numbering": {
            "mode": "global",
            "sequenceId": "WPSC_EQ",
            "chapterStyleLevel": None,
            "resetLevel": None,
            "prefix": "(",
            "suffix": ")",
        },
        "bookmarkName": "wpsc_eq_" + "a" * 24,
        "fallbackText": "x+y",
    }
    assert equation["failurePolicy"] == {
        "mode": "degrade",
        "recoverableCodes": ["EQUATION_INSERT_FAILED"],
        "fallback": "explicit-image-then-source-notice",
    }
    assert validate_generation_plan(plan.to_dict(), "writer") == plan


def test_fallback_resource_degradation_does_not_suppress_native_math() -> None:
    degradation = ResourceDegradation(
        "eq:one",
        FORMULA_FALLBACK_IMAGE_UNAVAILABLE,
        "Formula fallback image is unavailable.",
        "[FORMULA_FALLBACK_IMAGE_UNAVAILABLE 公式图像备选不可用]",
        "private/fallback.png",
    )
    plan = build_longform_plan(
        _semantic([_formula()]),
        _preflight(degradations=(degradation,)),
    )

    args = _ops(plan, "writer.add_equation")[0]["args"]
    assert "nativeMath" in args["content"]
    assert set(args["fallbackResource"]) == {"fallbackResourcePlannedDegradation"}
    descriptor = args["fallbackResource"]["fallbackResourcePlannedDegradation"]
    assert descriptor["code"] == FORMULA_FALLBACK_IMAGE_UNAVAILABLE
    assert "private/fallback.png" not in str(descriptor)


def test_preflight_invalid_formula_emits_planned_content_degradation_with_shell() -> None:
    plan = build_longform_plan(_semantic([_formula(valid=False)]), _preflight())

    equation = _ops(plan, "writer.add_equation")[0]
    assert set(equation["args"]["content"]) == {"plannedDegradation"}
    assert equation["args"]["content"]["plannedDegradation"]["code"] == "FORMULA_MALFORMED"
    assert equation["args"]["bookmarkName"] == "wpsc_eq_" + "a" * 24
    assert equation["args"]["fallbackText"] == "x+y"
    assert "source" not in equation["args"]


def test_multiline_formula_source_is_allowed_only_as_readable_fallback() -> None:
    source = "x +\n y"
    formula = _formula(valid=False)
    formula.source = source
    formula.raw_source = source
    plan = build_longform_plan(_semantic([formula]), _preflight())

    equation = _ops(plan, "writer.add_equation")[0]
    assert equation["args"]["fallbackText"] == source
    assert set(equation["args"]["content"]) == {"plannedDegradation"}
    assert validate_generation_plan(plan.to_dict(), "writer") == plan


def test_literal_reference_citation_and_degradation_runs_preserve_order_recursively() -> None:
    citation = CitationRun("para:one/cite:1", "ref:a", "ref:a", 1, "[1]")
    missing = InlineDegradationRun(
        "para:one/cite:2",
        "REFERENCE_UNRESOLVED",
        "[REFERENCE_UNRESOLVED 引用目标未解析]",
    )
    paragraph = Paragraph(
        [
            Span("before "),
            Span("[1]", citation=citation),
            Span(" middle "),
            Span(missing.fallback_text, inline_degradation=missing),
            Span(" after"),
        ],
        node_id="para:one",
    )
    quote_paragraph = Paragraph(
        [Span("quote "), Span("[1]", citation=CitationRun("quote:p/cite:1", "ref:a", "ref:a", 1, "[1]"))],
        node_id="quote:p",
    )
    break_paragraph = Paragraph(
        [Span("break "), Span("[1]", citation=CitationRun("break:p/cite:1", "ref:a", "ref:a", 1, "[1]"))],
        node_id="break:p",
    )
    plan = build_longform_plan(
        _semantic(
            [paragraph, BlockQuote([quote_paragraph]), PageBreakBlock("break:one", [break_paragraph])],
            [_bibliography()],
        ),
        _preflight(),
    )

    paragraphs = _ops(plan, "writer.add_cross_reference")
    assert [item["nodeId"] for item in paragraphs] == ["para:one", "quote:p", "break:p"]
    assert [run["type"] for run in paragraphs[0]["args"]["runs"]] == [
        "text", "citation", "text", "degradation", "text"
    ]
    operations = plan.to_dict()["operations"]
    page_break = next(i for i, item in enumerate(operations) if item["op"] == "writer.add_page_break")
    break_child = next(i for i, item in enumerate(operations) if item.get("nodeId") == "break:p")
    assert page_break < break_child
    assert validate_generation_plan(plan.to_dict(), "writer") == plan


def test_structured_bibliography_and_cell_degradations_are_emitted_without_empty_blocks() -> None:
    table = SemanticTableBlock(
        identifier="tab:one",
        node_id="tab:one",
        caption="Table",
        headers=["A"],
        rows=[["[REFERENCE_UNRESOLVED 引用目标未解析]"]],
        alignments=["left"],
        cell_degradations=[
            TableCellDegradation(
                2,
                1,
                "REFERENCE_UNRESOLVED",
                "[REFERENCE_UNRESOLVED 引用目标未解析]",
            )
        ],
    )
    plan = build_longform_plan(
        _semantic([table], [_bibliography(), ReferenceListBlock(node_id="bib:empty")]),
        _preflight(),
    )

    table_args = _ops(plan, "writer.add_semantic_table")[0]["args"]
    assert table_args["cellDegradations"] == [{
        "row": 2,
        "column": 1,
        "code": "REFERENCE_UNRESOLVED",
        "fallbackText": "[REFERENCE_UNRESOLVED 引用目标未解析]",
    }]
    bibliography = _ops(plan, "writer.add_bibliography")
    assert len(bibliography) == 1
    assert bibliography[0]["args"] == {
        "schemaVersion": 1,
        "entries": [
            {"id": "ref:a", "nodeId": "ref:a", "number": 1, "text": "Alpha.", "cited": True},
            {"id": "ref:b", "nodeId": "ref:b", "number": 2, "text": "Beta.", "cited": False},
        ],
        "style": "numeric",
        "hangingIndentPt": 18.0,
        "leftIndentPt": 18.0,
        "spaceAfterPt": 6.0,
    }


def test_quality_anchor_is_always_early_and_contains_initial_document_issues() -> None:
    issue = DocumentIssue("CONFIG_VALUE_INVALID", "Invalid configuration.", "document")
    plan = build_longform_plan(_semantic([], issues=(issue,)), _preflight()).to_dict()
    operations = plan["operations"]

    anchors = [item for item in operations if item["op"] == "writer.reserve_document_quality_anchor"]
    assert len(anchors) == 1
    assert anchors[0] == {
        "op": "writer.reserve_document_quality_anchor",
        "nodeId": "doc:quality",
        "args": {
            "title": "生成质量提示",
            "notices": [{
                "code": "CONFIG_VALUE_INVALID",
                "message": "Invalid configuration.",
                "fallbackText": "CONFIG_VALUE_INVALID",
                "placement": "document",
            }],
        },
        "failurePolicy": {"mode": "fail"},
    }
    anchor_index = operations.index(anchors[0])
    first_body = next(i for i, item in enumerate(operations) if item["op"] == "writer.configure_section")
    assert anchor_index < first_body
    assert not _ops(build_longform_plan(_semantic([]), _preflight()), "writer.add_document_quality_notice")


def test_malformed_bibliography_degradation_keeps_its_relative_position() -> None:
    citation = CitationRun("para:one/cite:1", "ref:a", "ref:a", 1, "[1]")
    paragraph = Paragraph(
        [Span("See "), Span("[1]", citation=citation)],
        node_id="para:one",
    )
    bibliography = ReferenceListBlock(
        node_id="bib:block",
        resolved_items=[
            BibliographyEntry("ref:a", "ref:a", "Alpha.", 1, 1, True),
            DegradationBlock(
                DocumentIssue(
                    "BIBLIOGRAPHY_ENTRY_MALFORMED",
                    "Malformed bibliography entry.",
                    "block",
                ),
                "ref:malformed",
                "[BIBLIOGRAPHY_ENTRY_MALFORMED 参考文献条目格式错误]",
            ),
            BibliographyEntry("ref:b", "ref:b", "Beta.", 2, 3, False),
        ],
    )
    plan = build_longform_plan(
        _semantic([paragraph], [bibliography]),
        _preflight(),
    )

    ordered = [
        item
        for item in plan.to_dict()["operations"]
        if item["op"] in {
            "writer.add_bibliography", "writer.add_degradation_notice"
        }
    ]
    assert [item["op"] for item in ordered] == [
        "writer.add_bibliography",
        "writer.add_degradation_notice",
        "writer.add_bibliography",
    ]
    assert ordered[0]["args"]["entries"][0]["number"] == 1
    assert ordered[1]["nodeId"] == "ref:malformed"
    assert ordered[2]["args"]["entries"][0]["number"] == 2
    assert validate_generation_plan(plan.to_dict(), "writer") == plan


def test_recording_composer_mirrors_m4_and_preserves_legacy_shapes() -> None:
    descriptor = convert_restricted_latex("x+y")
    with RecordingWriterComposer() as composer:
        composer.reserve_document_quality_anchor()
        composer.configure_section(role="body")
        composer.add_equation(
            node_id="eq:one",
            render_mode="native-m4",
            content={"nativeMath": {
                "syntax": descriptor.syntax,
                "linearText": descriptor.linear_text,
                "sourceHash": descriptor.source_hash,
            }},
        )
        composer.configure_section(role="bibliography")
        composer.add_bibliography(
            node_id="bib:block",
            schema_version=1,
            entries=[{
                "id": "ref:a",
                "nodeId": "ref:a",
                "number": 1,
                "text": "Alpha.",
                "cited": False,
            }],
            style="numeric",
            hanging_indent_pt=18.0,
            left_indent_pt=18.0,
            space_after_pt=6.0,
            failure_policy={
                "mode": "degrade",
                "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
                "fallback": "notice",
            },
        )
        composer.finalize_fields()
        m4 = composer.save_docx("ignored.docx").plan.to_dict()
    equation = next(item for item in m4["operations"] if item["op"] == "writer.add_equation")
    assert equation["args"]["renderMode"] == "native-m4"

    with RecordingWriterComposer() as composer:
        composer.add_equation(node_id="eq:legacy", source="x+y")
        composer.add_bibliography(
            node_id="bib:legacy",
            entries=["Legacy entry"],
            style="numbered",
        )
        composer.finalize_fields()
        legacy = composer.save_docx("ignored.docx").plan.to_dict()
    legacy_equation = next(item for item in legacy["operations"] if item["op"] == "writer.add_equation")
    legacy_bibliography = next(item for item in legacy["operations"] if item["op"] == "writer.add_bibliography")
    assert set(legacy_equation["args"]) == {
        "source", "numbering", "bookmarkName", "fallbackText"
    }
    assert legacy_bibliography["args"] == {
        "entries": ["Legacy entry"], "style": "numbered"
    }
