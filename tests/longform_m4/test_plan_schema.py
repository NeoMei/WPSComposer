from __future__ import annotations

import copy

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    OperationPlanError,
    validate_generation_plan,
)


_DIGEST = "sha256:" + "0" * 64
_SOURCE_HASH = "1" * 64


def _numbering() -> dict:
    return {
        "mode": "global",
        "sequenceId": "WPSC_EQ",
        "chapterStyleLevel": None,
        "resetLevel": None,
        "prefix": "(",
        "suffix": ")",
    }


def _degradation(code: str, *, placement: str = "block") -> dict:
    fallback_kind = "source" if code.startswith("FORMULA_") else "text"
    return {
        "code": code,
        "placement": placement,
        "objectLabel": "formula" if placement == "block" else "citation",
        "reason": "The native content cannot be emitted.",
        "fallbackText": "x+y" if placement == "block" else "[REFERENCE_UNRESOLVED 引用目标未解析]",
        "fallbackKind": fallback_kind,
    }


def _m4_equation() -> dict:
    return {
        "op": "writer.add_equation",
        "nodeId": "eq:one",
        "args": {
            "renderMode": "native-m4",
            "content": {
                "nativeMath": {
                    "syntax": "wps-linear-v1",
                    "linearText": "x+y",
                    "sourceHash": _SOURCE_HASH,
                }
            },
            "numbering": _numbering(),
            "bookmarkName": "wpsc_eq_" + "a" * 24,
            "fallbackText": "x+y",
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["EQUATION_INSERT_FAILED"],
            "fallback": "explicit-image-then-source-notice",
        },
    }


def _structured_bibliography() -> dict:
    return {
        "op": "writer.add_bibliography",
        "nodeId": "bib:block",
        "args": {
            "schemaVersion": 1,
            "entries": [
                {
                    "id": "ref:a",
                    "nodeId": "ref:a",
                    "number": 1,
                    "text": "Author. Title.",
                    "cited": True,
                }
            ],
            "style": "numeric",
            "hangingIndentPt": 18.0,
            "leftIndentPt": 18.0,
            "spaceAfterPt": 6.0,
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
            "fallback": "notice",
        },
    }


def _citation_paragraph() -> dict:
    return {
        "op": "writer.add_cross_reference",
        "nodeId": "para:one",
        "args": {
            "runs": [
                {"type": "text", "text": "See "},
                {
                    "type": "citation",
                    "nodeId": "para:one/cite:1",
                    "targetId": "ref:a",
                    "targetNodeId": "ref:a",
                    "number": 1,
                    "fallbackText": "[1]",
                },
            ]
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
    }


def _plan() -> dict:
    return {
        "component": "writer",
        "protocolVersion": 2,
        "semanticVersion": "longform-1",
        "resourceManifestVersion": 1,
        "resourceManifestDigest": _DIGEST,
        "operations": [
            {"op": "writer.reset", "args": {}},
            {
                "op": "writer.reserve_document_quality_anchor",
                "nodeId": "doc:quality",
                "args": {"title": "生成质量提示", "notices": []},
                "failurePolicy": {"mode": "fail"},
            },
            {
                "op": "writer.configure_section",
                "nodeId": "doc:section-1:body",
                "args": {"role": "body"},
            },
            _citation_paragraph(),
            _m4_equation(),
            {
                "op": "writer.configure_section",
                "nodeId": "doc:section-2:bibliography",
                "args": {"role": "bibliography"},
            },
            _structured_bibliography(),
            {
                "op": "writer.finalize_fields",
                "nodeId": "doc:finalize",
                "args": {"maxRounds": 3},
            },
        ],
    }


def test_closed_m4_formula_citation_bibliography_and_anchor_validate() -> None:
    plan = _plan()

    assert validate_generation_plan(plan, "writer").to_dict() == plan


def test_m3_equation_shell_and_legacy_bibliography_shapes_remain_unchanged() -> None:
    plan = _plan()
    plan["operations"].remove(_m4_equation())
    plan["operations"].remove(_structured_bibliography())
    plan["operations"].remove(_citation_paragraph())
    plan["operations"].insert(
        -2,
        {
            "op": "writer.add_equation",
            "nodeId": "eq:legacy",
            "args": {
                "source": "x+y",
                "numbering": _numbering(),
                "bookmarkName": "wpsc_eq_" + "b" * 24,
                "fallbackText": "x+y",
            },
            "failurePolicy": {"mode": "fail"},
        },
    )
    plan["operations"].insert(
        -1,
        {
            "op": "writer.add_bibliography",
            "nodeId": "bib:legacy",
            "args": {"entries": ["Legacy entry"], "style": "numbered"},
        },
    )

    validated = validate_generation_plan(plan, "writer").to_dict()
    equation = next(op for op in validated["operations"] if op.get("nodeId") == "eq:legacy")
    bibliography = next(op for op in validated["operations"] if op.get("nodeId") == "bib:legacy")
    assert equation["args"] == {
        "source": "x+y",
        "numbering": _numbering(),
        "bookmarkName": "wpsc_eq_" + "b" * 24,
        "fallbackText": "x+y",
    }
    assert bibliography["args"] == {"entries": ["Legacy entry"], "style": "numbered"}


@pytest.mark.parametrize(
    "mutate",
    [
        lambda op: op["args"]["content"].update({"plannedDegradation": _degradation("FORMULA_MALFORMED")}),
        lambda op: op["args"].update({"content": {}}),
        lambda op: op["args"]["content"]["nativeMath"].update({"source": r"\\frac{x}{y}"}),
        lambda op: op["args"]["content"]["nativeMath"].update({"sourceHash": "A" * 64}),
        lambda op: op["args"].update({"sourcePath": "/private/formula.tex"}),
        lambda op: op["args"].update({"payload": "data:image/png;base64,AAAA"}),
    ],
)
def test_native_equation_is_closed_and_has_exact_content_one_of(mutate) -> None:
    plan = _plan()
    equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
    mutate(equation)

    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


def test_formula_fallback_resource_is_an_orthogonal_closed_one_of() -> None:
    for fallback_resource in (
        {"fallbackResourceId": "wpsc-rsrc:opaque"},
        {
            "fallbackResourcePlannedDegradation": {
                **_degradation("FORMULA_FALLBACK_IMAGE_UNAVAILABLE"),
                "fallbackKind": "none",
            }
        },
    ):
        plan = _plan()
        equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
        equation["args"]["fallbackResource"] = fallback_resource
        assert validate_generation_plan(plan, "writer")

    plan = _plan()
    equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
    equation["args"]["fallbackResource"] = {
        "fallbackResourceId": "wpsc-rsrc:opaque",
        "fallbackResourcePlannedDegradation": {
            **_degradation("FORMULA_FALLBACK_IMAGE_UNAVAILABLE"),
            "fallbackKind": "none",
        },
    }
    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


def test_planned_formula_content_suppresses_native_math_but_keeps_shell() -> None:
    plan = _plan()
    equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
    equation["args"]["content"] = {
        "plannedDegradation": _degradation("FORMULA_MALFORMED")
    }

    assert validate_generation_plan(plan, "writer")


def test_citation_and_inline_degradation_runs_are_strict() -> None:
    plan = _plan()
    paragraph = next(op for op in plan["operations"] if op["op"] == "writer.add_cross_reference")
    paragraph["args"]["runs"].append({
        "type": "degradation",
        "nodeId": "para:one/degradation:1",
        "code": "REFERENCE_UNRESOLVED",
        "fallbackText": "[REFERENCE_UNRESOLVED 引用目标未解析]",
    })
    assert validate_generation_plan(plan, "writer")

    for field, value in (("number", 0), ("fieldCode", "REF secret"), ("path", "/tmp/x")):
        broken = copy.deepcopy(plan)
        citation = next(
            run
            for op in broken["operations"]
            if op["op"] == "writer.add_cross_reference"
            for run in op["args"]["runs"]
            if run["type"] == "citation"
        )
        citation[field] = value
        with pytest.raises(OperationPlanError):
            validate_generation_plan(broken, "writer")


def test_cell_degradations_are_bounded_to_their_table_grid() -> None:
    plan = _plan()
    table = {
        "op": "writer.add_semantic_table",
        "nodeId": "tab:one",
        "args": {
            "caption": "T",
            "numbering": {**_numbering(), "sequenceId": "WPSC_TAB", "prefix": "表 ", "suffix": ""},
            "bookmarkName": "wpsc_tab_" + "c" * 24,
            "indexable": True,
            "referenceable": True,
            "headers": ["A"],
            "rows": [["[REFERENCE_UNRESOLVED 引用目标未解析]"]],
            "alignments": ["left"],
            "style": "grid",
            "orientation": "portrait",
            "borderSpec": {key: 0.75 for key in ("top", "bottom", "headerBottom", "left", "right", "insideHorizontal", "insideVertical")},
            "merges": [],
            "repeatHeader": True,
            "allowRowSplit": False,
            "cellIndentPt": 0.0,
            "plannedDegradation": [],
            "cellDegradations": [{"row": 2, "column": 1, "code": "REFERENCE_UNRESOLVED", "fallbackText": "[REFERENCE_UNRESOLVED 引用目标未解析]"}],
            "keepCaptionWithFirstRow": True,
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED", "TABLE_ROW_FORCED_SPLIT", "TABLE_INSERT_FAILED"],
            "fallback": "grid-then-text",
        },
    }
    plan["operations"].insert(4, table)
    assert validate_generation_plan(plan, "writer")
    table["args"]["cellDegradations"][0]["row"] = 3
    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


def test_structured_bibliography_requires_unique_gap_free_order_and_fixed_geometry() -> None:
    for mutator in (
        lambda args: args["entries"].append(dict(args["entries"][0])),
        lambda args: args["entries"][0].update({"number": 2}),
        lambda args: args.update({"hangingIndentPt": -18.0}),
        lambda args: args.update({"style": "apa"}),
        lambda args: args["entries"][0].update({"unknown": True}),
    ):
        plan = _plan()
        bibliography = next(op for op in plan["operations"] if op["op"] == "writer.add_bibliography")
        mutator(bibliography["args"])
        with pytest.raises(OperationPlanError):
            validate_generation_plan(plan, "writer")


def test_state_matches_citations_and_bibliography_and_requires_backmatter() -> None:
    for mutator in (
        lambda plan: next(op for op in plan["operations"] if op["op"] == "writer.add_cross_reference")["args"]["runs"][1].update({"number": 2}),
        lambda plan: next(op for op in plan["operations"] if op["op"] == "writer.add_cross_reference")["args"]["runs"][1].update({"targetNodeId": "ref:other"}),
        lambda plan: next(op for op in plan["operations"] if op.get("nodeId") == "doc:section-2:bibliography")["args"].update({"role": "body"}),
    ):
        plan = _plan()
        mutator(plan)
        with pytest.raises(OperationPlanError):
            validate_generation_plan(plan, "writer")


def test_table_static_citation_can_be_the_only_cited_occurrence() -> None:
    plan = _plan()
    plan["operations"].remove(_citation_paragraph())
    table = {
        "op": "writer.add_semantic_table",
        "nodeId": "tab:static-citation",
        "args": {
            "caption": "T",
            "numbering": {
                **_numbering(),
                "sequenceId": "WPSC_TAB",
                "prefix": "表 ",
                "suffix": "",
            },
            "bookmarkName": "wpsc_tab_" + "d" * 24,
            "indexable": True,
            "referenceable": True,
            "headers": ["Citation"],
            "rows": [["[1]"]],
            "alignments": ["left"],
            "style": "grid",
            "orientation": "portrait",
            "borderSpec": {
                key: 0.75
                for key in (
                    "top", "bottom", "headerBottom", "left", "right",
                    "insideHorizontal", "insideVertical",
                )
            },
            "merges": [],
            "repeatHeader": True,
            "allowRowSplit": False,
            "cellIndentPt": 0.0,
            "plannedDegradation": [],
            "keepCaptionWithFirstRow": True,
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": [
                "TABLE_STYLE_APPLY_FAILED",
                "TABLE_MERGE_APPLY_FAILED",
                "TABLE_ROW_FORCED_SPLIT",
                "TABLE_INSERT_FAILED",
            ],
            "fallback": "grid-then-text",
        },
    }
    plan["operations"].insert(3, table)

    assert validate_generation_plan(plan, "writer")


def test_quality_anchor_is_unique_early_fail_hard_and_replaces_late_notice() -> None:
    for mutator in (
        lambda plan: plan["operations"].insert(2, copy.deepcopy(plan["operations"][1])),
        lambda plan: plan["operations"][1].update({"failurePolicy": {"mode": "degrade", "recoverableCodes": ["DEGRADATION_INSERT_FAILED"], "fallback": "notice"}}),
        lambda plan: plan["operations"].append(plan["operations"].pop(1)),
        lambda plan: plan["operations"].insert(-1, {"op": "writer.add_document_quality_notice", "nodeId": "doc:quality:late", "args": {"notices": []}}),
    ):
        plan = _plan()
        mutator(plan)
        with pytest.raises(OperationPlanError):
            validate_generation_plan(plan, "writer")
