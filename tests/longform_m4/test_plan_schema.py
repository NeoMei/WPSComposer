from __future__ import annotations

import copy
import hashlib
import json

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    OperationPlanError,
    validate_generation_plan,
)


_DIGEST = "sha256:" + "0" * 64
_SOURCE_HASH = "1" * 64


def _cell_citation_node_id(table: str, row: int, column: int, target: str) -> str:
    payload = json.dumps([table, row, column, target], ensure_ascii=False, separators=(",", ":"))
    digest = hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]
    return f"{table}/cell:{row}:{column}/cite:{digest}"


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


@pytest.mark.parametrize("linear_text", [
    r"\input{/Users/alice/private.tex}", r"\write18{calc.exe}", r"\madeup{x}",
    "/Users/alice/private.tex", r"C:\private\input.tex", "source=file:///tmp/private.tex",
    "secrets/private.tex", "assets/../secret/x.tex",
])
def test_native_math_linear_text_is_independently_trusted(linear_text) -> None:
    plan = _plan()
    equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
    equation["args"]["content"]["nativeMath"]["linearText"] = linear_text
    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


@pytest.mark.parametrize("linear_text", ["x)", "(x", "{■(x&y)", "([x)]"])
def test_native_math_linear_text_preserves_valid_asymmetric_delimiters(linear_text) -> None:
    plan = _plan()
    equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
    equation["args"]["content"]["nativeMath"]["linearText"] = linear_text
    assert validate_generation_plan(plan, "writer")


@pytest.mark.parametrize("fallback_text", [
    "path:/Users/alice/private.tex", r"path:C:\private\input.tex",
    "secrets/private.tex", "assets/../secret/x.tex", "A" * 76,
])
def test_native_equation_fallback_text_is_privacy_safe(fallback_text) -> None:
    plan = _plan()
    equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
    equation["args"]["fallbackText"] = fallback_text
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


@pytest.mark.parametrize(
    "private_text",
    [
        "/Users/alice/private/input.tex",
        r"C:\private\input.tex",
        "0" * 64,
        "ValueError('private input')",
        "data:image/png;base64,QUJDRA==",
        "path:/Users/alice/private.tex", "source=file:///Users/alice/private.tex",
        r"path:C:\private\input.tex", r"source=\\server\share\input.tex",
        "path:/home/alice/input.tex", "path:/tmp/input.tex", "path:~/input.tex",
        "secrets/private.tex", "assets/../secret/x.tex",
        "A" * 76, "A" * 76 + "\n" + "B" * 76, "blob:" + "A" * 76,
    ],
)
@pytest.mark.parametrize("field", ["reason", "fallbackText"])
def test_planned_formula_degradation_rejects_private_diagnostics(
    private_text, field
) -> None:
    plan = _plan()
    equation = next(op for op in plan["operations"] if op["op"] == "writer.add_equation")
    equation["args"]["content"] = {
        "plannedDegradation": _degradation("FORMULA_MALFORMED")
    }
    equation["args"]["content"]["plannedDegradation"][field] = private_text

    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


@pytest.mark.parametrize(
    "private_text",
    [
        "/Users/alice/private/input.tex",
        r"C:\private\input.tex",
        "0" * 64,
        "RuntimeError('private input')",
        "data:text/plain;base64,QUJDRA==",
        "path:/Users/alice/private.tex", "source=file:///Users/alice/private.tex",
        r"path:C:\private\input.tex", r"source=\\server\share\input.tex",
        "path:/home/alice/input.tex", "path:/tmp/input.tex", "path:~/input.tex",
        "secrets/private.tex", "assets/../secret/x.tex",
        "A" * 76, "A" * 76 + "\n" + "B" * 76, "blob:" + "A" * 76,
    ],
)
@pytest.mark.parametrize("field", ["message", "fallbackText"])
def test_quality_anchor_rejects_private_diagnostics(private_text, field) -> None:
    plan = _plan()
    anchor = next(
        op for op in plan["operations"]
        if op["op"] == "writer.reserve_document_quality_anchor"
    )
    notice = {
        "code": "CONFIG_VALUE_INVALID",
        "message": "Invalid configuration.",
        "fallbackText": "CONFIG_VALUE_INVALID",
        "placement": "document",
    }
    notice[field] = private_text
    anchor["args"]["notices"] = [notice]

    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


def test_quality_anchor_rejects_duplicate_stable_notice_keys() -> None:
    plan = _plan()
    anchor = next(
        op for op in plan["operations"]
        if op["op"] == "writer.reserve_document_quality_anchor"
    )
    notice = {
        "code": "CONFIG_VALUE_INVALID",
        "message": "Invalid configuration.",
        "fallbackText": "CONFIG_VALUE_INVALID",
        "placement": "document",
    }
    anchor["args"]["notices"] = [notice, dict(notice)]

    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


def test_quality_anchor_rejects_a_hash_disguised_as_notice_code() -> None:
    plan = _plan()
    anchor = next(
        op for op in plan["operations"]
        if op["op"] == "writer.reserve_document_quality_anchor"
    )
    anchor["args"]["notices"] = [{
        "code": "A" * 64,
        "message": "Invalid configuration.",
        "fallbackText": "controlled",
        "placement": "document",
    }]

    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


def test_privacy_checks_do_not_reject_short_text_or_logical_node_slashes() -> None:
    plan = _plan()
    citation = next(
        run for op in plan["operations"] if op["op"] == "writer.add_cross_reference"
        for run in op["args"]["runs"] if run["type"] == "citation"
    )
    bibliography = next(op for op in plan["operations"] if op["op"] == "writer.add_bibliography")
    citation["targetNodeId"] = "scheme:logical/entry:1"
    bibliography["args"]["entries"][0]["nodeId"] = "scheme:logical/entry:1"
    anchor = next(op for op in plan["operations"] if op["op"] == "writer.reserve_document_quality_anchor")
    anchor["args"]["notices"] = [{
        "code": "SAFE_NOTICE", "message": "abc123 and/or ratio a/b",
        "fallbackText": "short-safe", "placement": "document",
    }]
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


@pytest.mark.parametrize("node_id", [
    "/leading/node", "C:/private/node", r"C:\private\node",
    "file:///home/alice/node", "https://example.test/node", "//server/share/node", "~/node",
])
def test_m4_node_ids_reject_absolute_paths_and_uri_disguises(node_id) -> None:
    plan = _plan()
    citation = next(
        run for op in plan["operations"] if op["op"] == "writer.add_cross_reference"
        for run in op["args"]["runs"] if run["type"] == "citation"
    )
    citation["targetNodeId"] = node_id
    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


@pytest.mark.parametrize(
    "op_name,node_id",
    [
        ("writer.add_cross_reference", "/Users/alice/body.md"),
        ("writer.add_equation", "file:///Users/alice/formula.tex"),
        ("writer.configure_section", r"C:\private\section"),
    ],
)
def test_m4_top_level_semantic_and_section_node_ids_are_privacy_safe(
    op_name, node_id
) -> None:
    plan = _plan()
    operation = next(op for op in plan["operations"] if op["op"] == op_name)
    operation["nodeId"] = node_id
    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


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
    table["args"]["cellDegradations"][0]["row"] = 2
    table["args"]["rows"] = [["fallback was lost"]]
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


@pytest.mark.parametrize(
    "text",
    ["Author.\nTitle.", "Author.\rTitle.", "Author.\tTitle.", "Author.\u2029Title."],
)
def test_structured_bibliography_entry_text_is_one_trimmed_paragraph(text) -> None:
    plan = _plan()
    bibliography = next(op for op in plan["operations"] if op["op"] == "writer.add_bibliography")
    bibliography["args"]["entries"][0]["text"] = text

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


def test_explicit_table_cell_citation_can_be_the_only_cited_occurrence() -> None:
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
            "cellCitations": [{
                "row": 2,
                "column": 1,
                "nodeId": _cell_citation_node_id("tab:static-citation", 2, 1, "ref:a"),
                "targetId": "ref:a",
                "targetNodeId": "ref:a",
                "number": 1,
                "fallbackText": "[1]",
            }],
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


@pytest.mark.parametrize(
    "field,value",
    [
        ("row", 3),
        ("column", 2),
        ("number", 2),
        ("targetId", "ref:missing"),
        ("targetNodeId", "ref:other"),
        ("nodeId", "tab:cell-citation/cell:2:1/cite:forged"),
        ("unknown", True),
    ],
)
def test_table_cell_citation_metadata_is_closed_and_matches_grid_and_bibliography(
    field, value
) -> None:
    plan = _plan()
    plan["operations"].remove(_citation_paragraph())
    table = {
        "op": "writer.add_semantic_table",
        "nodeId": "tab:cell-citation",
        "args": {
            "caption": "T",
            "numbering": {
                **_numbering(), "sequenceId": "WPSC_TAB", "prefix": "表 ", "suffix": ""
            },
            "bookmarkName": "wpsc_tab_" + "e" * 24,
            "indexable": True,
            "referenceable": True,
            "headers": ["Citation", "Literal"],
            "rows": [["[1]", "ordinary [1] text"]],
            "alignments": ["left", "left"],
            "style": "grid",
            "orientation": "portrait",
            "borderSpec": {key: 0.75 for key in ("top", "bottom", "headerBottom", "left", "right", "insideHorizontal", "insideVertical")},
            "merges": [],
            "repeatHeader": True,
            "allowRowSplit": False,
            "cellIndentPt": 0.0,
            "plannedDegradation": [],
            "cellCitations": [{
                "row": 2,
                "column": 1,
                "nodeId": _cell_citation_node_id("tab:cell-citation", 2, 1, "ref:a"),
                "targetId": "ref:a",
                "targetNodeId": "ref:a",
                "number": 1,
                "fallbackText": "[1]",
            }],
            "keepCaptionWithFirstRow": True,
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED", "TABLE_ROW_FORCED_SPLIT", "TABLE_INSERT_FAILED"],
            "fallback": "grid-then-text",
        },
    }
    table["args"]["cellCitations"][0][field] = value
    plan["operations"].insert(3, table)

    with pytest.raises(OperationPlanError):
        validate_generation_plan(plan, "writer")


def test_plain_bracketed_number_without_cell_citation_metadata_remains_literal() -> None:
    plan = _plan()
    plan["operations"].remove(_citation_paragraph())
    bibliography = next(
        item for item in plan["operations"] if item["op"] == "writer.add_bibliography"
    )
    bibliography["args"]["entries"][0]["cited"] = False
    table = {
        "op": "writer.add_semantic_table",
        "nodeId": "tab:literal",
        "args": {
            "caption": "T",
            "numbering": {**_numbering(), "sequenceId": "WPSC_TAB", "prefix": "表 ", "suffix": ""},
            "bookmarkName": "wpsc_tab_" + "f" * 24,
            "indexable": True,
            "referenceable": True,
            "headers": ["Literal"],
            "rows": [["ordinary [999] text"]],
            "alignments": ["left"],
            "style": "grid",
            "orientation": "portrait",
            "borderSpec": {key: 0.75 for key in ("top", "bottom", "headerBottom", "left", "right", "insideHorizontal", "insideVertical")},
            "merges": [],
            "repeatHeader": True,
            "allowRowSplit": False,
            "cellIndentPt": 0.0,
            "plannedDegradation": [],
            "keepCaptionWithFirstRow": True,
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED", "TABLE_ROW_FORCED_SPLIT", "TABLE_INSERT_FAILED"],
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
        lambda plan: plan["operations"].insert(2, plan["operations"].pop(1)),
        lambda plan: plan["operations"].insert(-1, {"op": "writer.add_document_quality_notice", "nodeId": "doc:quality:late", "args": {"notices": []}}),
    ):
        plan = _plan()
        mutator(plan)
        with pytest.raises(OperationPlanError):
            validate_generation_plan(plan, "writer")
