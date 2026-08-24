from __future__ import annotations

import copy

import pytest

from skills.WPSComposer.scripts.generation_plan import OperationPlanError, validate_generation_plan


BOOKMARK = "wpsc_fig_" + "a" * 24
NUMBERING = {
    "mode": "chapter",
    "sequenceId": "WPSC_FIG",
    "chapterStyleLevel": 1,
    "resetLevel": 1,
    "prefix": "图 ",
    "suffix": "",
}


def _plan(operation: dict) -> dict:
    return {
        "protocolVersion": 2,
        "semanticVersion": "longform-1",
        "component": "writer",
        "resourceManifestVersion": 1,
        "resourceManifestDigest": "sha256:" + "0" * 64,
        "operations": [
            operation,
            {"op": "writer.finalize_fields", "nodeId": "doc:finalize", "args": {"maxRounds": 3}},
        ],
    }


def _figure() -> dict:
    return {
        "op": "writer.add_captioned_figure",
        "nodeId": "fig:one",
        "args": {
            "caption": "示例",
            "numbering": copy.deepcopy(NUMBERING),
            "bookmarkName": BOOKMARK,
            "indexable": True,
            "referenceable": True,
            "widthMode": "full",
            "orientation": "portrait",
            "kind": "diagram",
            "children": [{
                "nodeId": "fig:one/image:1",
                "resourceId": "image-1",
                "displayWidthPt": 240.0,
                "displayHeightPt": 120.0,
                "effectiveDpi": 144.0,
                "mediaType": "image/png",
                "normalizerId": "none-v1",
            }],
            "layout": "stack",
            "keepWithCaption": True,
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["IMAGE_INSERT_FAILED"],
            "fallback": "figure-child-stack-then-notice",
        },
    }


def _table() -> dict:
    return {
        "op": "writer.add_semantic_table",
        "nodeId": "tab:one",
        "args": {
            "caption": "表格",
            "numbering": {**copy.deepcopy(NUMBERING), "sequenceId": "WPSC_TAB", "prefix": "表 "},
            "bookmarkName": "wpsc_tab_" + "b" * 24,
            "indexable": True,
            "referenceable": True,
            "headers": ["A", "B"],
            "rows": [["1", "2"]],
            "alignments": ["left", "right"],
            "style": "three-line",
            "orientation": "portrait",
            "borderSpec": {
                "top": 1.5, "bottom": 1.5, "headerBottom": 0.75,
                "left": 0.0, "right": 0.0,
                "insideHorizontal": 0.0, "insideVertical": 0.0,
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
                "TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED",
                "TABLE_ROW_FORCED_SPLIT", "TABLE_INSERT_FAILED",
            ],
            "fallback": "grid-then-text",
        },
    }


def test_accepts_complete_closed_figure_and_table() -> None:
    validate_generation_plan(_plan(_figure()), "writer")
    validate_generation_plan(_plan(_table()), "writer")


@pytest.mark.parametrize(
    ("mutate", "match"),
    [
        (lambda op: op["args"]["children"].extend([copy.deepcopy(op["args"]["children"][0]), copy.deepcopy(op["args"]["children"][0])]), "children"),
        (lambda op: op["args"].update(layout="columns"), "columns"),
        (lambda op: op["args"]["children"][0].update(displayWidthPt=float("inf")), "JSON-compatible|displayWidthPt"),
        (lambda op: op["args"]["children"][0].update(path="/tmp/secret.png"), "unknown argument|path"),
        (lambda op: op["args"]["numbering"].update(sequenceId="SEQ USER"), "sequenceId"),
        (lambda op: op["args"]["numbering"].update(fieldCode="SEQ EVIL"), "unknown argument"),
        (lambda op: op.update(failurePolicy={"mode": "degrade", "recoverableCodes": ["OTHER"], "fallback": "notice"}), "failurePolicy"),
    ],
)
def test_rejects_invalid_or_injectable_figure_nesting(mutate, match: str) -> None:
    op = _figure()
    mutate(op)
    with pytest.raises(OperationPlanError, match=match):
        validate_generation_plan(_plan(op), "writer")


def test_columns_require_exactly_two_children_and_columns_two() -> None:
    op = _figure()
    op["args"]["layout"] = "columns"
    op["args"]["columns"] = 2
    op["args"]["children"].append({**op["args"]["children"][0], "nodeId": "fig:one/image:2", "resourceId": "image-2"})
    validate_generation_plan(_plan(op), "writer")

    op["args"]["columns"] = 1
    with pytest.raises(OperationPlanError, match="columns"):
        validate_generation_plan(_plan(op), "writer")


@pytest.mark.parametrize("bookmark", ["fig_bad", "wpsc_fig_ABC", "wpsc_tab_" + "a" * 24])
def test_figure_bookmark_regex_and_kind_are_strict(bookmark: str) -> None:
    op = _figure()
    op["args"]["bookmarkName"] = bookmark
    with pytest.raises(OperationPlanError, match="bookmarkName"):
        validate_generation_plan(_plan(op), "writer")


def test_table_requires_rectangular_data_and_legal_merges() -> None:
    op = _table()
    op["args"]["rows"] = [["1"]]
    with pytest.raises(OperationPlanError, match="rectangular|rows"):
        validate_generation_plan(_plan(op), "writer")

    op = _table()
    op["args"]["merges"] = [{"top": 1, "left": 1, "bottom": 3, "right": 2}]
    with pytest.raises(OperationPlanError, match="merges"):
        validate_generation_plan(_plan(op), "writer")


def test_table_policy_allowlist_is_exact() -> None:
    op = _table()
    op["failurePolicy"]["recoverableCodes"] = ["TABLE_INSERT_FAILED"]
    with pytest.raises(OperationPlanError, match="failurePolicy"):
        validate_generation_plan(_plan(op), "writer")


def test_reference_runs_are_closed_resolved_and_ordered() -> None:
    operation = {
        "op": "writer.add_cross_reference",
        "nodeId": "para:one",
        "args": {"runs": [
            {"type": "text", "text": "见"},
            {"type": "reference", "targetNodeId": "fig:one", "targetKind": "figure", "bookmarkName": BOOKMARK, "prefix": "图 ", "suffix": "", "fallbackText": "[图]"},
            {"type": "text", "text": "。"},
        ]},
        "failurePolicy": {"mode": "degrade", "recoverableCodes": ["CROSS_REFERENCE_FAILED"], "fallback": "inline-fallback"},
    }
    raw = _plan(operation)
    raw["operations"].insert(0, _figure())
    validate_generation_plan(raw, "writer")
    operation["args"]["runs"][1]["fieldCode"] = "REF evil \\h"
    with pytest.raises(OperationPlanError, match="unknown argument"):
        validate_generation_plan(raw, "writer")


def test_reference_list_paragraph_formatting_is_closed_and_allows_plain_items() -> None:
    operation = {
        "op": "writer.add_cross_reference",
        "nodeId": "para:list-item",
        "args": {
            "runs": [{"type": "text", "text": "•\tPlain companion"}],
            "listFormatting": {"kind": "bullet", "indentPt": 24.0},
        },
        "failurePolicy": {"mode": "degrade", "recoverableCodes": ["CROSS_REFERENCE_FAILED"], "fallback": "inline-fallback"},
    }
    validate_generation_plan(_plan(operation), "writer")

    operation["args"]["listFormatting"]["indentPt"] = 36.0
    with pytest.raises(OperationPlanError, match="listFormatting|indentPt"):
        validate_generation_plan(_plan(operation), "writer")

    operation["args"]["listFormatting"] = {"kind": "evil", "indentPt": 24.0}
    with pytest.raises(OperationPlanError, match="listFormatting|kind"):
        validate_generation_plan(_plan(operation), "writer")


def test_utf16_string_bound_is_enforced() -> None:
    operation = {
        "op": "writer.add_cross_reference",
        "nodeId": "para:one",
        "args": {"runs": [{"type": "text", "text": "😀" * 50_001}]},
        "failurePolicy": {"mode": "degrade", "recoverableCodes": ["CROSS_REFERENCE_FAILED"], "fallback": "inline-fallback"},
    }
    with pytest.raises(OperationPlanError, match="UTF-16|string"):
        validate_generation_plan(_plan(operation), "writer")


def test_sectioned_m3_plan_enforces_index_ownership_and_finalize_order() -> None:
    figure = _figure()
    index = {
        "op": "writer.insert_figure_index",
        "nodeId": "doc:figure-index",
        "args": {"title": "图目录", "sequenceId": "WPSC_FIG", "titleStyleId": "WPSC_INDEX_TITLE"},
    }
    section = {
        "op": "writer.configure_section",
        "nodeId": "doc:section-body",
        "args": {"role": "body"},
    }
    front_section = {
        "op": "writer.configure_section",
        "nodeId": "doc:section-front",
        "args": {"role": "front_matter"},
    }
    finalize = {"op": "writer.finalize_fields", "nodeId": "doc:finalize", "args": {"maxRounds": 3}}
    raw = _plan(section)
    raw["operations"] = [section, figure, index, finalize]
    with pytest.raises(OperationPlanError, match="indexes must precede"):
        validate_generation_plan(raw, "writer")

    raw["operations"] = [
        front_section,
        index,
        section,
        figure,
        copy.deepcopy(figure),
        finalize,
    ]
    with pytest.raises(OperationPlanError, match="owned more than once"):
        validate_generation_plan(raw, "writer")

    raw["operations"] = [section, index, figure, finalize, {"op": "writer.add_page_break", "args": {}}]
    with pytest.raises(OperationPlanError, match="exactly once and last"):
        validate_generation_plan(raw, "writer")
