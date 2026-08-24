from __future__ import annotations

import copy

import pytest

from skills.WPSComposer.scripts.generation_plan import OperationPlanError, validate_generation_plan


FIG_BOOKMARK = "wpsc_fig_" + "a" * 24


def _envelope(operations: list[dict], *, digest: str = "sha256:" + "0" * 64) -> dict:
    return {
        "protocolVersion": 2,
        "semanticVersion": "longform-1",
        "component": "writer",
        "resourceManifestVersion": 1,
        "resourceManifestDigest": digest,
        "operations": operations,
    }


def _finalize(max_rounds: int = 3) -> dict:
    return {"op": "writer.finalize_fields", "nodeId": "doc:finalize", "args": {"maxRounds": max_rounds}}


def _numbering(sequence="WPSC_FIG", prefix="图 ", suffix="") -> dict:
    return {
        "mode": "global", "sequenceId": sequence,
        "chapterStyleLevel": None, "resetLevel": None,
        "prefix": prefix, "suffix": suffix,
    }


def _m3_figure(node_id="fig:one", bookmark=FIG_BOOKMARK) -> dict:
    return {
        "op": "writer.add_captioned_figure", "nodeId": node_id,
        "args": {
            "caption": "Figure", "numbering": _numbering(),
            "bookmarkName": bookmark, "indexable": True, "referenceable": True,
            "widthMode": "auto", "orientation": "portrait", "kind": "diagram",
            "children": [{
                "nodeId": node_id + "/image:1", "resourceId": "image-1",
                "displayWidthPt": 100.0, "displayHeightPt": 50.0,
                "effectiveDpi": 144.0, "mediaType": "image/png",
                "normalizerId": "none-v1",
            }],
            "layout": "stack", "keepWithCaption": True,
        },
        "failurePolicy": {
            "mode": "degrade", "recoverableCodes": ["IMAGE_INSERT_FAILED"],
            "fallback": "figure-child-stack-then-notice",
        },
    }


def _m3_reference(*, target="fig:one", kind="figure", bookmark=FIG_BOOKMARK) -> dict:
    affixes = {"figure": ("图 ", ""), "table": ("表 ", ""), "equation": ("(", ")")}
    prefix, suffix = affixes[kind]
    return {
        "op": "writer.add_cross_reference", "nodeId": "para:ref",
        "args": {"runs": [{
            "type": "reference", "targetNodeId": target, "targetKind": kind,
            "bookmarkName": bookmark, "prefix": prefix, "suffix": suffix,
            "fallbackText": "[ref]",
        }]},
        "failurePolicy": {
            "mode": "degrade", "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
    }


def _legacy_figure(node_id="fig:legacy", resource_id="image-legacy") -> dict:
    return {
        "op": "writer.add_captioned_figure", "nodeId": node_id,
        "args": {
            "caption": "Legacy", "children": [{"nodeId": node_id + "/image:1", "resourceId": resource_id}],
            "layout": "stack",
        },
        "failurePolicy": {
            "mode": "degrade", "recoverableCodes": ["IMAGE_INSERT_FAILED"],
            "fallback": "notice",
        },
    }


def _legacy_reference() -> dict:
    return {
        "op": "writer.add_cross_reference", "nodeId": "para:legacy",
        "args": {"targetId": "fig:legacy", "kind": "figure", "fallbackText": "[figure]"},
        "failurePolicy": {
            "mode": "degrade", "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline",
        },
    }


def _m3_table() -> dict:
    return {
        "op": "writer.add_semantic_table", "nodeId": "tab:one",
        "args": {
            "caption": "Table", "numbering": _numbering("WPSC_TAB", "表 "),
            "bookmarkName": "wpsc_tab_" + "b" * 24,
            "indexable": True, "referenceable": True,
            "headers": ["A", "B"], "rows": [["", ""]],
            "alignments": ["left", "left"], "style": "three-line",
            "orientation": "portrait",
            "borderSpec": {"top": 1.5, "bottom": 1.5, "headerBottom": 0.75, "left": 0.0, "right": 0.0, "insideHorizontal": 0.0, "insideVertical": 0.0},
            "merges": [], "repeatHeader": True, "allowRowSplit": False,
            "cellIndentPt": 0.0, "plannedDegradation": [],
            "keepCaptionWithFirstRow": True,
        },
        "failurePolicy": {
            "mode": "degrade",
            "recoverableCodes": ["TABLE_STYLE_APPLY_FAILED", "TABLE_MERGE_APPLY_FAILED", "TABLE_ROW_FORCED_SPLIT", "TABLE_INSERT_FAILED"],
            "fallback": "grid-then-text",
        },
    }


def test_whole_plan_mode_rejects_m3_figure_with_legacy_reference_injection() -> None:
    legacy_ref = _legacy_reference()
    legacy_ref["args"].update(kind="REF EVIL \\h", targetId="ghost")
    legacy_ref["failurePolicy"] = {"mode": "degrade", "recoverableCodes": ["ANY_ERROR"], "fallback": "anything"}
    with pytest.raises(OperationPlanError, match="mix|mode|legacy"):
        validate_generation_plan(_envelope([_m3_figure(), legacy_ref, _finalize()]), "writer")


def test_whole_plan_mode_rejects_m3_figure_with_unsafe_legacy_figure() -> None:
    legacy = _legacy_figure("fig:legacy", "/tmp/secret.png")
    legacy["args"]["layout"] = "arbitrary"
    legacy["failurePolicy"] = {"mode": "degrade", "recoverableCodes": ["ANY_ERROR"], "fallback": "shell"}
    with pytest.raises(OperationPlanError, match="mix|mode|legacy"):
        validate_generation_plan(_envelope([_m3_figure(), legacy, _finalize()]), "writer")


def test_positive_whole_plan_m2_compatibility_is_safe() -> None:
    plan = validate_generation_plan(
        _envelope([_legacy_figure(), _legacy_reference(), _finalize()]), "writer"
    )
    assert [op.op for op in plan.operations] == [
        "writer.add_captioned_figure",
        "writer.add_cross_reference",
        "writer.finalize_fields",
    ]
    fatal = _legacy_figure()
    fatal["failurePolicy"] = {"mode": "fail"}
    validate_generation_plan(_envelope([fatal, _finalize()]), "writer")


def test_legacy_figure_child_cannot_mix_resource_and_degradation() -> None:
    figure = _legacy_figure()
    figure["args"]["children"][0]["plannedDegradation"] = {
        "code": "RESOURCE_NOT_FOUND", "message": "missing",
        "fallback": "[missing]", "placement": "block",
    }
    with pytest.raises(OperationPlanError, match="resource|degradation|exactly one"):
        validate_generation_plan(_envelope([figure]), "writer")


@pytest.mark.parametrize(
    "reference",
    [
        _m3_reference(target="fig:ghost"),
        _m3_reference(kind="table", bookmark="wpsc_tab_" + "b" * 24),
        _m3_reference(bookmark="wpsc_fig_" + "c" * 24),
    ],
)
def test_reference_target_node_kind_and_bookmark_must_match_exactly(reference: dict) -> None:
    with pytest.raises(OperationPlanError, match="target|reference|bookmark|kind"):
        validate_generation_plan(_envelope([_m3_figure(), reference, _finalize()]), "writer")


def test_m3_requires_one_finalizer_last_and_bounded_rounds() -> None:
    with pytest.raises(OperationPlanError, match="finalize"):
        validate_generation_plan(_envelope([_m3_figure()]), "writer")
    with pytest.raises(OperationPlanError, match="last"):
        validate_generation_plan(_envelope([_m3_figure(), _finalize(), {"op": "writer.add_page_break", "args": {}}]), "writer")
    for value in (0, 4):
        with pytest.raises(OperationPlanError, match="maxRounds"):
            validate_generation_plan(_envelope([_m3_figure(), _finalize(value)]), "writer")


def test_indexes_precede_native_body_without_section_marker() -> None:
    index = {"op": "writer.insert_figure_index", "nodeId": "doc:index", "args": {"title": "图目录", "sequenceId": "WPSC_FIG", "titleStyleId": "WPSC_INDEX_TITLE"}}
    with pytest.raises(OperationPlanError, match="indexes must precede"):
        validate_generation_plan(_envelope([_m3_figure(), index, _finalize()]), "writer")


@pytest.mark.parametrize("node_id", ["fig:one", "doc:collision"])
def test_plain_paragraph_and_native_object_cannot_share_node_owner(node_id: str) -> None:
    paragraph = {"op": "writer.add_paragraph", "nodeId": node_id, "args": {"text": "duplicate"}}
    figure = _m3_figure(node_id=node_id)
    with pytest.raises(OperationPlanError, match="owned more than once"):
        validate_generation_plan(_envelope([paragraph, figure, _finalize()]), "writer")


@pytest.mark.parametrize(("field", "value"), [("keepWithCaption", False)])
def test_figure_cohesion_is_mandatory(field: str, value: object) -> None:
    figure = _m3_figure()
    figure["args"][field] = value
    with pytest.raises(OperationPlanError, match=field):
        validate_generation_plan(_envelope([figure, _finalize()]), "writer")


def test_table_cohesion_is_mandatory() -> None:
    table = _m3_table()
    table["args"]["keepCaptionWithFirstRow"] = False
    with pytest.raises(OperationPlanError, match="keepCaptionWithFirstRow"):
        validate_generation_plan(_envelope([table, _finalize()]), "writer")


def test_table_merge_rejects_nonempty_covered_cell() -> None:
    table = _m3_table()
    table["args"]["merges"] = [{"top": 1, "left": 1, "bottom": 1, "right": 2}]
    with pytest.raises(OperationPlanError, match="covered|merge"):
        validate_generation_plan(_envelope([table, _finalize()]), "writer")


@pytest.mark.parametrize(
    "degradation",
    [
        {"code": "TABLE_MERGE_INVALID", "message": "x", "placement": "block", "insertAfter": "caption", "trigger": "row-exceeds-available-page", "recoveryScope": "row", "actions": ["allow-row-split"], "rowGroup": {"top": 2, "bottom": 2}},
        {"code": "TABLE_ROW_FORCED_SPLIT", "message": "x", "placement": "block", "insertAfter": "caption", "trigger": "invalid-merge-declaration", "recoveryScope": "complete-table", "actions": ["discard-all-merges", "preserve-complete-grid"]},
        {"code": "TABLE_ROW_FORCED_SPLIT", "message": "x", "placement": "block", "insertAfter": "caption", "trigger": "row-exceeds-available-page", "recoveryScope": "row", "actions": ["allow-row-split"], "rowGroup": {"top": 1, "bottom": 1}},
    ],
)
def test_table_degradation_code_has_one_exact_recovery_shape(degradation: dict) -> None:
    table = _m3_table()
    table["args"]["plannedDegradation"] = [degradation]
    with pytest.raises(OperationPlanError, match="degradation|trigger|rowGroup|recovery"):
        validate_generation_plan(_envelope([table, _finalize()]), "writer")


def test_table_degradation_matches_the_resolved_merge_state() -> None:
    table = _m3_table()
    table["args"]["merges"] = [{"top": 2, "left": 1, "bottom": 2, "right": 2}]
    table["args"]["plannedDegradation"] = [{
        "code": "TABLE_MERGE_INVALID", "message": "x", "placement": "block",
        "insertAfter": "caption", "trigger": "invalid-merge-declaration",
        "recoveryScope": "complete-table",
        "actions": ["discard-all-merges", "preserve-complete-grid"],
    }]
    with pytest.raises(OperationPlanError, match="degradation|merge"):
        validate_generation_plan(_envelope([table, _finalize()]), "writer")

    table = _m3_table()
    table["args"]["rows"] = [["", ""], ["", ""]]
    table["args"]["plannedDegradation"] = [{
        "code": "TABLE_ROW_FORCED_SPLIT", "message": "x", "placement": "block",
        "insertAfter": "caption",
        "trigger": "vertical-merge-group-exceeds-available-page",
        "recoveryScope": "complete-table",
        "actions": ["discard-all-merges", "apply-grid-style", "allow-row-split"],
        "rowGroup": {"top": 2, "bottom": 3},
    }]
    with pytest.raises(OperationPlanError, match="rowGroup|merge"):
        validate_generation_plan(_envelope([table, _finalize()]), "writer")


def test_node_id_and_manifest_digest_use_closed_string_validation() -> None:
    figure = _m3_figure(node_id="😀" * 50_001)
    with pytest.raises(OperationPlanError, match="UTF-16|string|nodeId"):
        validate_generation_plan(_envelope([figure, _finalize()]), "writer")
    with pytest.raises(OperationPlanError, match="resourceManifestDigest"):
        validate_generation_plan(_envelope([_m3_figure(), _finalize()], digest="not-a-digest"), "writer")


@pytest.mark.parametrize(
    ("media_type", "normalizer"),
    [("image/svg+xml", "none-v1"), ("image/jpeg", "svg-static-v1"), ("image/jpeg", "exif-transpose-png-v1")],
)
def test_media_type_and_normalizer_must_match(media_type: str, normalizer: str) -> None:
    figure = _m3_figure()
    child = figure["args"]["children"][0]
    child["mediaType"] = media_type
    child["normalizerId"] = normalizer
    with pytest.raises(OperationPlanError, match="mediaType|normalizer"):
        validate_generation_plan(_envelope([figure, _finalize()]), "writer")


def test_fatal_policy_cannot_smuggle_recovery_fields() -> None:
    equation = {
        "op": "writer.add_equation", "nodeId": "eq:one",
        "args": {"source": "x", "numbering": _numbering("WPSC_EQ", "(", ")"), "bookmarkName": "wpsc_eq_" + "d" * 24, "fallbackText": "x"},
        "failurePolicy": {"mode": "fail", "recoverableCodes": ["ANY"], "fallback": "continue"},
    }
    with pytest.raises(OperationPlanError, match="failurePolicy|fatal"):
        validate_generation_plan(_envelope([equation, _finalize()]), "writer")


def test_m3_generic_operation_cannot_invent_a_degradation_policy() -> None:
    paragraph = {
        "op": "writer.add_paragraph", "nodeId": "para:one",
        "args": {"text": "body"},
        "failurePolicy": {"mode": "degrade", "recoverableCodes": ["ANY_ERROR"], "fallback": "continue"},
    }
    with pytest.raises(OperationPlanError, match="failurePolicy|fatal|recoverable"):
        validate_generation_plan(_envelope([paragraph, _m3_figure(), _finalize()]), "writer")


def test_table_style_and_border_policy_cannot_diverge() -> None:
    table = _m3_table()
    table["args"]["borderSpec"]["insideVertical"] = 0.75
    with pytest.raises(OperationPlanError, match="borderSpec|three-line"):
        validate_generation_plan(_envelope([table, _finalize()]), "writer")


def _m3_figure_index() -> dict:
    return {
        "op": "writer.insert_figure_index", "nodeId": "doc:figure-index",
        "args": {"title": "图目录", "sequenceId": "WPSC_FIG", "titleStyleId": "WPSC_INDEX_TITLE"},
    }


def _configure(role: str, node_id: str) -> dict:
    return {
        "op": "writer.configure_section", "nodeId": node_id,
        "args": {"role": role},
    }


def test_native_index_must_be_inside_front_matter_section() -> None:
    operations = [
        _configure("body", "doc:body"),
        _m3_figure_index(),
        _m3_figure(),
        _finalize(),
    ]
    with pytest.raises(OperationPlanError, match="front.?matter|index"):
        validate_generation_plan(_envelope(operations), "writer")


def test_native_index_requires_one_nonempty_indexable_target() -> None:
    empty = _m3_figure()
    empty["args"].pop("bookmarkName")
    empty["args"].update(caption="", indexable=False, referenceable=False)
    operations = [
        _configure("front_matter", "doc:front"),
        _m3_figure_index(),
        _configure("body", "doc:body"),
        empty,
        _finalize(),
    ]
    with pytest.raises(OperationPlanError, match="indexable|caption|target"):
        validate_generation_plan(_envelope(operations), "writer")


def test_native_index_kind_occurs_at_most_once() -> None:
    operations = [
        _configure("front_matter", "doc:front"),
        _m3_figure_index(), copy.deepcopy(_m3_figure_index()),
        _configure("body", "doc:body"),
        _m3_figure(),
        _finalize(),
    ]
    operations[2]["nodeId"] = "doc:figure-index-duplicate"
    with pytest.raises(OperationPlanError, match="duplicate|at most one|index"):
        validate_generation_plan(_envelope(operations), "writer")


def test_bookmark_name_is_globally_owned_by_one_target() -> None:
    second = _m3_figure(node_id="fig:two", bookmark=FIG_BOOKMARK)
    second["args"]["children"][0]["resourceId"] = "image-2"
    with pytest.raises(OperationPlanError, match="bookmark|owner|unique"):
        validate_generation_plan(_envelope([_m3_figure(), second, _finalize()]), "writer")


@pytest.mark.parametrize(("field", "value"), [("allowRowSplit", True), ("cellIndentPt", 1.0)])
def test_initial_table_descriptor_cannot_preapply_forced_split(field: str, value: object) -> None:
    table = _m3_table()
    table["args"][field] = value
    with pytest.raises(OperationPlanError, match=field):
        validate_generation_plan(_envelope([table, _finalize()]), "writer")


def test_pure_text_v2_plan_rejects_arbitrary_degradation_policy() -> None:
    paragraph = {
        "op": "writer.add_paragraph", "nodeId": "para:text", "args": {"text": "body"},
        "failurePolicy": {"mode": "degrade", "recoverableCodes": ["ANY_ERROR"], "fallback": "continue"},
    }
    with pytest.raises(OperationPlanError, match="failurePolicy|recoverable|fatal"):
        validate_generation_plan(_envelope([paragraph, _finalize()]), "writer")


def test_pure_text_v2_plan_requires_finalizer_last() -> None:
    paragraph = {"op": "writer.add_paragraph", "nodeId": "para:text", "args": {"text": "body"}}
    with pytest.raises(OperationPlanError, match="finalize"):
        validate_generation_plan(_envelope([paragraph]), "writer")
    with pytest.raises(OperationPlanError, match="last"):
        validate_generation_plan(_envelope([_finalize(), paragraph]), "writer")
