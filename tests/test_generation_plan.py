import json

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    ALLOWED_OPERATIONS,
    MAX_PLAN_BYTES,
    GenerationOperation,
    GenerationPlan,
    GenerationResource,
    OperationPlanError,
    RecordedGeneration,
    validate_generation_plan,
)


def test_generation_plan_round_trips_valid_writer_operations():
    plan = GenerationPlan(
        "writer",
        (
            GenerationOperation("writer.reset", {}),
            GenerationOperation(
                "writer.add_paragraph",
                {"text": "macOS parity", "style": "Heading 1"},
            ),
        ),
    )
    assert validate_generation_plan(plan.to_dict(), "writer") == plan


def test_generation_plan_rejects_unknown_or_cross_component_operation():
    with pytest.raises(OperationPlanError, match="unsupported operation"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [{"op": "sheet.write_table", "args": {}}],
            },
            "writer",
        )


def test_recorded_resources_are_not_serialized_into_bridge_plan(tmp_path):
    plan = GenerationPlan(
        "writer",
        (GenerationOperation("writer.reset", {}),),
    )
    resource = GenerationResource("image-1", tmp_path / "figure.png", "image/png")
    recorded = RecordedGeneration(plan=plan, resources=(resource,))
    assert "source_path" not in json.dumps(recorded.plan.to_dict())
    assert recorded.resources == (resource,)


def test_allowed_operations_are_a_closed_component_scoped_set():
    assert ALLOWED_OPERATIONS == {
        "writer": {
            "writer.reset",
            "writer.configure_page",
            "writer.ensure_styles",
            "writer.add_paragraph",
            "writer.add_heading",
            "writer.add_list",
            "writer.add_table",
            "writer.add_image",
            "writer.add_page_break",
            "writer.add_section",
            "writer.add_horizontal_line",
            "writer.insert_toc",
            "writer.set_page_number",
            "writer.update_fields",
        },
        "spreadsheet": {
            "sheet.reset",
            "sheet.rename",
            "sheet.add",
            "sheet.select",
            "sheet.write_table",
            "sheet.set_column_width",
            "sheet.autofit",
        },
        "presentation": {
            "slide.reset",
            "slide.set_size",
            "slide.apply_preset",
            "slide.add_title",
            "slide.add_section",
            "slide.add_bullets",
            "slide.add_blank",
            "slide.add_image",
            "slide.add_table",
        },
    }


def test_allowed_operations_cannot_be_widened_at_runtime():
    try:
        with pytest.raises(TypeError):
            ALLOWED_OPERATIONS["rogue"] = {"eval"}
    finally:
        if isinstance(ALLOWED_OPERATIONS, dict):
            ALLOWED_OPERATIONS.pop("rogue", None)

    writer_operations = ALLOWED_OPERATIONS["writer"]
    try:
        with pytest.raises(AttributeError):
            writer_operations.add("writer.eval")
    finally:
        if isinstance(writer_operations, set):
            writer_operations.discard("writer.eval")


def test_generation_plan_requires_matching_component():
    with pytest.raises(OperationPlanError, match="component mismatch"):
        validate_generation_plan(
            {
                "component": "presentation",
                "operations": [{"op": "slide.reset", "args": {}}],
            },
            "writer",
        )


def test_generation_plan_requires_at_least_one_operation():
    with pytest.raises(OperationPlanError, match="at least one"):
        validate_generation_plan(
            {"component": "writer", "operations": []},
            "writer",
        )


def test_generation_plan_rejects_more_than_ten_thousand_operations():
    operation = {"op": "writer.reset", "args": {}}
    with pytest.raises(OperationPlanError, match="too many operations"):
        validate_generation_plan(
            {"component": "writer", "operations": [operation] * 10_001},
            "writer",
        )


def test_generation_plan_requires_exact_top_level_keys():
    with pytest.raises(OperationPlanError, match="exactly component and operations"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [{"op": "writer.reset", "args": {}}],
                "resources": [],
            },
            "writer",
        )


@pytest.mark.parametrize(
    "operation",
    [
        {"op": "writer.reset"},
        {"op": "writer.reset", "args": {}, "handler": "eval"},
    ],
)
def test_generation_operation_requires_exact_op_and_args_keys(operation):
    with pytest.raises(OperationPlanError, match="exactly op and args"):
        validate_generation_plan(
            {"component": "writer", "operations": [operation]},
            "writer",
        )


def test_generation_operation_requires_string_name():
    with pytest.raises(OperationPlanError, match="op must be a string"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [{"op": 7, "args": {}}],
            },
            "writer",
        )


def test_generation_plan_rejects_non_json_arguments():
    with pytest.raises(OperationPlanError, match="JSON-compatible"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [
                    {"op": "writer.add_paragraph", "args": {"text": object()}}
                ],
            },
            "writer",
        )


def test_generation_plan_normalizes_cyclic_arguments_to_protocol_error():
    cyclic = {}
    cyclic["self"] = cyclic
    with pytest.raises(OperationPlanError, match="JSON-compatible"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [
                    {"op": "writer.add_paragraph", "args": cyclic}
                ],
            },
            "writer",
        )


def test_generation_plan_rejects_more_than_two_million_serialized_bytes():
    assert MAX_PLAN_BYTES == 2_000_000
    chunks = ["x" * 100_000 for _ in range(21)]
    with pytest.raises(OperationPlanError, match="2,000,000 bytes"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [
                    {"op": "writer.add_paragraph", "args": {"chunks": chunks}}
                ],
            },
            "writer",
        )


def test_generation_plan_rejects_nested_strings_over_one_hundred_thousand_chars():
    with pytest.raises(OperationPlanError, match="100,000 characters"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [
                    {
                        "op": "writer.add_paragraph",
                        "args": {"spans": [{"text": "x" * 100_001}]},
                    }
                ],
            },
            "writer",
        )


@pytest.mark.parametrize(
    "op", ["writer.add_table", "sheet.write_table", "slide.add_table"]
)
def test_generation_plan_rejects_tables_over_ten_thousand_cells(op):
    component = {
        "writer.add_table": "writer",
        "sheet.write_table": "spreadsheet",
        "slide.add_table": "presentation",
    }[op]
    with pytest.raises(OperationPlanError, match="10,000 cells"):
        validate_generation_plan(
            {
                "component": component,
                "operations": [
                    {"op": op, "args": {"data": [[0] * 101 for _ in range(100)]}}
                ],
            },
            component,
        )


@pytest.mark.parametrize("op", ["writer.add_image", "slide.add_image"])
def test_generation_plan_requires_logical_image_id(op):
    component = "writer" if op.startswith("writer.") else "presentation"
    with pytest.raises(OperationPlanError, match="requires imageId"):
        validate_generation_plan(
            {
                "component": component,
                "operations": [{"op": op, "args": {"path": "/tmp/figure.png"}}],
            },
            component,
        )


def test_generation_plan_accepts_logical_image_id():
    raw = {
        "component": "writer",
        "operations": [
            {"op": "writer.add_image", "args": {"imageId": "image-1", "alt": "Chart"}}
        ],
    }
    assert validate_generation_plan(raw, "writer").to_dict() == raw


def test_generation_plan_rejects_nested_host_image_path():
    with pytest.raises(OperationPlanError, match="image paths are not allowed"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [
                    {
                        "op": "writer.add_image",
                        "args": {
                            "imageId": "image-1",
                            "metadata": {"sourcePath": "/tmp/figure.png"},
                        },
                    }
                ],
            },
            "writer",
        )


def test_generation_resource_resolves_host_path_and_rejects_unlisted_media_type(tmp_path):
    resource = GenerationResource(
        "image-1",
        tmp_path / "nested" / ".." / "figure.png",
        "image/png",
    )
    assert resource.source_path == (tmp_path / "figure.png").resolve()

    with pytest.raises(OperationPlanError, match="unsupported media type"):
        GenerationResource("image-2", tmp_path / "notes.txt", "text/plain")
