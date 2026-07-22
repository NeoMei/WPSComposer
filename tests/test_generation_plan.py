import json

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    ALLOWED_OPERATIONS,
    MAX_NESTING_DEPTH,
    MAX_PLAN_BYTES,
    MAX_SAFE_NUMBER,
    GenerationOperation,
    GenerationPlan,
    GenerationResource,
    OperationPlanError,
    RecordedGeneration,
    _validate_json_value,
    validate_generation_plan,
)


VALID_OPERATION_ARGS = {
    "writer.reset": {},
    "writer.configure_page": {
        "marginTop": 72,
        "marginBottom": 72,
        "marginLeft": 72,
        "marginRight": 72,
        "pageWidth": 612,
        "pageHeight": 792,
        "landscape": False,
        "columns": 1,
        "header": "Quarterly report",
        "footer": "Confidential",
    },
    "writer.ensure_styles": {
        "styles": [
            {
                "name": "Body Text",
                "type": "paragraph",
                "basedOn": "Normal",
                "fontName": "FangSong",
                "fontNameAscii": "Times New Roman",
                "fontSize": 12,
                "bold": False,
                "italic": False,
                "color": "#000000",
                "align": 3,
                "indentFirst": 24,
                "leftIndent": 0,
                "rightIndent": 0,
                "lineSpacing": 1.5,
                "lineSpacingRule": "one_and_half",
                "spaceBefore": 0,
                "spaceAfter": 0,
                "shading": "#FFFFFF",
                "leftBorder": False,
                "borderColor": "#CCCCCC",
                "outlineLevel": 1,
            }
        ]
    },
    "writer.add_paragraph": {
        "text": "Summary",
        "style": "Body Text",
        "spans": [
            {
                "text": "Summary",
                "bold": True,
                "italic": False,
                "strikethrough": False,
                "code": False,
                "link": "https://example.test",
            }
        ],
        "size": 12,
        "color": "#000000",
        "align": 3,
        "indentFirst": 24,
        "lineSpacing": 1.5,
        "lineSpacingRule": "one_and_half",
        "spaceBefore": 0,
        "spaceAfter": 0,
        "fontName": "FangSong",
        "fontNameAscii": "Times New Roman",
    },
    "writer.add_heading": {
        "text": "Results",
        "level": 1,
        "size": 16,
        "color": "#000000",
        "lineSpacing": 1.5,
        "lineSpacingRule": "one_and_half",
        "spaceAfter": 6,
    },
    "writer.add_list": {
        "items": ["First", "Second"],
        "ordered": False,
        "glyph": "•",
        "indent": 24,
    },
    "writer.add_table": {
        "rows": 2,
        "cols": 2,
        "data": [["Item", "Amount"], ["A", 10]],
        "shadeHeader": "#4472C4",
        "headerColor": "#FFFFFF",
        "fontSize": 10,
        "columnWidths": [120, 80],
        "alignments": ["left", "right"],
        "bandedRows": True,
        "autoFit": True,
        "repeatHeader": True,
        "borderColor": "#D0D0D0",
    },
    "writer.add_image": {
        "imageId": "image-1",
        "width": 400,
        "height": 300,
        "maxWidth": 500,
        "maxHeight": 500,
        "wrap": 0,
        "inline": True,
        "preserveAspect": True,
        "alt": "Chart",
    },
    "writer.add_page_break": {},
    "writer.add_section": {},
    "writer.add_horizontal_line": {},
    "writer.insert_toc": {"title": "Table of Contents"},
    "writer.set_page_number": {},
    "writer.update_fields": {},
    "sheet.reset": {},
    "sheet.rename": {"index": 1, "name": "Summary"},
    "sheet.add": {"name": "Details"},
    "sheet.select": {"index": 1},
    "sheet.write_table": {
        "startRow": 1,
        "startCol": 1,
        "values": [["Item", "Amount"], ["A", 10]],
        "headerBold": True,
        "headerShade": "#4472C4",
        "headerFontColor": "#FFFFFF",
        "fontSize": 11,
    },
    "sheet.set_column_width": {"column": "A", "width": 15},
    "sheet.autofit": {},
    "slide.reset": {},
    "slide.set_size": {"width": 960, "height": 540},
    "slide.apply_preset": {
        "preset": {
            "name": "business",
            "colors": {
                "primary": "#005294",
                "secondary": "#C82828",
                "accent": "#2DA050",
                "dark": "#2D2D30",
                "light": "#F8F9FA",
                "background": "#FFFFFF",
            },
            "fonts": {
                "title": {"family": "Arial", "size": 36, "color": "#005294"},
                "subtitle": {"family": "Arial", "size": 20, "color": "#5A5A5A"},
                "body": {"family": "Arial", "size": 18, "color": "#2D2D30"},
                "caption": {"family": "Arial", "size": 14, "color": "#8C8C8C"},
                "chinese": {"family": "Microsoft YaHei", "size": 18, "color": "#2D2D30"},
            },
            "spacing": {
                "margin": 70,
                "gap": 22,
                "cardPadding": 18,
                "lineHeight": 1.4,
            },
        }
    },
    "slide.add_title": {
        "title": "Quarterly results",
        "subtitle": "FY26 Q4",
        "titleSize": 40,
        "subtitleSize": 20,
        "titleColor": "#005294",
    },
    "slide.add_section": {"title": "Results"},
    "slide.add_bullets": {
        "title": "Highlights",
        "items": ["Revenue grew", "Margin expanded"],
        "titleSize": 32,
        "bodySize": 18,
    },
    "slide.add_blank": {},
    "slide.add_image": {
        "slide": 2,
        "imageId": "image-2",
        "left": 80,
        "top": 100,
        "width": 800,
        "height": 400,
    },
    "slide.add_table": {
        "slide": 2,
        "rows": 2,
        "cols": 2,
        "left": 60,
        "top": 120,
        "width": 840,
        "height": 380,
        "data": [["Item", "Amount"], ["A", 10]],
        "headerShade": "#4472C4",
        "headerFont": "#FFFFFF",
        "fontSize": 14,
    },
}


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
    assert set(VALID_OPERATION_ARGS) == set().union(*ALLOWED_OPERATIONS.values())


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


def _component_for_operation(op):
    return next(
        component
        for component, operations in ALLOWED_OPERATIONS.items()
        if op in operations
    )


@pytest.mark.parametrize("op", sorted(VALID_OPERATION_ARGS))
def test_every_allowed_operation_accepts_its_exact_argument_schema(op):
    component = _component_for_operation(op)
    raw = {
        "component": component,
        "operations": [{"op": op, "args": VALID_OPERATION_ARGS[op]}],
    }
    assert validate_generation_plan(raw, component).to_dict() == raw


@pytest.mark.parametrize("op", sorted(VALID_OPERATION_ARGS))
def test_every_allowed_operation_rejects_unknown_arguments(op):
    component = _component_for_operation(op)
    args = dict(VALID_OPERATION_ARGS[op], javascript="Application.eval('x')")
    with pytest.raises(OperationPlanError, match="unknown argument"):
        validate_generation_plan(
            {"component": component, "operations": [{"op": op, "args": args}]},
            component,
        )


@pytest.mark.parametrize(
    ("op", "missing"),
    [
        ("writer.configure_page", "marginTop"),
        ("writer.ensure_styles", "styles"),
        ("writer.add_paragraph", "text"),
        ("writer.add_heading", "text"),
        ("writer.add_list", "items"),
        ("writer.add_table", "data"),
        ("writer.add_image", "imageId"),
        ("writer.insert_toc", "title"),
        ("sheet.rename", "name"),
        ("sheet.add", "name"),
        ("sheet.select", "index"),
        ("sheet.write_table", "values"),
        ("sheet.set_column_width", "width"),
        ("slide.set_size", "width"),
        ("slide.apply_preset", "preset"),
        ("slide.add_title", "title"),
        ("slide.add_section", "title"),
        ("slide.add_bullets", "items"),
        ("slide.add_image", "imageId"),
        ("slide.add_table", "data"),
    ],
)
def test_operation_schemas_reject_missing_required_arguments(op, missing):
    component = _component_for_operation(op)
    args = dict(VALID_OPERATION_ARGS[op])
    del args[missing]
    with pytest.raises(
        OperationPlanError, match="missing required argument|requires imageId"
    ):
        validate_generation_plan(
            {"component": component, "operations": [{"op": op, "args": args}]},
            component,
        )


@pytest.mark.parametrize(
    ("op", "field", "invalid"),
    [
        ("writer.configure_page", "columns", True),
        ("writer.add_paragraph", "text", 7),
        ("writer.add_heading", "level", 1.5),
        ("writer.add_list", "items", ["ok", {"text": "escape"}]),
        ("writer.add_table", "rows", False),
        ("writer.add_image", "inline", "true"),
        ("sheet.rename", "index", True),
        ("sheet.write_table", "startRow", 1.5),
        ("sheet.set_column_width", "width", "15"),
        ("slide.set_size", "width", "960"),
        ("slide.add_title", "titleSize", True),
        ("slide.add_bullets", "items", [1]),
        ("slide.add_image", "left", "80"),
        ("slide.add_table", "rows", 2.5),
    ],
)
def test_operation_schemas_reject_wrong_argument_types(op, field, invalid):
    component = _component_for_operation(op)
    args = dict(VALID_OPERATION_ARGS[op])
    args[field] = invalid
    with pytest.raises(OperationPlanError, match="invalid argument"):
        validate_generation_plan(
            {"component": component, "operations": [{"op": op, "args": args}]},
            component,
        )


@pytest.mark.parametrize(
    ("op", "field"),
    [
        ("writer.add_paragraph", "size"),
        ("slide.set_size", "width"),
        ("slide.add_image", "left"),
    ],
)
@pytest.mark.parametrize("value", [2**53 - 1, float(2**53 - 1)])
def test_numeric_operation_fields_accept_exact_safe_bound(op, field, value):
    assert MAX_SAFE_NUMBER == 2**53 - 1
    component = _component_for_operation(op)
    args = dict(VALID_OPERATION_ARGS[op])
    args[field] = value
    raw = {"component": component, "operations": [{"op": op, "args": args}]}
    assert validate_generation_plan(raw, component).to_dict() == raw


@pytest.mark.parametrize(
    ("op", "field"),
    [
        ("writer.add_paragraph", "size"),
        ("slide.set_size", "width"),
        ("slide.add_image", "left"),
    ],
)
@pytest.mark.parametrize(
    "value",
    [-(2**53), 2**53, float(2**53), 10**400],
)
def test_numeric_operation_fields_reject_values_outside_safe_range(
    op, field, value
):
    component = _component_for_operation(op)
    args = dict(VALID_OPERATION_ARGS[op])
    args[field] = value
    with pytest.raises(OperationPlanError, match="safe range"):
        validate_generation_plan(
            {"component": component, "operations": [{"op": op, "args": args}]},
            component,
        )


@pytest.mark.parametrize("value", [-(2**53), 2**53, 10**400])
def test_integer_operation_fields_reject_values_outside_safe_range(value):
    args = dict(VALID_OPERATION_ARGS["slide.add_image"])
    args["slide"] = value
    with pytest.raises(OperationPlanError, match="safe range"):
        validate_generation_plan(
            {
                "component": "presentation",
                "operations": [{"op": "slide.add_image", "args": args}],
            },
            "presentation",
        )


@pytest.mark.parametrize("value", [-(2**53 - 1), 2**53 - 1])
def test_table_cells_accept_exact_safe_integer_bound(value):
    args = dict(VALID_OPERATION_ARGS["writer.add_table"])
    args["data"] = [["Item", "Amount"], ["A", value]]
    raw = {
        "component": "writer",
        "operations": [{"op": "writer.add_table", "args": args}],
    }
    assert validate_generation_plan(raw, "writer").to_dict() == raw


@pytest.mark.parametrize("value", [-(2**53), 2**53, 10**400])
def test_table_cells_reject_integers_outside_safe_range(value):
    args = dict(VALID_OPERATION_ARGS["writer.add_table"])
    args["data"] = [["Item", "Amount"], ["A", value]]
    with pytest.raises(OperationPlanError, match="safe range"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [{"op": "writer.add_table", "args": args}],
            },
            "writer",
        )


@pytest.mark.parametrize(
    ("op", "field", "value"),
    [
        (
            "writer.ensure_styles",
            "styles",
            [{"name": "Body Text", "propertyPath": "Font.Name"}],
        ),
        (
            "writer.add_paragraph",
            "spans",
            [{"text": "unsafe", "javascript": "eval('x')"}],
        ),
        (
            "writer.add_table",
            "data",
            [[{"source": "/tmp/private"}]],
        ),
        (
            "slide.apply_preset",
            "preset",
            {
                "name": "unsafe",
                "colors": {"primary": "#000000", "src": "/tmp/theme"},
                "fonts": {},
            },
        ),
    ],
)
def test_nested_operation_structures_reject_escape_fields(op, field, value):
    component = _component_for_operation(op)
    args = dict(VALID_OPERATION_ARGS[op])
    args[field] = value
    with pytest.raises(OperationPlanError, match="unknown|invalid"):
        validate_generation_plan(
            {"component": component, "operations": [{"op": op, "args": args}]},
            component,
        )


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


def _nested_lists(depth):
    value = "leaf"
    for _ in range(depth):
        value = [value]
    return value


def test_json_nesting_depth_accepts_limit_and_rejects_one_over():
    assert MAX_NESTING_DEPTH == 64
    _validate_json_value(_nested_lists(MAX_NESTING_DEPTH))
    with pytest.raises(OperationPlanError, match="nesting exceeds 64"):
        _validate_json_value(_nested_lists(MAX_NESTING_DEPTH + 1))


def test_generation_plan_rejects_deep_acyclic_arguments_without_recursion_error():
    with pytest.raises(OperationPlanError, match="nesting exceeds 64"):
        validate_generation_plan(
            {
                "component": "writer",
                "operations": [
                    {
                        "op": "writer.add_paragraph",
                        "args": {"text": "safe", "spans": _nested_lists(500)},
                    }
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


@pytest.mark.parametrize(
    ("op", "field", "value"),
    [
        ("writer.add_table", "rows", 0),
        ("writer.add_table", "cols", -5),
        ("writer.add_heading", "level", 0),
        ("writer.add_heading", "level", 10),
        ("sheet.write_table", "startRow", 0),
        ("sheet.select", "index", -1),
        ("slide.add_table", "slide", 0),
        ("slide.add_table", "left", -0.5),
        ("slide.add_image", "width", 0),
        ("slide.set_size", "height", -100),
        ("writer.add_image", "maxWidth", -1),
    ],
)
def test_bounded_numeric_fields_reject_out_of_range_values(op, field, value):
    component = _component_for_operation(op)
    args = dict(VALID_OPERATION_ARGS[op])
    args[field] = value
    with pytest.raises(OperationPlanError, match="invalid argument"):
        validate_generation_plan(
            {"component": component, "operations": [{"op": op, "args": args}]},
            component,
        )
