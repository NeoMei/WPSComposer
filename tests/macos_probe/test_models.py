from pathlib import Path

import pytest

from skills.WPSComposer.scripts.macos_probe.models import (
    METHOD_COMPONENT,
    PathPolicy,
    ProbeCommand,
    ProbeResult,
    ProtocolError,
    validate_longform_generation_request,
)


def test_methods_are_routed_to_one_component():
    assert METHOD_COMPONENT == {
        "probe_capabilities": None,
        "smoke_docx": "writer",
        "smoke_pdf": "writer",
        "smoke_pptx": "presentation",
        "smoke_xlsx": "spreadsheet",
        "convert_writer_pdf": "writer",
        "convert_workbook_pdf": "spreadsheet",
        "convert_presentation_pdf": "presentation",
        "generate_writer_document": "writer",
        "generate_spreadsheet_workbook": "spreadsheet",
        "generate_presentation_deck": "presentation",
        "inspect_presentation": "presentation",
        "edit_presentation": "presentation",
        "inspect_document": "writer",
        "inspect_workbook": "spreadsheet",
        "probe_longform_m0": "writer",
        "generate_longform_document": "writer",
        "mutate_longform_document": "writer",
        "patch_longform_quality_notices": "writer",
    }


def test_command_rejects_unknown_method():
    with pytest.raises(ProtocolError, match="Unsupported method"):
        ProbeCommand.create("writer", "eval", {})


def test_command_rejects_wrong_component():
    with pytest.raises(ProtocolError, match="requires presentation"):
        ProbeCommand.create("writer", "smoke_pptx", {})


def test_generation_command_is_component_typed():
    with pytest.raises(ProtocolError, match="requires writer"):
        ProbeCommand.create("spreadsheet", "generate_writer_document", {})


def test_longform_request_accepts_optional_activation_document_and_old_shape():
    base = {
        "plan": {"protocolVersion": 2, "component": "writer"},
        "outputPath": "/private/out.docx",
        "resources": {},
    }

    assert validate_longform_generation_request(base) == base
    claimed = dict(base, activationDocument="/private/wpscomposer-writer-blank.docx")
    assert validate_longform_generation_request(claimed) == claimed


@pytest.mark.parametrize("value", [None, "", False, 1])
def test_longform_request_rejects_invalid_activation_document(value):
    request = {
        "plan": {"protocolVersion": 2, "component": "writer"},
        "outputPath": "/private/out.docx",
        "resources": {},
        "activationDocument": value,
    }

    with pytest.raises(ProtocolError, match="request is invalid"):
        validate_longform_generation_request(request)


@pytest.mark.parametrize(
    ("component", "method"),
    [
        ("writer", "convert_writer_pdf"),
        ("spreadsheet", "convert_workbook_pdf"),
        ("presentation", "convert_presentation_pdf"),
    ],
)
def test_conversion_methods_are_scoped_to_one_component(component, method):
    assert ProbeCommand.create(component, method, {}).method == method
    wrong = next(
        name for name in ("writer", "spreadsheet", "presentation")
        if name != component
    )
    with pytest.raises(ProtocolError, match=f"requires {component}"):
        ProbeCommand.create(wrong, method, {})


def test_result_round_trip():
    result = ProbeResult("cmd-1", True, {"path": "/tmp/a.docx"}, None)
    assert ProbeResult.from_dict(result.to_dict()) == result


def test_path_policy_accepts_only_declared_roots(tmp_path: Path):
    runtime = tmp_path / "runtime"
    output = tmp_path / "output"
    runtime.mkdir()
    output.mkdir()
    policy = PathPolicy((runtime, output))

    assert policy.require_allowed(output / "smoke.docx") == (
        output / "smoke.docx"
    ).resolve()

    with pytest.raises(ProtocolError, match="outside allowed roots"):
        policy.require_allowed(tmp_path.parent / "escape.docx")
