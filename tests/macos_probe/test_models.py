from pathlib import Path

import pytest

from skills.WPSComposer.scripts.macos_probe.models import (
    METHOD_COMPONENT,
    PathPolicy,
    ProbeCommand,
    ProbeResult,
    ProtocolError,
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
    }


def test_command_rejects_unknown_method():
    with pytest.raises(ProtocolError, match="Unsupported method"):
        ProbeCommand.create("writer", "eval", {})


def test_command_rejects_wrong_component():
    with pytest.raises(ProtocolError, match="requires presentation"):
        ProbeCommand.create("writer", "smoke_pptx", {})


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
