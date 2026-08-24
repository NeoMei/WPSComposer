from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess
import tempfile

import pytest

from skills.WPSComposer.scripts.macos_probe.models import (
    METHOD_COMPONENT,
    ProbeCommand,
    ProtocolError,
)

ROOT = Path("macos/wps-jsapi-probe/addin")
EMPTY_MANIFEST_DIGEST = (
    "3516239310e9d6c0d9a736e816661d57d7e5378e334cfb4f03454bb3da0fc4ae"
)
SVG_SHA256 = "924f47ebd7a4c22393defac4103818aedced9f5dd0630bf27e1d6aa7ad30cbfc"


def run_node(script: str) -> None:
    descriptor, name = tempfile.mkstemp(suffix=".js")
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(script)
        subprocess.run(["node", name], check=True, capture_output=True, text=True)
    finally:
        os.unlink(name)


def valid_params() -> dict:
    return {
        "manifest": {
            "version": 1,
            "digest": EMPTY_MANIFEST_DIGEST,
            "entries": [
                {
                    "resourceId": "m0-static-svg",
                    "sourceSha256": SVG_SHA256,
                    "payloadSha256": SVG_SHA256,
                    "byteLength": 185,
                    "mediaType": "image/svg+xml",
                    "normalizerId": "identity-svg-m0-v1",
                }
            ],
        },
        "probeVersion": "0.8.0-m0.1",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "expectedWpsVersion": "12.1.fake",
        "stagedDocxPath": "/staged/probe.docx",
        "stagedPdfPath": "/staged/probe.pdf",
        "stagedSvgPath": "/staged/probe.svg",
    }


def test_longform_method_is_writer_only():
    assert METHOD_COMPONENT["probe_longform_m0"] == "writer"
    assert ProbeCommand.create("writer", "probe_longform_m0", {}).method == (
        "probe_longform_m0"
    )
    with pytest.raises(ProtocolError, match="requires writer"):
        ProbeCommand.create("spreadsheet", "probe_longform_m0", {})


def test_html_loads_longform_boundary_before_component():
    html = (ROOT / "index.html").read_text(encoding="utf-8")

    assert html.index("bridge-client.js") < html.index("writer-longform-m0.js")
    assert html.index("writer-longform-m0.js") < html.index("component.js")


def test_longform_asset_has_no_dynamic_execution_or_network_access():
    source = (ROOT / "writer-longform-m0.js").read_text(encoding="utf-8")

    for forbidden in (
        "eval(",
        "new Function",
        "Function(",
        "fetch(",
        "XMLHttpRequest",
        "WebSocket",
    ):
        assert forbidden not in source
    assert "M0_KEYS" in source
    assert EMPTY_MANIFEST_DIGEST in source
    assert "window.WPSComposerLongformM0" in source


def test_native_probe_has_closed_lifecycle_and_all_capability_assays():
    source = (ROOT / "writer-longform-m0.js").read_text(encoding="utf-8")

    for required in (
        "Application.Documents.Add()",
        "SaveAs2(envelope.stagedDocxPath, 12)",
        "Application.Documents.Open(envelope.stagedDocxPath",
        "document.Fields.Update()",
        "ExportAsFixedFormat(envelope.stagedPdfPath, 17",
        "Application.DisplayAlerts = previousAlerts",
        "Application.ScreenUpdating = previousScreenUpdating",
        "InlineShapes.AddPicture(envelope.stagedSvgPath",
        "TablesOfContents.Add",
        "TablesOfFigures.Add",
        "OMaths.Add",
        "ApplyListTemplateWithLevel",
    ):
        assert required in source
    assert source.count("\n    assay(document, capabilities,") == 13
    assert "for (let capabilityId = 3; capabilityId <= 15" in source


def test_numbering_assay_checks_exact_schemes_and_real_mutations():
    source = (ROOT / "writer-longform-m0.js").read_text(encoding="utf-8")

    for expected_format in (
        '"第%1章"',
        '"第%2节"',
        '"%3、"',
        '"（%4）"',
        '"%1.%2.%3.%4"',
        '"关键工法%4："',
    ):
        assert expected_format in source
    assert "numberingSnapshotMatches(document)" in source
    assert "range.ListFormat.ListLevelNumber = level" in source
    assert source.count("style: 253") == 2
    assert "listLevel.Legal" not in source
    assert ".Cut()" in source
    assert ".Paste()" in source
    assert "mutationRenumberChecks: 3" in source


def test_writer_registers_exactly_one_typed_longform_handler():
    writer = (ROOT / "writer.js").read_text(encoding="utf-8")

    assert writer.count('"probe_longform_m0"') == 1
    assert (
        '"probe_longform_m0": function (params) { '
        "return window.WPSComposerLongformM0.run(params); }"
    ) in writer
    bridge = (ROOT / "bridge-client.js").read_text(encoding="utf-8")
    assert '"PROTOCOL_MISMATCH"' in bridge
    assert '"RESOURCE_MANIFEST_MISMATCH"' in bridge


def test_valid_longform_envelope_is_accepted_without_native_mutation():
    asset = json.dumps(str((ROOT / "writer-longform-m0.js").resolve()))
    params = json.dumps(valid_params())
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
const validated = window.WPSComposerLongformM0.validateEnvelope({params});
assert.equal(validated.protocolVersion, 2);
    assert.equal(validated.manifest.entries.length, 1);
"""

    run_node(script)


@pytest.mark.parametrize(
    ("replacement", "expected_code"),
    [
        ({"protocolVersion": 1}, "PROTOCOL_MISMATCH"),
        ({"resourceManifestVersion": 2}, "PROTOCOL_MISMATCH"),
        ({"probeVersion": "0.8.0-m0.invalid"}, "PROTOCOL_MISMATCH"),
        ({"expectedWpsVersion": ""}, "PROTOCOL_MISMATCH"),
        ({"unexpected": True}, "PROTOCOL_MISMATCH"),
        ({"stagedDocxPath": ""}, "PROTOCOL_MISMATCH"),
        (
            {
                "manifest": {
                    "version": 1,
                    "digest": "c" * 64,
                    "entries": valid_params()["manifest"]["entries"],
                }
            },
            "RESOURCE_MANIFEST_MISMATCH",
        ),
        (
            {
                "manifest": {
                    "version": 1,
                    "digest": EMPTY_MANIFEST_DIGEST,
                    "entries": [{"resourceId": "unexpected"}],
                }
            },
            "RESOURCE_MANIFEST_MISMATCH",
        ),
    ],
)
def test_invalid_envelope_is_rejected_before_document_open(
    replacement: dict, expected_code: str
):
    params = valid_params()
    params.update(replacement)
    m0_asset = json.dumps(str((ROOT / "writer-longform-m0.js").resolve()))
    writer_asset = json.dumps(str((ROOT / "writer.js").resolve()))
    params_json = json.dumps(params)
    script = f"""
const assert = require("assert");
const fs = require("fs");
let openCalls = 0;
global.window = {{}};
global.Application = {{
  DisplayAlerts: 7,
  Documents: {{Count: 0, Open() {{openCalls += 1; throw new Error("opened");}}}},
  FontNames: {{Count: 0}}, Templates: {{Count: 0}}
}};
eval(fs.readFileSync({m0_asset}, "utf8"));
eval(fs.readFileSync({writer_asset}, "utf8"));
(async function () {{
  try {{
    await window.WPSComposerProbe.handleCommand({{
      method: "probe_longform_m0", params: {params_json}
    }});
    process.exit(2);
  }} catch (error) {{
    assert.equal(error.code, {json.dumps(expected_code)});
    assert.equal(openCalls, 0);
  }}
}})().catch(function (error) {{console.error(error); process.exit(1);}});
"""

    run_node(script)
