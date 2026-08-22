"""Tests for the macOS JSAPI long-form executor using mocks.

These tests never start WPS or bind loopback sockets. They verify bridge
command construction, plan serialization, add-in method routing, and
executor outcome validation.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Optional, Tuple

import json
import subprocess
import tempfile

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    GenerationOperation,
    GenerationPlan,
    GenerationResource,
)
from skills.WPSComposer.scripts.longform.executor import (
    ExecutionIssue,
    ExecutionOutcome,
    LongformExecutor,
    PaginationMap,
    PaginationNode,
    PaginationFragment,
)


@dataclass
class FakeBridgeCommand:
    component: str
    method: str
    params: Mapping[str, Any]
    id: str = "fake-cmd-1"


@dataclass
class FakeBridgeResult:
    ok: bool = True
    value: Mapping[str, Any] = None
    error: Optional[Mapping[str, Any]] = None

    def __post_init__(self):
        if self.value is None:
            self.value = {}


class FakeLoopbackBridge:
    """Records issued commands and returns predetermined results."""

    def __init__(self, result: Optional[FakeBridgeResult] = None):
        self.commands: list[FakeBridgeCommand] = []
        self._result = result or FakeBridgeResult(value={
            "outputPath": "/staged/output.docx",
            "appliedOperations": 3,
            "issueCodes": [],
            "paginationMap": {"version": "M2-stub", "nodes": []},
        })

    def issue(
        self, component: str, method: str, params: Mapping[str, Any]
    ) -> FakeBridgeCommand:
        cmd = FakeBridgeCommand(component=component, method=method, params=params)
        self.commands.append(cmd)
        return cmd

    def wait_result(self, command_id: str, timeout: float) -> FakeBridgeResult:
        return self._result


@pytest.fixture
def tmp_staging(tmp_path: Path) -> Path:
    staging = tmp_path / "staging"
    staging.mkdir()
    return staging


@pytest.fixture
def project_root() -> Path:
    return Path(__file__).parents[2]


def make_simple_plan() -> GenerationPlan:
    return GenerationPlan(
        component="writer",
        operations=(
            GenerationOperation(op="writer.reset", args={}, node_id="doc:reset"),
            GenerationOperation(
                op="writer.configure_section",
                args={
                    "role": "body",
                    "pageNumberFormat": "arabic",
                    "restartPageNumbering": True,
                    "startPageNumber": 1,
                    "headerText": "Header",
                    "footerText": "",
                    "linkToPreviousHeader": False,
                    "linkToPreviousFooter": False,
                },
                node_id="doc:section-1",
            ),
            GenerationOperation(
                op="writer.add_heading",
                args={"text": "Intro", "level": 1, "numbering": True, "numberingScheme": "decimal"},
                node_id="sec:1",
            ),
            GenerationOperation(
                op="writer.finalize_fields",
                args={"maxRounds": 3},
                node_id="doc:finalize",
            ),
        ),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )


def test_macos_executor_is_longform_executor(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
    executor = MacOSLongformExecutor(staging_dir=str(tmp_staging))
    assert isinstance(executor, LongformExecutor)


def test_execute_uses_generate_longform_document_method(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
    bridge = FakeLoopbackBridge()
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_staging))
    plan = make_simple_plan()
    outcome = executor.execute(plan, ())

    assert isinstance(outcome, ExecutionOutcome)
    assert len(bridge.commands) == 1
    cmd = bridge.commands[0]
    assert cmd.component == "writer"
    assert cmd.method == "generate_longform_document"


def test_execute_serializes_plan_into_command_params(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
    bridge = FakeLoopbackBridge()
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_staging))
    plan = make_simple_plan()
    executor.execute(plan, ())

    params = bridge.commands[0].params
    assert "plan" in params
    assert params["plan"]["component"] == "writer"
    assert params["plan"]["protocolVersion"] == 2
    assert len(params["plan"]["operations"]) == 4
    assert params["plan"]["operations"][0]["op"] == "writer.reset"


def test_execute_stages_resources_in_command_params(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
    bridge = FakeLoopbackBridge()
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_staging))
    plan = make_simple_plan()
    source = tmp_staging / "source.png"
    source.write_bytes(b"PNG")
    resource = GenerationResource(
        id="image-1",
        source_path=source,
        media_type="image/png",
    )
    executor.execute(plan, (resource,))

    params = bridge.commands[0].params
    assert "resources" in params
    assert Path(params["resources"]["image-1"]).parent == tmp_staging


def test_execute_returns_outcome_from_add_in_result(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
    result = FakeBridgeResult(value={
        "outputPath": "/staged/final.docx",
        "appliedOperations": 5,
        "issueCodes": [
            {"code": "DEGRADED", "message": "x", "placement": "document", "nodeId": "n:1"},
        ],
        "paginationMap": {
            "version": "M2-stub",
            "nodes": [
                {"nodeId": "sec:1", "pageStart": 1, "pageEnd": 1, "fragments": [{"page": 1}]},
            ],
        },
    })
    bridge = FakeLoopbackBridge(result=result)
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_staging))
    plan = make_simple_plan()
    outcome = executor.execute(plan, ())

    assert outcome.staged_artifact == "/staged/final.docx"
    assert len(outcome.issues) == 1
    assert outcome.issues[0].code == "DEGRADED"
    assert outcome.pagination_map.version == "M2-stub"
    assert len(outcome.pagination_map.nodes) == 1
    assert outcome.pagination_map.nodes[0].node_id == "sec:1"


def test_execute_raises_dedicated_host_error_when_bridge_missing(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSDedicatedHostUnavailableError,
        MACOS_DEDICATED_HOST_UNAVAILABLE,
    )
    executor = MacOSLongformExecutor(bridge=None, staging_dir=str(tmp_staging))
    plan = make_simple_plan()
    with pytest.raises(MacOSDedicatedHostUnavailableError) as exc_info:
        executor.execute(plan, ())
    assert exc_info.value.code == MACOS_DEDICATED_HOST_UNAVAILABLE


def test_execute_records_bridge_command_failure_as_issue(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
    result = FakeBridgeResult(
        ok=False,
        error={"code": "GENERATION_COMMAND_FAILED", "message": "add-in crashed"},
    )
    bridge = FakeLoopbackBridge(result=result)
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_staging))
    plan = make_simple_plan()
    outcome = executor.execute(plan, ())

    assert any(issue.code == "GENERATION_COMMAND_FAILED" for issue in outcome.issues)


def test_execute_validates_plan_before_sending(tmp_staging: Path):
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor
    bridge = FakeLoopbackBridge()
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_staging))
    plan = GenerationPlan(
        component="writer",
        operations=(GenerationOperation(op="writer.rogue", args={}),),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest="sha256:" + "0" * 64,
    )
    with pytest.raises(Exception):
        executor.execute(plan, ())
    assert bridge.commands == []


def test_addin_routes_generate_longform_document_to_v2(project_root: Path):
    """The writer.js component routes generate_longform_document to the v2 add-in."""
    addin_dir = project_root / "macos" / "wps-jsapi-probe" / "addin"
    writer_path = json.dumps(str(addin_dir / "writer.js"))
    js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{
  WPSComposerLongformV2: {{
    run: function(params) {{
      assert.equal(params.outputPath, "/staged/output.docx");
      assert.equal(params.plan.component, "writer");
      assert.equal(params.plan.operations.length, 1);
      return {{ outputPath: params.outputPath, appliedOperations: 1, issueCodes: [], paginationMap: {{ version: "M2-stub", nodes: [] }} }};
    }}
  }}
}};
global.Application = {{ DisplayAlerts: 7, ScreenUpdating: true }};
eval(fs.readFileSync({writer_path}, "utf8"));
(async function () {{
  const result = await window.WPSComposerProbe.handleCommand({{
    method: "generate_longform_document",
    params: {{
      outputPath: "/staged/output.docx",
      plan: {{ component: "writer", operations: [{{ op: "writer.reset", args: {{}}, nodeId: "doc:reset" }}] }}
    }}
  }});
  assert.equal(result.outputPath, "/staged/output.docx");
  assert.equal(result.appliedOperations, 1);
}})().catch(function (error) {{ console.error(error); process.exit(1); }});
"""
    path = Path(tempfile.mkdtemp()) / "route_test.js"
    path.write_text(js, encoding="utf-8")
    subprocess.run(["node", str(path)], check=True, capture_output=True, text=True)


def test_addin_applies_toc_density_from_configure_toc_styles(project_root: Path):
    """The add-in captures TOC density and applies it when insert_toc runs."""
    addin_dir = project_root / "macos" / "wps-jsapi-probe" / "addin"
    v2_path = json.dumps(str(addin_dir / "writer-longform-v2.js"))
    js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
let capturedTOCArgs = null;
function makeStyle(name) {{ return {{ Name: name, Font: {{ Size: null }}, ParagraphFormat: {{ SpaceBefore: null, SpaceAfter: null }} }}; }}
const styles = {{}};
const document = {{
  _wpscFirstSectionConfigured: false,
  Content: {{ End: 0, Text: "" }},
  PageSetup: {{}},
  Styles: {{
    Item: function(name) {{ return styles[name] = styles[name] || makeStyle(name); }},
    Add: function(name) {{ return styles[name] = makeStyle(name); }}
  }},
  TablesOfContents: {{ Add: function() {{ capturedTOCArgs = arguments; }} }},
  SaveAs2: function() {{}},
  Close: function() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  ScreenUpdating: true,
  Documents: {{ Add: function() {{ return document; }} }}
}};
eval(fs.readFileSync({v2_path}, "utf8"));
window.WPSComposerLongformV2.run({{
  outputPath: "/staged/output.docx",
  plan: {{
    component: "writer",
    operations: [
      {{op: "writer.configure_toc_styles", args: {{minFontSizePt: {{toc1: 10.5, toc2: 10, toc3: 9.5}}, minSpaceBeforePt: {{toc1: 0, toc2: 0, toc3: 0}}, minSpaceAfterPt: {{toc1: 0, toc2: 0, toc3: 0}}}}, nodeId: "doc:toc"}},
      {{op: "writer.insert_toc", args: {{title: "目录"}}, nodeId: "doc:toc"}}
    ]
  }}
}});
assert.equal(document.Styles.Item("TOC 1").Font.Size, 10.5);
assert.equal(document.Styles.Item("TOC 2").Font.Size, 10);
assert.equal(document.Styles.Item("TOC 3").Font.Size, 9.5);
"""
    path = Path(tempfile.mkdtemp()) / "toc_density_test.js"
    path.write_text(js, encoding="utf-8")
    subprocess.run(["node", str(path)], check=True, capture_output=True, text=True)


def test_addin_routes_add_heading_to_native_numbering(project_root: Path):
    """writer.add_heading applies native numbering via the v2 add-in."""
    addin_dir = project_root / "macos" / "wps-jsapi-probe" / "addin"
    v2_path = json.dumps(str(addin_dir / "writer-longform-v2.js"))
    js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
let linkedLevel = null;
function makeStyle() {{ return {{ Name: "", Font: {{}}, ParagraphFormat: {{}}, LinkToListTemplate: function(t, level) {{ linkedLevel = level; }} }}; }}
function makeListTemplate() {{ return {{ ListLevels: function(level) {{ return {{ NumberFormat: null }}; }} }}; }}
const document = {{
  _wpscFirstSectionConfigured: false,
  Content: {{ End: 0, Text: "" }},
  PageSetup: {{}},
  Styles: {{ Item: function(name) {{ return makeStyle(); }}, Add: function(name) {{ return makeStyle(); }} }},
  ListTemplates: {{ Add: function() {{ return makeListTemplate(); }} }},
  TablesOfContents: {{ Add: function() {{}} }},
  SaveAs2: function() {{}},
  Close: function() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  ScreenUpdating: true,
  Documents: {{ Add: function() {{ return document; }} }}
}};
eval(fs.readFileSync({v2_path}, "utf8"));
const doc = global.Application.Documents.Add();
window.WPSComposerLongformV2.run({{
  outputPath: "/staged/output.docx",
  plan: {{
    component: "writer",
    operations: [
      {{op: "writer.add_heading", args: {{text: "Intro", level: 1, numbering: true, numberingScheme: "decimal"}}, nodeId: "sec:1"}}
    ]
  }}
}});
assert.equal(linkedLevel, 1);
"""
    path = Path(tempfile.mkdtemp()) / "heading_native_test.js"
    path.write_text(js, encoding="utf-8")
    subprocess.run(["node", str(path)], check=True, capture_output=True, text=True)


def test_addin_handles_configure_front_matter(project_root: Path):
    """writer.configure_front_matter sets the page role in the add-in."""
    addin_dir = project_root / "macos" / "wps-jsapi-probe" / "addin"
    v2_path = json.dumps(str(addin_dir / "writer-longform-v2.js"))
    js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
let addedVarName = null;
let addedVarValue = null;
function makeSection() {{
  return {{
    Headers: {{ Item: makeHeader }},
    Footers: {{ Item: makeHeader }},
    Range: {{ DocumentVariables: {{ Add: function(name, value) {{ addedVarName = name; addedVarValue = value; }} }} }}
  }};
}}
function makeHeader() {{
  return {{ Range: {{ Text: "", ParagraphFormat: {{ Alignment: 0 }}, Collapse: function() {{}}, Fields: {{ Add: function() {{}} }} }} }};
}}
const document = {{
  _wpscFirstSectionConfigured: false,
  Content: {{ End: 0, Text: "" }},
  PageSetup: {{}},
  Sections: {{ Count: 1, Item: makeSection }},
  Styles: {{ Item: function() {{ return {{}}; }}, Add: function() {{ return {{}}; }} }},
  TablesOfContents: {{ Add: function() {{}} }},
  SaveAs2: function() {{}},
  Close: function() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  ScreenUpdating: true,
  Documents: {{ Add: function() {{ return document; }} }}
}};
eval(fs.readFileSync({v2_path}, "utf8"));
window.WPSComposerLongformV2.run({{
  outputPath: "/staged/output.docx",
  plan: {{
    component: "writer",
    operations: [
      {{op: "writer.configure_front_matter", args: {{role: "front_matter"}}, nodeId: "doc:front"}}
    ]
  }}
}});
assert.ok(addedVarName && addedVarName.indexOf("SectionRole_") !== -1);
assert.equal(addedVarValue, "front_matter");
"""
    path = Path(tempfile.mkdtemp()) / "front_matter_test.js"
    path.write_text(js, encoding="utf-8")
    subprocess.run(["node", str(path)], check=True, capture_output=True, text=True)


def test_addin_empty_footer_preserves_page_number(project_root: Path):
    """An empty footerText must not wipe the page-number field."""
    addin_dir = project_root / "macos" / "wps-jsapi-probe" / "addin"
    v2_path = json.dumps(str(addin_dir / "writer-longform-v2.js"))
    js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
let footerText = "initial";
let fieldAddedCount = 0;
function makeHeaderRange() {{ return {{ Text: "", ParagraphFormat: {{ Alignment: 0 }}, Collapse: function() {{}}, Fields: {{ Add: function() {{}} }} }}; }}
function makeFooterRange() {{ return {{ Text: footerText, ParagraphFormat: {{ Alignment: 0 }}, Collapse: function(dir) {{}}, Fields: {{ Add: function() {{ fieldAddedCount += 1; }} }} }}; }}
function makeSection() {{
  return {{
    Headers: {{ Item: function() {{ return {{ Range: makeHeaderRange() }}; }} }},
    Footers: {{ Item: function() {{ return {{ Range: makeFooterRange(), PageNumbers: {{ RestartNumberingAtSection: false, StartingNumber: 1, NumberStyle: 0 }} }}; }} }}
  }};
}}
const document = {{
  _wpscFirstSectionConfigured: false,
  Content: {{ End: 0, Text: "" }},
  PageSetup: {{}},
  Sections: {{ Count: 1, Item: makeSection }},
  Styles: {{ Item: function() {{ return {{}}; }}, Add: function() {{ return {{}}; }} }},
  TablesOfContents: {{ Add: function() {{}} }},
  SaveAs2: function() {{}},
  Close: function() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  ScreenUpdating: true,
  Documents: {{ Add: function() {{ return document; }} }}
}};
eval(fs.readFileSync({v2_path}, "utf8"));
window.WPSComposerLongformV2.run({{
  outputPath: "/staged/output.docx",
  plan: {{
    component: "writer",
    operations: [
      {{op: "writer.set_page_numbering", args: {{format: "arabic"}}, nodeId: "doc:pn"}},
      {{op: "writer.set_header_footer", args: {{headerText: "Header", footerText: ""}}, nodeId: "doc:hf"}}
    ]
  }}
}});
assert.notEqual(footerText, "");
assert.equal(fieldAddedCount, 1);
"""
    path = Path(tempfile.mkdtemp()) / "empty_footer_test.js"
    path.write_text(js, encoding="utf-8")
    subprocess.run(["node", str(path)], check=True, capture_output=True, text=True)


def test_addin_runs_field_convergence_before_save(project_root: Path):
    """Field convergence runs on the open document before SaveAs2/Close."""
    addin_dir = project_root / "macos" / "wps-jsapi-probe" / "addin"
    v2_path = json.dumps(str(addin_dir / "writer-longform-v2.js"))
    js = f"""
const fs = require("fs");
const assert = require("assert");
global.window = {{}};
let updateCount = 0;
let saveCount = 0;
const blankRange = {{
  Text: "",
  ParagraphFormat: {{ Alignment: 0 }},
  Collapse: function() {{}},
  Fields: {{ Add: function() {{}} }}
}};
const blankFooterRange = {{
  Text: "",
  ParagraphFormat: {{ Alignment: 0 }},
  Collapse: function() {{}},
  Fields: {{ Add: function() {{}} }},
  PageNumbers: {{ RestartNumberingAtSection: false, StartingNumber: 1, NumberStyle: 0 }}
}};
const header = {{ Range: blankRange }};
const footer = {{ Range: blankFooterRange }};
const section = {{
  Headers: {{ Item: function() {{ return header; }} }},
  Footers: {{ Item: function() {{ return footer; }} }}
}};
const document = {{
  _wpscFirstSectionConfigured: false,
  Content: {{ End: 0, Text: "" }},
  PageSetup: {{}},
  Sections: {{ Count: 1, Item: function() {{ return section; }} }},
  TablesOfContents: {{ Count: 0, Item: function(i) {{}}, Add: function() {{}} }},
  Fields: {{ Update: function() {{ updateCount += 1; }} }},
  ComputeStatistics: function() {{ return 1; }},
  SaveAs2: function() {{ saveCount += 1; assert.ok(updateCount > 0, "Fields.Update must run before SaveAs2"); }},
  Close: function() {{}}
}};
global.Application = {{
  DisplayAlerts: 7,
  ScreenUpdating: true,
  Documents: {{ Add: function() {{ return document; }} }}
}};
eval(fs.readFileSync({v2_path}, "utf8"));
const result = window.WPSComposerLongformV2.run({{
  outputPath: "/staged/output.docx",
  plan: {{
    component: "writer",
    operations: [
      {{op: "writer.finalize_fields", args: {{maxRounds: 3}}, nodeId: "doc:finalize"}}
    ]
  }}
}});
assert.equal(saveCount, 1);
assert.equal(result.fieldSnapshots.length, 4);
"""
    path = Path(tempfile.mkdtemp()) / "ordering_test.js"
    path.write_text(js, encoding="utf-8")
    subprocess.run(["node", str(path)], check=True, capture_output=True, text=True)

