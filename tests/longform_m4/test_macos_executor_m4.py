from __future__ import annotations

import hashlib
import json
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Optional

import pytest

from skills.WPSComposer.scripts.generation_plan import GenerationOperation, GenerationPlan
from skills.WPSComposer.scripts.longform.macos_executor import (
    MacOSLongformExecutor,
    MacOSLongformExecutorError,
)
from skills.WPSComposer.scripts.longform.resources import (
    ImageProfile,
    PreparedLongformResource,
)
from skills.WPSComposer.scripts.macos_probe.models import (
    ProbeResult,
    ProtocolError,
    validate_longform_generation_request,
    validate_longform_generation_value,
)


EMPTY_MANIFEST_DIGEST = (
    "sha256:dc7749a3af2a2bb77cad0700bddd3716d4b431bf91885cc20c6a1af68136f890"
)
ADDIN = Path("macos/wps-jsapi-probe/addin/writer-longform-v2.js")


def _manifest(resources: tuple[PreparedLongformResource, ...]) -> str:
    entries = [{
        "resourceId": item.id,
        "sourceSha256": item.source_sha256,
        "payloadSha256": item.payload_sha256,
        "byteLength": len(item.payload_bytes),
        "mediaType": item.media_type,
        "normalizerId": item.normalizer_id,
    } for item in resources]
    encoded = json.dumps(
        {"version": "1", "entries": sorted(entries, key=lambda item: item["resourceId"])},
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return "sha256:" + hashlib.sha256(encoded).hexdigest()


def _resource(payload: bytes = b"formula-image") -> PreparedLongformResource:
    return PreparedLongformResource(
        id="formula-image-1",
        media_type="image/png",
        source_sha256=hashlib.sha256(payload).hexdigest(),
        payload_sha256=hashlib.sha256(payload).hexdigest(),
        normalizer_id="none-v1",
        payload_bytes=payload,
        image_profile=ImageProfile(300, 100, 144.0, 144.0, 1, "PNG", False),
    )


def _equation(resource_id: Optional[str] = None) -> GenerationOperation:
    args: dict[str, Any] = {
        "renderMode": "native-m4",
        "content": {"nativeMath": {
            "syntax": "wps-linear-v1",
            "linearText": "x+y",
            "sourceHash": "2" * 64,
        }},
        "numbering": {
            "mode": "chapter", "sequenceId": "WPSC_EQ",
            "chapterStyleLevel": 1, "resetLevel": 1,
            "prefix": "(", "suffix": ")",
        },
        "bookmarkName": "wpsc_eq_" + "e" * 24,
        "fallbackText": "x+y",
    }
    if resource_id is not None:
        args["fallbackResource"] = {"fallbackResourceId": resource_id}
    return GenerationOperation(
        "writer.add_equation", args, node_id="eq:one",
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["EQUATION_INSERT_FAILED"],
            "fallback": "explicit-image-then-source-notice",
        },
    )


def _plan(
    *operations: GenerationOperation,
    resources: tuple[PreparedLongformResource, ...] = (),
) -> GenerationPlan:
    return GenerationPlan(
        component="writer",
        operations=(
            GenerationOperation(
                "writer.reserve_document_quality_anchor",
                {"title": "生成质量提示", "notices": []},
                node_id="doc:quality",
                failure_policy={"mode": "fail"},
            ),
            GenerationOperation(
                "writer.configure_section", {"role": "body"},
                node_id="doc:section:body",
            ),
            *operations,
            GenerationOperation(
            "writer.finalize_fields", {"maxRounds": 3}, node_id="doc:finalize"
            ),
        ),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest=_manifest(resources) if resources else EMPTY_MANIFEST_DIGEST,
    )


@dataclass
class _Command:
    id: str = "m4-command"


class _Bridge:
    def __init__(self, *, returned_path: Optional[str] = None) -> None:
        self.params: Optional[Mapping[str, Any]] = None
        self.returned_path = returned_path

    def issue(self, component: str, method: str, params: Mapping[str, Any]) -> _Command:
        assert (component, method) == ("writer", "generate_longform_document")
        self.params = params
        return _Command()

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        assert self.params is not None
        snapshot = {
            "stableKey": ["doc:finalize", "PAGE", 0],
            "fieldCategory": "page",
            "resultHash": hashlib.sha256(b"stable").hexdigest(),
            "tocPageCount": 0,
            "figureIndexPageCount": 0,
            "tableIndexPageCount": 0,
            "totalPages": 1,
        }
        return ProbeResult(command_id, True, {
            "outputPath": self.returned_path or self.params["outputPath"],
            "appliedOperations": len(self.params["plan"]["operations"]),
            "issueCodes": [],
            "fieldSnapshots": [[snapshot], [dict(snapshot)]],
            "childResults": [],
            "paginationMap": {"version": "M2-stub", "nodes": []},
        }, None)


class _FailureBridge(_Bridge):
    def __init__(self, error: Mapping[str, Any]) -> None:
        super().__init__()
        self.error = error

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        assert self.params is not None
        return ProbeResult(command_id, False, {}, self.error)


def test_closed_bridge_request_preserves_protocol_v2_shape() -> None:
    request = {
        "plan": _plan(_equation()).to_dict(),
        "outputPath": "/private/staged.docx",
        "resources": {},
    }
    assert validate_longform_generation_request(request) == request
    for changed in (
        {**request, "unexpected": True},
        {**request, "plan": {**request["plan"], "protocolVersion": 1}},
        {**request, "resources": {"r": 42}},
    ):
        with pytest.raises(ProtocolError, match="request is invalid"):
            validate_longform_generation_request(changed)


def test_formula_fallback_resource_must_bind_before_bridge(tmp_path: Path) -> None:
    bridge = _Bridge()
    with pytest.raises(MacOSLongformExecutorError, match="resource validation"):
        MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path)).execute(
            _plan(_equation("formula-image-1")), ()
        )
    assert bridge.params is None


def test_formula_fallback_is_staged_privately_and_cleaned(tmp_path: Path) -> None:
    resource = _resource()
    bridge = _Bridge()
    MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path)).execute(
        _plan(_equation(resource.id), resources=(resource,)), (resource,)
    )
    assert bridge.params is not None
    locator = Path(bridge.params["resources"][resource.id])
    assert locator.parent == tmp_path
    assert resource.id not in locator.name
    assert not locator.exists()


def test_bridge_response_cannot_redirect_output_path(tmp_path: Path) -> None:
    bridge = _Bridge(returned_path="/private/other.docx")
    with pytest.raises(MacOSLongformExecutorError, match="Bridge result was invalid"):
        MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path)).execute(
            _plan(_equation()), ()
        )


def test_changed_staged_payload_is_fatal_and_cleanup_still_runs(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    resource = _resource()
    bridge = _Bridge()
    executor = MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path))

    original_stage = executor._stage_resources

    def corrupt(resources, staged_docx):
        staged = original_stage(resources, staged_docx)
        staged[0][1].write_bytes(b"changed")
        return staged

    monkeypatch.setattr(executor, "_stage_resources", corrupt)
    with pytest.raises(MacOSLongformExecutorError, match="resource validation"):
        executor.execute(_plan(_equation(resource.id), resources=(resource,)), (resource,))
    assert bridge.params is None
    assert not tuple(tmp_path.glob("wpsc-resource-*"))


@pytest.mark.parametrize("private_code", [
    "/Users/alice/private.tex",
    "file:///tmp/private.tex",
    "A" * 64,
])
def test_bridge_error_code_and_detail_are_closed_and_privacy_safe(
    tmp_path: Path, private_code: str
) -> None:
    bridge = _FailureBridge({
        "code": private_code,
        "message": private_code + ":writer./Users/alice/private.tex",
    })
    with pytest.raises(MacOSLongformExecutorError) as captured:
        MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path)).execute(
            _plan(_equation()), ()
        )
    message = str(captured.value)
    assert "EXECUTION_ABORTED" in message
    assert private_code not in message
    assert "/Users/alice" not in message


def test_bridge_error_keeps_a_closed_stable_code_and_controlled_detail(
    tmp_path: Path,
) -> None:
    bridge = _FailureBridge({
        "code": "ENGINE_LOST",
        "message": "ENGINE_LOST:writer.OMathsAdd",
    })
    with pytest.raises(MacOSLongformExecutorError) as captured:
        MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path)).execute(
            _plan(_equation()), ()
        )
    assert str(captured.value).endswith(
        "(ENGINE_LOST: ENGINE_LOST:writer.OMathsAdd)"
    )


def _citation_degradation_plan() -> GenerationPlan:
    fallback = "[REFERENCE_UNRESOLVED 引用目标未解析]"
    return GenerationPlan(
        component="writer",
        operations=(
            GenerationOperation(
                "writer.reserve_document_quality_anchor",
                {"title": "生成质量提示", "notices": []},
                node_id="doc:quality", failure_policy={"mode": "fail"},
            ),
            GenerationOperation(
                "writer.configure_section", {"role": "body"},
                node_id="doc:section:body",
            ),
            GenerationOperation(
                "writer.add_cross_reference",
                {"runs": [
                    {"type": "text", "text": "before "},
                    {"type": "citation", "nodeId": "p/cite:0",
                     "targetId": "ref:a", "targetNodeId": "ref:a",
                     "number": 1, "fallbackText": "[1]"},
                    {"type": "text", "text": " then "},
                    {"type": "degradation", "nodeId": "p/cite:1",
                     "code": "REFERENCE_UNRESOLVED", "fallbackText": fallback},
                    {"type": "text", "text": " middle "},
                    {"type": "degradation", "nodeId": "p/cite:2",
                     "code": "REFERENCE_UNRESOLVED", "fallbackText": fallback},
                ]},
                node_id="p", failure_policy={
                    "mode": "degrade",
                    "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
                    "fallback": "inline-fallback",
                },
            ),
            GenerationOperation(
                "writer.configure_section", {"role": "bibliography"},
                node_id="doc:section:bibliography",
            ),
            GenerationOperation(
                "writer.add_bibliography",
                {"schemaVersion": 1, "entries": [{
                    "id": "ref:a", "nodeId": "ref:a", "number": 1,
                    "text": "Alpha.", "cited": True,
                }], "style": "numeric", "hangingIndentPt": 18.0,
                 "leftIndentPt": 18.0, "spaceAfterPt": 6.0},
                node_id="bib:block", failure_policy={
                    "mode": "degrade",
                    "recoverableCodes": ["BIBLIOGRAPHY_INSERT_FAILED"],
                    "fallback": "notice",
                },
            ),
            GenerationOperation(
                "writer.finalize_fields", {"maxRounds": 3},
                node_id="doc:finalize",
            ),
        ),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest=EMPTY_MANIFEST_DIGEST,
    )


def test_closed_result_preserves_citation_run_children_and_ownership() -> None:
    plan = _citation_degradation_plan()
    fallback_children = [
        {"nodeId": "p/cite:0", "status": "applied"},
        {"nodeId": "p/cite:1", "status": "degraded",
         "issueCode": "REFERENCE_UNRESOLVED"},
        {"nodeId": "p/cite:2", "status": "degraded",
         "issueCode": "REFERENCE_UNRESOLVED"},
    ]
    value = validate_longform_generation_value({
        "outputPath": "/private/staged.docx",
        "appliedOperations": len(plan.operations),
        "issueCodes": [{
            "code": "REFERENCE_UNRESOLVED",
            "message": "Citation used its planned fallback",
            "placement": "inline", "nodeId": node_id,
        } for node_id in ("p/cite:1", "p/cite:2")],
        "childResults": fallback_children,
        "paginationMap": {"version": "M2-stub", "nodes": []},
    })
    MacOSLongformExecutor._validate_result_references(value, plan)

    for changed in (
        {**value, "childResults": [
            {**fallback_children[0], "nodeId": "p/cite:1"},
            fallback_children[1],
            fallback_children[2],
        ]},
        {**value, "childResults": [
            fallback_children[0],
            {**fallback_children[1], "status": "applied", "issueCode": None},
            fallback_children[2],
        ]},
        {**value, "issueCodes": [
            {**value["issueCodes"][0], "nodeId": "p"},
            value["issueCodes"][1],
        ]},
    ):
        with pytest.raises(MacOSLongformExecutorError, match="result was invalid"):
            MacOSLongformExecutor._validate_result_references(changed, plan)


def test_addin_exposes_m4_native_handlers_and_bibliography_is_not_deferred() -> None:
    source = ADDIN.read_text(encoding="utf-8")
    for name in (
        "validateLongformRequest",
        "addEquationNativeM4",
        "addFormulaNativeFallback",
        "addCitationParagraph",
        "addBibliographyNative",
        "applyTableCellMetadata",
    ):
        assert f"function {name}" in source
    deferred = source.split("const LONGFORM_DEFERRED =", 1)[1].split("};", 1)[0]
    assert "writer.add_bibliography" not in deferred
    assert 'args.renderMode === "native-m4"' in source
    assert "OMaths.Add" in source
    assert ".BuildUp()" in source
    assert "_wpscNumericFormulaDebug" not in source
    assert "_wpscFieldNumericDebug" not in source
    assert "_wpscFormulaLocalDebug" not in source
    assert "failure.message =" not in source


def test_addin_m4_paths_remain_controller_owned_and_m3_formula_is_explicit() -> None:
    source = ADDIN.read_text(encoding="utf-8")
    equation_body = source.split("function addEquationNumberNative", 1)[1].split("\n  }", 1)[0]
    assert "OMaths" not in equation_body
    assert "fallbackResourcePlannedDegradation" in source
    assert "FORMULA_FALLBACK_IMAGE_UNAVAILABLE" in source
    assert "cellDegradations" in source
    assert "cellCitations" in source


def _run_node(body: str) -> None:
    asset = json.dumps(str(ADDIN.resolve()))
    subprocess.run(
        ["node", "-e", f'''const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
{body}
'''],
        check=True,
        capture_output=True,
        text=True,
    )


def test_js_native_formula_builds_one_omath_and_retains_number_shell() -> None:
    _run_node(r'''
let text = "";
const builds = [];
const added = [];
const bookmarks = [];
const paragraphFormats = [];
function makeRange(start, end) {
  const tabStops = [];
  const paragraphFormat = {TabStops: {Add: function(position, alignment, leader) {
    tabStops.push([position, alignment, leader]);
  }}, _tabStops: tabStops};
  paragraphFormats.push(paragraphFormat);
  return {
    Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: paragraphFormat,
    get Text() { return text.slice(start, end); },
    set Text(value) { text = text.slice(0, start) + String(value) + text.slice(end); },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() { text = text.slice(0, start) + text.slice(end); }
  };
}
const maths = {
  items: [],
  Count: 0,
  Add: function(range) {
    added.push([range.Start, range.End, range.Text]);
    this.Count += 1;
    const native = {Range: {Start: range.Start, End: range.End}, BuildUp: function() { builds.push(1); }};
    this.items.push(native);
    return {Start: range.Start, End: range.End, OMaths: {Count: 1, Item: function() { return native; }}};
  },
  Item: function(index) { return this.items[index - 1]; }
};
const document = {
  get Content() { return {End: text.length + 1, get Text() { return text; }}; },
  Range: makeRange,
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: maths,
  Fields: {Add: function(target) { const fieldStart = target.End;
    target.InsertAfter("1");
    return {Update: function(){}, Result: {Text: "1", Start: fieldStart, End: target.End}};
  }},
  Bookmarks: {Add: function(name) { bookmarks.push(name); }}
};
const args = {
  renderMode: "native-m4",
  content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y", sourceHash: "2".repeat(64)}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", chapterStyleLevel: null, resetLevel: null, prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"
};
window.WPSComposerLongformV2.__test.addEquationNativeM4(
  document, args, {}, {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true}
);
assert.deepEqual(added, [[1, 4, "x+y"]]);
assert.equal(builds.length, 1);
assert.equal(maths.Count, 1);
assert.equal(bookmarks.length, 1);
assert.ok(text.includes("\tx+y\t("));
const formulaFormat = paragraphFormats.find(function(format) { return format._tabStops.length === 2; });
assert.ok(formulaFormat);
assert.equal(formulaFormat.Alignment, 0);
assert.equal(formulaFormat.KeepTogether, -1);
assert.deepEqual(formulaFormat._tabStops, [[233.5, 1, 0], [467, 2, 0]]);
''')


def test_js_native_formula_preserves_canonical_linear_text_and_requires_complete_professional_structure() -> None:
    _run_node(r'''
function exercise(linearText, functionTypes) {
  let text = "", addedText = null;
  function range(start, end) {
    return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
      ParagraphFormat: {TabStops: {Add: function(){}}},
      get Text() { return text.slice(this.Start, this.End); },
      InsertAfter: function(value) {
        value = String(value);
        text = text.slice(0, this.End) + value + text.slice(this.End);
        this.End += value.length;
      }};
  }
  const native = {Range: null, Functions: {Count: functionTypes.length,
    Item: function(index) { return {Type: functionTypes[index - 1]}; }},
    BuildUp: function(){}};
  const maths = {Count: 0, Add: function(target) {
    addedText = target.Text;
    native.Range = {Start: target.Start, End: target.End};
    this.Count += 1;
    return {Start: target.Start, End: target.End,
      OMaths: {Count: 1, Item: function() { return native; }}};
  }};
  const document = {
    get Content() { return {End: text.length + 1, get Text() { return text; }}; },
    Range: range, OMaths: maths,
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
    Fields: {Add: function(target) { const start = target.End; target.InsertAfter("1");
      return {Update: function(){}, Result: {Text: "1", Start: start, End: target.End}};
    }},
    Bookmarks: {Add: function(){}}
  };
  const args = {content: {nativeMath: {syntax: "wps-linear-v1", linearText: linearText}},
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: linearText};
  return {run: function() {
    window.WPSComposerLongformV2.__test.addEquationNativeM4(
      document, args, {},
      {ownerNodeId: "eq:matrix", issues: [], childResults: [], controllerOwned: true}
    );
    return addedText;
  }};
}
assert.equal(exercise("(■(a&b@c&d))", [12, 5]).run(), "(■(a&b@c&d))");
assert.throws(function() { exercise("■(a&b@c&d)", [12]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise('"a/b^c_d"', [20]).run(), '"a/b^c_d"');
assert.throws(function() { exercise("x^2", []).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.throws(function() { exercise("x^2", [20]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.throws(function() { exercise("x^2+y^3", [19]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise("x^2+y^3", [19, 19]).run(), "x^2+y^3");
assert.throws(function() { exercise("√(x)+√(y)", [16]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise("√(x)+√(y)", [16, 16]).run(), "√(x)+√(y)");
assert.throws(function() { exercise("(a)/(b)+(c)/(d)", [7]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise("(a)/(b)+(c)/(d)", [7, 7]).run(), "(a)/(b)+(c)/(d)");
assert.throws(function() { exercise("∑_i^n i+∏_j^m j", [13]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise("∑_i^n i+∏_j^m j", [13, 13]).run(), "∑_i^n i+∏_j^m j");
assert.equal(exercise("(√(x_i^2))/(∑_j^n j)", [7, 16, 18, 13]).run(),
  "(√(x_i^2))/(∑_j^n j)");
assert.throws(function() { exercise("(■(a&b@c&d))", [12]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise("{■(x&x>0@-x&x≤0)", [12, 5]).run(),
  "{■(x&x>0@-x&x≤0)");
assert.throws(function() { exercise("((n)¦(k))", []).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise("((n)¦(k))", [7, 5]).run(), "((n)¦(k))");
assert.throws(function() { exercise("((n)¦(k))+((a)¦(b))", [7, 5]).run(); },
  function(error) { return error.code === "EQUATION_INSERT_FAILED"; });
assert.equal(exercise("((n)¦(k))+((a)¦(b))", [7, 7, 5, 5]).run(),
  "((n)¦(k))+((a)¦(b))");
''')


def test_js_formula_controller_attempts_image_once_then_places_source_notice_and_continues() -> None:
    _run_node(r'''
let text = "";
let imageAttempts = 0;
let bookmarkNames = [];
let nativeMathCount = 0;
const selection = {Document: null, Range: null, OMaths: null};
function makeRange(start, end) {
  const value = {
    Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
    ParagraphFormat: {TabStops: {Add: function(){}}},
    OMaths: {Add: function(target) {
      nativeMathCount += 1;
      const native = {Range: {Start: target.Start, End: target.End},
        Functions: {Count: 1, Item: function() { return {Type: 20}; }},
        BuildUp: function(){}};
      const added = makeRange(target.Start, target.End);
      added.OMaths = {Count: 1, Item: function() { return native; }};
      return added;
    }},
    Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
    get Text() { return text.slice(start, end); },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() { text = text.slice(0, start) + text.slice(end);
      bookmarkNames = []; nativeMathCount = 0; }
  };
  return value;
}
const document = {
  Name: "Recovery.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: makeRange,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: {get Count() { return nativeMathCount; },
    Add: function() { throw new Error("global add forbidden"); },
    Item: function() { throw new Error("global item forbidden"); }},
  InlineShapes: {AddPicture: function() { imageAttempts += 1;
    const error = new Error("image failed"); error.code = "IMAGE_INSERT_FAILED"; throw error; }},
  Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }},
  Bookmarks: {get Count() { return bookmarkNames.length; },
    Add: function(name) { bookmarkNames.push(name); }},
  _wpscRecoveryController: window.WPSComposerLongformV2.__test.createLocalRecoveryController()
};
selection.Document = document;
const issues = [], children = [];
const equation = {
  op: "writer.add_equation", nodeId: "eq:one",
  args: {renderMode: "native-m4", content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x^2", sourceHash: "2".repeat(64)}},
    fallbackResource: {fallbackResourceId: "formula-image-1"}, fallbackText: "x^2",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", chapterStyleLevel: null, resetLevel: null, prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"], fallback: "explicit-image-then-source-notice"}
};
window.WPSComposerLongformV2.__test.runOperation(document, equation, {"formula-image-1": "/private/staged.png"}, issues, children);
window.WPSComposerLongformV2.__test.runOperation(document, {op: "writer.add_paragraph", nodeId: "p:later", args: {text: "later"}}, {}, issues, children);
assert.equal(imageAttempts, 1);
assert.equal(issues.filter(x => x.code === "EQUATION_INSERT_FAILED").length, 1);
assert.ok(text.includes("x^2"));
assert.ok(text.includes("later"));
assert.ok(!JSON.stringify(issues).includes("/private"));
assert.equal(nativeMathCount, 0);
assert.equal(document._wpscNativeFields.length, 1);
assert.equal(document.Bookmarks.Count, 1);
''')


def test_js_partial_professional_formula_degrades_then_keeps_number_reference_and_later_flow() -> None:
    _run_node(r'''
function exercise(imageSucceeds, firstFunctionType, expectLatch) {
  let text = "", mathCount = 0, mathAttempts = 0, shapeCount = 0, fieldCount = 0;
  let fieldRefreshes = 0;
  const bookmarkEntries = Object.create(null);
  const selection = {Document: null, Range: null, OMaths: null};
  function paragraphCount() { return 1 + (text.match(/\r/g) || []).length; }
  function range(start, end) {
    const value = {Start: start, End: end,
      Font: {Italic: 0, Color: 0},
      Shading: {BackgroundPatternColor: 0},
      ParagraphFormat: {TabStops: {Add: function(){}}},
      OMaths: null,
      Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
      get Text() { return text.slice(this.Start, this.End); },
      InsertAfter: function(inserted) {
        inserted = String(inserted);
        text = text.slice(0, this.End) + inserted + text.slice(this.End);
        this.End += inserted.length;
      },
      Delete: function() {
        text = text.slice(0, this.Start) + text.slice(this.End);
        mathCount = 0; shapeCount = 0; fieldCount = 0;
        Object.keys(bookmarkEntries).forEach(function(name) { delete bookmarkEntries[name]; });
      }};
    value.OMaths = {Add: function(target) {
      mathAttempts += 1;
      mathCount += 1;
      const native = {Range: {Start: target.Start, End: target.End},
        Functions: {Count: 1, Item: function() {
          return {Type: mathAttempts === 1 ? firstFunctionType : 20};
        }},
        BuildUp: function(){}};
      const added = range(target.Start, target.End);
      added.OMaths = {Count: 1, Item: function() { return native; }};
      return added;
    }};
    return value;
  }
  const document = {
    Name: imageSucceeds ? "Image.docx" : "Source.docx",
    ActiveWindow: {Selection: selection},
    get Content() { return {End: text.length + 1, get Text() { return text; }}; },
    Range: range,
    get Paragraphs() { return {Count: paragraphCount(), Item: function() {
      const last = text.lastIndexOf("\r");
      return {Range: {Start: last < 0 ? 0 : last + 1, End: text.length + 1}};
    }}; },
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
    OMaths: {get Count() { return mathCount; },
      Add: function() { throw new Error("global OMath add forbidden"); },
      Item: function() { throw new Error("global OMath item forbidden"); }},
    InlineShapes: {get Count() { return shapeCount; }, AddPicture: function(a, b, c, target) {
      if (!imageSucceeds) {
        const failure = new Error("image failed");
        failure.code = "IMAGE_INSERT_FAILED";
        throw failure;
      }
      const start = target.End;
      target.InsertAfter("I");
      shapeCount += 1;
      return {Range: {Start: start, End: target.End, ParagraphFormat: {}}};
    }},
    Fields: {get Count() { return fieldCount; }, Add: function(target) {
      const start = target.End;
      target.InsertAfter("1");
      fieldCount += 1;
      return {Update: function() { fieldRefreshes += 1; },
        Range: {Start: start, End: target.End},
        Result: {Text: "1", Start: start, End: target.End}};
    }},
    Bookmarks: {get Count() { return Object.keys(bookmarkEntries).length; },
      Add: function(name) { bookmarkEntries[name] = {Range: {Fields: {Count: 1}}}; },
      Exists: function(name) { return Boolean(bookmarkEntries[name]); },
      Item: function(name) { return bookmarkEntries[name]; }}
  };
  selection.Document = document;
  const bookmarkName = "wpsc_eq_" + "e".repeat(24);
  const plainBookmarkName = "wpsc_eq_" + "f".repeat(24);
  const issues = [], children = [];
  const equation = {op: "writer.add_equation", nodeId: "eq:one", args: {
    renderMode: "native-m4",
    content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x^2+y^3"}},
    fallbackResource: {fallbackResourceId: "image"}, fallbackText: "x^2+y^3",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: bookmarkName},
    failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
      fallback: "explicit-image-then-source-notice"}};
  const api = window.WPSComposerLongformV2.__test;
  api.runOperation(document, equation, {image: "/private/staged.png"}, issues, children);
  api.runOperation(document, {op: "writer.add_equation", nodeId: "eq:plain", args: {
    renderMode: "native-m4",
    content: {nativeMath: {syntax: "wps-linear-v1", linearText: "α+β"}},
    fallbackText: "α+β",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: plainBookmarkName},
    failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
      fallback: "explicit-image-then-source-notice"}}, {}, issues, children);
  api.runOperation(document, {op: "writer.add_cross_reference", nodeId: "p:ref", args: {runs: [
    {type: "reference", prefix: "(", bookmarkName: bookmarkName,
      suffix: ")", fallbackText: "1"},
    {type: "text", text: " then "},
    {type: "reference", prefix: "(", bookmarkName: plainBookmarkName,
      suffix: ")", fallbackText: "2"}
  ]}, failurePolicy: {mode: "degrade", recoverableCodes: ["CROSS_REFERENCE_FAILED"],
    fallback: "inline-fallback"}}, {}, issues, children);
  api.runOperation(document, {op: "writer.add_paragraph", nodeId: "p:later",
    args: {text: "later"}}, {}, issues, children);
  assert.equal(mathCount, expectLatch ? 0 : 1,
    "the failed formula is rolled back while a later formula may remain native");
  assert.equal(mathAttempts, expectLatch ? 1 : 2,
    "only a text-only professional no-op latches for the run");
  assert.equal(document.Bookmarks.Count, 2, "only the two fallback number bookmarks remain");
  assert.equal(fieldCount, 4, "two fallback numbers and two REF fields remain");
  assert.equal(document._wpscNativeFields.length, 4);
  assert.equal(document._wpscNativeFields[0].fieldKind, "SEQ_EQ");
  assert.equal(document._wpscNativeFields[1].fieldKind, "SEQ_EQ");
  assert.equal(document._wpscNativeFields[2].fieldKind, "REF");
  assert.equal(document._wpscNativeFields[3].fieldKind, "REF");
  document._wpscNativeFields[2].native.Update();
  assert.equal(fieldRefreshes, 1);
  assert.equal(issues.filter(function(issue) {
    return issue.code === "EQUATION_INSERT_FAILED";
  }).length, expectLatch ? 2 : 1);
  assert.ok(text.includes("later"));
  assert.ok(text.includes("(1)"));
  assert.equal(text.includes("[EQUATION_INSERT_FAILED: α+β]"), expectLatch);
  return {shapeCount: shapeCount, text: text};
}
const image = exercise(true, 20, true);
assert.equal(image.shapeCount, 1);
assert.ok(image.text.includes("formula image fallback"));
const source = exercise(false, 20, true);
assert.equal(source.shapeCount, 0);
assert.ok(source.text.includes("[EQUATION_INSERT_FAILED: x^2+y^3]"));
const partialStructure = exercise(false, 19, false);
assert.equal(partialStructure.shapeCount, 0);
assert.ok(partialStructure.text.includes("α+β"));
''')


def test_js_citations_bibliography_and_table_metadata_are_paragraph_and_cell_local() -> None:
    _run_node(r'''
let text = "";
const formats = [];
function makeRange(start, end) {
  const format = {};
  formats.push(format);
  return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: format,
    InsertAfter: function(value) { text += String(value); this.End = text.length; }};
}
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: makeRange};
const context = {ownerNodeId: "p:one", issues: [], childResults: [], controllerOwned: true};
window.WPSComposerLongformV2.__test.addCitationParagraph(document, {runs: [
  {type: "text", text: "See "},
  {type: "citation", nodeId: "p/c:1", targetId: "a", targetNodeId: "ref:a", number: 1, fallbackText: "[1]"},
  {type: "text", text: "."}
]}, {}, context);
window.WPSComposerLongformV2.__test.addBibliographyNative(document, {
  schemaVersion: 1, style: "numeric", hangingIndentPt: 18, leftIndentPt: 18, spaceAfterPt: 6,
  entries: [{id: "a", nodeId: "ref:a", number: 1, text: "Alpha.", cited: true},
            {id: "b", nodeId: "ref:b", number: 2, text: "Beta.", cited: false}]
});
assert.equal(text, "See [1].\r[1] Alpha.\r[2] Beta.\r");
assert.equal(formats.filter(f => f.LeftIndent === 18 && f.FirstLineIndent === -18 && f.SpaceAfter === 6).length, 2);
const cells = {"2:1": {Range: {Text: "[1]", Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}}},
               "2:2": {Range: {Text: "[REFERENCE_UNRESOLVED \u5f15\u7528\u76ee\u6807\u672a\u89e3\u6790]", Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}}}};
const table = {Cell: function(row, col) { return cells[row + ":" + col]; }};
window.WPSComposerLongformV2.__test.applyTableCellMetadata(table, {
  cellCitations: [{row: 2, column: 1}],
  cellDegradations: [{row: 2, column: 2}]
});
assert.equal(cells["2:1"].Range.Text, "[1]");
assert.equal(cells["2:2"].Range.Text, "[REFERENCE_UNRESOLVED \u5f15\u7528\u76ee\u6807\u672a\u89e3\u6790]");
assert.equal(cells["2:1"].Range.Font.Italic, 0);
assert.equal(cells["2:2"].Range.Font.Italic, -1);
assert.ok(cells["2:2"].Range.Shading.BackgroundPatternColor !== undefined);
''')


def test_js_m4_request_rejects_protocol_capability_and_unknown_shape() -> None:
    _run_node(r'''
const validate = window.WPSComposerLongformV2.__test.validateLongformRequest;
const operation = {op: "writer.add_equation", args: {renderMode: "native-m4"}};
const base = {plan: {component: "writer", protocolVersion: 2, operations: [operation]}, outputPath: "/private/out.docx", resources: {}};
assert.equal(validate(base), base);
assert.throws(() => validate({...base, unexpected: true}), error => error.code === "PROTOCOL_MISMATCH");
assert.throws(() => validate({...base, plan: {...base.plan, protocolVersion: 1}}), error => error.code === "PROTOCOL_MISMATCH");
const document = {Content: {End: 1}, Range: function() { return {Start: 0, End: 0, InsertAfter: function(){}, Delete: function(){}}; }};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(
  document, {content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x"}}}, {},
  {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true}
), error => error.code === "CAPABILITY_MISMATCH");
''')


def test_js_formula_selection_requires_target_document_window_and_owner() -> None:
    _run_node(r'''
function makeDocument(mode) {
  let text = "", addCalls = 0;
  const selection = {Document: {Name: "Foreign.docx"}, Range: null, OMaths: null};
  function range(start, end) {
    const reportedStart = mode === "clamped-input" && start !== end ? start + 1 : start;
    const value = {Start: reportedStart, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
      ParagraphFormat: {TabStops: {Add: function(){}}},
      OMaths: {Add: function() { addCalls += 1; return null; }},
      Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
      get Text() { return text.slice(this.Start, this.End); },
      InsertAfter: function(raw) {
        const inserted = String(raw);
        text = text.slice(0, this.End) + inserted + text.slice(this.End);
        this.End += inserted.length;
      },
      Delete: function() { text = text.slice(0, this.Start) + text.slice(this.End); }
    };
    return value;
  }
  const document = {Name: "Target.docx",
    get Content() { return {End: text.length + 1, get Text() { return text; }}; },
    Range: range, Paragraphs: {Count: 1, Item: function() {
      return {Range: {Start: 0, End: text.length + 1}};
    }}, PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
    OMaths: {Count: 0, Add: function() { throw new Error("global add forbidden"); }}}
  if (mode !== "missing-window") document.ActiveWindow = {Selection: selection};
  if (mode === "clamped-input") selection.Document = document;
  document._wpscRunOwnsAppendCursor = true;
  document._wpscAppendCursorRange = range(0, 0);
  return {document: document, addCalls: function() { return addCalls; }};
}
const args = {content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"};
const context = {ownerNodeId: "eq:owner", issues: [], childResults: [], controllerOwned: true};
for (const mode of ["foreign-owner", "missing-window", "clamped-input"]) {
  const state = makeDocument(mode);
  assert.throws(function() {
    window.WPSComposerLongformV2.__test.addEquationNativeM4(
      state.document, args, {}, context
    );
  }, function(error) { return error.code === "CAPABILITY_MISMATCH"; });
  assert.equal(state.addCalls(), 0);
}
''')


def test_js_omath_verification_rejects_escape_and_does_not_normalize_unknown_api_errors() -> None:
    _run_node(r'''
let text = "";
function range(start, end) { return {Start: start, End: end,
  ParagraphFormat: {TabStops: {Add: function(){}}},
  get Text() { return text.slice(start, end); },
  InsertAfter: function(value) { text += String(value); }}; }
const args = {content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"};
const context = {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true};
const native = {Range: {Start: 1, End: 4}, BuildUp: function() { this.Range.End = 99; }};
const escaping = {Count: 0, Item: function() { return native; },
  Add: function() { this.Count = 1; return {Start: 1, End: 4, OMaths: {Count: 1, Item: function() { return native; }}}; }};
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range, OMaths: escaping,
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(document, args, {}, context),
  error => error.code === "EQUATION_INSERT_FAILED");
const unknown = new Error("unknown OMath getter failure");
text = "";
document.OMaths = {get Count() { throw unknown; }, Add: function() {}};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(document, args, {}, context),
  error => error === unknown);
''')


def test_js_planned_formula_uses_validated_image_without_entering_omath() -> None:
    _run_node(r'''
let text = "";
let imageAttempts = 0;
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
  ParagraphFormat: {TabStops: {Add: function(){}}},
  get Text() { return text.slice(start, end); },
  InsertAfter: function(value) { value = String(value); text += value; this.End += value.length; },
  Delete: function() {}}; }
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  InlineShapes: {AddPicture: function(locator) { imageAttempts += 1; assert.equal(locator, "/private/formula.png");
    return {Range: {Start: 1, End: 2, ParagraphFormat: {}}}; }},
  Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }},
  Bookmarks: {Add: function() {}},
  get OMaths() { throw new Error("planned content must not enter OMath"); }};
const context = {ownerNodeId: "eq:planned", issues: [], childResults: [], controllerOwned: true};
window.WPSComposerLongformV2.__test.addEquationNativeM4(document, {
  renderMode: "native-m4",
  content: {plannedDegradation: {code: "FORMULA_MALFORMED", placement: "block", fallbackKind: "source"}},
  fallbackResource: {fallbackResourceId: "formula-image-1"}, fallbackText: "x+y",
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)
}, {"formula-image-1": "/private/formula.png"}, context);
assert.equal(imageAttempts, 1);
assert.equal(context.issues.length, 1);
assert.equal(context.issues[0].code, "FORMULA_MALFORMED");
assert.ok(text.includes("formula image fallback"));
const noticeAt = text.indexOf("formula image fallback");
assert.ok(noticeAt > 0);
assert.ok(text.slice(0, noticeAt).includes("\r"));
assert.ok(text.indexOf("\t") < text.indexOf("\r"));
''')


def test_js_degradation_styling_is_scoped_to_the_exact_visible_span() -> None:
    _run_node(r'''
let text = "";
const normal = {italic: 0, color: 17, background: 23, texture: 0};
const style = Object.assign({}, normal);
function styledRange(start, end) {
  const font = {};
  Object.defineProperties(font, {
    Italic: {get: function() { return style.italic; }, set: function(value) { style.italic = value; }},
    Color: {get: function() { return style.color; }, set: function(value) { style.color = value; }}
  });
  const shading = {};
  Object.defineProperties(shading, {
    BackgroundPatternColor: {get: function() { return style.background; },
      set: function(value) { style.background = value; }},
    Texture: {get: function() { return style.texture; }, set: function(value) { style.texture = value; }}
  });
  return {Start: start, End: end, Font: font, Shading: shading, ParagraphFormat: {},
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) { value = String(value); text += value; this.End += value.length; }};
}
const document = {
  get Content() { return {End: text.length + 1, get Text() { return text; }}; },
  Range: styledRange
};
const context = {ownerNodeId: "p:one", issues: [], childResults: [], controllerOwned: true};
window.WPSComposerLongformV2.__test.addCitationParagraph(document, {runs: [
  {type: "degradation", nodeId: "p:one/c:0", code: "REFERENCE_UNRESOLVED",
    fallbackText: "[REFERENCE_UNRESOLVED 引用目标未解析]"}
]}, {}, context);
assert.deepEqual(style, normal);
window.WPSComposerLongformV2.__test.addBibliographyNative(document, {
  schemaVersion: 1, style: "numeric", hangingIndentPt: 18, leftIndentPt: 18,
  spaceAfterPt: 6,
  entries: [{id: "a", nodeId: "ref:a", number: 1, text: "Alpha.", cited: true}]
});
assert.deepEqual(style, normal);
assert.ok(text.endsWith("[1] Alpha.\r"));
''')


def test_js_inline_degradation_styles_retrospectively_after_paragraph_boundary() -> None:
    _run_node(r'''
let text = "";
let paragraphLatch = {italic: 0, color: 17, background: 23};
let nextParagraphStyle = null;
let noticeStyled = 0;
const selection = {Document: null, Range: null};
function range(start, end) {
  const font = {};
  const shading = {};
  function record(name, value) {
    if (start < end) {
      noticeStyled += 1;
      if (text.indexOf("\r") < 0) paragraphLatch[name] = value;
    }
  }
  Object.defineProperties(font, {
    Italic: {get: function() { return paragraphLatch.italic; },
      set: function(value) { record("italic", value); }},
    Color: {get: function() { return paragraphLatch.color; },
      set: function(value) { record("color", value); }}
  });
  Object.defineProperty(shading, "BackgroundPatternColor", {
    get: function() { return paragraphLatch.background; },
    set: function(value) { record("background", value); }
  });
  return {Start: start, End: end, Document: document,
    Font: font, Shading: shading, ParagraphFormat: {},
    Select: function() { selection.Range = this; selection.Document = document; },
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
      if (value.indexOf("\r") >= 0) {
        nextParagraphStyle = Object.assign({}, paragraphLatch);
      }
    }};
}
const document = {Name: "Deferred-style.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; },
  Range: range};
selection.Document = document;
document._wpscRunOwnsAppendCursor = true;
document._wpscAppendCursorRange = document.Range(0, 0);
window.WPSComposerLongformV2.__test.addCitationParagraph(document, {runs: [
  {type: "text", text: "before "},
  {type: "degradation", nodeId: "p/c:0", code: "REFERENCE_UNRESOLVED",
    fallbackText: "[REFERENCE_UNRESOLVED 引用目标未解析]"},
  {type: "text", text: " after."}
]}, {}, {ownerNodeId: "p", issues: [], childResults: [], controllerOwned: true});
assert.ok(noticeStyled >= 3, "the exact notice span must be styled");
assert.deepEqual(nextParagraphStyle, {italic: 0, color: 17, background: 23});
assert.equal(paragraphLatch.italic, 0);
assert.equal(paragraphLatch.color, 17);
assert.equal(paragraphLatch.background, 23);
''')


def test_js_degradation_style_restore_is_target_local_owner_bound_and_closed() -> None:
    _run_node(r'''
let text = "BODY";
const writes = [];
const selection = {Document: null, Range: null};
function range(start, end) {
  const anchorStyle = start <= 1
    ? {italic: 0, color: 11, background: 22}
    : {italic: -1, color: 33, background: 44};
  const font = {Italic: anchorStyle.italic, Color: anchorStyle.color};
  const shading = {BackgroundPatternColor: anchorStyle.background};
  return {Start: start, End: end, Font: font, Shading: shading, ParagraphFormat: {},
    Select: function() { selection.Range = this; },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    }};
}
const document = {Name: "Style.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; },
  Range: function(start, end) {
    const value = range(start, end);
    const font = value.Font, shading = value.Shading;
    Object.defineProperties(font, {
      Italic: {get: function() { return this._italic; }, set: function(raw) {
        this._italic = raw; writes.push([start, "italic", raw]); }},
      Color: {get: function() { return this._color; }, set: function(raw) {
        this._color = raw; writes.push([start, "color", raw]); }}
    });
    font._italic = start <= 1 ? 0 : -1;
    font._color = start <= 1 ? 11 : 33;
    Object.defineProperty(shading, "BackgroundPatternColor", {
      get: function() { return this._background; }, set: function(raw) {
        this._background = raw; writes.push([start, "background", raw]); }
    });
    shading._background = start <= 1 ? 22 : 44;
    return value;
  }};
selection.Document = document;
document._wpscRunOwnsAppendCursor = true;
document._wpscAppendCursorRange = document.Range(4, 4);
const target = document.Range(0, 0);
window.WPSComposerLongformV2.__test.insertStyledDegradationAtRange(
  document, target, "[NOTICE]"
);
assert.ok(writes.some(function(item) {
  return item[0] === 8 && item[1] === "color" && item[2] === 11;
}));
assert.ok(writes.some(function(item) {
  return item[0] === 8 && item[1] === "background" && item[2] === 22;
}));

const missingStyleSelection = {Document: null, Range: null};
const missingStyle = {Name: "MissingStyle.docx",
  ActiveWindow: {Selection: missingStyleSelection},
  get Content() { return {End: 1, Text: ""}; },
  Range: function(start, end) { return {Start: start, End: end, Font: null, Shading: null,
    Select: function() { missingStyleSelection.Range = this; },
    InsertAfter: function(value){ this.End += String(value).length; }}; }};
missingStyleSelection.Document = missingStyle;
missingStyle._wpscRunOwnsAppendCursor = true;
missingStyle._wpscAppendCursorRange = missingStyle.Range(0, 0);
assert.throws(function() {
  window.WPSComposerLongformV2.__test.addCitationParagraph(missingStyle, {runs: [
    {type: "degradation", nodeId: "p/c:0", code: "NOTICE", fallbackText: "safe"}
  ]}, {}, {ownerNodeId: "p", issues: [], childResults: [], controllerOwned: true});
}, function(error) { return error.code === "CAPABILITY_MISMATCH"; });

const missingSelection = {Name: "MissingSelection.docx",
  get Content() { return {End: 1, Text: ""}; },
  Range: function(start, end) { return {Start: start, End: end,
    Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0},
    InsertAfter: function(value) { this.End += String(value).length; }}; }};
missingSelection._wpscRunOwnsAppendCursor = true;
missingSelection._wpscAppendCursorRange = missingSelection.Range(0, 0);
assert.throws(function() {
  window.WPSComposerLongformV2.__test.addCitationParagraph(missingSelection, {runs: [
    {type: "degradation", nodeId: "p/c:0", code: "NOTICE", fallbackText: "safe"}
  ]}, {}, {ownerNodeId: "p", issues: [], childResults: [], controllerOwned: true});
}, function(error) { return error.code === "CAPABILITY_MISMATCH"; });
''')


def test_js_formula_image_rung_only_recovers_named_image_failures() -> None:
    _run_node(r'''
function makeDocument(addPicture) {
  let text = "", rollbacks = 0;
  function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
    ParagraphFormat: {TabStops: {Add: function(){}}},
    get Text() { return text.slice(start, end); },
    InsertAfter: function(value) { text += String(value); },
    Delete: function() { rollbacks += 1; text = ""; }}; }
  return {document: {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
    Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
    InlineShapes: {AddPicture: function() { return addPicture(function(value) { text += value; }); }},
    Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }},
    Bookmarks: {Add: function() {}}},
    text: function() { return text; }, rollbacks: function() { return rollbacks; }};
}
const args = {fallbackResource: {fallbackResourceId: "formula-image-1"}, fallbackText: "x+y",
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)};
const context = {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true};
const unknown = new Error("raw engine exception");
let state = makeDocument(function(mutate) { mutate("partial"); throw unknown; });
assert.throws(() => window.WPSComposerLongformV2.__test.addFormulaNativeFallback(
  state.document, args, {"formula-image-1": "/private/formula.png"}, context, "EQUATION_INSERT_FAILED"
), error => error === unknown);
assert.equal(state.rollbacks(), 1);
assert.equal(state.text(), "");
const engine = new Error("engine lost"); engine.code = "ENGINE_LOST";
state = makeDocument(function(mutate) { mutate("partial"); throw engine; });
assert.throws(() => window.WPSComposerLongformV2.__test.addFormulaNativeFallback(
  state.document, args, {"formula-image-1": "/private/formula.png"}, context, "EQUATION_INSERT_FAILED"
), error => error.code === "ENGINE_LOST");
state = makeDocument(function(mutate) { mutate("partial"); return {}; });
window.WPSComposerLongformV2.__test.addFormulaNativeFallback(
  state.document, args, {"formula-image-1": "/private/formula.png"}, context, "EQUATION_INSERT_FAILED"
);
assert.equal(state.rollbacks(), 1);
assert.ok(state.text().includes("x+y"));
''')


def test_js_formula_inner_rollback_and_number_failures_remain_exact_fatal_codes() -> None:
    _run_node(r'''
let text = "", deletes = 0;
const selection = {Document: null, Range: null, OMaths: null};
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
  ParagraphFormat: {TabStops: {Add: function(){}}},
  OMaths: {Add: function() { return null; }},
  Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
  get Text() { return text.slice(start, end); },
  InsertAfter: function(value) {
    value = String(value); text += value; this.End += value.length;
  },
  Delete: function() { deletes += 1; if (deletes === 2) { const e = new Error("rollback"); e.code = "LOCAL_MUTATION_ROLLBACK_FAILED"; throw e; } text = ""; }}; }
const document = {Name: "Rollback.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: {Count: 0, Add: function() { return null; }, Item: function() {}},
  InlineShapes: {AddPicture: function() { text += "partial"; return {}; }},
  Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }}, Bookmarks: {Add: function() {}}};
selection.Document = document;
const operation = {op: "writer.add_equation", nodeId: "eq:one", args: {
  renderMode: "native-m4", content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  fallbackResource: {fallbackResourceId: "formula-image-1"}, fallbackText: "x+y",
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"], fallback: "explicit-image-then-source-notice"}};
assert.throws(() => window.WPSComposerLongformV2.__test.runOperation(
  document, operation, {"formula-image-1": "/private/formula.png"}, [], []
), error => error.code === "LOCAL_MUTATION_ROLLBACK_FAILED");

text = "";
let rollbackCount = 0;
function range2(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
  ParagraphFormat: {TabStops: {Add: function(){}}},
  get Text() { return text.slice(start, end); },
  InsertAfter: function(value) { value = String(value); text += value; this.End += value.length; },
  Delete: function() { rollbackCount += 1; text = ""; }}; }
const fieldError = new Error("field failed"); fieldError.code = "FIELD_REFRESH_FAILED";
const document2 = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range2,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  InlineShapes: {AddPicture: function(locator, link, save, target) {
    const start = target.End; target.InsertAfter("I");
    return {Range: {Start: start, End: target.End, ParagraphFormat: {}}};
  }},
  Fields: {Add: function() { throw fieldError; }}, Bookmarks: {Add: function() {}}};
assert.throws(() => window.WPSComposerLongformV2.__test.addFormulaNativeFallback(
  document2, operation.args, {"formula-image-1": "/private/formula.png"},
  {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true}, "EQUATION_INSERT_FAILED"
), error => error.code === "FIELD_REFRESH_FAILED");
assert.equal(rollbackCount, 1);
''')


def test_js_omath_post_buildup_rechecks_local_and_global_ranges() -> None:
    _run_node(r'''
let text = "";
function range(start, end) { return {Start: start, End: end,
  ParagraphFormat: {TabStops: {Add: function(){}}},
  get Text() { return text.slice(start, end); }, InsertAfter: function(value) { text += String(value); }}; }
const args = {content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"};
const context = {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true};
let local = null, globalItem = null;
const original = {Range: {Start: 1, End: 4}, BuildUp: function() {
  this.Range.End = 99;
}};
local = original; globalItem = original;
const localCollection = {Count: 1, Item: function() { return local; }};
const globalCollection = {Count: 0, Add: function() { this.Count = 1; return {Start: 1, End: 4, OMaths: localCollection}; },
  Item: function() { return globalItem; }};
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range, OMaths: globalCollection,
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(document, args, {}, context),
  error => error.code === "EQUATION_INSERT_FAILED");

const unknown = new Error("global getter failed");
text = "";
let failRange = false;
const unknownRange = {get Start() { if (failRange) throw unknown; return 1; },
  get End() { if (failRange) throw unknown; return 4; }};
const unknownMath = {Range: unknownRange, BuildUp: function() { failRange = true; }};
const unknownLocal = {Count: 1, Item: function() { return unknownMath; }};
const unknownGlobal = {Count: 0, Add: function() { this.Count = 1; return {Start: 1, End: 4, OMaths: unknownLocal}; },
  Item: function() { return unknownMath; }};
const unknownDocument = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range, OMaths: unknownGlobal,
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(unknownDocument, args, {}, context),
  error => error === unknown);
''')


def test_js_omath_accepts_distinct_wps_proxy_wrappers_with_stable_ranges() -> None:
    _run_node(r'''
let text = "", buildUps = 0;
function format() { return {TabStops: {Add: function(){}}}; }
function range(start, end) { return {Start: start, End: end, ParagraphFormat: format(),
  get Text() { return text.slice(start, end); }, InsertAfter: function(value) { text += String(value); }}; }
const state = {start: 1, end: 4};
function wrapper() { return {get Range() { return {Start: state.start, End: state.end}; },
  BuildUp: function() { buildUps += 1; }}; }
function globalWrapper() { return {get Range() { return {Start: 0, End: 10}; }}; }
const local = {Count: 1, Item: function() { return wrapper(); }};
const maths = {Count: 0, Add: function() { this.Count = 1; return {
  Start: state.start, End: state.end, OMaths: local
}; }, Item: function() { return globalWrapper(); }};
const bookmarks = [], bookmarkByName = Object.create(null);
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}, OMaths: maths,
  Fields: {Add: function(target, type, code) {
    if (code.indexOf("REF ") === 0) {
      const bookmark = bookmarkByName[code.split(/\s+/)[1]];
      text += text.slice(bookmark.start, bookmark.end);
    } else {
      text += code.indexOf("STYLEREF") === 0 ? "1" : "9";
    }
    return {Update: function(){}, Result: {Text: "1"}};
  }},
  Styles: {Item: function() { return {NameLocal: "Heading 1"}; }},
  Bookmarks: {Add: function(name, target) {
    const value = {name: name, start: target.Start, end: target.End};
    bookmarks.push(value); bookmarkByName[name] = value;
  }}};
const args = {content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "chapter", sequenceId: "WPSC_EQ", chapterStyleLevel: 1,
    resetLevel: 1, prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"};
window.WPSComposerLongformV2.__test.addEquationNativeM4(document, args, {},
  {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true});
assert.equal(buildUps, 1);
assert.equal(bookmarks.length, 1);
assert.equal(text.slice(bookmarks[0].start, bookmarks[0].end), "1-9");
assert.equal(text[bookmarks[0].start - 1], "(");
assert.equal(text[bookmarks[0].end], ")");
window.WPSComposerLongformV2.__test.addCrossReferenceParagraph(document, {runs: [
  {type: "text", text: "Formula ref "},
  {type: "reference", prefix: "(", bookmarkName: args.bookmarkName,
    suffix: ")", fallbackText: "1-9"}
]}, {}, {ownerNodeId: "p:ref", issues: [], childResults: [], controllerOwned: true});
assert.ok(text.includes("Formula ref (1-9)"));
''')


def test_js_omath_recovers_fresh_exact_local_and_never_builds_global_proxy() -> None:
    _run_node(r'''
let text = "", buildUps = 0, freshEnabled = "item";
const freshMath = {Range: {Start: 1, End: 4}, BuildUp: function() { buildUps += 1; }};
function range(start, end) {
  const result = {Start: start, End: end, Font: {},
    ParagraphFormat: {TabStops: {Add: function(){}}},
    get Text() { return text.slice(start, end); },
    InsertAfter: function(value) { text += String(value); }};
  if (freshEnabled === "item" && start === 1 && end === 4 && document.OMaths.Count === 1) {
    result.OMaths = {Count: 0, Item: function(index) {
      if (index === 1) return freshMath;
      throw new Error("no second local math");
    }};
  }
  return result;
}
const local = {Count: 0, Item: function() { throw new Error("missing local proxy"); }};
const maths = {Count: 0, Add: function() { this.Count = 1;
  return {Start: 1, End: 4, OMaths: local};
}, Item: function() { throw new Error("global Item must not be read"); }};
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}, OMaths: maths,
  Fields: {Add: function(target) { text += "1";
    return {Update: function(){}, Result: {Text: "1"}};
  }}, Bookmarks: {Add: function() {}}};
window.WPSComposerLongformV2.__test.addEquationNativeM4(document, {
  content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"
}, {}, {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true});
assert.equal(buildUps, 1);
assert.ok(text.includes("x+y"));

text = "";
freshEnabled = false;
let localBuildUps = 0;
const truthfulLocalMath = {Range: {Start: 1, End: 4},
  BuildUp: function() { localBuildUps += 1; }};
const broadGlobal = {Range: {Start: 0, End: 10},
  BuildUp: function() { throw new Error("global proxy must not build"); }};
document.OMaths = {Count: 0, Add: function() { this.Count = 1;
  return {Start: 1, End: 4, OMaths: {
    Count: 0, Item: function() { return truthfulLocalMath; }
  }};
}, Item: function() { return broadGlobal; }};
window.WPSComposerLongformV2.__test.addEquationNativeM4(document, {
  content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "f".repeat(24), fallbackText: "x+y"
}, {}, {ownerNodeId: "eq:local-liar", issues: [], childResults: [], controllerOwned: true});
assert.equal(localBuildUps, 1);

text = "";
freshEnabled = false;
const wrongGlobal = {Range: {Start: 0, End: 10}};
document.OMaths = {Count: 0, Add: function() { this.Count = 1;
  return {Start: 1, End: 4, OMaths: {Count: 0}};
}, Item: function() { return wrongGlobal; }};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(document, {
  content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "a".repeat(24), fallbackText: "x+y"
}, {}, {ownerNodeId: "eq:wrong", issues: [], childResults: [], controllerOwned: true}),
error => error.code === "EQUATION_INSERT_FAILED");
''')


def test_js_omath_success_commits_host_expansion_before_number_and_next_operation() -> None:
    _run_node(r'''
let text = "";
const writes = [];
let localMath = null;
const selection = {Document: null, Range: null, OMaths: null};
function range(start, end) {
  const value = {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, Style: null,
    ParagraphFormat: {TabStops: {Add: function(){}}},
    OMaths: {Add: function(target) { return maths.Add(target); }},
    Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) {
      value = String(value);
      writes.push([this.End, value]);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() { text = text.slice(0, this.Start) + text.slice(this.End); }
  };
  if (localMath && start === localMath.Range.Start && end === localMath.Range.End) {
    value.OMaths = {Count: 0, Item: function() { return localMath; }};
  }
  return value;
}
const globalMath = {Range: {Start: 0, End: 20}, BuildUp: function() {
  throw new Error("global proxy must not build");
}};
const maths = {Count: 0, Add: function(target) {
  this.Count = 1;
  localMath = {Range: {Start: target.Start, End: target.End}, BuildUp: function() {
  // Real WPS may advance Content.End while the collapsed insertion Range
  // remains at the smaller end returned by OMaths.Add.
  text = text.slice(0, 4) + "<OMATHPAD>" + text.slice(4);
  }};
  const result = {Start: target.Start, End: target.End,
    OMaths: {Count: 0, Item: function() { throw new Error("no local proxy"); }}};
  result.Select = function() { selection.Range = result; selection.OMaths = {
    Count: 0, Item: function() { return localMath; }
  }; };
  return result;
}, Item: function() { return globalMath; }};
const document = {
  Name: "SelectionOwned.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  get Paragraphs() { return {Count: 1, Item: function() {
    return {Range: {Start: 0, End: text.length + 1}};
  }}; },
  Styles: {Item: function() { return {Font: {}, ParagraphFormat: {}}; }},
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}, OMaths: maths,
  Fields: {Add: function(target) { target.InsertAfter("1");
    return {Update: function(){}, Result: {Text: "1", Start: target.Start, End: target.End}};
  }}, Bookmarks: {Add: function() {}}
};
selection.Document = document;
const api = window.WPSComposerLongformV2.__test;
api.runOperation(document, {op: "writer.add_equation", nodeId: "eq:1", args: {
  renderMode: "native-m4",
  content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  fallbackText: "x+y",
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)
}}, {}, [], []);
const numberTab = writes.find(function(item) {
  return item[1] === "\t" && item[0] === "\tx+y".length;
});
assert.ok(numberTab, JSON.stringify(writes));
api.runOperation(document, {op: "writer.add_paragraph", nodeId: "p:after",
  args: {text: "After"}}, {}, [], []);
const afterWrite = writes.find(function(item) { return item[1] === "After\r"; });
assert.ok(afterWrite && afterWrite[0] > numberTab[0], JSON.stringify(writes));
assert.ok(text.endsWith("\t(1)\rAfter\r"), text);
''')


def test_js_omath_success_uses_each_growing_host_endpoint_independently() -> None:
    _run_node(r'''
function runCase(staleSource, recoverFirst) {
  let text = "", stale = false, failNext = recoverFirst;
  const writes = [], issues = [], ranges = [];
  let freshMath = null;
  const selection = {Document: null, Range: null, OMaths: null};
  function reportedEnd(source) {
    return stale && source === staleSource ? 100 : text.length + 1;
  }
  function range(start, end) {
    const value = {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, Style: null,
      ParagraphFormat: {TabStops: {Add: function(){}}},
      OMaths: {Add: function(target) { return maths.Add(target); }},
      Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
      get Text() { return text.slice(this.Start, this.End); },
      InsertAfter: function(value) {
        value = String(value);
        writes.push([this.End, value]);
        text = text.slice(0, this.End) + value + text.slice(this.End);
        this.End += value.length;
      },
      Delete: function() { text = text.slice(0, this.Start) + text.slice(this.End); }
    };
    if (freshMath && start === freshMath.Range.Start && end === freshMath.Range.End) {
      value.OMaths = {Count: 0, Item: function() { return freshMath; }};
    }
    ranges.push(value);
    return value;
  }
  let globalMath = null;
  const maths = {Count: 0, Add: function(target) {
    if (failNext) { failNext = false; return null; }
    this.Count += 1;
      freshMath = {Range: {Start: target.Start, End: target.End}, BuildUp: function() {
        text = text.slice(0, target.End) + "<OMATHPAD>" + text.slice(target.End);
        ranges.forEach(function(item) {
          if (item.Start === target.End && item.End === target.End) {
            item.Start += 10; item.End += 10;
          }
        });
      }};
    globalMath = {Range: {Start: 0, End: target.End + 20}, BuildUp: function() {
      throw new Error("global proxy must not build");
    }};
    const result = {Start: target.Start, End: target.End,
      OMaths: {Count: 0, Item: function() { throw new Error("no local proxy"); }}};
    result.Select = function() { selection.Range = result; selection.OMaths = {
      Count: 0, Item: function() { return freshMath; }
    }; };
    return result;
  }, Item: function() { return globalMath; }};
  const document = {
    Name: "SelectionOwned.docx", ActiveWindow: {Selection: selection},
    get Content() { return {End: reportedEnd("content"), get Text() { return text; }}; }, Range: range,
    get Paragraphs() { return {
      get Count() { return (text.match(/\r/g) || []).length + 1; },
      Item: function() {
        const start = text.lastIndexOf("\r") + 1;
        return {Range: {Start: start, End: reportedEnd("paragraph")}};
      }
    }; },
    Styles: {Item: function() { return {Font: {}, ParagraphFormat: {}}; }},
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}, OMaths: maths,
    Fields: {Add: function(target) { target.InsertAfter("1");
      return {Update: function(){}, Result: {Text: "1", Start: target.Start, End: target.End}};
    }}, Bookmarks: {Add: function() {}}
  };
  selection.Document = document;
  const api = window.WPSComposerLongformV2.__test;
  api.runOperation(document, {op: "writer.add_paragraph", nodeId: "p:prefix",
    args: {text: "Prefix"}}, {}, issues, []);
  stale = true;
  const operation = function(nodeId) { return {op: "writer.add_equation", nodeId: nodeId,
    args: {renderMode: "native-m4",
      content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
      fallbackText: "x+y",
      numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
      bookmarkName: "wpsc_eq_" + (nodeId === "eq:first" ? "e" : "f").repeat(24)},
    failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
      fallback: "explicit-image-then-source-notice"}}; };
  if (recoverFirst) api.runOperation(document, operation("eq:first"), {}, issues, []);
  const beforeSuccess = text.length;
  api.runOperation(document, operation("eq:second"), {}, issues, []);
  const numberTab = writes.find(function(item) {
    return item[1] === "\t" && item[0] === beforeSuccess + "\tx+y".length;
  });
  assert.ok(numberTab, staleSource + ":" + JSON.stringify(writes));
  api.runOperation(document, {op: "writer.add_paragraph", nodeId: "p:after",
    args: {text: "After"}}, {}, issues, []);
  assert.ok(text.endsWith("\t(1)\rAfter\r"), staleSource + ":" + text);
  if (recoverFirst) assert.deepEqual(issues.map(function(item) { return item.code; }),
    ["EQUATION_INSERT_FAILED"]);
  else assert.deepEqual(issues, []);
}
runCase("paragraph", false);
''')


def test_js_omath_accepts_wps_content_and_range_coordinate_remap() -> None:
    _run_node(r'''
function runCase(corruptHeldPrefix) {
  let text = "Prefix\r", built = false, formulaStart = -1;
  const ranges = [], writes = [];
  let freshMath = null;
  const selection = {Document: null, Range: null, OMaths: null};
  function range(start, end) {
    const snapshot = text.slice(start, end);
    const heldBeforeBuild = !built;
    const value = {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, Style: null,
      ParagraphFormat: {TabStops: {Add: function(){}}},
      OMaths: {Add: function(target) { return maths.Add(target); }},
      Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
      get Text() {
        if (heldBeforeBuild && this.Start === 0 && built) {
          if (this.End === formulaStart) return snapshot + "<ABSORBED>";
          if (corruptHeldPrefix && this.End === formulaStart - 1) {
            return snapshot.slice(0, -1);
          }
          return snapshot;
        }
        // Rebuilding the old numeric prefix after BuildUp is deliberately
        // wrong, matching WPS's remapped character coordinates.
        if (!heldBeforeBuild && built && this.Start === 0) {
          return text.slice(0, Math.max(0, this.End - 2));
        }
        return text.slice(this.Start, this.End);
      },
      InsertAfter: function(raw) {
        const inserted = String(raw);
        writes.push([this.End, inserted]);
        text = text.slice(0, this.End) + inserted + text.slice(this.End);
        this.End += inserted.length;
      },
      Delete: function() { text = text.slice(0, this.Start) + text.slice(this.End); }
    };
    if (freshMath && start === freshMath.Range.Start && end === freshMath.Range.End) {
      value.OMaths = {Count: 0, Item: function() { return freshMath; }};
    }
    ranges.push(value);
    return value;
  }
  let globalMath = null;
  const maths = {Count: 0, Add: function(target) {
    this.Count = 1;
    formulaStart = target.Start;
    freshMath = {Range: {Start: target.Start, End: target.End}, BuildUp: function() {
      const oldEnd = target.End;
      text = text.slice(0, oldEnd) + "<OMATHPAD>" + text.slice(oldEnd);
      built = true;
      ranges.forEach(function(item) {
        if (item.Start === oldEnd && item.End === oldEnd) {
          item.Start += 10; item.End += 10;
        }
      });
    }};
    globalMath = {Range: {Start: 0, End: target.End + 20}, BuildUp: function() {
      throw new Error("global proxy must not build");
    }};
    const result = {Start: target.Start, End: target.End,
      OMaths: {Count: 0, Item: function() { throw new Error("no local proxy"); }}};
    result.Select = function() { selection.Range = result; selection.OMaths = {
      Count: 0, Item: function() { return freshMath; }
    }; };
    return result;
  }, Item: function() { return globalMath; }};
  const document = {get Content() { return {End: text.length + 1,
    get Text() { return corruptHeldPrefix && built ? "X" + text.slice(1) : text; }}; },
    Range: range,
    get Paragraphs() { return {Count: 2, Item: function() {
      return {Range: {Start: 7, End: text.length + 1}};
    }}; },
    Styles: {Item: function() { return {Font: {}, ParagraphFormat: {}}; }},
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}, OMaths: maths,
    Fields: {Add: function(target) { target.InsertAfter("1");
      return {Update: function(){}, Result: {Text: "1", Start: target.Start, End: target.End}};
    }}, Bookmarks: {Add: function() {}}
  };
  document.Name = "SelectionOwned.docx";
  document.ActiveWindow = {Selection: selection};
  selection.Document = document;
  const operation = {op: "writer.add_equation", nodeId: "eq:remap", args: {
    renderMode: "native-m4",
    content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
    fallbackText: "x+y",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24)}, failurePolicy: {mode: "fail"}};
  window.WPSComposerLongformV2.__test.runOperation(document, operation, {}, [], []);
  assert.ok(writes.some(function(item) { return item[0] === 11 && item[1] === "\t"; }),
    JSON.stringify(writes));
}
runCase(false);
runCase(true);
''')


def test_js_run_owned_insert_and_number_field_reject_silent_noops() -> None:
    _run_node(r'''
let text = "";
function silentRange(start, end) { return {
  Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
  InsertAfter: function() {}, Delete: function() {}
}; }
const silentDocument = {
  get Content() { return {End: 1, get Text() { return text; }}; },
  Range: silentRange,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: 1}}; }}
};
silentDocument._wpscRunOwnsAppendCursor = true;
silentDocument._wpscAppendCursorRange = silentRange(0, 0);
assert.throws(function() {
  window.WPSComposerLongformV2.__test.addNativeNumberShell(silentDocument, {
    mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"
  }, "wpsc_eq_" + "d".repeat(24), "eq:insert-no-op", true);
}, function(error) { return error.code === "CAPABILITY_MISMATCH"; });
assert.equal(text, "");

function range(start, end) { return {
  Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
  InsertAfter: function(raw) {
    const value = String(raw);
    text = text.slice(0, this.End) + value + text.slice(this.End);
    this.End += value.length;
  }
}; }
const fields = {Count: 0, Add: function(target) {
  return {Update: function(){}, Result: {Start: target.End, End: target.End + 1, Text: "1"}};
}};
const fieldDocument = {
  get Content() { return {End: text.length + 1, get Text() { return text; }}; },
  Range: range, Fields: fields, Bookmarks: {Add: function() {}}
};
fieldDocument._wpscRunOwnsAppendCursor = true;
fieldDocument._wpscAppendCursorRange = range(0, 0);
assert.throws(function() {
  window.WPSComposerLongformV2.__test.addNativeNumberShell(fieldDocument, {
    mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"
  }, "wpsc_eq_" + "e".repeat(24), "eq:no-op", true);
}, function(error) { return error.code === "FIELD_REFRESH_FAILED"; });
assert.equal(fields.Count, 0);

text = "";
const writes = [];
let provenContentPosition = 0;
function provenRange(start, end) { return {
  Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
  InsertAfter: function(raw) {
    const value = String(raw);
    writes.push([this.End, value]);
    text = text.slice(0, this.End) + value + text.slice(this.End);
    this.End += value.length;
    provenContentPosition = Math.max(provenContentPosition, this.End);
  }
}; }
const provenFields = {Count: 0, Add: function(target) {
  const start = target.End;
  text = text.slice(0, start) + "1" + text.slice(start);
  provenContentPosition = 52;
  this.Count += 1;
  // Real WPS may omit Field.Range and place Result after internal field code.
  // The result remains bounded by the same Content growth caused by Add.
  return {Update: function(){}, Result: {Start: 50, End: 51, Text: "1"}};
}};
const provenDocument = {
  get Content() { return {End: provenContentPosition + 1, get Text() { return text; }}; },
  Range: provenRange, Fields: provenFields, Bookmarks: {Add: function() {}}
};
provenDocument._wpscRunOwnsAppendCursor = true;
provenDocument._wpscAppendCursorRange = provenRange(0, 0);
window.WPSComposerLongformV2.__test.addNativeNumberShell(provenDocument, {
  mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"
}, "wpsc_eq_" + "f".repeat(24), "eq:bound", true);
assert.deepEqual(writes, [[0, "("], [52, ")"]]);
window.WPSComposerLongformV2.__test.runOperation(provenDocument, {
  op: "writer.add_paragraph", nodeId: "p:after-field", args: {text: "After"}
}, {}, [], []);
assert.equal(text, "(1)After\r");
''')


def test_js_formula_uses_last_paragraph_end_when_content_end_is_stale() -> None:
    _run_node(r'''
const citation = "Inline citation remains complete in this paragraph.\r";
let text = citation;
const added = [];
function paragraphFormat() { return {TabStops: {Add: function(){}}}; }
function range(start, end) {
  return {
    Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: paragraphFormat(),
    get Text() { return text.slice(start, end); },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() { text = text.slice(0, start) + text.slice(end); }
  };
}
const maths = {
  Count: 0,
  items: [],
  Add: function(target) {
    added.push([target.Start, target.End, target.Text]);
    const bounds = {start: target.Start, end: target.End};
    const native = {get Range() { return {Start: bounds.start, End: bounds.end}; },
      BuildUp: function(){}};
    this.items.push(native);
    this.Count += 1;
    return {Start: bounds.start, End: bounds.end,
      OMaths: {Count: 1, Item: function() { return native; }}};
  },
  Item: function(index) { return this.items[index - 1]; }
};
const document = {
  // Real WPS can lag by the length of the most recently inserted run.
  get Content() { return {End: Math.max(1, text.length + 1 - 9), get Text() { return text; }}; },
  Range: range,
  get Paragraphs() { return {
    get Count() { return (text.match(/\r/g) || []).length + 1; },
    Item: function(index) {
      assert.equal(index, this.Count);
      const start = text.lastIndexOf("\r") + 1;
      return {Range: {Start: start, End: text.length + 1, Text: text.slice(start)}};
    }
  }; },
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: maths,
  Fields: {Add: function(target) {
    target.InsertAfter("1");
    return {Update: function(){}, Result: {Text: "1"}};
  }},
  Bookmarks: {Add: function() {}}
};
window.WPSComposerLongformV2.__test.addEquationNativeM4(document, {
  content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"
}, {}, {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true});
assert.equal(added.length, 1);
assert.equal(added[0][0], citation.length + 1);
assert.ok(text.startsWith(citation));
assert.ok(text.slice(citation.length).startsWith("\tx+y\t(1)\r"));
''')


def test_js_runtime_cursor_survives_content_and_paragraph_collection_lag() -> None:
    _run_node(r'''
let text = "";
let forbiddenGlobalAdds = 0, forbiddenGlobalItems = 0;
let mathBounds = null;
const numberFieldTargets = [];
const fieldCodes = [], bookmarkNames = [];
const events = [];
const selection = {Document: null, Range: null, OMaths: null, InMath: false};
const liveRanges = [];
function paragraphFormat() { return {TabStops: {Add: function(){}}}; }
function range(start, end) {
  const value = {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: paragraphFormat(),
    Style: null, ListFormat: {},
    OMaths: {Add: function(target) { return maths.Add(target); }},
    Select: function() { selection.Range = this; selection.OMaths = this.OMaths;
      selection.InMath = false;
      if (this.Text === "x+y") mathBounds = {Start: this.Start, End: this.End}; },
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) {
      value = String(value);
      const insertionPoint = this.End;
      liveRanges.forEach(function(item) {
        if (item !== this && item.End === insertionPoint) item.End += value.length;
      }, this);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() { text = text.slice(0, start) + text.slice(end); }
  };
  liveRanges.push(value);
  return value;
}
const maths = {Count: 0, items: [], Add: function(target) {
  assert.equal(target.Text, "x+y");
  assert.equal(text.slice(target.End, target.End + 1), "\t");
  assert.deepEqual(mathBounds, {Start: target.Start, End: target.End});
  const native = {Range: {Start: target.Start, End: target.End}, BuildUp: function(){
    events.push("build");
  }};
  this.items.push(native); this.Count += 1;
  const result = {Start: target.Start, End: target.End,
    OMaths: {Count: 1, Item: function() { return native; }}};
  result.Select = function() { selection.Range = result; selection.OMaths = result.OMaths;
    selection.InMath = true; };
  return result;
}, Item: function(index) { return this.items[index - 1]; }};
const globalMaths = {get Count() { return maths.Count; }, Add: function() {
  forbiddenGlobalAdds += 1; throw new Error("global add must not run");
}, Item: function() {
  forbiddenGlobalItems += 1; throw new Error("global item must not run");
}};
const document = {
  Name: "SelectionOwned.docx", ActiveWindow: {Selection: selection},
  // Both collection views remain stale for the whole session. Only the actual
  // insertion Range returned by WPS advances its End.
  Content: {End: 1, get Text() { return text; }},
  Range: range,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: 1}}; }},
  Styles: {Item: function() { return {Font: {}, ParagraphFormat: {}}; }},
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: globalMaths,
  Fields: {Add: function(target, type, code) {
    assert.equal(selection.InMath, false);
    assert.ok(mathBounds && target.Start >= mathBounds.End + 1);
    numberFieldTargets.push(target.Start);
    fieldCodes.push(code);
    events.push("field:" + code);
    target.InsertAfter("1");
    return {Update: function(){}, Result: {Text: "1", Start: target.Start, End: target.End}};
  }},
  Bookmarks: {Add: function(name) { bookmarkNames.push(name); },
    Exists: function(name) { return bookmarkNames.indexOf(name) !== -1; },
    Item: function() { return {Range: {Fields: {Count: 1}}}; }}
};
selection.Document = document;
const api = window.WPSComposerLongformV2.__test;
const issues = [], children = [];
api.runOperation(document, {op: "writer.add_heading", nodeId: "h:1",
  args: {text: "Heading", level: 1, numbering: false}}, {}, issues, children);
api.runOperation(document, {op: "writer.add_paragraph", nodeId: "p:1",
  args: {text: "Plain"}}, {}, issues, children);
api.runOperation(document, {op: "writer.add_cross_reference", nodeId: "p:2",
  args: {runs: [{type: "text", text: "See "},
    {type: "citation", nodeId: "p:2/c:0", targetId: "a", targetNodeId: "ref:a",
      number: 1, fallbackText: "[1]"}, {type: "text", text: "."}]}}, {}, issues, children);
api.runOperation(document, {op: "writer.add_equation", nodeId: "eq:1",
  args: {renderMode: "native-m4",
    content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
    fallbackText: "x+y",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
    fallback: "explicit-image-then-source-notice"}}, {}, issues, children);
api.runOperation(document, {op: "writer.add_cross_reference", nodeId: "p:ref",
  args: {runs: [{type: "text", text: "Ref "}, {type: "reference",
    targetNodeId: "eq:1", targetKind: "equation",
    bookmarkName: "wpsc_eq_" + "e".repeat(24), prefix: "(", suffix: ")",
    fallbackText: "1"}]}}, {}, issues, children);
assert.equal(text, "Heading\rPlain\rSee [1].\r\tx+y\t(1)\rRef (1)\r");
assert.deepEqual(issues, []);
assert.deepEqual(children, [{nodeId: "p:2/c:0", status: "applied"}]);
assert.equal(forbiddenGlobalAdds, 0);
assert.equal(forbiddenGlobalItems, 0);
assert.equal(numberFieldTargets.length, 2);
assert.ok(bookmarkNames.includes("wpsc_eq_" + "e".repeat(24)));
assert.ok(fieldCodes.includes("REF wpsc_eq_" + "e".repeat(24) + " \\h"));
assert.ok(events.indexOf("field:SEQ WPSC_EQ \\* ARABIC") < events.indexOf("build"));
assert.ok(events.indexOf("field:REF wpsc_eq_" + "e".repeat(24) + " \\h") >
  events.indexOf("build"));
assert.equal(selection.InMath, false);
''')


def test_js_runtime_cursor_advances_after_table_image_and_break_objects() -> None:
    _run_node(r'''
function makeDocument() {
  let text = "";
  function range(start, end) {
    return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
      get Text() { return text.slice(start, end); },
      InsertAfter: function(value) {
        value = String(value);
        text = text.slice(0, this.End) + value + text.slice(this.End);
        this.End += value.length;
      },
      InsertBreak: function() { this.InsertAfter("<BREAK>"); },
      Delete: function() { text = text.slice(0, start) + text.slice(end); }
    };
  }
  const document = {Content: {End: 1}, Range: range,
    Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: 1}}; }},
    Styles: {Item: function() { return {Font: {}, ParagraphFormat: {}}; }},
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}};
  return {document: document, text: function() { return text; }};
}
const api = window.WPSComposerLongformV2.__test;

let state = makeDocument();
api.runOperation(state.document, {op: "writer.reset", args: {}}, {}, [], []);
function border() { return {}; }
function rows() { return {Borders: border}; }
rows.AllowBreakAcrossPages = 0;
state.document.Tables = {Add: function(target) {
  const start = target.End;
  target.InsertAfter("<TABLE>");
  const cell = {Range: {Text: "", Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {}},
    Merge: function(){}};
  return {Range: {Start: start, End: target.End}, Cell: function() { return cell; },
    Rows: rows, Borders: border};
}};
api.runOperation(state.document, {op: "writer.add_semantic_table", args: {
  headers: ["A"], rows: [["B"]], keepCaptionWithFirstRow: true,
  orientation: "portrait", plannedDegradation: [],
  alignments: ["left"], cellIndentPt: 0, allowRowSplit: false,
  repeatHeader: false, borderSpec: {top: 0, left: 0, bottom: 0, right: 0,
    insideHorizontal: 0, insideVertical: 0, headerBottom: 0}, merges: [],
  cellCitations: [], cellDegradations: []}}, {}, [], []);
api.runOperation(state.document, {op: "writer.add_paragraph", args: {text: "After"}}, {}, [], []);
assert.equal(state.text(), "<TABLE>\rAfter\r");

state = makeDocument();
api.runOperation(state.document, {op: "writer.reset", args: {}}, {}, [], []);
state.document.InlineShapes = {Count: 0, AddPicture: function(locator, link, save, target) {
  const start = target.End;
  target.InsertAfter("<IMAGE>");
  this.Count += 1;
  return {Range: {Start: start, End: target.End, ParagraphFormat: {}}};
}, Item: function(){}};
api.runOperation(state.document, {op: "writer.add_captioned_figure", args: {
  keepWithCaption: true,
  orientation: "portrait", layout: "stack", children: [{nodeId: "fig:1/image:1",
    resourceId: "image", displayWidthPt: 100, displayHeightPt: 50}]},
  nodeId: "fig:1"}, {image: "/private/image.png"}, [], []);
api.runOperation(state.document, {op: "writer.add_paragraph", args: {text: "After"}}, {}, [], []);
assert.equal(state.text(), "<IMAGE>\rAfter\r");

state = makeDocument();
api.runOperation(state.document, {op: "writer.reset", args: {}}, {}, [], []);
api.runOperation(state.document, {op: "writer.add_page_break", args: {}}, {}, [], []);
api.runOperation(state.document, {op: "writer.add_paragraph", args: {text: "After"}}, {}, [], []);
assert.equal(state.text(), "<BREAK>After\r");

state = makeDocument();
api.runOperation(state.document, {op: "writer.reset", args: {}}, {}, [], []);
const footer = {PageNumbers: {}, Range: {Text: "", Collapse: function(){},
  Fields: {Add: function(){}}}};
const section = {PageSetup: {}, Range: {DocumentVariables: {Add: function(){}}},
  Headers: {Item: function() { return {}; }},
  Footers: {Item: function() { return footer; }}};
state.document.Sections = {Count: 1, Item: function() { return section; }};
state.document._wpscFirstSectionConfigured = true;
api.runOperation(state.document, {op: "writer.configure_section", args: {
  role: "body", pageNumberFormat: "none"
}}, {}, [], []);
api.runOperation(state.document, {op: "writer.add_paragraph", args: {text: "After"}}, {}, [], []);
assert.equal(state.text(), "<BREAK>After\r");
''')


def test_js_run_owned_native_partial_writes_rollback_from_live_target_range() -> None:
    _run_node(r'''
function makeDocument() {
  let text = "PREFIX\r";
  function range(start, end) {
    return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
      get Text() { return text.slice(start, end); },
      InsertAfter: function(value) {
        value = String(value);
        text = text.slice(0, this.End) + value + text.slice(this.End);
        this.End += value.length;
      },
      Delete: function() { text = text.slice(0, this.Start) + text.slice(this.End); }
    };
  }
  const document = {
    get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
    Paragraphs: {Count: 2, Item: function() {
      return {Range: {Start: 7, End: text.length + 1}};
    }},
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64}
  };
  return {document: document, text: function() { return text; }};
}
const api = window.WPSComposerLongformV2.__test;

let state = makeDocument();
state.document.InlineShapes = {Count: 0, AddPicture: function(a, b, c, target) {
  target.InsertAfter("<PARTIAL-IMAGE>");
  const error = new Error("private image failure");
  error.code = "IMAGE_INSERT_FAILED";
  throw error;
}};
assert.throws(() => api.runOperation(state.document, {
  op: "writer.add_captioned_figure", nodeId: "fig:partial",
  failurePolicy: {mode: "fail"},
  args: {keepWithCaption: true, orientation: "portrait", layout: "stack",
    children: [{nodeId: "fig:partial/image:1", resourceId: "image",
      displayWidthPt: 100, displayHeightPt: 50}]}
}, {image: "/private/image.png"}, [], []),
error => error.code === "EXECUTION_ABORTED");
assert.equal(state.text(), "PREFIX\r");

state = makeDocument();
state.document.Tables = {Add: function(target) {
  target.InsertAfter("<PARTIAL-TABLE>");
  const error = new Error("private table failure");
  error.code = "TABLE_INSERT_FAILED";
  throw error;
}};
assert.throws(() => api.runOperation(state.document, {
  op: "writer.add_semantic_table", nodeId: "table:partial",
  failurePolicy: {mode: "fail"},
  args: {headers: ["A"], rows: [["B"]], keepCaptionWithFirstRow: true,
    orientation: "portrait", plannedDegradation: [], alignments: ["left"],
    cellIndentPt: 0, allowRowSplit: false, repeatHeader: false,
    borderSpec: {top: 0, left: 0, bottom: 0, right: 0,
      insideHorizontal: 0, insideVertical: 0, headerBottom: 0},
    merges: [], cellCitations: [], cellDegradations: []}
}, {}, [], []), error => error.code === "EXECUTION_ABORTED");
assert.equal(state.text(), "PREFIX\r");

state = makeDocument();
state.document.InlineShapes = {Count: 0, AddPicture: function() {
  this.Count += 1;
  return {};
}};
assert.throws(() => api.runOperation(state.document, {
  op: "writer.add_captioned_figure", nodeId: "fig:no-range",
  failurePolicy: {mode: "degrade", recoverableCodes: ["IMAGE_INSERT_FAILED"],
    fallback: "figure-child-stack-then-notice"},
  args: {keepWithCaption: true, orientation: "portrait", layout: "stack",
    children: [{nodeId: "fig:no-range/image:1", resourceId: "image",
      displayWidthPt: 100, displayHeightPt: 50}]}
}, {image: "/private/image.png"}, [], []),
error => error.code === "CAPABILITY_MISMATCH");
assert.equal(state.text(), "PREFIX\r");

state = makeDocument();
state.document.Tables = {Add: function() { return {}; }};
assert.throws(() => api.runOperation(state.document, {
  op: "writer.add_semantic_table", nodeId: "table:no-range",
  failurePolicy: {mode: "degrade", recoverableCodes: ["TABLE_INSERT_FAILED"],
    fallback: "grid-then-text"},
  args: {headers: ["A"], rows: [["B"]], keepCaptionWithFirstRow: true,
    orientation: "portrait", plannedDegradation: [], alignments: ["left"],
    cellIndentPt: 0, allowRowSplit: false, repeatHeader: false,
    borderSpec: {top: 0, left: 0, bottom: 0, right: 0,
      insideHorizontal: 0, insideVertical: 0, headerBottom: 0},
    merges: [], cellCitations: [], cellDegradations: []}
}, {}, [], []), error => error.code === "CAPABILITY_MISMATCH");
assert.equal(state.text(), "PREFIX\r");
''')


def test_js_formula_recovery_never_deletes_before_authoritative_checkpoint() -> None:
    _run_node(r'''
const citation = "Inline citation remains complete in this paragraph.\r";
let text = citation;
const selection = {Document: null, Range: null, OMaths: null};
function paragraphFormat() { return {TabStops: {Add: function(){}}}; }
function range(start, end) {
  return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: paragraphFormat(),
    OMaths: {Add: function() { return null; }},
    Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
    get Text() { return text.slice(start, end); },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() { text = text.slice(0, start) + text.slice(end); }
  };
}
const document = {
  Name: "Recovery.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: Math.max(1, text.length + 1 - 9), get Text() { return text; }}; },
  Range: range,
  get Paragraphs() { return {
    get Count() { return (text.match(/\r/g) || []).length + 1; },
    Item: function(index) {
      assert.equal(index, this.Count);
      const start = text.lastIndexOf("\r") + 1;
      return {Range: {Start: start, End: text.length + 1}};
    }
  }; },
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: {Count: 0, Add: function() { return null; }, Item: function() { return null; }},
  Fields: {Add: function(target) {
    target.InsertAfter("1");
    return {Update: function(){}, Result: {Text: "1"}};
  }},
  Bookmarks: {Add: function() {}}
};
selection.Document = document;
const issues = [];
window.WPSComposerLongformV2.__test.runOperation(document, {
  op: "writer.add_equation", nodeId: "eq:one",
  args: {renderMode: "native-m4",
    content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
    fallbackText: "x+y",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
    fallback: "explicit-image-then-source-notice"}
}, {}, issues, []);
assert.equal(issues.length, 1);
assert.equal(issues[0].code, "EQUATION_INSERT_FAILED");
assert.ok(text.startsWith(citation));
assert.equal(text.slice(citation.length), "\t[EQUATION_INSERT_FAILED: x+y]\t(1)\r");
''')


def test_js_formula_rollback_accepts_empty_checkpoint_when_collapsed_range_view_drifts() -> None:
    _run_node(r'''
let text = "P\r";
const selection = {Document: null, Range: null, OMaths: null};
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
  ParagraphFormat: {TabStops: {Add: function(){}}},
  OMaths: {Add: function() { return null; }},
  Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
  get Text() {
    if (this.Start === this.End && text.length > this.Start) {
      return text.slice(this.Start, this.Start + 1);
    }
    return text.slice(this.Start, this.End);
  },
  InsertAfter: function(value) {
    value = String(value);
    text = text.slice(0, this.End) + value + text.slice(this.End);
    this.End += value.length;
  },
  Delete: function() { text = text.slice(0, this.Start) + text.slice(this.End); }
}; }
const document = {
  Name: "Recovery.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  get Paragraphs() { return {Count: (text.match(/\r/g) || []).length + 1,
    Item: function() {
      const start = text.lastIndexOf("\r") + 1;
      return {Range: {Start: start, End: text.length + 1}};
    }}; },
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: {Count: 0, Add: function() { return null; }, Item: function() {}},
  Fields: {Add: function(target) { target.InsertAfter("1");
    return {Update: function(){}, Result: {Text: "1", Start: target.Start, End: target.End}};
  }}, Bookmarks: {Add: function() {}}
};
selection.Document = document;
const issues = [];
window.WPSComposerLongformV2.__test.runOperation(document, {
  op: "writer.add_equation", nodeId: "eq:collapsed", args: {
    renderMode: "native-m4",
    content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
    fallbackText: "x+y",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
    fallback: "explicit-image-then-source-notice"}
}, {}, issues, []);
assert.equal(text, "P\r\t[EQUATION_INSERT_FAILED: x+y]\t(1)\r");
assert.deepEqual(issues.map(function(issue) { return issue.code; }), ["EQUATION_INSERT_FAILED"]);
''')


def test_js_formula_image_recovery_ignores_stale_ahead_host_end_after_rollback() -> None:
    _run_node(r'''
let text = "", reportedEnd = 1;
const selection = {Document: null, Range: null, OMaths: null};
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
  ParagraphFormat: {TabStops: {Add: function(){}}},
  OMaths: {Add: function() { return null; }},
  Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
  get Text() { return text.slice(start, end); },
  InsertAfter: function(value) {
    value = String(value);
    text = text.slice(0, this.End) + value + text.slice(this.End);
    this.End += value.length;
  },
  Delete: function() { text = text.slice(0, start) + text.slice(end); }
}; }
const document = {Name: "Recovery.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: reportedEnd, get Text() { return text; }}; }, Range: range,
  get Paragraphs() { return {Count: 1, Item: function() {
    return {Range: {Start: 0, End: reportedEnd}};
  }}; },
  Styles: {Item: function() { return {Font: {}, ParagraphFormat: {}}; }},
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: {Count: 0, Add: function() { return null; }, Item: function() { return null; }},
  InlineShapes: {AddPicture: function() {
    const error = new Error("image failed"); error.code = "IMAGE_INSERT_FAILED"; throw error;
  }},
  Fields: {Add: function(target) { target.InsertAfter("1");
    return {Update: function(){}, Result: {Text: "1", Start: target.Start, End: target.End}};
  }}, Bookmarks: {Add: function(){}}};
selection.Document = document;
const api = window.WPSComposerLongformV2.__test;
api.runOperation(document, {op: "writer.add_paragraph", args: {text: "Prefix"}}, {}, [], []);
reportedEnd = 100;
const issues = [];
api.runOperation(document, {op: "writer.add_equation", nodeId: "eq:1", args: {
  renderMode: "native-m4", content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  fallbackResource: {fallbackResourceId: "image"}, fallbackText: "x+y",
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)}, failurePolicy: {mode: "degrade",
    recoverableCodes: ["EQUATION_INSERT_FAILED"], fallback: "explicit-image-then-source-notice"}
}, {image: "/private/image.png"}, issues, []);
assert.equal(text, "Prefix\r\t[EQUATION_INSERT_FAILED: x+y]\t(1)\r");
assert.equal(issues.length, 1);
''')


def test_js_formula_image_partial_write_is_fatal_and_postprocess_rollback_covers_tail() -> None:
    _run_node(r'''
function makeDocument() {
  let text = "Prefix\r", reportedEnd = null;
  const selection = {Document: null, Range: null, OMaths: null};
  function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
    ParagraphFormat: {TabStops: {Add: function(){}}},
    OMaths: {Add: function() { return null; }},
    Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() { text = text.slice(0, this.Start) + text.slice(this.End); }
  }; }
  const document = {
    Name: "Recovery.docx", ActiveWindow: {Selection: selection},
    get Content() { return {End: reportedEnd === null ? text.length + 1 : reportedEnd, get Text() { return text; }}; }, Range: range,
    get Paragraphs() { return {Count: 2, Item: function() {
      return {Range: {Start: 7, End: text.length + 1}};
    }}; },
    PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
    OMaths: {Count: 0, Add: function() { return null; }, Item: function() {}},
    Bookmarks: {Add: function() {}}
  };
  selection.Document = document;
  return {document: document, text: function() { return text; },
    appendUntracked: function(value) { text += value; },
    freezeHostEnd: function() { reportedEnd = text.length + 1; }};
}
const api = window.WPSComposerLongformV2.__test;
function operation() { return {op: "writer.add_equation", nodeId: "eq:image", args: {
  renderMode: "native-m4",
  content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  fallbackResource: {fallbackResourceId: "image"}, fallbackText: "x+y",
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
    fallback: "explicit-image-then-source-notice"}}; }

let state = makeDocument();
state.freezeHostEnd();
state.document.InlineShapes = {Count: 0, AddPicture: function() {
  state.appendUntracked("<UNRANGED-IMAGE>");
  this.Count += 1;
  const error = new Error("private partial image");
  error.code = "IMAGE_INSERT_FAILED";
  throw error;
}};
state.document.Fields = {Add: function(target) { const start = target.End;
  target.InsertAfter("1"); return {Result: {Text: "1", Start: start, End: target.End}};
}};
assert.throws(() => api.runOperation(
  state.document, operation(), {image: "/private/image.png"}, [], []
), error => error.code === "LOCAL_MUTATION_ROLLBACK_FAILED");
assert.ok(state.text().includes("<UNRANGED-IMAGE>"));

state = makeDocument();
const unknown = new Error("private unknown image failure");
state.document.InlineShapes = {Count: 0, AddPicture: function() {
  state.appendUntracked("<HOST-RANGED-IMAGE>");
  throw unknown;
}};
state.document.Fields = {Add: function(target) { const start = target.End;
  target.InsertAfter("1"); return {Result: {Text: "1", Start: start, End: target.End}};
}};
assert.throws(() => api.runOperation(
  state.document, operation(), {image: "/private/image.png"}, [], []
), error => error === unknown);
assert.equal(state.text(), "Prefix\r");

state = makeDocument();
state.document.InlineShapes = {Count: 0, AddPicture: function(a, b, c, target) {
  const start = target.End;
  target.InsertAfter("<IMAGE>");
  this.Count += 1;
  return {Range: {Start: start, End: target.End, ParagraphFormat: {}}};
}};
const deleteRange = state.document.Range;
state.document.Range = function(start, end) {
  const value = deleteRange(start, end);
  const remove = value.Delete;
  value.Delete = function() {
    remove.call(value);
    state.document.InlineShapes.Count = 0;
  };
  return value;
};
let fieldCalls = 0;
state.document.Fields = {Add: function(target) {
  fieldCalls += 1;
  if (fieldCalls === 1) { const start = target.End; target.InsertAfter("1");
    return {Result: {Text: "1", Start: start, End: target.End}}; }
  const error = new Error("private field failure");
  error.code = "FIELD_REFRESH_FAILED"; throw error;
}};
assert.throws(() => api.runOperation(
  state.document, operation(), {image: "/private/image.png"}, [], []
), error => error.code === "FIELD_REFRESH_FAILED");
assert.equal(state.text(), "Prefix\r");
''')


def test_js_formula_image_postprocess_failure_restores_bookmark_field_and_object_transaction() -> None:
    _run_node(r'''
let text = "Prefix\r";
let shapeCount = 0, fieldCount = 0, bookmarkCount = 0, tableCount = 0;
let bookmarkCountReads = 0, failLayout = false;
const selection = {Document: null, Range: null, OMaths: null};
const postFailure = new Error("postprocess failed");
postFailure.code = "FIELD_REFRESH_FAILED";
function range(start, end) {
  return {Start: start, End: end,
    Font: {Italic: 0, Color: 0},
    Shading: {BackgroundPatternColor: 0},
    ParagraphFormat: {TabStops: {Add: function() {
      if (failLayout) throw postFailure;
    }}},
    OMaths: {Add: function() { return null; }},
    Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) {
      value = String(value);
      text = text.slice(0, this.End) + value + text.slice(this.End);
      this.End += value.length;
    },
    Delete: function() {
      text = text.slice(0, this.Start) + text.slice(this.End);
      shapeCount = 0; fieldCount = 0; bookmarkCount = 0; tableCount = 0;
    }};
}
const document = {
  Name: "Image-postprocess.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; },
  Range: range,
  get Paragraphs() { return {Count: 2, Item: function() {
    return {Range: {Start: 7, End: text.length + 1}};
  }}; },
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: {Count: 0, Add: function() { return null; }, Item: function() {}},
  InlineShapes: {get Count() { return shapeCount; }, AddPicture: function(a, b, c, target) {
    const start = target.End;
    target.InsertAfter("<IMAGE>");
    shapeCount += 1;
    return {Range: {Start: start, End: target.End, ParagraphFormat: {}}};
  }},
  Fields: {get Count() { return fieldCount; }, Add: function(target) {
    const start = target.End;
    target.InsertAfter("1");
    fieldCount += 1;
    return {Update: function(){}, Result: {Text: "1", Start: start, End: target.End}};
  }},
  Bookmarks: {get Count() { bookmarkCountReads += 1; return bookmarkCount; },
    Add: function() { bookmarkCount += 1; failLayout = true; }},
  Tables: {get Count() { return tableCount; }}
};
selection.Document = document;
document._wpscRunOwnsAppendCursor = true;
document._wpscAppendCursorRange = document.Range(7, 7);
document._wpscNativeFields = [];
const args = {
  fallbackResource: {fallbackResourceId: "image"}, fallbackText: "x^2",
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)};
assert.throws(function() {
  window.WPSComposerLongformV2.__test.addFormulaNativeFallback(
    document, args, {image: "/private/image.png"},
    {ownerNodeId: "eq:image", issues: [], childResults: [], controllerOwned: true},
    "EQUATION_INSERT_FAILED"
  );
}, function(error) { return error === postFailure; });
assert.equal(text, "Prefix\r");
assert.equal(shapeCount, 0);
assert.equal(fieldCount, 0);
assert.equal(bookmarkCount, 0);
assert.equal(tableCount, 0);
assert.equal(document._wpscNativeFields.length, 0);
assert.ok(bookmarkCountReads >= 2, "bookmarks must be captured and revalidated");
''')


def test_js_formula_recovery_is_fatal_if_rollback_consumes_prior_paragraph_boundary() -> None:
    _run_node(r'''
let text = "citation\r", paragraphCount = 2, rollbackDeletes = 0;
const selection = {Document: null, Range: null, OMaths: null};
function paragraphFormat() { return {TabStops: {Add: function(){}}}; }
function range(start, end) {
  const boundedStart = Math.min(start, text.length);
  const boundedEnd = Math.min(end, text.length);
  return {Start: boundedStart, End: boundedEnd, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
  ParagraphFormat: paragraphFormat(),
  OMaths: {Add: function() { return null; }},
  Select: function() { selection.Range = this; selection.OMaths = this.OMaths; },
  get Text() { return text.slice(boundedStart, boundedEnd); },
      InsertAfter: function(value) {
        value = String(value); text += value;
        this.End += value.length;
        paragraphCount += (value.match(/\r/g) || []).length;
      },
  Delete: function() {
    rollbackDeletes += 1;
    text = text.slice(0, start);
    if (text.endsWith("\r")) { text = text.slice(0, -1); paragraphCount -= 1; }
  }}; }
const document = {Name: "Recovery.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  get Paragraphs() { return {Count: paragraphCount, Item: function(index) {
    assert.equal(index, paragraphCount);
    const start = text.lastIndexOf("\r") + 1;
    return {Range: {Start: start, End: text.length + 1}};
  }}; },
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  OMaths: {Count: 0, Add: function() { return null; }, Item: function() {}},
  Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }},
  Bookmarks: {Add: function() {}}};
selection.Document = document;
const operation = {op: "writer.add_equation", nodeId: "eq:one", args: {
  renderMode: "native-m4", content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  fallbackText: "x+y", numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"],
    fallback: "explicit-image-then-source-notice"}};
assert.throws(() => window.WPSComposerLongformV2.__test.runOperation(
  document, operation, {}, [], []
), error => error.code === "LOCAL_MUTATION_ROLLBACK_FAILED");
assert.ok(rollbackDeletes >= 1);
assert.equal(text, "citation");
''')


def test_js_formula_image_rollback_requires_a_paragraph_boundary_checkpoint() -> None:
    _run_node(r'''
let text = "", imageAttempts = 0;
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0},
  ParagraphFormat: {TabStops: {Add: function(){}}},
  InsertAfter: function(value) { text += String(value); }, Delete: function() { text = ""; }}; }
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  PageSetup: {PageWidth: 595, LeftMargin: 64, RightMargin: 64},
  InlineShapes: {AddPicture: function() { imageAttempts += 1; return {}; }},
  Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }},
  Bookmarks: {Add: function() {}}};
assert.throws(() => window.WPSComposerLongformV2.__test.addFormulaNativeFallback(
  document, {fallbackResource: {fallbackResourceId: "image"}, fallbackText: "x+y",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24)}, {image: "/private/image.png"},
  {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true},
  "EQUATION_INSERT_FAILED"
), error => error.code === "CAPABILITY_MISMATCH");
assert.equal(imageAttempts, 0);
assert.equal(text, "");
''')


def test_js_mixed_citation_recovery_rebuilds_run_results_and_planned_issues() -> None:
    _run_node(r'''
let text = "";
const selection = {Document: null, Range: null};
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
  get Text() { return text.slice(this.Start, this.End); },
  Select: function() { selection.Range = this; },
  InsertAfter: function(value) {
    value = String(value); text += value; this.End += value.length;
  }, Delete: function() { text = ""; document.Fields.Count = 0; }}; }
const document = {Name: "CitationRecovery.docx", ActiveWindow: {Selection: selection},
  get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
  Bookmarks: {Exists: function() { return true; },
    Item: function() { return {Range: {Fields: {Count: 2}}}; }},
  Fields: {Count: 0, Add: function(target) {
    const start = target.End;
    if (this.Count === 0) {
      target.InsertAfter("1"); this.Count = 2;
      return {Result: {Text: "1", Start: start, End: target.End}};
    }
    const hostStart = target.End;
    text += "<GHOST>"; this.Count = 4;
    return {Result: {Text: "", Start: hostStart + 1, End: text.length},
      Range: {Start: 0, End: text.length}};
  }}};
selection.Document = document;
const fallback = "[REFERENCE_UNRESOLVED 引用目标未解析]";
const issues = [], children = [];
window.WPSComposerLongformV2.__test.runOperation(document, {
  op: "writer.add_cross_reference", nodeId: "p", args: {runs: [
    {type: "citation", nodeId: "p/cite:0", targetId: "a", targetNodeId: "ref:a",
      number: 1, fallbackText: "[1]"},
    {type: "text", text: " then "},
    {type: "degradation", nodeId: "p/cite:1", code: "REFERENCE_UNRESOLVED", fallbackText: fallback},
    {type: "text", text: " and "},
    {type: "reference", prefix: "(", bookmarkName: "wpsc_eq_" + "e".repeat(24),
      suffix: ")", fallbackText: "1-1"},
    {type: "text", text: " and "},
    {type: "reference", prefix: "(", bookmarkName: "wpsc_eq_" + "f".repeat(24),
      suffix: ")", fallbackText: "1-2"}
  ]}, failurePolicy: {mode: "degrade", recoverableCodes: ["CROSS_REFERENCE_FAILED"],
    fallback: "inline-fallback"}
}, {}, issues, children);
assert.equal(text, "[1] then " + fallback + " and (1-1) and (1-2)\r");
assert.equal(document.Fields.Count, 0);
assert.equal(document._wpscNativeFields.length, 0);
assert.deepEqual(children, [
  {nodeId: "p/cite:0", status: "applied"},
  {nodeId: "p/cite:1", status: "degraded", issueCode: "REFERENCE_UNRESOLVED"}
]);
assert.deepEqual(issues.map(function(issue) { return [issue.code, issue.nodeId]; }), [
  ["REFERENCE_UNRESOLVED", "p/cite:1"],
  ["CROSS_REFERENCE_FAILED", "p"]
]);
''')


def test_js_partial_reference_without_a_bounded_target_is_fatal() -> None:
    _run_node(r'''
let text = "";
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
  get Text() { return text.slice(this.Start, this.End); },
  InsertAfter: function(value) { value = String(value); text += value; this.End += value.length; },
  Delete: function() { throw new Error("rollback must not guess"); }}; }
const fieldError = new Error("private partial reference"); fieldError.code = "CROSS_REFERENCE_FAILED";
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
  Bookmarks: {Exists: function() { return true; },
    Item: function() { return {Range: {Fields: {Count: 1}}}; }},
  Fields: {Count: 0, Add: function() {
    text += "<UNBOUNDED-GHOST>"; this.Count = 1; throw fieldError;
  }}};
assert.throws(() => window.WPSComposerLongformV2.__test.runOperation(document, {
  op: "writer.add_cross_reference", nodeId: "p", args: {runs: [
    {type: "reference", prefix: "(", bookmarkName: "wpsc_eq_" + "e".repeat(24),
      suffix: ")", fallbackText: "1-1"}
  ]}, failurePolicy: {mode: "degrade", recoverableCodes: ["CROSS_REFERENCE_FAILED"],
    fallback: "inline-fallback"}
}, {}, [], []), error => error.code === "LOCAL_MUTATION_ROLLBACK_FAILED");
assert.ok(text.includes("<UNBOUNDED-GHOST>"));
assert.equal(document.Fields.Count, 1);
''')


def test_js_returned_reference_without_field_count_proof_is_fatal() -> None:
    _run_node(r'''
let text = "";
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
  get Text() { return text.slice(this.Start, this.End); },
  InsertAfter: function(value) { value = String(value); text += value; this.End += value.length; },
  Delete: function() { throw new Error("rollback must not guess"); }}; }
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
  Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
  Bookmarks: {Exists: function() { return true; },
    Item: function() { return {Range: {Fields: {Count: 1}}}; }},
  Fields: {Count: 0, Add: function(target) {
    const start = target.End; text += "<UNPROVEN-FIELD>";
    return {Result: {Text: "", Start: start, End: text.length}};
  }}};
assert.throws(() => window.WPSComposerLongformV2.__test.runOperation(document, {
  op: "writer.add_cross_reference", nodeId: "p", args: {runs: [
    {type: "reference", prefix: "(", bookmarkName: "wpsc_eq_" + "e".repeat(24),
      suffix: ")", fallbackText: "1-1"}
  ]}, failurePolicy: {mode: "degrade", recoverableCodes: ["CROSS_REFERENCE_FAILED"],
    fallback: "inline-fallback"}
}, {}, [], []), error => error.code === "LOCAL_MUTATION_ROLLBACK_FAILED");
assert.ok(text.includes("<UNPROVEN-FIELD>"));
assert.equal(document.Fields.Count, 0);
''')


def test_js_reference_field_delta_must_match_a_closed_number_shell() -> None:
    _run_node(r'''
[0, 3].forEach(function(sourceFieldCount) {
  let text = "", addCalls = 0;
  function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) { value = String(value); text += value; this.End += value.length; },
    Delete: function() { text = ""; }}; }
  const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
    Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
    Bookmarks: {Exists: function() { return true; },
      Item: function() { return {Range: {Fields: {Count: sourceFieldCount}}}; }},
    Fields: {Count: 0, Add: function() { addCalls += 1; return {}; }}};
  assert.throws(() => window.WPSComposerLongformV2.__test.runOperation(document, {
    op: "writer.add_cross_reference", nodeId: "p", args: {runs: [
      {type: "reference", prefix: "(", bookmarkName: "wpsc_eq_" + "e".repeat(24),
        suffix: ")", fallbackText: "1-1"}
    ]}, failurePolicy: {mode: "degrade", recoverableCodes: ["CROSS_REFERENCE_FAILED"],
      fallback: "inline-fallback"}
  }, {}, [], []), error => error.code === "CAPABILITY_MISMATCH");
  assert.equal(addCalls, 0);
  assert.equal(text, "");
});
''')


def test_js_reference_engine_and_unknown_failures_never_degrade() -> None:
    _run_node(r'''
function exercise(error, partial) {
  let text = "";
  function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
    get Text() { return text.slice(this.Start, this.End); },
    InsertAfter: function(value) { value = String(value); text += value; this.End += value.length; },
    Delete: function() { text = ""; document.Fields.Count = 0; }}; }
  const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range,
    Paragraphs: {Count: 1, Item: function() { return {Range: {Start: 0, End: text.length + 1}}; }},
    Bookmarks: {Exists: function() { return true; },
      Item: function() { return {Range: {Fields: {Count: 1}}}; }},
    Fields: {Count: 0, Add: function(target) {
      if (partial) { target.InsertAfter("<BOUNDED-GHOST>"); this.Count = 1; }
      throw error;
    }}};
  const issues = [];
  assert.throws(() => window.WPSComposerLongformV2.__test.runOperation(document, {
    op: "writer.add_cross_reference", nodeId: "p", args: {runs: [
      {type: "reference", prefix: "(", bookmarkName: "wpsc_eq_" + "e".repeat(24),
        suffix: ")", fallbackText: "1-1"}
    ]}, failurePolicy: {mode: "degrade", recoverableCodes: ["CROSS_REFERENCE_FAILED"],
      fallback: "inline-fallback"}
  }, {}, issues, []), failure => failure.code === (error.code || "EXECUTION_ABORTED"));
  assert.equal(text, "");
  assert.equal(document.Fields.Count, 0);
  assert.equal(document._wpscNativeFields.length, 0);
  assert.deepEqual(issues, []);
}
const engine = new Error("private engine detail"); engine.code = "ENGINE_LOST";
exercise(engine, false);
exercise(engine, true);
exercise(new Error("private raw detail"), false);
exercise(new Error("private raw detail"), true);
''')


def test_js_citation_degradations_are_literal_and_keep_run_occurrences() -> None:
    _run_node(r'''
let text = "";
function range(start, end) { return {Start: start, End: end, Font: {Italic: 0, Color: 0}, Shading: {BackgroundPatternColor: 0, Texture: 0}, ParagraphFormat: {},
  InsertAfter: function(value) { text += String(value); }}; }
const fallback = "[REFERENCE_UNRESOLVED 引用目标未解析]";
const context = {ownerNodeId: "p", issues: [], childResults: [], controllerOwned: true};
window.WPSComposerLongformV2.__test.addCitationParagraph({
  get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range
}, {runs: [
  {type: "text", text: "before "},
  {type: "citation", nodeId: "p/cite:0", targetId: "ref:a", targetNodeId: "ref:a",
    number: 1, fallbackText: "[1]"},
  {type: "text", text: " then "},
  {type: "degradation", nodeId: "p/cite:1", code: "REFERENCE_UNRESOLVED", fallbackText: fallback},
  {type: "text", text: " middle "},
  {type: "degradation", nodeId: "p/cite:2", code: "REFERENCE_UNRESOLVED", fallbackText: fallback}
]}, {}, context);
assert.equal(text, "before [1] then " + fallback + " middle " + fallback + "\r");
assert.deepEqual(context.issues.map(function(issue) { return issue.nodeId; }), ["p/cite:1", "p/cite:2"]);
assert.deepEqual(context.childResults, [
  {nodeId: "p/cite:0", status: "applied"},
  {nodeId: "p/cite:1", status: "degraded", issueCode: "REFERENCE_UNRESOLVED"},
  {nodeId: "p/cite:2", status: "degraded", issueCode: "REFERENCE_UNRESOLVED"}
]);
''')


def test_js_bibliography_preserves_legacy_text_and_has_closed_failure_boundary() -> None:
    _run_node(r'''
let text = "";
const formats = [];
function range(start, end) {
  const format = {};
  formats.push(format);
  return {Start: start, End: end, ParagraphFormat: format,
    InsertAfter: function(value) { text += String(value); }};
}
const document = {get Content() { return {End: text.length + 1, get Text() { return text; }}; }, Range: range};
const api = window.WPSComposerLongformV2.__test;
api.addBibliographyNative(document, {entries: ["Legacy entry", "[2] Already numbered"], style: "numbered"});
assert.equal(text, "Legacy entry\r[2] Already numbered\r");
text = ""; formats.length = 0;
api.addBibliographyNative(document, {schemaVersion: 1, entries: [
  {id: "a", nodeId: "ref:a", number: 1, text: "Alpha.", cited: true}
], style: "numeric", hangingIndentPt: 18, leftIndentPt: 18, spaceAfterPt: 6});
assert.equal(text, "[1] Alpha.\r");
const structured = formats.find(function(format) { return format.LeftIndent === 18; });
assert.equal(structured.FirstLineIndent, -18);
assert.equal(structured.SpaceBefore, 0);
assert.equal(structured.SpaceAfter, 6);
assert.equal(structured.KeepTogether, -1);
assert.throws(() => api.addBibliographyNative(document, {}), error => error.code === "CONFIGURATION_INVALID");
assert.throws(() => api.addBibliographyNative(document, {schemaVersion: 1, entries: []}),
  error => error.code === "CONFIGURATION_INVALID");
const beforeEmptyLegacy = text;
api.addBibliographyNative(document, {entries: [], style: "numbered"});
assert.equal(text, beforeEmptyLegacy);

const unknown = new Error("unknown insert API failure");
const unknownDocument = {Content: {End: 1}, Range: function(start, end) { return {
  Start: start, End: end, ParagraphFormat: {}, InsertAfter: function() { throw unknown; }
}; }};
assert.throws(() => api.addBibliographyNative(unknownDocument, {entries: ["Legacy"]}),
  error => error === unknown);
const engine = new Error("engine lost"); engine.code = "ENGINE_LOST";
const engineDocument = {Content: {End: 1}, Range: function(start, end) { return {
  Start: start, End: end, InsertAfter: function(value){ this.End += String(value).length; },
  get ParagraphFormat() { throw engine; }
}; }};
assert.throws(() => api.addBibliographyNative(engineDocument, {entries: ["Legacy"]}),
  error => error === engine);
const recoverable = new Error("named format failure"); recoverable.code = "BIBLIOGRAPHY_INSERT_FAILED";
const recoverableDocument = {Content: {End: 1}, Range: function(start, end) { return {
  Start: start, End: end, InsertAfter: function(value){ this.End += String(value).length; },
  get ParagraphFormat() { throw recoverable; }
}; }};
assert.throws(() => api.addBibliographyNative(recoverableDocument, {entries: ["Legacy"]}),
  error => error === recoverable);
''')
