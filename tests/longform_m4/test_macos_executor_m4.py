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
function makeRange(start, end) {
  return {
    Start: start, End: end, Font: {}, Shading: {}, ParagraphFormat: {},
    get Text() { return text.slice(start, end); },
    set Text(value) { text = text.slice(0, start) + String(value) + text.slice(end); },
    InsertAfter: function(value) { text += String(value); this.End = text.length; },
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
  get Content() { return {End: text.length + 1}; },
  Range: makeRange,
  OMaths: maths,
  Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }},
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
assert.deepEqual(added, [[0, 3, "x+y"]]);
assert.equal(builds.length, 1);
assert.equal(maths.Count, 1);
assert.equal(bookmarks.length, 1);
assert.ok(text.includes("x+y\t("));
''')


def test_js_formula_controller_attempts_image_once_then_places_source_notice_and_continues() -> None:
    _run_node(r'''
let text = "";
let imageAttempts = 0;
function makeRange(start, end) {
  return {
    Start: start, End: end, Font: {}, Shading: {}, ParagraphFormat: {},
    InsertAfter: function(value) { text += String(value); this.End = text.length; },
    Delete: function() { text = text.slice(0, start) + text.slice(end); }
  };
}
const document = {
  get Content() { return {End: text.length + 1}; }, Range: makeRange,
  OMaths: {Count: 0, Add: function() { return null; }, Item: function() { return null; }},
  InlineShapes: {AddPicture: function() { imageAttempts += 1; throw new Error("image failed"); }},
  Fields: {Add: function() { return {Update: function(){}, Result: {Text: "1"}}; }},
  Bookmarks: {Add: function() {}},
  _wpscRecoveryController: window.WPSComposerLongformV2.__test.createLocalRecoveryController()
};
const issues = [], children = [];
const equation = {
  op: "writer.add_equation", nodeId: "eq:one",
  args: {renderMode: "native-m4", content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y", sourceHash: "2".repeat(64)}},
    fallbackResource: {fallbackResourceId: "formula-image-1"}, fallbackText: "x+y",
    numbering: {mode: "global", sequenceId: "WPSC_EQ", chapterStyleLevel: null, resetLevel: null, prefix: "(", suffix: ")"},
    bookmarkName: "wpsc_eq_" + "e".repeat(24)},
  failurePolicy: {mode: "degrade", recoverableCodes: ["EQUATION_INSERT_FAILED"], fallback: "explicit-image-then-source-notice"}
};
window.WPSComposerLongformV2.__test.runOperation(document, equation, {"formula-image-1": "/private/staged.png"}, issues, children);
window.WPSComposerLongformV2.__test.runOperation(document, {op: "writer.add_paragraph", nodeId: "p:later", args: {text: "later"}}, {}, issues, children);
assert.equal(imageAttempts, 1);
assert.equal(issues.filter(x => x.code === "EQUATION_INSERT_FAILED").length, 1);
assert.ok(text.includes("x+y"));
assert.ok(text.includes("later"));
assert.ok(!JSON.stringify(issues).includes("/private"));
''')


def test_js_citations_bibliography_and_table_metadata_are_paragraph_and_cell_local() -> None:
    _run_node(r'''
let text = "";
const formats = [];
function makeRange(start, end) {
  const format = {};
  formats.push(format);
  return {Start: start, End: end, Font: {}, Shading: {}, ParagraphFormat: format,
    InsertAfter: function(value) { text += String(value); this.End = text.length; }};
}
const document = {get Content() { return {End: text.length + 1}; }, Range: makeRange};
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
const cells = {"2:1": {Range: {Text: "[1]", Font: {}, Shading: {}}},
               "2:2": {Range: {Text: "[REFERENCE_UNRESOLVED \u5f15\u7528\u76ee\u6807\u672a\u89e3\u6790]", Font: {}, Shading: {}}}};
const table = {Cell: function(row, col) { return cells[row + ":" + col]; }};
window.WPSComposerLongformV2.__test.applyTableCellMetadata(table, {
  cellCitations: [{row: 2, column: 1}],
  cellDegradations: [{row: 2, column: 2}]
});
assert.equal(cells["2:1"].Range.Text, "[1]");
assert.equal(cells["2:2"].Range.Text, "[REFERENCE_UNRESOLVED \u5f15\u7528\u76ee\u6807\u672a\u89e3\u6790]");
assert.equal(cells["2:1"].Range.Font.Italic, undefined);
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


def test_js_omath_verification_rejects_escape_and_does_not_normalize_unknown_api_errors() -> None:
    _run_node(r'''
let text = "";
function range(start, end) { return {Start: start, End: end, ParagraphFormat: {},
  get Text() { return text.slice(start, end); },
  InsertAfter: function(value) { text += String(value); }}; }
const args = {content: {nativeMath: {syntax: "wps-linear-v1", linearText: "x+y"}},
  numbering: {mode: "global", sequenceId: "WPSC_EQ", prefix: "(", suffix: ")"},
  bookmarkName: "wpsc_eq_" + "e".repeat(24), fallbackText: "x+y"};
const context = {ownerNodeId: "eq:one", issues: [], childResults: [], controllerOwned: true};
const native = {Range: {Start: 0, End: 3}, BuildUp: function() { this.Range.End = 99; }};
const escaping = {Count: 0, Item: function() { return native; },
  Add: function() { this.Count = 1; return {Start: 0, End: 3, OMaths: {Count: 1, Item: function() { return native; }}}; }};
const document = {get Content() { return {End: text.length + 1}; }, Range: range, OMaths: escaping};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(document, args, {}, context),
  error => error.code === "EQUATION_INSERT_FAILED");
const unknown = new Error("unknown OMath getter failure");
document.OMaths = {get Count() { throw unknown; }, Add: function() {}};
assert.throws(() => window.WPSComposerLongformV2.__test.addEquationNativeM4(document, args, {}, context),
  error => error === unknown);
''')


def test_js_planned_formula_uses_validated_image_without_entering_omath() -> None:
    _run_node(r'''
let text = "";
let imageAttempts = 0;
function range(start, end) { return {Start: start, End: end, Font: {}, Shading: {}, ParagraphFormat: {},
  InsertAfter: function(value) { text += String(value); }, Delete: function() {}}; }
const document = {get Content() { return {End: text.length + 1}; }, Range: range,
  InlineShapes: {AddPicture: function(locator) { imageAttempts += 1; assert.equal(locator, "/private/formula.png"); return {}; }},
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
''')
