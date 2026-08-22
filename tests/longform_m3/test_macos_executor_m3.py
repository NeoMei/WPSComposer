from __future__ import annotations

import hashlib
import json
import os
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Optional

import pytest

from skills.WPSComposer.scripts.generation_plan import (
    GenerationOperation,
    GenerationPlan,
)
from skills.WPSComposer.scripts.longform.resources import (
    ImageProfile,
    PreparedLongformResource,
)
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult


ROOT = Path("macos/wps-jsapi-probe/addin")


def _manifest_digest(resources: tuple[PreparedLongformResource, ...]) -> str:
    entries = [
        {
            "resourceId": resource.id,
            "sourceSha256": resource.source_sha256,
            "payloadSha256": resource.payload_sha256,
            "byteLength": len(resource.payload_bytes),
            "mediaType": resource.media_type,
            "normalizerId": resource.normalizer_id,
        }
        for resource in resources
    ]
    raw = json.dumps(
        {"version": "1", "entries": sorted(entries, key=lambda item: item["resourceId"])},
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return "sha256:" + hashlib.sha256(raw).hexdigest()


def _resource(resource_id: str = "image-1", payload: bytes = b"PNG") -> PreparedLongformResource:
    return PreparedLongformResource(
        id=resource_id,
        media_type="image/png",
        source_sha256=hashlib.sha256(payload).hexdigest(),
        payload_sha256=hashlib.sha256(payload).hexdigest(),
        normalizer_id="none-v1",
        payload_bytes=payload,
        image_profile=ImageProfile(2, 1, None, None, 1, "PNG", False),
    )


def _plan(resources: tuple[PreparedLongformResource, ...] = ()) -> GenerationPlan:
    operations = [
        GenerationOperation(op="writer.reset", args={}, node_id="doc:reset"),
    ]
    if resources:
        operations.append(
            GenerationOperation(
                op="writer.add_captioned_figure",
                node_id="fig:one",
                args={
                    "caption": "Figure",
                    "numbering": {
                        "mode": "global",
                        "sequenceId": "WPSC_FIG",
                        "chapterStyleLevel": None,
                        "resetLevel": None,
                        "prefix": "图 ",
                        "suffix": "",
                    },
                    "bookmarkName": "wpsc_fig_" + "a" * 24,
                    "indexable": True,
                    "referenceable": True,
                    "widthMode": "full",
                    "orientation": "portrait",
                    "kind": "diagram",
                    "children": [
                        {
                            "nodeId": "fig:one/image:1",
                            "resourceId": resources[0].id,
                            "displayWidthPt": 120.0,
                            "displayHeightPt": 60.0,
                            "effectiveDpi": 144.0,
                            "mediaType": "image/png",
                            "normalizerId": "none-v1",
                        }
                    ],
                    "layout": "stack",
                    "keepWithCaption": True,
                },
                failure_policy={
                    "mode": "degrade",
                    "recoverableCodes": ["IMAGE_INSERT_FAILED"],
                    "fallback": "figure-child-stack-then-notice",
                },
            )
        )
    operations.append(
        GenerationOperation(
            op="writer.finalize_fields",
            args={"maxRounds": 3},
            node_id="doc:finalize",
        )
    )
    return GenerationPlan(
        component="writer",
        operations=tuple(operations),
        protocol_version=2,
        semantic_version="longform-1",
        resource_manifest_version=1,
        resource_manifest_digest=_manifest_digest(resources),
    )


@dataclass
class _Command:
    id: str = "m3-command"


class _Bridge:
    def __init__(self, *, error: Optional[Exception] = None) -> None:
        self.error = error
        self.params: Optional[Mapping[str, Any]] = None

    def issue(self, component: str, method: str, params: Mapping[str, Any]) -> _Command:
        assert component == "writer"
        assert method == "generate_longform_document"
        self.params = params
        return _Command()

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        if self.error is not None:
            raise self.error
        assert self.params is not None
        child_results = [
            {"nodeId": child["nodeId"], "status": "applied"}
            for operation in self.params["plan"]["operations"]
            if operation["op"] == "writer.add_captioned_figure"
            for child in operation["args"].get("children", [])
        ]
        return ProbeResult(
            command_id,
            True,
            {
                "outputPath": self.params["outputPath"],
                "appliedOperations": len(self.params["plan"]["operations"]),
                "issueCodes": [],
                "fieldSnapshots": [
                    [{
                        "stableKey": ["doc:finalize", "PAGE", 0],
                        "fieldCategory": "page",
                        "resultHash": hashlib.sha256(b"stable").hexdigest(),
                        "tocPageCount": 0,
                        "figureIndexPageCount": 0,
                        "tableIndexPageCount": 0,
                        "totalPages": 1,
                    }],
                    [{
                        "stableKey": ["doc:finalize", "PAGE", 0],
                        "fieldCategory": "page",
                        "resultHash": hashlib.sha256(b"stable").hexdigest(),
                        "tocPageCount": 0,
                        "figureIndexPageCount": 0,
                        "tableIndexPageCount": 0,
                        "totalPages": 1,
                    }],
                ],
                "childResults": child_results,
                "paginationMap": {"version": "M2-stub", "nodes": []},
            },
            None,
        )


def test_executor_validates_manifest_stages_private_payload_and_cleans_success(tmp_path: Path) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor

    resource = _resource()
    bridge = _Bridge()
    outcome = MacOSLongformExecutor(
        bridge=bridge, staging_dir=str(tmp_path)
    ).execute(_plan((resource,)), (resource,))

    assert outcome.applied_operations == 3
    assert bridge.params is not None
    assert set(bridge.params) == {"plan", "outputPath", "resources"}
    locator = Path(bridge.params["resources"][resource.id])
    assert locator.parent == tmp_path
    assert resource.id not in locator.name
    assert not locator.exists()
    serialized = json.dumps(outcome.to_dict(), ensure_ascii=False)
    assert str(locator) not in serialized
    assert resource.payload_sha256 not in serialized


def test_executor_rejects_payload_hash_before_bridge_and_leaves_no_resource(tmp_path: Path) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    good = _resource()
    bad = PreparedLongformResource(
        id=good.id,
        media_type=good.media_type,
        source_sha256=good.source_sha256,
        payload_sha256="0" * 64,
        normalizer_id=good.normalizer_id,
        payload_bytes=good.payload_bytes,
        image_profile=good.image_profile,
    )
    bridge = _Bridge()
    with pytest.raises(MacOSLongformExecutorError, match="resource validation"):
        MacOSLongformExecutor(bridge=bridge, staging_dir=str(tmp_path)).execute(
            _plan((good,)), (bad,)
        )
    assert bridge.params is None
    assert not tuple(tmp_path.glob("wpsc-resource-*"))


@pytest.mark.parametrize("native_error", [TimeoutError("secret timeout /private/a.png"), RuntimeError("secret /private/a.png")])
def test_executor_cleans_private_payload_and_sanitizes_bridge_failures(
    tmp_path: Path, native_error: Exception
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    resource = _resource()
    bridge = _Bridge(error=native_error)
    with pytest.raises(MacOSLongformExecutorError, match="Bridge command failed") as exc_info:
        MacOSLongformExecutor(
            bridge=bridge, staging_dir=str(tmp_path)
        ).execute(_plan((resource,)), (resource,))
    assert "/private" not in str(exc_info.value)
    assert not tuple(tmp_path.glob("wpsc-resource-*"))


def test_executor_does_not_downgrade_private_staging_failure_to_bridge_issue(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    resource = _resource()
    executor = MacOSLongformExecutor(bridge=_Bridge(), staging_dir=str(tmp_path))

    def fail_stage(*_args: Any, **_kwargs: Any) -> Any:
        raise MacOSLongformExecutorError("Private resource staging failed")

    monkeypatch.setattr(executor, "_stage_resources", fail_stage)
    with pytest.raises(MacOSLongformExecutorError, match="staging failed"):
        executor.execute(_plan((resource,)), (resource,))


def test_executor_treats_permanent_private_cleanup_failure_as_fatal(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    resource = _resource()
    executor = MacOSLongformExecutor(bridge=_Bridge(), staging_dir=str(tmp_path))

    def fail_cleanup(*_args: Any, **_kwargs: Any) -> None:
        raise OSError("secret /private/locked.png")

    monkeypatch.setattr(executor, "_cleanup_resources", fail_cleanup)
    with pytest.raises(MacOSLongformExecutorError, match="cleanup failed") as exc_info:
        executor.execute(_plan((resource,)), (resource,))
    assert exc_info.value.cleanup_failed is True
    assert "/private" not in str(exc_info.value)


def test_executor_preserves_primary_native_failure_when_private_cleanup_also_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    class FailedBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            return ProbeResult(
                command_id,
                False,
                {},
                {"code": "SAVE_FAILED", "message": "secret /private/output.docx"},
            )

    executor = MacOSLongformExecutor(
        bridge=FailedBridge(), staging_dir=str(tmp_path)
    )
    monkeypatch.setattr(
        executor,
        "_cleanup_resources",
        lambda *_args, **_kwargs: (_ for _ in ()).throw(
            OSError("secret /private/locked.png")
        ),
    )
    with pytest.raises(MacOSLongformExecutorError, match="Execution aborted") as exc_info:
        executor.execute(_plan((_resource(),)), (_resource(),))
    assert exc_info.value.cleanup_failed is True
    assert "/private" not in str(exc_info.value)


@pytest.mark.parametrize(
    "code",
    [
        "EXECUTION_ABORTED",
        "FIELD_REFRESH_FAILED",
        "SAVE_FAILED",
        "LOCAL_MUTATION_ROLLBACK_FAILED",
        "DEGRADATION_FALLBACK_FAILED",
        "GENERATION_COMMAND_FAILED",
        "UNEXPECTED_NATIVE_FAILURE",
    ],
)
def test_executor_keeps_required_native_api_failures_fatal_and_redacted(
    tmp_path: Path, code: str
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    class FailedBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            return ProbeResult(
                command_id,
                False,
                {},
                {"code": code, "message": "secret /private/resource.png"},
            )

    with pytest.raises(MacOSLongformExecutorError) as exc_info:
        MacOSLongformExecutor(
            bridge=FailedBridge(), staging_dir=str(tmp_path)
        ).execute(_plan(), ())
    assert "/private" not in str(exc_info.value)


@pytest.mark.parametrize("private_surface", ["issue", "pagination", "code"])
def test_executor_rejects_private_response_references_before_public_outcome(
    tmp_path: Path, private_surface: str
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    class PrivateBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            snapshot = {
                "stableKey": ["doc:finalize", "PAGE", 0],
                "fieldCategory": "page",
                "resultHash": hashlib.sha256(b"stable").hexdigest(),
                "tocPageCount": 0,
                "figureIndexPageCount": 0,
                "tableIndexPageCount": 0,
                "totalPages": 1,
            }
            value = {
                "outputPath": str(tmp_path / "out.docx"),
                "appliedOperations": 2,
                "issueCodes": [],
                "fieldSnapshots": [[snapshot], [snapshot]],
                "childResults": [],
                "paginationMap": {"version": "M2-stub", "nodes": []},
            }
            if private_surface in {"issue", "code"}:
                value["issueCodes"] = [{
                    "code": "A" * 64 if private_surface == "code" else "BIBLIOGRAPHY_INSERT_FAILED",
                    "message": "controlled",
                    "placement": "document",
                    "nodeId": (
                        "/private/secret.png"
                        if private_surface == "issue"
                        else "doc:reset"
                    ),
                }]
            else:
                value["paginationMap"] = {
                    "version": "M2-stub",
                    "nodes": [{
                        "nodeId": "/private/secret.png",
                        "fragments": [{"page": 1}],
                    }],
                }
            return ProbeResult(command_id, True, value, None)

    with pytest.raises(MacOSLongformExecutorError, match="Bridge result was invalid") as exc_info:
        MacOSLongformExecutor(
            bridge=PrivateBridge(), staging_dir=str(tmp_path)
        ).execute(_plan(), ())
    assert "/private" not in str(exc_info.value)


def test_executor_requires_field_history_for_m3_native_plan(tmp_path: Path) -> None:
    from skills.WPSComposer.scripts.longform.field_contract import (
        NativeFieldContractError,
    )
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor

    class MissingHistoryBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            result = super().wait_result(command_id, timeout)
            value = dict(result.value)
            del value["fieldSnapshots"]
            return ProbeResult(command_id, True, value, None)

    resource = _resource()
    with pytest.raises(NativeFieldContractError, match="field_history"):
        MacOSLongformExecutor(
            bridge=MissingHistoryBridge(), staging_dir=str(tmp_path)
        ).execute(_plan((resource,)), (resource,))


def test_executor_keeps_m2_native_object_shape_compatible_without_history(
    tmp_path: Path,
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor

    plan = GenerationPlan(
        component="writer",
        operations=(
            GenerationOperation(
                op="writer.add_equation",
                args={"source": "E=mc^2", "number": "1", "fallbackText": "E=mc^2"},
                node_id="eq:legacy",
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
        resource_manifest_digest=_manifest_digest(()),
    )

    class LegacyBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            assert self.params is not None
            return ProbeResult(command_id, True, {
                "outputPath": self.params["outputPath"],
                "appliedOperations": 2,
                "issueCodes": [],
                "childResults": [],
                "paginationMap": {"version": "M2-stub", "nodes": []},
            }, None)

    outcome = MacOSLongformExecutor(
        bridge=LegacyBridge(), staging_dir=str(tmp_path)
    ).execute(plan, ())

    assert outcome.applied_operations == 2


def test_executor_rejects_success_that_omits_planned_operations(tmp_path: Path) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    class IncompleteBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            result = super().wait_result(command_id, timeout)
            return ProbeResult(
                command_id,
                True,
                {**result.value, "appliedOperations": 0},
                None,
            )

    with pytest.raises(MacOSLongformExecutorError, match="Bridge result was invalid"):
        MacOSLongformExecutor(
            bridge=IncompleteBridge(), staging_dir=str(tmp_path)
        ).execute(_plan(), ())


def test_executor_rejects_success_that_omits_planned_figure_children(
    tmp_path: Path,
) -> None:
    from skills.WPSComposer.scripts.longform.macos_executor import (
        MacOSLongformExecutor,
        MacOSLongformExecutorError,
    )

    class IncompleteBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            result = super().wait_result(command_id, timeout)
            return ProbeResult(
                command_id,
                True,
                {**result.value, "childResults": []},
                None,
            )

    resource = _resource()
    with pytest.raises(MacOSLongformExecutorError, match="Bridge result was invalid"):
        MacOSLongformExecutor(
            bridge=IncompleteBridge(), staging_dir=str(tmp_path)
        ).execute(_plan((resource,)), (resource,))


@pytest.mark.parametrize(
    "child",
    [
        {"nodeId": "image:1", "status": "applied", "issueCode": "IMAGE_INSERT_FAILED"},
        {"nodeId": "image:1", "status": "degraded"},
    ],
)
def test_result_model_rejects_inconsistent_child_status_issue_code(
    child: dict[str, str]
) -> None:
    from skills.WPSComposer.scripts.macos_probe.models import (
        ProtocolError,
        validate_longform_generation_value,
    )

    with pytest.raises(ProtocolError, match="result is invalid"):
        validate_longform_generation_value({
            "outputPath": "/private/output.docx",
            "appliedOperations": 1,
            "issueCodes": [],
            "fieldSnapshots": [],
            "childResults": [child],
            "paginationMap": {"version": "M2-stub", "nodes": []},
        })


def test_executor_rejects_legacy_non_sha256_field_history(tmp_path: Path) -> None:
    from skills.WPSComposer.scripts.longform.field_contract import (
        NativeFieldContractError,
    )
    from skills.WPSComposer.scripts.longform.macos_executor import MacOSLongformExecutor

    class LegacyHashBridge(_Bridge):
        def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
            result = super().wait_result(command_id, timeout)
            value = dict(result.value)
            value["fieldSnapshots"] = [
                [dict(result.value["fieldSnapshots"][0][0], resultHash="stable")],
                [dict(result.value["fieldSnapshots"][1][0], resultHash="stable")],
            ]
            return ProbeResult(command_id, True, value, None)

    with pytest.raises(NativeFieldContractError, match="snapshot_fields"):
        MacOSLongformExecutor(
            bridge=LegacyHashBridge(), staging_dir=str(tmp_path)
        ).execute(_plan(), ())


def _run_node(script: str) -> None:
    descriptor, name = tempfile.mkstemp(suffix=".js")
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as stream:
            stream.write(script)
        subprocess.run(["node", name], check=True, capture_output=True, text=True)
    finally:
        os.unlink(name)


def test_addin_exposes_m3_native_handlers_and_only_bibliography_is_deferred() -> None:
    source = (ROOT / "writer-longform-v2.js").read_text(encoding="utf-8")
    for name in (
        "addCaptionedFigureNative",
        "addSemanticTableNative",
        "addEquationNumberNative",
        "addCrossReferenceParagraph",
        "insertCaptionIndexNative",
        "repaginateAndUpdateNumbering",
        "refreshBookmarksAndReferences",
        "refreshIndexes",
        "repaginateAndUpdatePageFields",
        "snapshotFields",
    ):
        assert f"function {name}" in source
    deferred_body = source.split("const LONGFORM_DEFERRED =", 1)[1].split("};", 1)[0]
    assert "writer.add_bibliography" in deferred_body
    for operation in (
        "writer.add_captioned_figure",
        "writer.add_semantic_table",
        "writer.add_equation",
        "writer.add_cross_reference",
    ):
        assert operation not in deferred_body


def test_native_insert_mutation_accepts_an_ordered_operation_group() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
let text = "";
const document = {{
  Content: {{get End() {{ return text.length + 1; }}}},
  Range: function(start, end) {{
    return {{Start: start, End: end, Font: {{}}, ParagraphFormat: {{}},
      InsertAfter: function(value) {{ text += String(value); this.End = text.length; }} }};
  }}
}};
window.WPSComposerLongformV2.__test.applyLongformMutation(document, {{
  type: "insert",
  operations: [
    {{op: "writer.add_paragraph", nodeId: "p:one", args: {{text: "first"}}}},
    {{op: "writer.add_paragraph", nodeId: "p:two", args: {{text: "second"}}}}
  ]
}}, {{}}, [], []);
assert.strictEqual(text, "first\\rsecond\\r");
"""
    _run_node(script)


def test_addin_number_shell_and_cross_reference_use_controlled_native_fields() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
const calls = [];
let position = 0;
const document = {{
  Content: {{get End() {{ return position + 1; }}}},
  Styles: {{Item: function(id) {{ assert.equal(id, -2); return {{NameLocal: "标题 1"}}; }}}},
  Range: function(start, end) {{
    return {{Start: start, End: end, ParagraphFormat: {{}},
      InsertAfter: function(text) {{ position += String(text).length; this.End = position; }},
      Delete: function() {{ position = start; }} }};
  }},
  Fields: {{Add: function(range, type, code, preserve) {{
    calls.push(["field", type, code, preserve]);
    position += 1;
    return {{Result: {{Text: "1"}}, Range: {{Start: position - 1, End: position}}, Update: function() {{}}}};
  }}}},
  Bookmarks: {{Count: 0, Add: function(name, range) {{ calls.push(["bookmark", name, range.Start, range.End]); }}}}
}};
const t = window.WPSComposerLongformV2.__test;
t.addNativeNumberShell(document, {{mode: "chapter", sequenceId: "WPSC_FIG", chapterStyleLevel: 1, resetLevel: 1, prefix: "图 ", suffix: ""}}, "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa", "fig:one");
t.addCrossReferenceParagraph(document, {{runs: [
  {{type: "text", text: "见"}},
  {{type: "reference", targetNodeId: "fig:one", targetKind: "figure", bookmarkName: "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa", prefix: "图 ", suffix: "", fallbackText: "[图]"}},
  {{type: "text", text: "。"}}
]}}, {{}}, {{ownerNodeId: "para:one", issues: [], childResults: []}});
assert.deepEqual(calls.filter(x => x[0] === "field").map(x => x[2]), [
  'STYLEREF "标题 1" \\\\s', "SEQ WPSC_FIG \\\\* ARABIC \\\\s 1", "REF wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa \\\\h"
]);
assert.equal(calls.filter(x => x[0] === "bookmark").length, 1);
assert.ok(calls.find(x => x[0] === "bookmark")[3] > calls.find(x => x[0] === "bookmark")[2]);
"""
    _run_node(script)


def test_addin_field_adapter_runs_five_phases_and_freezes_fourth_snapshot() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
const events = [];
let snapshot = 0;
const adapter = {{
  repaginateAndUpdateNumbering: function() {{ events.push("numbering"); }},
  refreshBookmarksAndReferences: function() {{ events.push("references"); }},
  refreshIndexes: function() {{ events.push("indexes"); }},
  repaginateAndUpdatePageFields: function() {{ events.push("pages"); }},
  snapshotFields: function() {{ events.push("snapshot"); snapshot += 1; return [{{stableKey: ["doc:finalize", "PAGE", 0], fieldCategory: "page", resultHash: String(snapshot), tocPageCount: 0, figureIndexPageCount: 0, tableIndexPageCount: 0, totalPages: snapshot}}]; }}
}};
const issues = [];
const history = window.WPSComposerLongformV2.__test.runNativeFieldConvergence(adapter, 3, issues);
assert.equal(history.length, 4);
assert.equal(events.filter(x => x === "numbering").length, 3);
assert.equal(events.filter(x => x === "snapshot").length, 4);
assert.equal(issues.filter(x => x.code === "FIELD_REFRESH_UNSTABLE").length, 1);
"""
    _run_node(script)


def test_addin_field_adapter_treats_missing_index_count_as_fatal() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
const document = {{
  Repaginate: function() {{}},
  Fields: {{Update: function() {{}}}},
  Bookmarks: {{Count: 0}},
  TablesOfContents: {{}},
  TablesOfFigures: {{Count: 0}},
  ComputeStatistics: function() {{ return 1; }}
}};
assert.throws(function() {{
  window.WPSComposerLongformV2.__test.runNativeFieldConvergence(
    window.WPSComposerLongformV2.__test.nativeFieldAdapter(document), 2, []
  );
}}, function(error) {{ return error.code === "FIELD_REFRESH_FAILED"; }});
"""
    _run_node(script)


def test_addin_page_phase_and_snapshots_match_windows_story_and_page_span_semantics() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
function collection(items) {{
  return {{Count: items.length, Item: function(index) {{ return items[index - 1]; }}}};
}}
function storyField(code) {{
  return {{Code: {{Text: code}}, updates: 0, Update: function() {{ this.updates += 1; }}}};
}}
const page = storyField(" PAGE ");
const title = storyField("TITLE");
const numPages = storyField("NUMPAGES");
const secondPage = storyField("PAGE");
const sections = [
  {{Headers: collection([{{Exists: true, Range: {{Fields: collection([page, title])}}}}]), Footers: collection([{{Exists: true, Range: {{Fields: collection([numPages])}}}}])}},
  {{Headers: collection([{{Exists: true, Range: {{Fields: collection([secondPage])}}}}]), Footers: collection([])}}
];
function pagedRange(start, end, firstPage, lastPage) {{
  return {{
    Start: start,
    End: end,
    get Duplicate() {{
      return {{
        position: start,
        SetRange: function(left) {{ this.position = left; }},
        Information: function() {{ return this.position === start ? firstPage : lastPage; }}
      }};
    }}
  }};
}}
function native(kind, start, end, firstPage, lastPage) {{
  return {{Result: {{Text: kind}}, Range: pagedRange(start, end, firstPage, lastPage), Update: function() {{}}}};
}}
const document = {{
  Repaginate: function() {{}},
  Sections: collection(sections),
  TablesOfContents: {{Count: 1}},
  TablesOfFigures: {{Count: 2}},
  ComputeStatistics: function() {{ return 7; }},
  _wpscNativeFields: [
    {{ownerNodeId: "doc:toc", fieldKind: "TOC", native: native("toc", 1, 5, 1, 2), category: "index"}},
    {{ownerNodeId: "doc:fig", fieldKind: "TOF_FIG", native: native("fig", 10, 15, 3, 4), category: "index"}},
    {{ownerNodeId: "doc:tab", fieldKind: "TOF_TAB", native: native("tab", 20, 20, 5, 5), category: "index"}}
  ]
}};
const adapter = window.WPSComposerLongformV2.__test.nativeFieldAdapter(document);
adapter.repaginateAndUpdatePageFields();
assert.deepEqual([page.updates, title.updates, numPages.updates, secondPage.updates], [1, 0, 1, 1]);
const snapshot = adapter.snapshotFields();
assert.ok(snapshot.length);
assert.deepEqual(snapshot.filter(function(item) {{
  return item.stableKey[1] === "PAGE" || item.stableKey[1] === "NUMPAGES";
}}).map(function(item) {{ return item.stableKey.join("|"); }}).sort(), [
  "section:1/footers:1|NUMPAGES|0",
  "section:1/headers:1|PAGE|0",
  "section:2/headers:1|PAGE|0"
]);
snapshot.forEach(function(item) {{
  assert.deepEqual(
    [item.tocPageCount, item.figureIndexPageCount, item.tableIndexPageCount, item.totalPages],
    [2, 2, 1, 7]
  );
}});
"""
    _run_node(script)


def test_addin_visible_hash_matches_shared_sha256_normalization() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    expected = hashlib.sha256("A\n中文".encode("utf-8")).hexdigest()
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
assert.equal(
  window.WPSComposerLongformV2.__test.hashVisible("A\\r\\n中文"),
  {json.dumps(expected)}
);
"""
    _run_node(script)


def test_addin_required_page_field_insertion_failure_is_fatal() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
const footer = {{
  PageNumbers: {{}},
  Range: {{Text: "", Collapse: function() {{}}, Fields: {{Add: function() {{ throw new Error("/private/page"); }}}}}}
}};
const document = {{
  Sections: {{Count: 1, Item: function() {{ return {{Footers: {{Item: function() {{ return footer; }}}}}}; }}}}
}};
assert.throws(function() {{
  window.WPSComposerLongformV2.__test.runOperation(document, {{
    op: "writer.set_page_numbering", nodeId: "section:1",
    failurePolicy: {{mode: "fail"}}, args: {{format: "arabic"}}
  }}, {{}}, [], []);
}}, function(error) {{ return error.code === "EXECUTION_ABORTED" && !String(error).includes("/private"); }});
"""
    _run_node(script)


def test_addin_figure_columns_use_exact_gap_and_preserve_child_order() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
let position = 0;
const widths = [];
const pictures = [];
function makeRange(start, end) {{ return {{Start: start, End: end, ParagraphFormat: {{}},
  get Duplicate() {{ return makeRange(this.Start, this.End); }},
  InsertAfter: function(text) {{ position += String(text).length; }}, Delete: function() {{ position = start; }}}}; }}
const columns = function(index) {{ return {{SetWidth: function(width) {{ widths[index - 1] = width; }}}}; }};
columns.Item = columns;
const borders = {{}};
const table = {{
  Columns: columns,
  Range: {{End: 1, ParagraphFormat: {{}}}},
  Borders: function(id) {{ return borders[id] = borders[id] || {{LineStyle: null}}; }},
  Cell: function(row, column) {{ return {{Range: makeRange(column * 10, column * 10 + 1)}}; }}
}};
const document = {{
  Content: {{get End() {{ return position + 1; }}}},
  Range: makeRange,
  Tables: {{Add: function() {{ return table; }}}},
  InlineShapes: {{AddPicture: function(path, link, save, target) {{ pictures.push([path, target.Start, target.End]); position += 1; return {{Range: {{ParagraphFormat: {{}}}}}}; }}}}
}};
const children = [
  {{nodeId: "image:1", resourceId: "r1", displayWidthPt: 100, displayHeightPt: 50}},
  {{nodeId: "image:2", resourceId: "r2", displayWidthPt: 110, displayHeightPt: 55}}
];
window.WPSComposerLongformV2.__test.createFigureColumns(document, children, {{r1: "/private/one.png", r2: "/private/two.png"}}, {{ownerNodeId: "fig:1", issues: [], childResults: []}});
assert.deepEqual(widths, [100, 12, 110]);
assert.deepEqual(pictures, [["/private/one.png", 10, 10], ["/private/two.png", 30, 30]]);
assert.deepEqual(Object.values(borders).map(function(border) {{ return border.LineStyle; }}), [0, 0, 0, 0, 0, 0]);
assert.equal(table.Range.ParagraphFormat.KeepTogether, -1);
assert.equal(table.Range.ParagraphFormat.KeepWithNext, -1);
"""
    _run_node(script)


def test_addin_table_applies_resolved_borders_alignment_header_and_merge_order() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
let position = 0;
const mergeOrder = [];
const borders = {{}};
const cells = {{}};
function cell(row, column) {{
  const key = row + ":" + column;
  if (!cells[key]) cells[key] = {{Range: {{Text: "", ParagraphFormat: {{}}}}, Merge: function(target) {{ mergeOrder.push([key, target.key]); }}, key: key}};
  return cells[key];
}}
const headerRow = {{Borders: function(id) {{ const key = "header:" + id; return borders[key] = borders[key] || {{}}; }}, HeadingFormat: 0}};
function row(index) {{ return index === 1 ? headerRow : {{Borders: function() {{ return {{}}; }}}}; }}
const rows = function(index) {{ return row(index); }};
rows.AllowBreakAcrossPages = null;
const table = {{
  Rows: rows,
  Cell: cell,
  Borders: function(id) {{ return borders[id] = borders[id] || {{}}; }}
}};
const document = {{
  Content: {{get End() {{ return position + 1; }}}},
  Range: function(start, end) {{ return {{Start: start, End: end, InsertAfter: function(text) {{ position += String(text).length; }}}}; }},
  Tables: {{Add: function(range, rowCount, columnCount) {{ assert.equal(rowCount, 3); assert.equal(columnCount, 2); return table; }}}}
}};
window.WPSComposerLongformV2.__test.createNativeTable(document, {{
  headers: ["A", "B"], rows: [["1", "2"], ["3", "4"]], alignments: ["left", "right"],
  borderSpec: {{top: 1.5, bottom: 1.5, headerBottom: 0.75, left: 0, right: 0, insideHorizontal: 0, insideVertical: 0}},
  merges: [{{top: 2, left: 1, bottom: 3, right: 1}}, {{top: 1, left: 1, bottom: 1, right: 2}}],
  repeatHeader: true, allowRowSplit: false, cellIndentPt: 0
}});
assert.equal(borders[-1].LineWidth, 12);
assert.equal(borders[-3].LineWidth, 12);
assert.equal(borders["header:-3"].LineWidth, 6);
assert.equal(borders[-6].LineStyle, 0);
assert.equal(rows.AllowBreakAcrossPages, 0);
assert.equal(headerRow.HeadingFormat, -1);
assert.equal(cells["1:1"].Range.ParagraphFormat.FirstLineIndent, 0);
assert.equal(cells["1:1"].Range.ParagraphFormat.Alignment, 0);
assert.equal(cells["1:2"].Range.ParagraphFormat.Alignment, 2);
assert.deepEqual(mergeOrder, [["2:1", "3:1"], ["1:1", "1:2"]]);
"""
    _run_node(script)


def test_addin_local_rollback_failure_escalates_without_native_error_text() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
assert.throws(function() {{
  window.WPSComposerLongformV2.__test.rollbackMutation({{Range: function() {{ return {{Delete: function() {{ throw new Error("/private/secret.png"); }}}}; }}}}, 1, 2);
}}, function(error) {{ return error.code === "LOCAL_MUTATION_ROLLBACK_FAILED" && !String(error).includes("/private"); }});
let position = 2;
const document = {{
  Content: {{get End() {{ return position + 1; }}}},
  Range: function(start, end) {{ return {{Start: start, End: end, Delete: function() {{ throw new Error("/private/secret.png"); }}}}; }},
  InlineShapes: {{AddPicture: function() {{ position = 3; throw new Error("native insert"); }}}}
}};
assert.throws(function() {{
  window.WPSComposerLongformV2.__test.runOperation(document, {{
    op: "writer.add_captioned_figure", nodeId: "fig:1",
    failurePolicy: {{mode: "fail"}},
    args: {{keepWithCaption: true, orientation: "portrait", layout: "stack", children: [
      {{nodeId: "image:1", resourceId: "r1", displayWidthPt: 10, displayHeightPt: 10}}
    ]}}
  }}, {{r1: "/private/secret.png"}}, [], []);
}}, function(error) {{ return error.code === "LOCAL_MUTATION_ROLLBACK_FAILED" && !String(error).includes("/private"); }});
"""
    _run_node(script)


def test_addin_cross_reference_failure_keeps_fallback_in_same_paragraph() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
let text = "";
function range(start, end) {{ return {{Start: start, End: end, InsertAfter: function(value) {{ text += String(value); }}, Delete: function() {{ text = text.slice(0, start); }}}}; }}
const document = {{
  Content: {{get End() {{ return text.length + 1; }}}}, Range: range,
  Fields: {{Add: function() {{ throw new Error("native /private/field"); }}}}
}};
const context = {{ownerNodeId: "para:1", issues: [], childResults: []}};
window.WPSComposerLongformV2.__test.addCrossReferenceParagraph(document, {{runs: [
  {{type: "text", text: "见"}},
  {{type: "reference", bookmarkName: "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa", prefix: "图 ", suffix: "", fallbackText: "[图]"}},
  {{type: "text", text: "。"}}
]}}, {{}}, context);
assert.equal(text, "见图 [图]。\\r");
assert.deepEqual(context.issues.map(x => x.code), ["CROSS_REFERENCE_FAILED"]);
assert.ok(!JSON.stringify(context).includes("/private"));
"""
    _run_node(script)


def test_addin_two_child_stack_rolls_back_failed_child_and_keeps_order() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
let text = "";
const pictures = [];
function range(start, end) {{ return {{Start: start, End: end, ParagraphFormat: {{}}, InsertAfter: function(value) {{ text += String(value); }}, Delete: function() {{ text = text.slice(0, start); }}}}; }}
const document = {{
  Content: {{get End() {{ return text.length + 1; }}}}, Range: range,
  InlineShapes: {{AddPicture: function(path) {{ pictures.push(path); if (path.includes("two")) throw new Error("fail"); text += "I"; return {{Range: {{ParagraphFormat: {{}}}}}}; }}}},
  Fields: {{Add: function(r, type, code) {{ text += "F"; return {{Result: {{Text: "1"}}, Update: function() {{}}}}; }}}},
  Bookmarks: {{Count: 0, Add: function() {{}}}}
}};
const context = {{ownerNodeId: "fig:1", issues: [], childResults: []}};
window.WPSComposerLongformV2.__test.addCaptionedFigureNative(document, {{
  caption: "示例", numbering: {{mode: "global", sequenceId: "WPSC_FIG", chapterStyleLevel: null, resetLevel: null, prefix: "图 ", suffix: ""}},
  bookmarkName: "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa", orientation: "portrait", layout: "stack", keepWithCaption: true,
  children: [
    {{nodeId: "image:1", resourceId: "r1", displayWidthPt: 100, displayHeightPt: 50}},
    {{nodeId: "image:2", resourceId: "r2", displayWidthPt: 100, displayHeightPt: 50}}
  ]
}}, {{r1: "/private/one.png", r2: "/private/two.png"}}, context);
assert.deepEqual(context.childResults, [
  {{nodeId: "image:1", status: "applied"}},
  {{nodeId: "image:2", status: "degraded", issueCode: "IMAGE_INSERT_FAILED"}}
]);
assert.deepEqual(pictures, ["/private/one.png", "/private/two.png", "/private/two.png"]);
assert.equal(context.issues.filter(x => x.code === "IMAGE_INSERT_FAILED").length, 1);
"""
    _run_node(script)


def test_addin_table_falls_back_grid_then_text_without_partial_native_table() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
let text = "";
let tableAttempts = 0;
function range(start, end) {{ return {{Start: start, End: end, ParagraphFormat: {{}}, InsertAfter: function(value) {{ text += String(value); }}, Delete: function() {{ text = text.slice(0, start); }}}}; }}
const document = {{
  Content: {{get End() {{ return text.length + 1; }}}}, Range: range,
  Fields: {{Add: function() {{ text += "F"; return {{Result: {{Text: "1"}}, Update: function() {{}}}}; }}}},
  Bookmarks: {{Count: 0, Add: function() {{}}}},
  Tables: {{Add: function() {{ tableAttempts += 1; throw new Error("native table failure"); }}}}
}};
const context = {{ownerNodeId: "tab:1", issues: [], childResults: []}};
window.WPSComposerLongformV2.__test.addSemanticTableNative(document, {{
  caption: "表格", numbering: {{mode: "global", sequenceId: "WPSC_TAB", chapterStyleLevel: null, resetLevel: null, prefix: "表 ", suffix: ""}},
  bookmarkName: "wpsc_tab_bbbbbbbbbbbbbbbbbbbbbbbb", orientation: "portrait", keepCaptionWithFirstRow: true,
  headers: ["A", "B"], rows: [["1", "2"]], alignments: ["left", "right"], style: "three-line",
  borderSpec: {{top: 1.5, bottom: 1.5, headerBottom: 0.75, left: 0, right: 0, insideHorizontal: 0, insideVertical: 0}},
  merges: [], repeatHeader: true, allowRowSplit: false, cellIndentPt: 0, plannedDegradation: []
}}, {{}}, context);
assert.equal(tableAttempts, 2);
assert.ok(text.includes("A | B\\r1 | 2\\r"));
assert.deepEqual(context.issues.map(x => x.code), ["TABLE_INSERT_FAILED"]);
"""
    _run_node(script)


def test_addin_vertical_overflow_grid_failure_rolls_back_then_uses_text() -> None:
    asset = json.dumps(str((ROOT / "writer-longform-v2.js").resolve()))
    script = f"""
const assert = require("assert");
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({asset}, "utf8"));
let text = "";
let tableAttempts = 0;
function range(start, end) {{ return {{Start: start, End: end, ParagraphFormat: {{}}, InsertAfter: function(value) {{ text += String(value); }}, Delete: function() {{ text = text.slice(0, start); }}}}; }}
const rowOne = {{HeadingFormat: 0, Borders: function() {{ return {{}}; }}}};
const rows = function(index) {{ return index === 1 ? rowOne : {{Borders: function() {{ return {{}}; }}}}; }};
rows.AllowBreakAcrossPages = null;
function cell(row, column) {{
  return {{
    Range: {{Text: "", ParagraphFormat: {{}}, Information: function() {{ return row === 3 ? 2 : 1; }}}},
    Merge: function() {{}}
  }};
}}
const nativeTable = {{
  Rows: rows,
  Cell: cell,
  Borders: function() {{ return {{}}; }}
}};
const document = {{
  Content: {{get End() {{ return text.length + 1; }}}}, Range: range,
  Fields: {{Add: function() {{ text += "F"; return {{Result: {{Text: "1"}}, Update: function() {{}}}}; }}}},
  Bookmarks: {{Count: 0, Add: function() {{}}}},
  Tables: {{Add: function() {{
    tableAttempts += 1;
    if (tableAttempts === 1) {{ text += "T"; return nativeTable; }}
    text += "G";
    throw new Error("grid /private/failure");
  }}}}
}};
const context = {{ownerNodeId: "tab:1", issues: [], childResults: []}};
window.WPSComposerLongformV2.__test.addSemanticTableNative(document, {{
  caption: "表格", numbering: {{mode: "global", sequenceId: "WPSC_TAB", chapterStyleLevel: null, resetLevel: null, prefix: "表 ", suffix: ""}},
  bookmarkName: "wpsc_tab_bbbbbbbbbbbbbbbbbbbbbbbb", orientation: "portrait", keepCaptionWithFirstRow: true,
  headers: ["A", "B"], rows: [["1", "2"], ["3", "4"]], alignments: ["left", "right"], style: "three-line",
  borderSpec: {{top: 1.5, bottom: 1.5, headerBottom: 0.75, left: 0, right: 0, insideHorizontal: 0, insideVertical: 0}},
  merges: [{{top: 2, left: 1, bottom: 3, right: 1}}], repeatHeader: true, allowRowSplit: false, cellIndentPt: 0, plannedDegradation: []
}}, {{}}, context);
assert.equal(tableAttempts, 2);
assert.ok(text.includes("A | B\\r1 | 2\\r3 | 4\\r"));
assert.ok(!text.includes("G"));
assert.deepEqual(context.issues.map(function(issue) {{ return issue.code; }}), [
  "TABLE_ROW_FORCED_SPLIT", "TABLE_INSERT_FAILED"
]);
"""
    _run_node(script)


def test_runtime_profile_verifies_longform_asset_before_copy(tmp_path: Path) -> None:
    from skills.WPSComposer.scripts.macos_probe.runtime import build_profile
    from skills.WPSComposer.scripts.macos_probe.templates import AddinAssetError

    assets = tmp_path / "assets"
    assets.mkdir()
    for name in ("index.html", "manifest.xml", "ribbon.xml", "bridge-client.js", "writer-longform-m0.js", "writer-longform-v2.js", "writer.js"):
        (assets / name).write_text(name, encoding="utf-8")
    (assets / "asset-manifest.json").write_text(
        json.dumps({"version": 1, "assets": {"writer-longform-v2.js": "0" * 64}}),
        encoding="utf-8",
    )
    with pytest.raises(AddinAssetError, match="writer-longform-v2"):
        build_profile(
            assets,
            tmp_path / "profiles",
            "writer",
            "http://127.0.0.1:1",
            "client",
            "capability",
        )


def test_committed_addin_asset_manifest_matches_template_generator() -> None:
    from skills.WPSComposer.scripts.macos_probe.templates import verify_addin_assets

    verify_addin_assets(ROOT)
