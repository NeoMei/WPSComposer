from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.longform.degradation import (
    FATAL_BOUNDARY_CODES,
    RECOVERY_MATRIX,
    DegradationDescriptor,
    LocalRecoveryController,
    RecoveryFatalError,
    decide_recovery,
    render_js_recovery_matrix,
)
from skills.WPSComposer.scripts.longform.executor import ExecutionIssue
from skills.WPSComposer.scripts.generation_plan import GenerationOperation
from skills.WPSComposer.scripts.longform.windows_executor import (
    WindowsLongformExecutor,
)
from skills.WPSComposer.scripts.longform.macos_executor import _execution_issue
from skills.WPSComposer.scripts.macos_probe.models import (
    ProtocolError,
    validate_longform_generation_value,
)
from skills.WPSComposer.scripts.writer import WriterComposer
from skills.WPSComposer.scripts.update_longform_recovery_matrix import (
    update_recovery_matrix,
)


ROOT = Path(__file__).resolve().parents[2]
ADDIN = ROOT / "macos/wps-jsapi-probe/addin/writer-longform-v2.js"
UPDATER = ROOT / "skills/WPSComposer/scripts/update_longform_recovery_matrix.py"


EXPECTED_RECOVERY = {
    "BIBLIOGRAPHY_INSERT_FAILED": {"notice": "block"},
    "CROSS_REFERENCE_FAILED": {"inline-fallback": "inline"},
    "DEGRADATION_INSERT_FAILED": {"inline": "inline", "notice": "block"},
    "EQUATION_INSERT_FAILED": {"explicit-image-then-source-notice": "block"},
    "FIELD_REFRESH_UNSTABLE": {"document-quality-notice": "document"},
    "IMAGE_INSERT_FAILED": {"figure-child-stack-then-notice": "block"},
    "TABLE_INSERT_FAILED": {"grid-then-text": "block"},
    "TABLE_MERGE_APPLY_FAILED": {"grid-then-text": "block"},
    "TABLE_ROW_FORCED_SPLIT": {"grid-then-text": "block"},
    "TABLE_STYLE_APPLY_FAILED": {"grid-then-text": "block"},
}


def test_recovery_matrix_is_closed_and_every_branch_allows_one_fallback() -> None:
    assert {
        code: {rule.fallback_kind: rule.placement for rule in rules}
        for code, rules in RECOVERY_MATRIX.items()
    } == EXPECTED_RECOVERY
    for code, branches in EXPECTED_RECOVERY.items():
        for fallback_kind, placement in branches.items():
            decision = decide_recovery(code, fallback_kind, placement)
            assert decision.recoverable is True
            assert decision.fallback_attempts == 1
            assert decision.fallback_kind == fallback_kind
            assert decision.placement == placement


@pytest.mark.parametrize(
    "code",
    sorted(
        FATAL_BOUNDARY_CODES
        | {
            "UNKNOWN_EXCEPTION",
            "ENGINE_LOST",
            "SAVE_FAILED",
            "FIELD_REFRESH_FAILED",
            "INDEX_REFRESH_FAILED",
            "CONFIGURATION_INVALID",
            "PRIVATE_RESOURCE_CLEANUP_FAILED",
            "DEGRADATION_FALLBACK_FAILED",
            "LOCAL_MUTATION_ROLLBACK_FAILED",
        }
    ),
)
def test_unknown_and_boundary_failures_never_receive_recovery(code: str) -> None:
    with pytest.raises(RecoveryFatalError) as exc_info:
        decide_recovery(code, "notice", "block")
    assert exc_info.value.code == code
    assert exc_info.value.__cause__ is None


def test_wrong_declared_fallback_or_placement_is_fatal() -> None:
    for fallback_kind, placement in (("notice", "block"), ("inline-fallback", "block")):
        with pytest.raises(RecoveryFatalError, match="contract"):
            decide_recovery("CROSS_REFERENCE_FAILED", fallback_kind, placement)


class _NativeFailure(RuntimeError):
    def __init__(self, code: str) -> None:
        super().__init__("private /Users/me/secret.png")
        self.code = code


def _descriptor() -> DegradationDescriptor:
    return DegradationDescriptor(
        code="CROSS_REFERENCE_FAILED",
        placement="inline",
        object_label="引用",
        reason="原生引用更新失败",
        fallback_text="[引用目标未解析]",
        fallback_kind="inline-fallback",
    )


def test_local_state_machine_rolls_back_then_falls_back_and_notices_once() -> None:
    controller = LocalRecoveryController()
    events: list[object] = []

    def native_attempt() -> None:
        events.append("native")
        raise _NativeFailure("CROSS_REFERENCE_FAILED")

    decision = controller.execute(
        node_id="para:1",
        descriptor=_descriptor(),
        checkpoint=lambda: events.append("checkpoint") or 17,
        native_attempt=native_attempt,
        rollback=lambda token: events.append(("rollback", token)),
        fallback_attempt=lambda descriptor: events.append(("fallback", descriptor.fallback_text)),
        insert_notice=lambda node_id, descriptor: events.append(
            ("notice", node_id, descriptor.placement)
        ),
    )

    assert events == [
        "checkpoint",
        "native",
        ("rollback", 17),
        ("fallback", "[引用目标未解析]"),
        ("notice", "para:1", "inline"),
    ]
    assert decision.recoverable is True
    assert decision.fallback_attempts == 1
    assert controller.decisions == (decision,)

    again = controller.execute(
        node_id="para:1",
        descriptor=_descriptor(),
        checkpoint=lambda: events.append("unexpected checkpoint"),
        native_attempt=lambda: events.append("unexpected native"),
        rollback=lambda token: events.append("unexpected rollback"),
        fallback_attempt=lambda descriptor: events.append("unexpected fallback"),
        insert_notice=lambda node_id, descriptor: events.append("unexpected notice"),
    )
    assert again == decision
    assert len(events) == 5


@pytest.mark.parametrize("failing_stage", ("rollback", "fallback", "notice"))
def test_rollback_fallback_and_notice_failures_are_fatal(failing_stage: str) -> None:
    controller = LocalRecoveryController()

    def fail() -> None:
        raise RuntimeError("/private/secret deadbeef")

    callbacks = {
        "rollback": lambda token: fail() if failing_stage == "rollback" else None,
        "fallback": lambda descriptor: fail() if failing_stage == "fallback" else None,
        "notice": lambda node_id, descriptor: fail() if failing_stage == "notice" else None,
    }
    expected = {
        "rollback": "LOCAL_MUTATION_ROLLBACK_FAILED",
        "fallback": "DEGRADATION_FALLBACK_FAILED",
        "notice": "DEGRADATION_INSERT_FAILED",
    }[failing_stage]
    with pytest.raises(RecoveryFatalError) as exc_info:
        controller.execute(
            node_id="quote:1/para:1",
            descriptor=_descriptor(),
            checkpoint=lambda: 1,
            native_attempt=lambda: (_ for _ in ()).throw(
                _NativeFailure("CROSS_REFERENCE_FAILED")
            ),
            rollback=callbacks["rollback"],
            fallback_attempt=callbacks["fallback"],
            insert_notice=callbacks["notice"],
        )
    assert exc_info.value.code == expected
    assert "/private" not in str(exc_info.value)
    assert exc_info.value.__cause__ is None


def test_descriptor_and_execution_issue_serialization_redacts_private_evidence() -> None:
    secrets = (
        "/Users/me/private/input.png",
        r"C:\private\input.png",
        "wpsc-rsrc:opaque-private-id",
        "sha256:" + "a" * 64,
        "wpsc_fig_aaaaaaaaaaaaaaaaaaaaaaaa",
        "RuntimeError('private payload')",
        "fieldResult=SecretFieldValue",
        "bookmarkMap=WPSC_PRIVATE_MAP",
    )
    raw = " | ".join(secrets)
    descriptor = DegradationDescriptor(
        code="IMAGE_INSERT_FAILED",
        placement="block",
        object_label="图像",
        reason=raw,
        fallback_text=raw,
        fallback_kind="figure-child-stack-then-notice",
    )
    issue = ExecutionIssue(
        code="IMAGE_INSERT_FAILED",
        message=raw,
        placement="block",
        node_id="fig:1",
        stage="native",
        fallback="figure-child-stack-then-notice",
        recoverable=True,
    )
    rendered = json.dumps(
        {"descriptor": descriptor.to_dict(), "issue": issue.to_dict()},
        ensure_ascii=False,
        sort_keys=True,
    )
    assert all(secret not in rendered for secret in secrets)
    assert "<redacted>" in rendered
    assert issue.to_dict() == ExecutionIssue.from_dict(issue.to_dict()).to_dict()
    private_repr = repr(descriptor) + repr(issue)
    assert all(secret not in private_repr for secret in secrets)
    assert "<redacted>" in private_repr


def test_execution_issue_old_payload_round_trips_without_new_fields() -> None:
    old = {
        "code": "IMAGE_INSERT_FAILED",
        "message": "Native image insertion failed",
        "placement": "block",
        "nodeId": "fig:1",
    }
    assert ExecutionIssue.from_dict(old).to_dict() == old


def _marked_block(text: str) -> str:
    begin = "  // BEGIN WPSCOMPOSER GENERATED RECOVERY MATRIX\n"
    end = "  // END WPSCOMPOSER GENERATED RECOVERY MATRIX\n"
    start = text.index(begin)
    stop = text.index(end, start) + len(end)
    return text[start:stop]


def test_generated_js_matrix_is_byte_exact_and_regeneration_is_idempotent(tmp_path: Path) -> None:
    tracked = ADDIN.read_text(encoding="utf-8")
    assert _marked_block(tracked) == render_js_recovery_matrix()

    copy = tmp_path / ADDIN.name
    copy.write_text(tracked, encoding="utf-8")
    before = copy.read_bytes()
    assert update_recovery_matrix(copy, check=False) is False
    assert copy.read_bytes() == before
    assert update_recovery_matrix(copy, check=True) is False

    completed = subprocess.run(
        [sys.executable, str(UPDATER), "--check"],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr


def test_javascript_decisions_match_python_matrix_and_fatal_boundaries() -> None:
    script = f"""
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({json.dumps(str(ADDIN))}, "utf8"));
const api = window.WPSComposerLongformV2.__test;
const cases = {json.dumps([
        [code, fallback, placement]
        for code, branches in EXPECTED_RECOVERY.items()
        for fallback, placement in branches.items()
    ])};
const accepted = cases.map(function(item) {{
  return api.recoveryDecision(item[0], item[1], item[2]);
}});
let fatal = null;
try {{ api.recoveryDecision("SAVE_FAILED", "notice", "block"); }}
catch (error) {{ fatal = error.code; }}
process.stdout.write(JSON.stringify({{accepted: accepted, fatal: fatal}}));
"""
    completed = subprocess.run(
        ["node", "-e", script],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=True,
    )
    actual = json.loads(completed.stdout)
    expected = [
        decide_recovery(code, fallback, placement).to_dict()
        for code, branches in EXPECTED_RECOVERY.items()
        for fallback, placement in branches.items()
    ]
    assert actual == {"accepted": expected, "fatal": "SAVE_FAILED"}


def test_javascript_local_state_machine_matches_python_order_and_deduplication() -> None:
    script = f"""
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({json.dumps(str(ADDIN))}, "utf8"));
const api = window.WPSComposerLongformV2.__test;
const controller = api.createLocalRecoveryController();
const events = [];
const spec = {{
  nodeId: "para:1",
  descriptor: {{
    code: "CROSS_REFERENCE_FAILED", placement: "inline",
    fallbackKind: "inline-fallback", fallbackText: "[missing]"
  }},
  checkpoint: function() {{ events.push("checkpoint"); return 17; }},
  nativeAttempt: function() {{ events.push("native"); const e = new Error("private"); e.code = "CROSS_REFERENCE_FAILED"; throw e; }},
  rollback: function(token) {{ events.push("rollback:" + token); }},
  fallbackAttempt: function() {{ events.push("fallback"); }},
  insertNotice: function(nodeId, descriptor) {{ events.push("notice:" + nodeId + ":" + descriptor.placement); }}
}};
const first = api.runLocalRecovery(controller, spec);
const second = api.runLocalRecovery(controller, spec);
let fallbackFatal = null;
try {{
  api.runLocalRecovery(api.createLocalRecoveryController(), Object.assign({{}}, spec, {{
    nodeId: "para:2", fallbackAttempt: function() {{ throw new Error("/private/fallback"); }}
  }}));
}} catch (error) {{ fallbackFatal = error.code; }}
process.stdout.write(JSON.stringify({{
  events: events, first: first, second: second,
  decisions: controller.decisions, fallbackFatal: fallbackFatal
}}));
"""
    completed = subprocess.run(
        ["node", "-e", script],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=True,
    )
    actual = json.loads(completed.stdout)
    expected = decide_recovery(
        "CROSS_REFERENCE_FAILED", "inline-fallback", "inline"
    ).to_dict()
    assert actual == {
        "events": [
            "checkpoint", "native", "rollback:17", "fallback", "notice:para:1:inline",
            "checkpoint", "native", "rollback:17",
        ],
        "first": expected,
        "second": expected,
        "decisions": [expected],
        "fallbackFatal": "DEGRADATION_FALLBACK_FAILED",
    }


class _Range:
    def __init__(self, start: int = 0, end: int = 0) -> None:
        self.Start = start
        self.End = end
        self.Text = ""
        self.Font = SimpleNamespace(Italic=False, Color=0)
        self.Shading = SimpleNamespace(BackgroundPatternColor=0)
        self.ParagraphFormat = SimpleNamespace(
            SpaceBefore=0, SpaceAfter=0, KeepTogether=False, OutlineLevel=0
        )


class _Selection:
    def __init__(self) -> None:
        self.End = 4
        self.Font = SimpleNamespace(Italic=False, Color=0)
        self.ParagraphFormat = SimpleNamespace()
        self.typed: list[str] = []
        self.paragraphs = 0

    def TypeText(self, value: str) -> None:
        self.typed.append(value)
        self.End += len(value)

    def TypeParagraph(self) -> None:
        self.paragraphs += 1
        self.End += 1


class _Tables:
    def __init__(self) -> None:
        self.calls: list[tuple[_Range, int, int]] = []
        self.tables: list[SimpleNamespace] = []

    def Add(self, target: _Range, rows: int, columns: int) -> SimpleNamespace:
        self.calls.append((target, rows, columns))
        cell_range = _Range(target.Start, target.End + 1)
        cell = SimpleNamespace(Range=cell_range, Shading=cell_range.Shading)
        table = SimpleNamespace(
            Range=cell_range,
            Cell=lambda row, column: cell,
            Borders=SimpleNamespace(Enable=True),
            Rows=SimpleNamespace(AllowBreakAcrossPages=True),
        )
        self.tables.append(table)
        return table


class _Document:
    def __init__(self) -> None:
        self.Tables = _Tables()
        self.ranges: list[_Range] = []
        self.Bookmarks = SimpleNamespace(Add=lambda name, target: None)

    def Range(self, start: int, end: int) -> _Range:
        result = _Range(start, end)
        self.ranges.append(result)
        return result


def _writer_double() -> tuple[WriterComposer, _Selection, _Document]:
    selection = _Selection()
    document = _Document()
    writer = WriterComposer.__new__(WriterComposer)
    writer._app = SimpleNamespace(Selection=selection)
    writer._doc = document
    writer._reset_selection_to_normal = lambda: None
    return writer, selection, document


def test_writer_inline_notice_stays_in_current_paragraph_and_styles_inserted_range() -> None:
    writer, selection, document = _writer_double()
    writer.add_inline_degradation(
        "CROSS_REFERENCE_FAILED", "private diagnostics", "[引用目标未解析]"
    )

    assert selection.paragraphs == 0
    assert selection.typed == ["[CROSS_REFERENCE_FAILED: [引用目标未解析]]"]
    styled = document.ranges[-1]
    assert (styled.Start, styled.End) == (4, selection.End)
    assert styled.Font.Italic is True
    assert styled.Font.Color != 0
    assert styled.Shading.BackgroundPatternColor != 0


def test_writer_block_notice_is_a_restrained_one_cell_box() -> None:
    writer, selection, document = _writer_double()
    writer.add_degradation_notice(
        "TABLE_INSERT_FAILED", "private diagnostics", "Name | Value", "block"
    )

    assert selection.paragraphs == 0
    assert len(document.Tables.calls) == 1
    _, rows, columns = document.Tables.calls[0]
    assert (rows, columns) == (1, 1)
    cell_range = document.Tables.tables[0].Cell(1, 1).Range
    assert cell_range.Text == "[TABLE_INSERT_FAILED] Name | Value"
    assert cell_range.Font.Italic is True
    assert cell_range.Shading.BackgroundPatternColor != 0


def test_writer_quality_anchor_is_reserved_when_empty_and_upserts_once() -> None:
    writer, selection, document = _writer_double()
    writer.add_paragraph = lambda text, **kwargs: (
        selection.TypeText(text), selection.TypeParagraph()
    )

    writer.reserve_document_quality_anchor("生成质量提示", [])
    assert writer._quality_notice_anchor_position == selection.End
    assert writer._quality_notice_seen == set()
    assert selection.typed == []
    assert selection.paragraphs == 0

    issue = ExecutionIssue(
        code="FIELD_REFRESH_UNSTABLE",
        message="Field refresh did not converge",
        placement="document",
        stage="field-refresh",
        fallback="document-quality-notice",
        recoverable=True,
    )
    writer.upsert_document_quality_notice(issue)
    writer.upsert_document_quality_notice(issue)

    assert len(document.Tables.calls) == 1
    assert document.Tables.tables[0].Cell(1, 1).Range.Text.startswith(
        "生成质量提示\r[FIELD_REFRESH_UNSTABLE]"
    )


def test_writer_quality_anchor_api_is_required() -> None:
    writer, _, document = _writer_double()
    document.Bookmarks = None
    with pytest.raises(Exception) as exc_info:
        writer.reserve_document_quality_anchor()
    assert getattr(exc_info.value, "code", None) == "DEGRADATION_INSERT_FAILED"


def test_javascript_styled_primitives_keep_inline_local_and_quality_notice_deduped() -> None:
    script = f"""
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({json.dumps(str(ADDIN))}, "utf8"));
const api = window.WPSComposerLongformV2.__test;
const written = [];
const tables = [];
function makeRange(start, end) {{
  return {{
    Start: start, End: end, Text: "", Font: {{}}, Shading: {{}}, ParagraphFormat: {{}},
    InsertAfter: function(value) {{ this.Text += value; written.push(value); }}
  }};
}}
const document = {{
  Content: {{End: 5}},
  Range: function(start, end) {{ return makeRange(start, end); }},
  Tables: {{Add: function(target, rows, columns) {{
    const cellRange = makeRange(target.Start, target.End + 1);
    const table = {{
      Range: cellRange, Rows: {{}},
      Cell: function() {{ return {{Range: cellRange, Shading: cellRange.Shading}}; }}
    }};
    tables.push({{rows: rows, columns: columns, table: table}});
    return table;
  }}}},
  Bookmarks: {{Add: function() {{}}}}
}};
api.addInlineDegradation(document, {{
  code: "CROSS_REFERENCE_FAILED", fallbackText: "[missing]"
}});
const inline = written.slice();
api.addDegradationNotice(document, {{
  code: "TABLE_INSERT_FAILED", fallbackText: "A | B", placement: "block"
}});
api.reserveDocumentQualityAnchor(document, {{title: "生成质量提示", notices: []}});
const afterReserve = document._wpscQualityAnchor.empty;
const issue = {{
  code: "FIELD_REFRESH_UNSTABLE", message: "not stable", placement: "document"
}};
api.upsertDocumentQualityNotice(document, issue);
api.upsertDocumentQualityNotice(document, issue);
process.stdout.write(JSON.stringify({{
  inline: inline,
  tables: tables.map(function(item) {{ return {{
    rows: item.rows, columns: item.columns, text: item.table.Cell().Range.Text,
    italic: item.table.Cell().Range.Font.Italic,
    shade: item.table.Cell().Range.Shading.BackgroundPatternColor
  }}; }}),
  anchor: document._wpscQualityAnchor,
  afterReserve: afterReserve,
  notices: Object.keys(document._wpscQualityNoticeSeen || {{}})
}}));
"""
    completed = subprocess.run(
        ["node", "-e", script], cwd=ROOT, text=True, capture_output=True, check=False
    )
    assert completed.returncode == 0, completed.stderr
    actual = json.loads(completed.stdout)
    assert actual["inline"] == ["[CROSS_REFERENCE_FAILED: [missing]]"]
    assert len(actual["tables"]) == 2
    assert actual["tables"][0]["rows"] == actual["tables"][0]["columns"] == 1
    assert actual["tables"][0]["text"] == "[TABLE_INSERT_FAILED] A | B"
    assert actual["tables"][0]["italic"] == -1
    assert actual["tables"][0]["shade"] != 0
    assert actual["anchor"]["title"] == "生成质量提示"
    assert actual["afterReserve"] is True
    assert len(actual["notices"]) == 1


def test_javascript_quality_anchor_api_is_required() -> None:
    script = f"""
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({json.dumps(str(ADDIN))}, "utf8"));
let code = null;
try {{
  window.WPSComposerLongformV2.__test.reserveDocumentQualityAnchor({{
    Content: {{End: 1}}, Range: function(start, end) {{ return {{Start: start, End: end}}; }}
  }}, {{title: "生成质量提示", notices: []}});
}} catch (error) {{ code = error.code; }}
process.stdout.write(JSON.stringify(code));
"""
    completed = subprocess.run(
        ["node", "-e", script], cwd=ROOT, text=True, capture_output=True, check=True
    )
    assert json.loads(completed.stdout) == "DEGRADATION_INSERT_FAILED"


class _RecoveryComposer:
    def __init__(self) -> None:
        self.events: list[object] = []

    def degradation_checkpoint(self):
        self.events.append("checkpoint")
        return 17

    def rollback_degradation_checkpoint(self, checkpoint):
        self.events.append(("rollback", checkpoint))

    def add_cross_reference_paragraph(self, **kwargs):
        self.events.append("native")
        raise _NativeFailure("CROSS_REFERENCE_FAILED")

    def add_inline_degradation(self, **kwargs):
        self.events.append(("inline", kwargs["fallback_text"]))

    def add_degradation_notice(self, **kwargs):
        self.events.append(("block", kwargs["code"]))


def test_windows_executor_routes_declared_recovery_through_closed_contract() -> None:
    executor = WindowsLongformExecutor()
    composer = _RecoveryComposer()
    operation = GenerationOperation(
        op="writer.add_cross_reference",
        node_id="para:1",
        args={"runs": [{"type": "text", "text": "See [1]."}]},
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
    )

    executor._run_op(composer, operation)

    assert composer.events == [
        "checkpoint",
        "native",
        ("rollback", 17),
        ("inline", "See [1]."),
    ]
    assert [issue.to_dict() for issue in executor._issues] == [{
        "code": "CROSS_REFERENCE_FAILED",
        "message": "writer.add_cross_reference used its declared native fallback",
        "placement": "inline",
        "nodeId": "para:1",
        "stage": "native",
        "fallback": "inline-fallback",
        "recoverable": True,
    }]


@pytest.mark.parametrize(
    "composer",
    [
        SimpleNamespace(
            rollback_degradation_checkpoint=lambda checkpoint: None,
            add_cross_reference_paragraph=lambda **kwargs: (_ for _ in ()).throw(
                AssertionError("native attempt must not start")
            ),
        ),
        SimpleNamespace(
            degradation_checkpoint=lambda: 17,
            add_cross_reference_paragraph=lambda **kwargs: (_ for _ in ()).throw(
                AssertionError("native attempt must not start")
            ),
        ),
        SimpleNamespace(
            degradation_checkpoint=lambda: (_ for _ in ()).throw(
                RuntimeError("/private/checkpoint")
            ),
            rollback_degradation_checkpoint=lambda checkpoint: None,
            add_cross_reference_paragraph=lambda **kwargs: (_ for _ in ()).throw(
                AssertionError("native attempt must not start")
            ),
        ),
    ],
)
def test_windows_recoverable_operation_requires_checkpoint_and_rollback_before_native(
    composer,
) -> None:
    executor = WindowsLongformExecutor()
    operation = GenerationOperation(
        op="writer.add_cross_reference",
        node_id="para:1",
        args={"runs": [{"type": "text", "text": "See [1]."}]},
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
    )

    with pytest.raises(Exception) as exc_info:
        executor._run_op(composer, operation)

    assert getattr(exc_info.value, "op_name", None) == "local-checkpoint"
    assert executor._issues == []


def test_windows_failed_rollback_is_fatal_and_never_attempts_fallback() -> None:
    events: list[object] = []
    composer = SimpleNamespace(
        degradation_checkpoint=lambda: events.append("checkpoint") or 17,
        rollback_degradation_checkpoint=lambda checkpoint: (_ for _ in ()).throw(
            RuntimeError("/private/rollback")
        ),
        add_cross_reference_paragraph=lambda **kwargs: (_ for _ in ()).throw(
            _NativeFailure("CROSS_REFERENCE_FAILED")
        ),
        add_inline_degradation=lambda **kwargs: events.append("fallback"),
    )
    executor = WindowsLongformExecutor()
    operation = GenerationOperation(
        op="writer.add_cross_reference",
        node_id="para:1",
        args={"runs": [{"type": "text", "text": "See [1]."}]},
        failure_policy={
            "mode": "degrade",
            "recoverableCodes": ["CROSS_REFERENCE_FAILED"],
            "fallback": "inline-fallback",
        },
    )

    with pytest.raises(Exception) as exc_info:
        executor._run_op(composer, operation)

    assert getattr(exc_info.value, "op_name", None) == operation.op
    assert events == ["checkpoint"]
    assert executor._issues == []


def test_windows_executor_unknown_operation_is_fatal_without_notice() -> None:
    executor = WindowsLongformExecutor()
    composer = _RecoveryComposer()
    operation = GenerationOperation(op="writer.not_declared", args={}, node_id="x:1")

    with pytest.raises(Exception) as exc_info:
        executor._run_op(composer, operation)

    assert getattr(exc_info.value, "op_name", None) == "writer.not_declared"
    assert composer.events == []
    assert executor._issues == []


def test_windows_table_dispatch_accepts_plan_only_cell_citations_metadata() -> None:
    executor = WindowsLongformExecutor()
    captured: list[dict] = []
    composer = SimpleNamespace(
        add_semantic_table_native=lambda **kwargs: captured.append(kwargs) or {}
    )
    operation = GenerationOperation(
        op="writer.add_semantic_table",
        node_id="tab:1",
        args={
            "numbering": {},
            "cellDegradations": [{"row": 2, "column": 1, "code": "REFERENCE_UNRESOLVED", "fallbackText": "[missing]"}],
            "cellCitations": [{"row": 2, "column": 1, "targetId": "a", "targetNodeId": "ref:a", "number": 1, "fallbackText": "[1]"}],
        },
    )

    executor._dispatch_one(composer, operation)

    assert captured[0]["cellDegradations"][0]["code"] == "REFERENCE_UNRESOLVED"
    assert "cellCitations" not in captured[0]


def test_javascript_operation_recovery_uses_closed_controller_and_unknown_is_fatal() -> None:
    script = f"""
const fs = require("fs");
global.window = {{}};
eval(fs.readFileSync({json.dumps(str(ADDIN))}, "utf8"));
const api = window.WPSComposerLongformV2.__test;
function range(start, end) {{ return {{Start: start, End: end, Font: {{}}, Shading: {{}}, ParagraphFormat: {{}}}}; }}
const tables = [];
const document = {{
  Content: {{End: 3}},
  Range: range,
  Tables: {{Add: function(target, rows, columns) {{
    const cellRange = range(target.Start, target.End + 1); cellRange.Text = "";
    const table = {{Range: cellRange, Rows: {{}}, Cell: function() {{ return {{Range: cellRange}}; }}}};
    tables.push(table); return table;
  }}}}
}};
const issues = [];
api.runOperation(document, {{
  op: "writer.add_bibliography", nodeId: "bib:1", args: {{fallbackText: "Alpha."}},
  failurePolicy: {{mode: "degrade", recoverableCodes: ["BIBLIOGRAPHY_INSERT_FAILED"], fallback: "notice"}}
}}, {{}}, issues, []);
let fatal = null;
try {{ api.runOperation(document, {{op: "writer.not_declared", nodeId: "x:1", args: {{}}}}, {{}}, issues, []); }}
catch (error) {{ fatal = error.code; }}
process.stdout.write(JSON.stringify({{issues: issues, fatal: fatal, text: tables[0].Cell().Range.Text}}));
"""
    completed = subprocess.run(
        ["node", "-e", script], cwd=ROOT, text=True, capture_output=True, check=True
    )
    actual = json.loads(completed.stdout)
    assert actual == {
        "issues": [{
            "code": "BIBLIOGRAPHY_INSERT_FAILED",
            "message": "writer.add_bibliography used its declared fallback",
            "placement": "block",
            "nodeId": "bib:1",
            "stage": "native",
            "fallback": "notice",
            "recoverable": True,
        }],
        "fatal": "UNKNOWN_OPERATION",
        "text": "[BIBLIOGRAPHY_INSERT_FAILED] Alpha.",
    }


def _mac_result(issue: dict[str, object]) -> dict[str, object]:
    return {
        "outputPath": "/private/output.docx",
        "appliedOperations": 1,
        "issueCodes": [issue],
        "fieldSnapshots": [],
        "childResults": [],
        "paginationMap": {"version": "M2-stub", "nodes": []},
    }


def test_macos_result_schema_accepts_new_issue_fields_and_old_payload() -> None:
    new = {
        "code": "IMAGE_INSERT_FAILED",
        "message": "Native image used fallback",
        "placement": "block",
        "nodeId": "fig:1",
        "stage": "native",
        "fallback": "figure-child-stack-then-notice",
        "recoverable": True,
    }
    assert validate_longform_generation_value(_mac_result(new))["issueCodes"] == [new]
    parsed = _execution_issue(new).to_dict()
    assert parsed["stage"] == "native"
    assert parsed["fallback"] == "figure-child-stack-then-notice"
    assert parsed["recoverable"] is True
    old = {key: new[key] for key in ("code", "message", "placement", "nodeId")}
    assert validate_longform_generation_value(_mac_result(old))["issueCodes"] == [old]


@pytest.mark.parametrize(
    "change",
    [
        {"stage": "/private/native"},
        {"fallback": "sha256:" + "a" * 64},
        {"recoverable": "true"},
        {"message": "/Users/me/private/input.png"},
        {"unknown": True},
    ],
)
def test_macos_result_schema_rejects_invalid_or_private_issue_fields(
    change: dict[str, object],
) -> None:
    issue: dict[str, object] = {
        "code": "IMAGE_INSERT_FAILED",
        "message": "Native image used fallback",
        "placement": "block",
        "nodeId": "fig:1",
        "stage": "native",
        "fallback": "figure-child-stack-then-notice",
        "recoverable": True,
    }
    issue.update(change)
    with pytest.raises(ProtocolError, match="result is invalid"):
        validate_longform_generation_value(_mac_result(issue))
