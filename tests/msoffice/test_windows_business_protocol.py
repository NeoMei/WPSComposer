from __future__ import annotations

import datetime as dt
import io
import json
from pathlib import Path
from types import SimpleNamespace

import pytest

from skills.WPSComposer.scripts.design_presets import DesignPreset
from skills.WPSComposer.scripts.document_model import (
    CitationRun,
    CrossReferenceRun,
    InlineDegradationRun,
    Span,
)
from skills.WPSComposer.scripts.layout_templates import LayoutTemplate
from skills.WPSComposer.scripts.longform.executor import ExecutionIssue, FieldSnapshot
from skills.WPSComposer.scripts.longform.quality import (
    QualityConfidence,
    QualityEvidence,
    QualityFinding,
    QualityReport,
    QualitySeverity,
)
from skills.WPSComposer.scripts.msoffice import windows_session_proxy as proxy
from skills.WPSComposer.scripts.msoffice import windows_session_worker as worker
from skills.WPSComposer.scripts.msoffice.windows_business_protocol import (
    DTO_TAG,
    decode_value,
)


WORD_ADDITIONS = {
    "add_bibliography_legacy", "add_bibliography_native",
    "add_captioned_figure_fallback", "add_captioned_figure_native",
    "add_citation_paragraph", "add_cross_reference_fallback",
    "add_cross_reference_paragraph", "add_degradation_notice",
    "add_document_quality_notice", "add_equation_native",
    "add_equation_native_fallback", "add_equation_number_native",
    "add_heading_level_native", "add_inline_degradation",
    "add_landscape_section_before_pending_heading", "add_merged_table",
    "add_quality_notice_at_bookmark", "add_semantic_table_fallback",
    "add_semantic_table_native", "degradation_checkpoint", "finalize_fields",
    "insert_caption_index_native", "insert_figure_index", "insert_table_index",
    "insert_toc", "insert_toc_with_styles", "pagination_fragment_for_bookmark",
    "pagination_map_for_ranges", "reserve_document_quality_anchor", "reset",
    "rollback_degradation_checkpoint", "upsert_document_quality_notice",
}


def _request(number, method, *, kind="writer", args=(), kwargs=None):
    return {
        "protocol": 1,
        "id": number,
        "kind": kind,
        "method": method,
        "args": list(args),
        "kwargs": kwargs or {},
        "deadline": 1_000_000_000.0,
        "remaining_seconds": 500,
    }


def _serve(tmp_path, frames, factory):
    incoming = io.BytesIO(
        b"".join(json.dumps(frame).encode("utf-8") + b"\n" for frame in frames)
    )
    outgoing = io.BytesIO()
    worker.serve(incoming, outgoing, tmp_path, factory=factory)
    return [json.loads(line) for line in outgoing.getvalue().splitlines()]


def test_frozen_windows_business_surface_is_fully_allowlisted():
    assert WORD_ADDITIONS <= worker.BUSINESS_METHODS["writer"]
    assert {"apply_design_preset", "apply_layout_template"} <= worker.BUSINESS_METHODS["slide"]
    assert not worker.UNSUPPORTED_BUSINESS.intersection(
        set().union(*worker.BUSINESS_METHODS.values())
    )


def test_closed_input_dtos_preserve_exact_public_types_through_real_serve(tmp_path):
    cross_reference = CrossReferenceRun(
        "run:1", "fig:x", "fig:node", "figure", "fig_x", "Figure X"
    )
    span = Span(
        "中文😀", bold=True, cross_reference=cross_reference,
        citation=CitationRun("cite:1", "ref:x", "ref:node", 2, "[2]"),
        inline_degradation=InlineDegradationRun("deg:1", "REF_MISSING", "[?]"),
    )
    issue = ExecutionIssue("REF_MISSING", "reference unavailable", node_id="fig:node")
    finding = QualityFinding(
        "LOW_DPI", QualitySeverity.WARNING, QualityConfidence.HIGH,
        "image resolution is low", node_id="fig:node", page=2,
        evidence=QualityEvidence.from_mapping({"minimumDpi": 90}),
    )
    received = {}

    class Session:
        staging_root = tmp_path

        def add_rich_paragraph(self, spans, style):
            received["paragraph"] = (spans, style)

        def upsert_document_quality_notice(self, value):
            received["issue"] = value

        def add_document_quality_notice(self, notices):
            received["quality"] = notices

        def close(self, save_changes=False):
            pass

    owner = object()
    frames = [
        _request(1, "open_document", args=("source.docx",)),
        _request(
            2,
            "add_rich_paragraph",
            args=[proxy._arguments(item, owner) for item in ([span], "Body Text")],
        ),
        _request(
            3,
            "upsert_document_quality_notice",
            args=[proxy._arguments(issue, owner)],
        ),
        _request(
            4,
            "add_document_quality_notice",
            args=[proxy._arguments([finding], owner)],
        ),
        _request(5, "close"),
    ]
    responses = _serve(tmp_path, frames, lambda *args: Session())

    actual_spans, style = received["paragraph"]
    assert actual_spans == [span] and style == "Body Text"
    assert type(actual_spans[0]) is Span
    assert type(actual_spans[0].cross_reference) is CrossReferenceRun
    assert type(received["issue"]) is ExecutionIssue
    assert received["issue"] == issue
    assert type(received["quality"][0]) is QualityFinding
    assert received["quality"] == [finding]
    assert all(response["status"] == "ok" for response in responses)


def test_refresh_fields_and_snapshot_fields_return_exact_field_snapshots(tmp_path):
    snapshot = FieldSnapshot(
        ("toc", "TOC", 1), "native", "b" * 64, 3, 0, 0, 7
    )
    calls = []

    class Session:
        staging_root = tmp_path

        def refresh_fields(self, round_index):
            calls.append(round_index)
            return (snapshot,)

        def snapshot_fields(self):
            return (snapshot,)

        def close(self, save_changes=False):
            pass

    frames = [
        _request(1, "new_document"),
        _request(2, "refresh_fields", args=(2,)),
        _request(3, "snapshot_fields"),
        _request(4, "close"),
    ]
    owner = object()
    responses = _serve(tmp_path, frames, lambda *args: Session())
    assert proxy._result(responses[1]["value"], owner) == (snapshot,)
    assert proxy._result(responses[2]["value"], owner) == (snapshot,)
    assert type(proxy._result(responses[1]["value"], owner)[0]) is FieldSnapshot
    assert calls == [2]


def test_temporal_values_roundtrip_as_date_and_datetime_including_subclass(tmp_path):
    class ComDateTime(dt.datetime):
        pass

    class Session:
        staging_root = tmp_path

        def inspect_selection(self):
            return {
                "date": dt.date(2026, 9, 9),
                "datetime": ComDateTime(2026, 9, 9, 8, 7, 6, 500000),
            }

        def close(self, save_changes=False):
            pass

    responses = _serve(
        tmp_path,
        [
            _request(1, "open_document", kind="sheet", args=("source.xlsx",)),
            _request(2, "inspect_selection", kind="sheet"),
            _request(3, "close", kind="sheet"),
        ],
        lambda *args: Session(),
    )
    result = proxy._result(responses[1]["value"], object())
    assert result == {
        "date": dt.date(2026, 9, 9),
        "datetime": dt.datetime(2026, 9, 9, 8, 7, 6, 500000),
    }
    assert type(result["date"]) is dt.date
    assert type(result["datetime"]) is dt.datetime


def test_powerpoint_dtos_reconstruct_before_bound_native_methods(tmp_path):
    preset = DesignPreset(
        "Exact", {"bg": "#123456"}, {"title": ("Arial", 31, "#FFFFFF")},
        {"margin": 12.5}, {"max_colors": 2},
    )
    layout = LayoutTemplate(
        "Bound", "content", "one slide",
        [{"type": "line", "x": 1, "y": 2, "w": 3, "h": 4, "color": "bg"}],
        {"minimum_items": 1},
    )
    received = []

    class Session:
        staging_root = tmp_path

        def apply_design_preset(self, value):
            received.append(value)

        def apply_layout_template(self, value, value_preset=None):
            received.append((value, value_preset))

        def close(self, save_changes=False):
            pass

    owner = object()
    frames = [
        _request(1, "new_document", kind="slide"),
        _request(2, "apply_design_preset", kind="slide", args=[proxy._arguments(preset, owner)]),
        _request(3, "apply_layout_template", kind="slide", args=[
            proxy._arguments(layout, owner), proxy._arguments(preset, owner),
        ]),
        _request(4, "close", kind="slide"),
    ]
    responses = _serve(tmp_path, frames, lambda *args: Session())
    assert all(item["status"] == "ok" for item in responses)
    assert type(received[0]) is DesignPreset and received[0].to_dict() == preset.to_dict()
    actual_layout, actual_preset = received[1]
    assert type(actual_layout) is LayoutTemplate
    assert actual_layout.to_dict() == layout.to_dict()
    assert actual_layout.structural_rules == {"minimum_items": 1}
    assert type(actual_preset) is DesignPreset and actual_preset.to_dict() == preset.to_dict()


class Collection:
    def __init__(self, values=()):
        self.values = list(values)

    @property
    def Count(self):
        return len(self.values)

    def Item(self, index):
        return self.values[index - 1]


def _writer_registry():
    doc = SimpleNamespace(token="doc")
    table = SimpleNamespace(token="table")
    toc = SimpleNamespace(token="toc")
    tof = SimpleNamespace(token="tof")
    rng = SimpleNamespace(token="range", Document=doc, Start=2, End=8)
    doc.Content = SimpleNamespace(End=20)
    doc.Tables = Collection([table])
    doc.Shapes = Collection()
    doc.InlineShapes = Collection()
    doc.TablesOfContents = Collection([toc])
    doc.TablesOfFigures = Collection([tof])
    doc.Paragraphs = Collection()
    session = SimpleNamespace(
        kind="writer", _verify=lambda: None,
        _composer=SimpleNamespace(
            _doc=doc, _deps=SimpleNamespace(identity=lambda value: value.token)
        ),
    )
    return worker.NativeHandleRegistry(session), doc, table, toc, tof, rng


def test_word_business_native_results_use_live_typed_handles():
    registry, _, table, toc, tof, rng = _writer_registry()
    cases = [
        ("add_merged_table", table, "table"),
        ("add_degradation_notice", table, "table"),
        ("add_inline_degradation", rng, "range"),
        ("add_quality_notice_at_bookmark", table, "table"),
        ("insert_toc", toc, "toc"),
        ("insert_caption_index_native", tof, "table_of_figures"),
    ]
    for method, native, expected in cases:
        encoded = registry.encode_result(method, native, [], {})
        assert encoded["__wpscomposer_handle__"]["type"] == expected
    inline = registry.encode_result(
        "add_degradation_notice", rng, ["CODE", "message", "fallback", "inline"], {}
    )
    assert inline["__wpscomposer_handle__"]["type"] == "range"


def test_degradation_box_fallback_registers_its_live_range():
    registry, _, _, _, _, rng = _writer_registry()
    encoded = registry.encode_result(
        "add_quality_notice_at_bookmark", SimpleNamespace(Range=rng), [], {}
    )
    assert encoded["__wpscomposer_handle__"]["type"] == "range"


def test_pagination_accepts_only_live_session_range_handles():
    registry, doc, _, _, _, rng = _writer_registry()
    reference = registry.register(rng, "range")
    args, kwargs = registry.prepare_call(
        "pagination_map_for_ranges",
        [({"nodeId": "n1", "op": "writer.add_heading", "role": "body", "range": reference},)],
        {},
    )
    assert args[0][0]["range"] is rng and kwargs == {}
    rng.End = 21
    with pytest.raises(ValueError, match="stale"):
        registry.prepare_call(
            "pagination_map_for_ranges",
            [[{"nodeId": "n1", "range": reference}]],
            {},
        )
    unknown = {"__wpscomposer_handle__": {"id": "f" * 32, "type": "range"}}
    with pytest.raises(ValueError, match="another session"):
        registry.prepare_call(
            "pagination_map_for_ranges", [[{"nodeId": "n", "range": unknown}]], {}
        )


def test_pagination_preserves_the_public_keyword_argument_contract():
    registry, _, _, _, _, rng = _writer_registry()
    reference = registry.register(rng, "range")
    args, kwargs = registry.prepare_call(
        "pagination_map_for_ranges",
        [],
        {"tracked_ranges": ({"nodeId": "n1", "range": reference},)},
    )
    assert args == []
    assert kwargs["tracked_ranges"][0]["range"] is rng


def test_new_handle_results_reserve_capacity_before_native_mutation():
    registry, _, table, _, _, _ = _writer_registry()
    registry.records.update({str(index): (table, "table") for index in range(1024)})
    for method in (
        "add_merged_table",
        "add_degradation_notice",
        "add_inline_degradation",
        "add_quality_notice_at_bookmark",
        "insert_toc",
        "insert_caption_index_native",
    ):
        with pytest.raises(ValueError, match="limit exceeded before mutation"):
            args = [[["nonempty"]]] if method == "add_merged_table" else []
            registry.prepare_call(method, args, {})


def test_unknown_objects_and_malformed_dto_are_rejected_before_request():
    owner = object()
    with pytest.raises(TypeError, match="closed business protocol"):
        proxy._arguments((object(),), owner)
    malformed = {"__wpscomposer_dto__": {"type": "Span", "value": {"text": "x"}}}
    with pytest.raises(ValueError, match="DTO"):
        proxy._result(malformed, owner)


def test_closed_codec_accepts_only_business_method_dto_classes():
    with pytest.raises(TypeError, match="closed business protocol"):
        proxy._arguments((QualityReport(),), object())


@pytest.mark.parametrize(
    "name,payload",
    [
        ("ExecutionIssue", {"code": "REF_MISSING", "message": "missing"}),
        ("FieldSnapshot", {
            "stableKey": ["toc", "TOC", 1], "fieldCategory": "native",
            "resultHash": "a" * 64, "tocPageCount": 1,
            "figureIndexPageCount": 0, "tableIndexPageCount": 0,
            "totalPages": 2,
        }),
        ("QualityFinding", {
            "code": "LOW_DPI", "severity": "warning", "confidence": "high",
            "message": "low resolution",
        }),
    ],
)
def test_mapping_backed_dtos_reject_every_undeclared_field(name, payload):
    payload["unexpected"] = "must not be ignored"
    with pytest.raises(ValueError, match="DTO fields"):
        decode_value({DTO_TAG: {"type": name, "value": payload}})


@pytest.mark.parametrize(
    "name,payload,missing",
    [
        ("ExecutionIssue", {"code": "REF_MISSING", "message": "missing"}, "message"),
        ("FieldSnapshot", {
            "stableKey": ["toc", "TOC", 1], "fieldCategory": "native",
            "resultHash": "a" * 64, "tocPageCount": 1,
            "figureIndexPageCount": 0, "tableIndexPageCount": 0,
            "totalPages": 2,
        }, "totalPages"),
        ("QualityFinding", {
            "code": "LOW_DPI", "severity": "warning", "confidence": "high",
            "message": "low resolution",
        }, "message"),
    ],
)
def test_mapping_backed_dtos_require_only_their_declared_required_fields(
    name, payload, missing
):
    valid = decode_value({DTO_TAG: {"type": name, "value": payload}})
    assert type(valid).__name__ == name
    del payload[missing]
    with pytest.raises(ValueError, match="DTO fields"):
        decode_value({DTO_TAG: {"type": name, "value": payload}})


def test_proxy_arguments_preserve_tuple_tags_with_nested_handles_and_paths(tmp_path):
    owner = object()
    handle = proxy.NativeSessionHandle(owner, "a" * 32, "range")
    encoded = proxy._arguments(
        (handle, tmp_path / "image.png", ("nested", [handle, tmp_path / "two.png"])),
        owner,
    )

    assert encoded[DTO_TAG]["type"] == "tuple"
    decoded = decode_value(encoded)
    assert type(decoded) is tuple
    assert decoded[0] == {"__wpscomposer_handle__": {"id": "a" * 32, "type": "range"}}
    assert decoded[1] == str(tmp_path / "image.png")
    assert type(decoded[2]) is tuple
    assert decoded[2][1][0] == decoded[0]
    assert decoded[2][1][1] == str(tmp_path / "two.png")
