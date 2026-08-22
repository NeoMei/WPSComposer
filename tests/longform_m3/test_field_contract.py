"""Behavior contract for M3 native-field convergence.

These tests are platform-pure.  Native COM/JSAPI adapters are added by the
following tasks; this module fixes the shared ordering, snapshot, failure, and
privacy semantics they must implement.
"""

from __future__ import annotations

import hashlib
import traceback
from dataclasses import dataclass
from typing import Tuple

import pytest

from skills.WPSComposer.scripts.longform.executor import FieldSnapshot
from skills.WPSComposer.scripts.longform.field_contract import (
    FIELD_KINDS,
    NativeFieldAdapter,
    NativeFieldContractError,
    evaluate_field_snapshot_history,
    finalize_native_fields,
    snapshot_visible_field,
)


@dataclass(frozen=True)
class _VisibleField:
    owner: str
    kind: str
    ordinal: int
    text: str


class _Adapter:
    def __init__(self, rounds: Tuple[Tuple[_VisibleField, ...], ...]) -> None:
        self.rounds = rounds
        self.round_index = 0
        self.calls: list[str] = []
        self.bookmark_by_owner = {"fig:1": "WPSC_FIG_1"}

    def repaginate_and_update_numbering(self) -> None:
        self.calls.append("numbering")

    def refresh_bookmarks_and_references(self) -> None:
        self.calls.append("references")

    def refresh_indexes(self) -> None:
        self.calls.append("indexes")

    def repaginate_and_update_page_fields(self) -> None:
        self.calls.append("pages")

    def snapshot_fields(self) -> tuple[FieldSnapshot, ...]:
        self.calls.append("snapshot")
        selected = self.rounds[min(self.round_index, len(self.rounds) - 1)]
        self.round_index += 1
        return tuple(
            snapshot_visible_field(
                owner_node_id=item.owner,
                field_kind=item.kind,
                ordinal_within_node=item.ordinal,
                visible_result=item.text,
                field_category="native",
                total_pages=5,
            )
            for item in selected
        )


def _field(owner: str, kind: str, ordinal: int, text: str) -> _VisibleField:
    return _VisibleField(owner, kind, ordinal, text)


def test_native_field_adapter_is_runtime_checkable_and_field_kinds_are_closed():
    adapter = _Adapter(((_field("fig:1", "SEQ_FIG", 0, "1"),),))
    assert isinstance(adapter, NativeFieldAdapter)
    assert FIELD_KINDS == frozenset(
        {
            "STYLEREF",
            "SEQ_FIG",
            "SEQ_TAB",
            "SEQ_EQ",
            "REF",
            "TOC",
            "TOF_FIG",
            "TOF_TAB",
            "PAGE",
            "NUMPAGES",
        }
    )


def test_each_round_uses_exact_five_phase_order_and_converges_on_second_digest():
    fields = (
        _field("table:1", "SEQ_TAB", 0, "1-1"),
        _field("fig:1", "SEQ_FIG", 0, "1-1"),
        _field("front:figures", "TOF_FIG", 0, "图 1-1\t3"),
        _field("p:ref", "REF", 0, "1-1"),
    )
    adapter = _Adapter((fields, tuple(reversed(fields))))

    result = finalize_native_fields(adapter)

    one_round = ["numbering", "references", "indexes", "pages", "snapshot"]
    assert adapter.calls == one_round * 2
    assert result.rounds == 2
    assert result.issues == ()
    assert tuple(item.stable_key for item in result.snapshot) == tuple(
        sorted(item.stable_key for item in result.snapshot)
    )


def test_snapshot_hash_normalizes_nfc_and_all_newline_spellings_without_text_storage():
    variants = ("Cafe\u0301\r\n第一行\r第二行", "Café\n第一行\n第二行")
    snapshots = [
        snapshot_visible_field(
            owner_node_id="p:1",
            field_kind="REF",
            ordinal_within_node=0,
            visible_result=text,
            field_category="reference",
            total_pages=1,
        )
        for text in variants
    ]

    expected = hashlib.sha256(variants[1].encode("utf-8")).hexdigest()
    assert snapshots[0].result_hash == snapshots[1].result_hash == expected
    serialized = snapshots[0].to_dict()
    assert "visibleResult" not in serialized
    assert variants[0] not in repr(snapshots[0])


def test_three_changing_rounds_get_one_frozen_fourth_snapshot_without_mutation():
    rounds = tuple(
        (_field("fig:1", "SEQ_FIG", 0, str(number)),)
        for number in ("1", "2", "3", "4")
    )
    adapter = _Adapter(rounds)

    result = finalize_native_fields(adapter, max_rounds=3)

    mutating_phases = [
        call for call in adapter.calls if call in {"numbering", "references", "indexes", "pages"}
    ]
    assert len(mutating_phases) == 12
    assert adapter.calls[-2:] == ["snapshot", "snapshot"]
    assert result.rounds == 4
    assert [issue.code for issue in result.issues] == ["FIELD_REFRESH_UNSTABLE"]
    assert result.snapshot[0].result_hash == snapshot_visible_field(
        owner_node_id="fig:1",
        field_kind="SEQ_FIG",
        ordinal_within_node=0,
        visible_result="4",
        field_category="native",
        total_pages=5,
    ).result_hash


@pytest.mark.parametrize(
    "missing",
    (
        "repaginate_and_update_numbering",
        "refresh_bookmarks_and_references",
        "refresh_indexes",
        "repaginate_and_update_page_fields",
        "snapshot_fields",
    ),
)
def test_missing_required_native_api_is_fatal(missing: str):
    adapter = _Adapter(((_field("fig:1", "SEQ_FIG", 0, "1"),),))
    setattr(adapter, missing, None)

    with pytest.raises(NativeFieldContractError, match=missing):
        finalize_native_fields(adapter)


def test_throwing_required_native_api_is_fatal_and_contains_no_private_evidence():
    secret_text = "内部标题文本"
    secret_bookmark = "WPSC_PRIVATE_BOOKMARK"

    class _Throwing(_Adapter):
        def refresh_indexes(self) -> None:
            raise RuntimeError(f"{secret_text} {secret_bookmark}")

    adapter = _Throwing(((_field("fig:1", "SEQ_FIG", 0, "1"),),))
    with pytest.raises(NativeFieldContractError) as exc_info:
        finalize_native_fields(adapter)

    assert secret_text not in str(exc_info.value)
    assert secret_bookmark not in str(exc_info.value)
    assert exc_info.value.phase == "refresh_indexes"


def test_adapter_exception_chain_and_formatted_traceback_are_fully_sanitized():
    secret = "/Users/private/标题 WPSC_SECRET deadbeef"

    class _SecretContractError(_Adapter):
        def refresh_indexes(self) -> None:
            raise NativeFieldContractError(secret, secret)

    with pytest.raises(NativeFieldContractError) as exc_info:
        finalize_native_fields(
            _SecretContractError(((_field("fig:1", "SEQ_FIG", 0, "1"),),))
        )

    exc = exc_info.value
    rendered = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    assert exc.__cause__ is None
    assert exc.__context__ is None
    assert secret not in rendered
    assert exc.phase == "refresh_indexes"


def test_malformed_snapshot_objects_are_fatal_without_exception_chain_or_secrets():
    secret = "/private/resource.png 可见文本 WPSC_BOOKMARK cafebabe"

    class _EvilSnapshot:
        @property
        def stable_key(self):
            raise RuntimeError(secret)

    malformed_values = (
        (_EvilSnapshot(),),
        (FieldSnapshot(("fig:1", [], 0), "native", "0" * 64, 0, 0, 0, 1),),
        (FieldSnapshot(("fig:1", "SEQ_FIG", 0), "native", None, 0, 0, 0, 1),),
    )

    class _Malformed(_Adapter):
        def __init__(self, value) -> None:
            super().__init__(((),))
            self.value = value

        def snapshot_fields(self):
            return self.value

    for value in malformed_values:
        with pytest.raises(NativeFieldContractError) as exc_info:
            finalize_native_fields(_Malformed(value))
        exc = exc_info.value
        rendered = "".join(
            traceback.format_exception(type(exc), exc, exc.__traceback__)
        )
        assert exc.__cause__ is None
        assert exc.__context__ is None
        assert secret not in rendered
        assert "cafebabe" not in rendered
        assert exc.phase == "snapshot_fields"


@pytest.mark.parametrize("value", (4, 0, -1, True, 2.9, "3", None))
def test_max_rounds_accepts_only_non_bool_int_from_one_through_three(value):
    adapter = _Adapter(((_field("fig:1", "SEQ_FIG", 0, "1"),),))
    with pytest.raises(NativeFieldContractError, match="max_rounds") as exc_info:
        finalize_native_fields(adapter, max_rounds=value)
    assert exc_info.value.__cause__ is None
    assert exc_info.value.__context__ is None
    assert adapter.calls == []


def test_unstable_issue_evidence_has_counts_only_no_text_hash_or_bookmark_mapping():
    secret = "可见内容-绝密"
    rounds = tuple(
        (_field("fig:1", "SEQ_FIG", 0, secret + str(index)),)
        for index in range(4)
    )
    adapter = _Adapter(rounds)
    result = finalize_native_fields(adapter)
    issue_json = result.issues[0].to_dict()
    evidence = repr(issue_json)

    assert secret not in evidence
    assert adapter.bookmark_by_owner["fig:1"] not in evidence
    assert all(item.result_hash not in evidence for item in result.snapshot)
    assert set(issue_json) <= {"code", "message", "placement", "nodeId"}


@pytest.mark.parametrize("mutation", ("move", "insert", "delete"))
def test_mutations_change_seq_ref_snapshot_without_changing_owner_or_bookmark(mutation: str):
    before = (
        _field("fig:1", "SEQ_FIG", 0, "1"),
        _field("p:ref", "REF", 0, "1"),
    )
    changed_number = {"move": "2", "insert": "3", "delete": "1"}[mutation]
    changed_ref = {"move": "2", "insert": "3", "delete": "错误! 未找到引用源。"}[mutation]
    after = (
        _field("fig:1", "SEQ_FIG", 0, changed_number),
        _field("p:ref", "REF", 0, changed_ref),
    )
    class _BookmarkResolvedAdapter(_Adapter):
        def __init__(self, rounds):
            super().__init__(rounds)
            self.owner_by_bookmark = {"WPSC_FIG_1": "fig:1", "WPSC_REF_1": "p:ref"}
            self.bookmark_order = ("WPSC_FIG_1", "WPSC_REF_1")

        def snapshot_fields(self) -> tuple[FieldSnapshot, ...]:
            self.calls.append("snapshot")
            selected = self.rounds[min(self.round_index, len(self.rounds) - 1)]
            self.round_index += 1
            return tuple(
                snapshot_visible_field(
                    owner_node_id=self.owner_by_bookmark[bookmark],
                    field_kind=item.kind,
                    ordinal_within_node=item.ordinal,
                    visible_result=item.text,
                    field_category="native",
                    total_pages=5,
                )
                for bookmark, item in zip(self.bookmark_order, selected)
            )

    adapter = _BookmarkResolvedAdapter((before, after, after))
    original_mapping = dict(adapter.owner_by_bookmark)

    result = finalize_native_fields(adapter)

    assert result.rounds == 3
    assert adapter.owner_by_bookmark == original_mapping
    assert [item.stable_key for item in result.snapshot] == [
        ("fig:1", "SEQ_FIG", 0),
        ("p:ref", "REF", 0),
    ]
    assert tuple(item.result_hash for item in result.snapshot) != tuple(
        snapshot_visible_field(
            owner_node_id=item.owner,
            field_kind=item.kind,
            ordinal_within_node=item.ordinal,
            visible_result=item.text,
            field_category="native",
            total_pages=5,
        ).result_hash
        for item in before
    )


def test_unknown_field_kind_and_duplicate_stable_key_are_fatal():
    invalid = FieldSnapshot(
        stable_key=("fig:1", "USER_FIELD", 0),
        field_category="native",
        result_hash="0" * 64,
        toc_page_count=0,
        figure_index_page_count=0,
        table_index_page_count=0,
        total_pages=1,
    )

    class _Raw(_Adapter):
        def snapshot_fields(self) -> tuple[FieldSnapshot, ...]:
            return (invalid, invalid)

    with pytest.raises(NativeFieldContractError, match="field snapshot"):
        finalize_native_fields(_Raw(((),)))


def test_remote_history_stops_only_on_its_final_adjacent_equal_pair():
    a = (snapshot_visible_field(owner_node_id="fig:1", field_kind="SEQ_FIG", ordinal_within_node=0, visible_result="1", field_category="native"),)
    b = (snapshot_visible_field(owner_node_id="fig:1", field_kind="SEQ_FIG", ordinal_within_node=0, visible_result="2", field_category="native"),)

    converged = evaluate_field_snapshot_history((a, b, b), max_rounds=3)
    assert converged.rounds == 3
    assert converged.snapshot == b
    assert converged.issues == ()

    with pytest.raises(NativeFieldContractError, match="history") as exc_info:
        evaluate_field_snapshot_history((a, a, b, b), max_rounds=3)
    assert exc_info.value.__cause__ is None
    assert exc_info.value.__context__ is None


def test_remote_unstable_history_uses_the_read_only_fourth_snapshot_not_third():
    rounds = tuple(
        (
            snapshot_visible_field(
                owner_node_id="fig:1",
                field_kind="SEQ_FIG",
                ordinal_within_node=0,
                visible_result=str(index),
                field_category="native",
                total_pages=index,
            ),
        )
        for index in (1, 2, 3, 4)
    )

    result = evaluate_field_snapshot_history(rounds, max_rounds=3)

    assert result.rounds == 4
    assert result.snapshot == rounds[3]
    assert result.snapshot != rounds[2]
    assert [issue.code for issue in result.issues] == ["FIELD_REFRESH_UNSTABLE"]


def test_remote_history_rejects_truncation_and_post_convergence_mutation():
    a = (snapshot_visible_field(owner_node_id="fig:1", field_kind="SEQ_FIG", ordinal_within_node=0, visible_result="1", field_category="native"),)
    b = (snapshot_visible_field(owner_node_id="fig:1", field_kind="SEQ_FIG", ordinal_within_node=0, visible_result="2", field_category="native"),)
    for history in ((), (a,), (a, b), (a, a, b), (a, b, b, a)):
        with pytest.raises(NativeFieldContractError, match="history"):
            evaluate_field_snapshot_history(history, max_rounds=3)


def test_digest_rejects_field_snapshot_subclass_without_leaking_overridden_serializer():
    secret = "/private/field.docx WPSC_SECRET visible field hash"

    class _EvilFieldSnapshot(FieldSnapshot):
        def to_dict(self):
            raise RuntimeError(secret)

    evil = _EvilFieldSnapshot(
        stable_key=("fig:1", "SEQ_FIG", 0),
        field_category="native",
        result_hash="0" * 64,
        toc_page_count=0,
        figure_index_page_count=0,
        table_index_page_count=0,
        total_pages=1,
    )

    class _EvilAdapter(_Adapter):
        def snapshot_fields(self):
            return (evil,)

    with pytest.raises(NativeFieldContractError) as exc_info:
        finalize_native_fields(_EvilAdapter(((),)))
    exc = exc_info.value
    rendered = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    assert exc.__cause__ is None
    assert exc.__context__ is None
    assert secret not in rendered
