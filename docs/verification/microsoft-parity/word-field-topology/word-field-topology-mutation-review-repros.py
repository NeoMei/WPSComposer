"""Pure reviewer reproductions for Mac Word field topology ownership."""

from __future__ import annotations

from copy import deepcopy

import pytest

from skills.WPSComposer.scripts.msoffice.macos_word_session import (
    MacWordSession,
    NativeWordError,
)


BOOKMARK = "wpsc_fig_" + "a" * 24
REFERENCE_CODE = f" REF {BOOKMARK} \\h "
REFERENCE_RUN = {
    "type": "reference",
    "bookmarkName": BOOKMARK,
    "prefix": "见😀(",
    "suffix": ")尾",
    "fallbackText": "静态",
}


def _session():
    session = MacWordSession()
    session.staging_root = None
    return session


def _existing_ref(*, start=1, visible="1"):
    return [
        ["stats", 3],
        ["field", "main", 1, "REF", " REF existing ", start, visible, 0, 0, 1],
    ]


def test_snapshot_accepts_reference_created_by_same_session(monkeypatch):
    session = _session()
    before = _existing_ref()
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(before))
    original = session.snapshot_fields()

    def insert(lines):
        identity = next(
            line for line in lines if "make new bookmark" in line
        ).split('name:"', 1)[1].split('"', 1)[0]
        return [
            ["literal", 0, 20, 24, "见😀("],
            ["reference", 0, 1, 2, 25, 66, 67, REFERENCE_CODE, identity],
            ["literal", 0, 68, 70, ")尾"],
            ["literal", 1, 70, 71, "\r"],
            ["complete", 1, 2, 20, 71],
        ]

    monkeypatch.setattr(session, "_execute", insert)
    session.add_cross_reference_paragraph(
        runs=[REFERENCE_RUN], owner_node_id="owner:new"
    )
    handle, code = session._tracked_references[0]
    after = _existing_ref() + [
        ["identity", handle.bookmark, 25],
        ["field", "main", 2, "REF", code, 25, "2", 0, 0, 1],
    ]
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(after))

    current = session.snapshot_fields()
    assert [item.stable_key for item in current] == [
        original[0].stable_key,
        ("owner:new", "REF", 0),
    ]


def test_snapshot_accepts_toc_created_by_same_session(monkeypatch):
    session = _session()
    before = _existing_ref()
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(before))
    original = session.snapshot_fields()

    toc_code = ' TOC \\o "1-3" \\h \\z \\* MERGEFORMAT '
    monkeypatch.setattr(
        session,
        "_execute",
        lambda _lines: [["index", 30, 60, toc_code]],
    )
    handle = session.insert_toc()
    code = session._tracked_indexes[0][1]
    after = _existing_ref() + [
        ["identity", handle.bookmark, 30],
        ["field", "main", 2, "INDEX", code, 30, "Heading\t1", 1, 1, 20],
    ]
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(after))

    current = session.snapshot_fields()
    assert [item.stable_key for item in current] == [
        original[0].stable_key,
        ("doc:toc", "TOC", 0),
    ]


def test_snapshot_rebases_after_same_session_structural_insert(monkeypatch):
    session = _session()
    before = _existing_ref(start=12)
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(before))
    original = session.snapshot_fields()

    monkeypatch.setattr(session, "_execute", lambda _lines: [["ok"]])
    session.apply_structural_op(
        {
            "op": "insert",
            "type": "paragraph",
            "position": "start",
            "props": {"text": "prefix"},
        }
    )
    after = _existing_ref(start=19)
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(after))

    current = session.snapshot_fields()
    assert current[0].stable_key == original[0].stable_key


def test_snapshot_still_rejects_external_topology_drift_without_owned_mutation(
    monkeypatch,
):
    session = _session()
    rows = _existing_ref(start=12)
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(rows))
    session.snapshot_fields()
    rows[1][5] = 19

    with pytest.raises(NativeWordError) as caught:
        session.snapshot_fields()

    assert caught.value.code == "NATIVE_WORD_FIELD_IDENTITY_STALE"


def test_tracked_toc_growth_still_shifts_following_reference_exactly(monkeypatch):
    session = _session()
    before = [
        ["stats", 3],
        ["field", "main", 1, "INDEX", ' TOC \\o "1-3" ', 10, "A", 1, 1, 10],
        ["field", "main", 2, "REF", " REF target ", 30, "X", 0, 0, 1],
    ]
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(before))
    original = session.snapshot_fields()
    after = deepcopy(before)
    after[1][6] = "ABCDEFGHIJK"
    after[1][9] = 20
    after[2][5] = 40
    monkeypatch.setattr(session, "_execute", lambda _lines: deepcopy(after))

    current = session.snapshot_fields()
    assert [item.stable_key for item in current] == [
        item.stable_key for item in original
    ]
