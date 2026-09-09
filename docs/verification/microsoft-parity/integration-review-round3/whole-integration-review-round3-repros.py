"""Reviewer-owned RED for current Microsoft-parity documentation consistency."""

from pathlib import Path
import io
import json
import re
from types import SimpleNamespace

import pytest


ROOT = Path(__file__).resolve().parents[3]


def test_current_status_matches_linked_word_direct_inventory():
    status = (ROOT / "docs/verification/microsoft-parity/status.md").read_text()
    remaining = (
        ROOT / "docs/verification/microsoft-parity/word-direct-remaining.md"
    ).read_text()

    inventory = re.search(
        r"(\d+) names are declared by MacWordSession and (\d+) remain absent",
        remaining,
    )
    assert inventory is not None
    declared, absent = inventory.groups()
    assert f"{declared} of the frozen 80 direct Word method names are declared" in status
    assert f"{absent} [remaining methods](word-direct-remaining.md)" in status


def test_windows_proxy_verified_save_failure_keeps_session_for_discard(monkeypatch):
    """A completed save error must not make native cleanup unverified."""
    from skills.WPSComposer.scripts.msoffice import windows_session_proxy as module

    class Child:
        stdin = type("Input", (), {"close": lambda self: None})()
        returncode = 0

        @staticmethod
        def wait(timeout):
            return 0

    session = module.ProxyWordSession.__new__(module.ProxyWordSession)
    session._closed = session._uncertain = False
    session._child = Child()
    session._lock = type("Lock", (), {"close": lambda self: None})()
    session._stop_io = lambda: None
    calls = []

    def completed_call(method, *args, **kwargs):
        calls.append((method, kwargs))
        if len(calls) == 1:
            raise RuntimeError("Source changed since open")
        return {"closed": True}

    session._call = completed_call
    session._abort = lambda reason: setattr(session, "_uncertain", True)
    session._error = lambda suffix: RuntimeError(suffix)
    session._remaining = lambda: 10

    with pytest.raises(RuntimeError, match="Source changed since open"):
        session.close(save_changes=True)

    assert not session._uncertain
    assert not session._closed
    session.close(save_changes=False)
    assert session._closed
    assert calls == [
        ("close", {"save_changes": True}),
        ("close", {"save_changes": False}),
    ]


def test_windows_worker_preserves_handles_after_acknowledged_saving_close_error(
    tmp_path,
):
    """A retryable saving-close error must not invalidate the live registry."""
    from skills.WPSComposer.scripts.msoffice import windows_session_worker as worker

    class Collection:
        def __init__(self, values=()):
            self.values = list(values)

        @property
        def Count(self):
            return len(self.values)

        def Item(self, index):
            return self.values[index - 1]

    document = SimpleNamespace(token="document")
    table = SimpleNamespace(token="table")
    document.Content = SimpleNamespace(End=10)
    document.Tables = Collection([table])
    document.Shapes = Collection()
    document.InlineShapes = Collection()
    document.TablesOfContents = Collection()
    document.TablesOfFigures = Collection()
    document.Paragraphs = Collection()
    calls = []

    class Session:
        staging_root = tmp_path
        kind = "writer"

        def __init__(self):
            self._composer = SimpleNamespace(
                _doc=document,
                _deps=SimpleNamespace(identity=lambda value: value.token),
            )

        def _verify(self):
            return None

        def add_table(self, *args, **kwargs):
            return table

        def apply_format_patch(self, target, **kwargs):
            calls.append((target, kwargs))
            return {"target": target, "ok": True}

        def save_current(self):
            raise RuntimeError("Source changed since open")

        def close(self, save_changes=False):
            return None

    def request(number, method, *, args=(), kwargs=None):
        return {
            "protocol": 1,
            "id": number,
            "kind": "writer",
            "method": method,
            "args": list(args),
            "kwargs": kwargs or {},
            "deadline": 1_000_000_000.0,
            "remaining_seconds": 500,
        }

    first = request(1, "new_document")
    create = request(2, "add_table", args=([['cell']],))
    # A single serve stream is required for handle identity. Build it manually
    # after deriving the deterministic request envelope, not the random token.
    # NativeHandleRegistry registration is exercised below, while serve owns
    # the lifecycle ordering under review.
    registry = worker.NativeHandleRegistry(Session())
    handle = registry.encode_result("add_table", table, [], {})
    frames = [
        first,
        create,
        request(3, "close", kwargs={"save_changes": True}),
        request(
            4,
            "apply_format_patch",
            args=(handle,),
            kwargs={"bold": True},
        ),
        request(5, "close", kwargs={"save_changes": False}),
    ]
    incoming = io.BytesIO(
        b"".join(json.dumps(row).encode() + b"\n" for row in frames)
    )
    outgoing = io.BytesIO()
    # Seed serve's registry token deterministically so request 4 references
    # exactly the handle returned by request 2.
    original_register = worker.NativeHandleRegistry.register

    def fixed_register(self, native, kind):
        self._location(native, kind)
        token = handle["__wpscomposer_handle__"]["id"]
        self.records[token] = (native, kind)
        return handle

    worker.NativeHandleRegistry.register = fixed_register
    try:
        worker.serve(incoming, outgoing, tmp_path, factory=lambda *args: Session())
    finally:
        worker.NativeHandleRegistry.register = original_register
    responses = [json.loads(line) for line in outgoing.getvalue().splitlines()]

    assert responses[1]["status"] == "ok"
    assert responses[2]["status"] == "error"
    assert responses[2]["error"]["message"] == "Source changed since open"
    assert responses[3]["status"] == "ok"
    assert calls == [("table:1", {"bold": True})]
    assert responses[4]["status"] == "ok"


def test_real_windows_session_retains_save_failure_evidence_after_worker_discard(
    tmp_path,
):
    """The split worker save phase must preserve _WindowsSession evidence."""
    from skills.WPSComposer.scripts.msoffice import windows_document_api as api
    from skills.WPSComposer.scripts.msoffice import windows_session_worker as worker

    class Composer:
        def __init__(self):
            self.closed = []

        def close(self, save_changes=False):
            self.closed.append(save_changes)

    session = api.WindowsWordSession.__new__(api.WindowsWordSession)
    session.staging_root = tmp_path / "real-session"
    session.staging_root.mkdir()
    session._attached = False
    session._closed = session._failed = False
    session._composer = Composer()
    session._source = session._logical_path = tmp_path / "source.docx"
    session._source_digest = session._logical_digest = "before"
    session._deadline = 1_000_000_000.0
    session._verify = lambda **kwargs: None

    def changed():
        raise RuntimeError("Source changed since open")

    session._check_current_unchanged = changed

    def request(number, method, *, kwargs=None):
        return {
            "protocol": 1,
            "id": number,
            "kind": "writer",
            "method": method,
            "args": [],
            "kwargs": kwargs or {},
            "deadline": 1_000_000_000.0,
            "remaining_seconds": 500,
        }

    frames = [
        request(1, "new_document"),
        request(2, "close", kwargs={"save_changes": True}),
        request(3, "close", kwargs={"save_changes": False}),
    ]
    incoming = io.BytesIO(
        b"".join(json.dumps(row).encode() + b"\n" for row in frames)
    )
    outgoing = io.BytesIO()
    worker.serve(incoming, outgoing, tmp_path, factory=lambda *args: session)
    responses = [json.loads(line) for line in outgoing.getvalue().splitlines()]

    assert responses[1]["status"] == "error"
    assert responses[1]["session_state"] == "open"
    assert responses[2]["status"] == "ok"
    assert session._failed and session._closed
    assert session.staging_root.is_dir()
    failures = list(session.staging_root.glob("failure-*.json"))
    assert failures
    assert json.loads(failures[0].read_text())["native_cleanup_verified"] is False


def test_worker_does_not_call_failed_identity_check_an_open_session(tmp_path):
    """`_closed is False` does not prove the original native binding is live."""
    from skills.WPSComposer.scripts.msoffice import windows_document_api as api
    from skills.WPSComposer.scripts.msoffice import windows_session_worker as worker

    class Composer:
        def close(self, save_changes=False):
            return None

    session = api.WindowsWordSession.__new__(api.WindowsWordSession)
    session.staging_root = tmp_path / "identity-session"
    session.staging_root.mkdir()
    session._attached = False
    session._closed = session._failed = False
    session._composer = Composer()
    session._source = session._logical_path = tmp_path / "source.docx"
    session._source_digest = session._logical_digest = "before"
    session._deadline = 1_000_000_000.0

    def stale(**kwargs):
        raise api.DocumentIdentityError("Microsoft document session identity changed")

    session._verify = stale

    def request(number, method, *, kwargs=None):
        return {
            "protocol": 1,
            "id": number,
            "kind": "writer",
            "method": method,
            "args": [],
            "kwargs": kwargs or {},
            "deadline": 1_000_000_000.0,
            "remaining_seconds": 500,
        }

    frames = [
        request(1, "new_document"),
        request(2, "close", kwargs={"save_changes": True}),
    ]
    incoming = io.BytesIO(
        b"".join(json.dumps(row).encode() + b"\n" for row in frames)
    )
    outgoing = io.BytesIO()
    worker.serve(incoming, outgoing, tmp_path, factory=lambda *args: session)
    responses = [json.loads(line) for line in outgoing.getvalue().splitlines()]

    assert responses[1]["status"] == "error"
    assert "identity changed" in responses[1]["error"]["message"]
    assert "session_state" not in responses[1]
