"""Rejected POSTs must deliver their error before bounded transport teardown."""

import io
import json
import socket
from http.client import HTTPConnection, HTTPResponse
from threading import Event

import pytest

from skills.WPSComposer.scripts.macos_probe import bridge as module


ORIGIN = "https://allowed.example"


def handler_double(monkeypatch, *, command="POST", length="999999999", fault=None):
    events = []

    class Writer(io.BytesIO):
        def flush(self):
            events.append(("flush", self.getvalue()))
            if fault == "flush":
                raise BrokenPipeError("peer closed")
            super().flush()

    class Reader:
        closed = False

        def read1(self, count):
            events.append(("read", count))
            if fault == "read":
                raise ConnectionResetError("peer reset")
            if fault == "missing":
                raise socket.timeout("no body")
            return b"x" * count

        def close(self):
            self.closed = True

    class Connection:
        def shutdown(self, direction):
            events.append(("shutdown", direction))
            if fault == "shutdown":
                raise OSError("peer closed")

        def settimeout(self, timeout):
            events.append(("timeout", timeout))
            if fault == "timeout":
                raise OSError("peer closed")

    handler_class = module._handler_class(
        module.BridgeState(session_nonce="nonce"), frozenset({ORIGIN})
    )
    handler = object.__new__(handler_class)
    handler.command = command
    handler.request_version = "HTTP/1.1"
    handler.requestline = command + " /v1/register HTTP/1.1"
    handler.headers = {"Origin": ORIGIN}
    if length is not None:
        handler.headers["Content-Length"] = length
    handler.wfile = Writer()
    handler.rfile = Reader()
    handler.connection = Connection()
    handler.close_connection = False
    monkeypatch.setattr(module, "monotonic", lambda: 0)
    return handler, events


@pytest.mark.parametrize("length", [None, "-1", "invalid", "999999999"])
def test_rejection_flushes_then_half_closes_and_caps_discard(monkeypatch, length):
    handler, events = handler_double(monkeypatch, length=length)
    handler._send_error_json(403, "ORIGIN_REJECTED", "Origin is not allowed")
    handler.finish()

    assert handler.close_connection
    assert events[0][0] == "flush"
    response = events[0][1]
    assert b"403 Forbidden\r\n" in response
    assert b"Connection: close\r\n" in response
    assert json.loads(response.split(b"\r\n\r\n", 1)[1])["error"]["code"] == "ORIGIN_REJECTED"
    assert events[1] == ("shutdown", socket.SHUT_WR)
    reads = [value for kind, value in events if kind == "read"]
    assert sum(reads) == module.MAX_BODY_BYTES + 1
    assert max(reads) <= 65536
    assert handler.wfile.closed and handler.rfile.closed


def test_slow_body_uses_one_absolute_deadline(monkeypatch):
    handler, events = handler_double(monkeypatch)
    clock = iter([0, 0, 0.08, 0.16, 0.24])
    monkeypatch.setattr(module, "monotonic", lambda: next(clock))

    def trickle(count):
        events.append(("read", 1))
        return b"x"

    handler.rfile.read1 = trickle
    handler._send_error_json(401, "UNAUTHORIZED", "Denied")
    handler.finish()
    timeouts = [value for kind, value in events if kind == "timeout"]
    assert timeouts == pytest.approx([0.2, 0.12, 0.04])
    assert sum(value for kind, value in events if kind == "read") == 3
    assert handler.wfile.closed and handler.rfile.closed


@pytest.mark.parametrize("fault", ["flush", "shutdown", "timeout", "read", "missing"])
def test_broken_or_missing_peer_does_not_prevent_stream_close(monkeypatch, fault):
    handler, events = handler_double(monkeypatch, fault=fault)
    handler._send_error_json(401, "UNAUTHORIZED", "Denied")
    handler.finish()
    assert handler.wfile.closed and handler.rfile.closed
    assert b"401 Unauthorized\r\n" in events[0][1]
    expected_phase = "read" if fault == "missing" else fault
    assert any(kind == expected_phase for kind, value in events)


def test_peer_eof_ends_discard_immediately(monkeypatch):
    handler, events = handler_double(monkeypatch)

    def eof(count):
        events.append(("read", 0))
        return b""

    handler.rfile.read1 = eof
    handler._send_error_json(401, "UNAUTHORIZED", "Denied")
    handler.finish()
    assert [value for kind, value in events if kind == "read"] == [0]
    assert handler.wfile.closed and handler.rfile.closed


@pytest.mark.parametrize("command,status", [("POST", 200), ("POST", 400), ("GET", 403), ("OPTIONS", 403)])
def test_other_responses_do_not_linger(monkeypatch, command, status):
    handler, events = handler_double(monkeypatch, command=command)
    handler._send_error_json(status, "UNCHANGED", "Unchanged")
    handler.finish()
    assert not any(kind in {"shutdown", "timeout", "read"} for kind, value in events)
    assert b"Connection: close\r\n" not in events[0][1]


@pytest.mark.parametrize("path,authenticated,origin,status,code", [
    ("/v1/register", True, "https://forbidden.example", 403, "ORIGIN_REJECTED"),
    ("/v1/register", False, ORIGIN, 401, "UNAUTHORIZED"),
    ("/v1/session", False, "https://forbidden.example", 403, "ORIGIN_REJECTED"),
])
def test_large_rejected_post_delivers_exact_error_without_parsing(monkeypatch, path, authenticated, origin, status, code):
    original_factory = module._handler_class

    def factory(state, origins):
        class Handler(original_factory(state, origins)):
            def _read_json(self):
                pytest.fail("rejected body must not be parsed")
        return Handler

    monkeypatch.setattr(module, "_handler_class", factory)
    body = b'{"padding":"' + b"x" * (module.MAX_BODY_BYTES - 14) + b'"}'
    assert len(body) == module.MAX_BODY_BYTES
    with module.LoopbackBridge({ORIGIN}) as bridge:
        credentials = bridge.bootstrap_credentials("writer")
        token = bridge.state.claim_session("writer", credentials["clientId"], credentials["capability"])
        headers = {"Origin": origin, "Content-Type": "application/json"}
        if authenticated:
            headers["Authorization"] = "Bearer " + token
        for _ in range(5):
            connection = HTTPConnection(*bridge._server.server_address, timeout=2)
            try:
                connection.request("POST", path, body, headers)
                response = connection.getresponse()
                assert response.status == status
                assert response.getheader("Connection") == "close"
                assert json.loads(response.read())["error"]["code"] == code
            finally:
                connection.close()
        assert not bridge.state._registered


def test_missing_body_gets_response_before_bounded_finish(monkeypatch):
    finished = Event()
    original_factory = module._handler_class

    def factory(state, origins):
        class Handler(original_factory(state, origins)):
            def finish(self):
                try:
                    super().finish()
                finally:
                    finished.set()
        return Handler

    monkeypatch.setattr(module, "_handler_class", factory)
    with module.LoopbackBridge({ORIGIN}) as bridge:
        with socket.create_connection(bridge._server.server_address, timeout=2) as connection:
            connection.sendall(b"POST /v1/register HTTP/1.1\r\nHost: localhost\r\nContent-Length: 1048576\r\n\r\n")
            response = HTTPResponse(connection)
            response.begin()
            assert response.status == 401
            assert response.getheader("Connection") == "close"
            assert json.loads(response.read())["error"]["code"] == "UNAUTHORIZED"
            assert finished.wait(1), "missing body must not hold the handler indefinitely"
            response.close()


@pytest.mark.parametrize("length", ["invalid", "-1", str(module.MAX_BODY_BYTES + 1)])
@pytest.mark.parametrize("peer", ["over-cap", "slow", "missing"])
def test_early_length_rejection_uses_bounded_discard_without_body_parsing(monkeypatch, length, peer):
    handler, events = handler_double(
        monkeypatch, length=length, fault="missing" if peer == "missing" else None
    )
    handler.path = "/v1/register"
    handler._authorize = lambda: ("writer", "client")
    handler.rfile.read = lambda count: pytest.fail("invalid length must reject before parsing")
    if peer == "slow":
        clock = iter([0, 0, 0.08, 0.16, 0.24])
        monkeypatch.setattr(module, "monotonic", lambda: next(clock))

        def trickle(count):
            events.append(("read", 1))
            return b"x"

        handler.rfile.read1 = trickle
    handler.do_POST()
    handler.finish()
    assert handler.close_connection
    assert events[0][0] == "flush"
    response = events[0][1]
    assert b"400 Bad Request\r\n" in response
    assert b"Connection: close\r\n" in response
    assert json.loads(response.split(b"\r\n\r\n", 1)[1])["error"] == {
        "code": "INVALID_REQUEST",
        "message": "Invalid Content-Length" if length == "invalid" else "Request body is too large",
    }
    assert events[1] == ("shutdown", socket.SHUT_WR)
    reads = [value for kind, value in events if kind == "read"]
    timeouts = [value for kind, value in events if kind == "timeout"]
    if peer == "over-cap":
        assert sum(reads) == module.MAX_BODY_BYTES + 1
        assert max(reads) <= 65536
    elif peer == "slow":
        assert reads == [1, 1, 1]
        assert timeouts == pytest.approx([0.2, 0.12, 0.04])
    else:
        assert len(reads) == 1
        assert timeouts == pytest.approx([0.2])
    assert handler.wfile.closed and handler.rfile.closed


@pytest.mark.parametrize("length", ["invalid", "-1", str(module.MAX_BODY_BYTES + 1)])
def test_max_plus_one_body_gets_exact_early_400_without_registration(length):
    with module.LoopbackBridge({ORIGIN}) as bridge:
        credentials = bridge.bootstrap_credentials("writer")
        token = bridge.state.claim_session("writer", credentials["clientId"], credentials["capability"])
        body = json.dumps({"component": "writer", "clientId": credentials["clientId"], "padding": ""}).encode()
        size = module.MAX_BODY_BYTES + 1
        body = body[:-2] + b"x" * (size - len(body)) + body[-2:]
        assert len(body) == size and isinstance(json.loads(body), dict)
        for _ in range(3):
            connection = HTTPConnection(*bridge._server.server_address, timeout=2)
            try:
                connection.request("POST", "/v1/register", body, {
                    "Origin": ORIGIN, "Authorization": "Bearer " + token,
                    "Content-Type": "application/json", "Content-Length": length,
                })
                response = connection.getresponse()
                assert response.status == 400
                assert response.getheader("Connection") == "close"
                assert json.loads(response.read())["error"] == {
                    "code": "INVALID_REQUEST",
                    "message": "Invalid Content-Length" if length == "invalid" else "Request body is too large",
                }
            finally:
                connection.close()
            assert not bridge.state._registered


@pytest.mark.parametrize("body", [b"not json", b"[]", b'{"component":"sheet","clientId":"client"}'])
def test_consumed_json_and_business_errors_do_not_linger(monkeypatch, body):
    handler, events = handler_double(monkeypatch, length=str(len(body)))
    handler.path = "/v1/register"
    handler._authorize = lambda: ("writer", "client")

    def read(count):
        assert count == len(body)
        events.append(("body-read", count))
        return body

    handler.rfile.read = read
    handler.do_POST()
    handler.finish()
    assert events[0] == ("body-read", len(body))
    response = next(value for kind, value in events if kind == "flush")
    assert b"400 Bad Request\r\n" in response
    assert json.loads(response.split(b"\r\n\r\n", 1)[1])["error"]["code"] == "INVALID_REQUEST"
    assert b"Connection: close\r\n" not in response
    assert not any(kind in {"shutdown", "timeout", "read"} for kind, value in events)
