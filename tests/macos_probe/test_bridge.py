import json
from concurrent.futures import ThreadPoolExecutor
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pytest

from skills.WPSComposer.scripts.macos_probe.bridge import BridgeState, LoopbackBridge
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult, ProtocolError


def request(bridge, method, path, body=None, token=None, origin=None):
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if origin:
        headers["Origin"] = origin
    if path.startswith("/v1/next") and "sessionNonce=" not in path:
        separator = "&" if "?" in path else "?"
        path += f"{separator}sessionNonce={bridge.session_nonce}"
    if body is not None and path.split("?", 1)[0] in {
        "/v1/session",
        "/v1/register",
        "/v1/result",
    }:
        body = dict(body)
        body.setdefault("sessionNonce", bridge.session_nonce)
    data = None if body is None else json.dumps(body).encode()
    req = Request(bridge.url + path, data=data, headers=headers, method=method)
    with urlopen(req, timeout=2) as response:
        raw = response.read()
        return response.status, None if not raw else json.loads(raw)


def test_bridge_rejects_missing_token():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        with pytest.raises(HTTPError) as error:
            request(bridge, "POST", "/v1/register", {"component": "writer"})
        assert error.value.code == 401


def test_session_token_claim_is_reload_safe():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        status, body = request(
            bridge, "POST", "/v1/session",
            {"component": "writer"}, origin=origin,
        )
        assert status == 200
        assert body == {"token": bridge.token, "leaseSeconds": 15.0}
        # re-claim is allowed so a reloaded add-in webview can recover
        status, body = request(
            bridge, "POST", "/v1/session",
            {"component": "writer"}, origin=origin,
        )
        assert status == 200
        assert body == {"token": bridge.token, "leaseSeconds": 15.0}
        # invalid component is rejected
        with pytest.raises(HTTPError) as error:
            request(
                bridge, "POST", "/v1/session",
                {"component": "bogus"}, origin=origin,
            )
        assert error.value.code == 400
        # unlisted origin is rejected
        with pytest.raises(HTTPError) as error:
            request(
                bridge, "POST", "/v1/session",
                {"component": "presentation"}, origin="http://example.com",
            )
        assert error.value.code == 403


def test_non_ascii_authorization_header_returns_401():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        req = Request(
            bridge.url + "/v1/register",
            data=json.dumps({"component": "writer"}).encode(),
            headers={
                "Content-Type": "application/json",
                "Origin": origin,
                "Authorization": "Bearer caf\u00e9",
            },
            method="POST",
        )
        with pytest.raises(HTTPError) as error:
            with urlopen(req, timeout=2):
                pass
        assert error.value.code == 401


def test_bridge_rejects_unlisted_origin():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        with pytest.raises(HTTPError) as error:
            request(
                bridge,
                "POST",
                "/v1/register",
                {"component": "writer"},
                bridge.token,
                "http://example.com",
            )
        assert error.value.code == 403


def test_command_round_trip():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        status, _ = request(
            bridge,
            "POST",
            "/v1/register",
            {"component": "writer"},
            bridge.token,
            origin,
        )
        assert status == 204
        command = bridge.issue(
            "writer", "smoke_docx", {"outputPath": "/tmp/a.docx"}
        )
        status, payload = request(
            bridge,
            "GET",
            "/v1/next?component=writer",
            token=bridge.token,
            origin=origin,
        )
        assert status == 200
        assert payload == command.to_dict()

        status, _ = request(
            bridge,
            "POST",
            "/v1/result",
            {
                "id": command.id,
                "ok": True,
                "value": {"saved": True},
                "error": None,
            },
            bridge.token,
            origin,
        )
        assert status == 204
        assert bridge.wait_result(command.id, 1).value == {"saved": True}


def test_wait_registered_reports_independent_components():
    origins = {
        "http://127.0.0.1:3891",
        "http://127.0.0.1:3892",
        "http://127.0.0.1:3893",
    }
    with LoopbackBridge(origins) as bridge, ThreadPoolExecutor() as pool:
        for component, port in (
            ("writer", 3891),
            ("presentation", 3892),
            ("spreadsheet", 3893),
        ):
            pool.submit(
                request,
                bridge,
                "POST",
                "/v1/register",
                {"component": component},
                bridge.token,
                f"http://127.0.0.1:{port}",
            )
        bridge.wait_registered({"writer", "presentation", "spreadsheet"}, 2)
        assert bridge.registered_components() == {
            "writer",
            "presentation",
            "spreadsheet",
        }


def test_completion_is_idempotent_and_cancellation_ignores_late_result():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        bridge.state.register("writer", bridge.session_nonce)
        first = bridge.issue("writer", "smoke_docx", {})
        result = ProbeResult(first.id, True, {"saved": True}, None)
        bridge.state.complete(result, bridge.session_nonce)
        bridge.state.complete(result, bridge.session_nonce)
        assert bridge.wait_result(first.id, 1) == result

        second = bridge.issue("writer", "smoke_pdf", {})
        bridge.state.cancel(second.id)
        bridge.state.complete(
            ProbeResult(second.id, True, {"saved": True}, None),
            bridge.session_nonce,
        )
        with pytest.raises(TimeoutError):
            bridge.wait_result(second.id, 0.01)


def test_conflicting_duplicate_result_is_rejected():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        bridge.state.register("writer", bridge.session_nonce)
        command = bridge.issue("writer", "smoke_docx", {})
        bridge.state.complete(
            ProbeResult(command.id, True, {"saved": True}, None),
            bridge.session_nonce,
        )
        with pytest.raises(ProtocolError, match="Conflicting duplicate"):
            bridge.state.complete(
                ProbeResult(command.id, False, {}, {"code": "failed"}),
                bridge.session_nonce,
            )


def test_stale_client_lease_is_not_registered():
    now = [10.0]
    state = BridgeState(
        session_nonce="session-a",
        lease_seconds=5,
        clock=lambda: now[0],
    )
    state.register("writer", "session-a")
    assert state.registered_components() == {"writer"}

    now[0] = 16.0

    assert state.registered_components() == set()
    with pytest.raises(TimeoutError, match="writer"):
        state.wait_registered({"writer"}, 0)
    with pytest.raises(ProtocolError, match="fresh client lease"):
        state.issue("writer", "smoke_docx", {})

    assert state.next("writer", "session-a", timeout=0) is None
    assert state.registered_components() == {"writer"}


def test_bridge_rejects_client_from_an_old_session_nonce():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        with pytest.raises(HTTPError) as error:
            request(
                bridge,
                "POST",
                "/v1/session",
                {"component": "writer", "sessionNonce": "stale-session"},
                origin=origin,
            )
        assert error.value.code == 400


def test_valid_result_renews_lease_for_next_command():
    now = [10.0]
    state = BridgeState(
        session_nonce="session-a",
        lease_seconds=5,
        clock=lambda: now[0],
    )
    state.register("writer", "session-a")
    first = state.issue("writer", "probe_capabilities", {})
    assert state.next("writer", "session-a", timeout=0) == first

    now[0] = 16.0
    state.complete(
        ProbeResult(first.id, True, {"supported": True}, None),
        "session-a",
    )

    second = state.issue("writer", "smoke_docx", {})
    assert second.component == "writer"
