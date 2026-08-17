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


def claim_token(bridge, component, origin):
    credentials = bridge.bootstrap_credentials(component)
    _, body = request(
        bridge,
        "POST",
        "/v1/session",
        {"component": component, **credentials},
        origin=origin,
    )
    return credentials["clientId"], body["token"]


def test_bridge_rejects_missing_token():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        with pytest.raises(HTTPError) as error:
            request(bridge, "POST", "/v1/register", {"component": "writer"})
        assert error.value.code == 401


def test_bootstrap_capability_is_one_time_and_bound_to_component_client():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        credentials = bridge.bootstrap_credentials("writer")
        status, body = request(
            bridge, "POST", "/v1/session",
            {"component": "writer", **credentials}, origin=origin,
        )
        assert status == 200
        assert body["token"] != bridge.token
        assert body["leaseSeconds"] == 15.0
        # A capability is consumed after its first exchange.
        with pytest.raises(HTTPError) as error:
            request(
                bridge, "POST", "/v1/session",
                {"component": "writer", **credentials}, origin=origin,
            )
        assert error.value.code == 400
        # A client token cannot cross its component/client binding.
        with pytest.raises(HTTPError) as error:
            request(
                bridge,
                "POST",
                "/v1/register",
                {
                    "component": "presentation",
                    "clientId": credentials["clientId"],
                },
                body["token"],
                origin,
            )
        assert error.value.code in {400, 403}


def test_public_session_nonce_cannot_be_exchanged_for_bearer_token():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        with pytest.raises(HTTPError) as error:
            request(
                bridge,
                "POST",
                "/v1/session",
                {
                    "component": "writer",
                    "sessionNonce": bridge.session_nonce,
                },
                origin=origin,
            )
        assert error.value.code == 400


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
    allowed = "http://127.0.0.1:3891"
    with LoopbackBridge({allowed}) as bridge:
        client_id, token = claim_token(bridge, "writer", allowed)
        with pytest.raises(HTTPError) as error:
            request(
                bridge,
                "POST",
                "/v1/register",
                {"component": "writer", "clientId": client_id},
                token,
                "http://example.com",
            )
        assert error.value.code == 403


def test_command_round_trip():
    origin = "http://127.0.0.1:3891"
    with LoopbackBridge({origin}) as bridge:
        client_id, token = claim_token(bridge, "writer", origin)
        status, _ = request(
            bridge,
            "POST",
            "/v1/register",
            {"component": "writer", "clientId": client_id},
            token,
            origin,
        )
        assert status == 204
        command = bridge.issue(
            "writer", "smoke_docx", {"outputPath": "/tmp/a.docx"}
        )
        status, payload = request(
            bridge,
            "GET",
            f"/v1/next?component=writer&clientId={client_id}",
            token=token,
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
                "component": "writer",
                "clientId": client_id,
            },
            token,
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
        claims = {
            component: claim_token(
                bridge, component, f"http://127.0.0.1:{port}"
            )
            for component, port in (
                ("writer", 3891),
                ("presentation", 3892),
                ("spreadsheet", 3893),
            )
        }
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
                {"component": component, "clientId": claims[component][0]},
                claims[component][1],
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
