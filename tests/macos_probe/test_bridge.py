import json
from concurrent.futures import ThreadPoolExecutor
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pytest

from skills.WPSComposer.scripts.macos_probe.bridge import LoopbackBridge
from skills.WPSComposer.scripts.macos_probe.models import ProbeResult, ProtocolError


def request(bridge, method, path, body=None, token=None, origin=None):
    headers = {"Content-Type": "application/json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if origin:
        headers["Origin"] = origin
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
        first = bridge.issue("writer", "smoke_docx", {})
        result = ProbeResult(first.id, True, {"saved": True}, None)
        bridge.state.complete(result)
        bridge.state.complete(result)
        assert bridge.wait_result(first.id, 1) == result

        second = bridge.issue("writer", "smoke_pdf", {})
        bridge.state.cancel(second.id)
        bridge.state.complete(ProbeResult(second.id, True, {"saved": True}, None))
        with pytest.raises(TimeoutError):
            bridge.wait_result(second.id, 0.01)


def test_conflicting_duplicate_result_is_rejected():
    with LoopbackBridge({"http://127.0.0.1:3891"}) as bridge:
        command = bridge.issue("writer", "smoke_docx", {})
        bridge.state.complete(ProbeResult(command.id, True, {"saved": True}, None))
        with pytest.raises(ProtocolError, match="Conflicting duplicate"):
            bridge.state.complete(
                ProbeResult(command.id, False, {}, {"code": "failed"})
            )
