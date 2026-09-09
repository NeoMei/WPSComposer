"""At most ten isolated loopback requests; no production monkeypatch or native app."""
from collections import Counter
import hashlib
import http.client
import json
from pathlib import Path
import time

from skills.WPSComposer.scripts.macos_probe import bridge as module


def main():
    print(json.dumps({"source_sha256": hashlib.sha256(Path(module.__file__).read_bytes()).hexdigest()}), flush=True)
    outcomes = Counter()
    origin = "http://127.0.0.1:3891"
    with module.LoopbackBridge({origin}) as bridge:
        credentials = bridge.bootstrap_credentials("writer")
        token = bridge.state.claim_session("writer", credentials["clientId"], credentials["capability"])
        small = json.dumps({"component": "writer", "clientId": credentials["clientId"], "padding": ""}).encode()
        size = module.MAX_BODY_BYTES + 1
        large = small[:-2] + b"x" * (size - len(small)) + small[-2:]
        assert len(large) == size and isinstance(json.loads(large), dict)
        cases = [("invalid-small", "invalid", small)] * 2
        cases += [("oversize-body", str(size), large)] * 4
        cases += [("invalid-large", "invalid", large)] * 4
        assert len(cases) == 10
        for index, (case, length, body) in enumerate(cases, 1):
            connection = http.client.HTTPConnection(*bridge._server.server_address, timeout=1)
            result = {"request": index, "case": case, "declared_length": length, "body_bytes": len(body)}
            start = time.monotonic()
            phase = "headers"
            try:
                connection.putrequest("POST", "/v1/register")
                connection.putheader("Origin", origin)
                connection.putheader("Authorization", "Bearer " + token)
                connection.putheader("Content-Type", "application/json")
                connection.putheader("Content-Length", length)
                connection.endheaders()
                phase = "send-body"
                connection.send(body)
                phase = "read-status"
                response = connection.getresponse()
                phase = "read-error-body"
                payload = json.loads(response.read())
                assert response.status == 400
                assert payload["error"]["code"] == "INVALID_REQUEST"
                result.update(outcome="HTTP400", payload=payload, connection=response.getheader("Connection"))
            except (OSError, http.client.HTTPException) as exc:
                result.update(outcome=type(exc).__name__, errno=getattr(exc, "errno", None), message=str(exc), phase=phase)
            finally:
                connection.close()
            result["elapsed_seconds"] = time.monotonic() - start
            outcomes[(case, result["outcome"])] += 1
            assert not bridge.state._registered
            print(json.dumps(result), flush=True)
    print(json.dumps({"summary": [{"case": case, "outcome": outcome, "count": count} for (case, outcome), count in sorted(outcomes.items())], "total_requests": 10, "registered": False}), flush=True)


if __name__ == "__main__":
    main()
