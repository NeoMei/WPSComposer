from __future__ import annotations

import json
import secrets
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from queue import Empty, Queue
from threading import Condition
from time import monotonic
from typing import Any, Mapping, Optional, Set
from urllib.parse import parse_qs, urlparse

from .models import COMPONENTS, ProbeCommand, ProbeResult, ProtocolError


class BridgeState:
    def __init__(self):
        self._queues = {component: Queue() for component in COMPONENTS}
        self._registered: Set[str] = set()
        self._results: dict[str, ProbeResult] = {}
        self._canceled: Set[str] = set()
        self._last_seen: dict[str, float] = {}
        self._condition = Condition()

    def register(self, component: str) -> None:
        if component not in COMPONENTS:
            raise ProtocolError(f"Unsupported component: {component}")
        with self._condition:
            self._registered.add(component)
            self._last_seen[component] = monotonic()
            self._condition.notify_all()

    def issue(self, component: str, method: str, params: Mapping[str, Any]) -> ProbeCommand:
        command = ProbeCommand.create(component, method, params)
        self._queues[component].put(command)
        return command

    def next(self, component: str, timeout: float = 10.0) -> Optional[ProbeCommand]:
        if component not in COMPONENTS:
            raise ProtocolError(f"Unsupported component: {component}")
        with self._condition:
            self._last_seen[component] = monotonic()
        try:
            return self._queues[component].get(timeout=timeout)
        except Empty:
            return None

    def complete(self, result: ProbeResult) -> None:
        with self._condition:
            if result.id in self._canceled:
                return
            existing = self._results.get(result.id)
            if existing is not None and existing != result:
                raise ProtocolError(f"Conflicting duplicate result: {result.id}")
            self._results[result.id] = result
            self._condition.notify_all()

    def cancel(self, command_id: str) -> None:
        with self._condition:
            self._canceled.add(command_id)
            self._results.pop(command_id, None)
            self._condition.notify_all()

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        deadline = monotonic() + timeout
        with self._condition:
            while command_id not in self._results:
                remaining = deadline - monotonic()
                if remaining <= 0:
                    raise TimeoutError(f"Timed out waiting for result {command_id}")
                self._condition.wait(remaining)
            return self._results.pop(command_id)

    def wait_registered(self, expected: Set[str], timeout: float) -> None:
        deadline = monotonic() + timeout
        with self._condition:
            while not expected <= self._registered:
                remaining = deadline - monotonic()
                if remaining <= 0:
                    missing = sorted(expected - self._registered)
                    raise TimeoutError(f"Timed out waiting for components: {missing}")
                self._condition.wait(remaining)

    def last_seen(self, component: str) -> Optional[float]:
        with self._condition:
            return self._last_seen.get(component)


class LoopbackBridge:
    def __init__(self, allowed_origins: Set[str]):
        self.state = BridgeState()
        self.token = secrets.token_urlsafe(32)
        self._allowed_origins = allowed_origins
        self._server: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None

    @property
    def url(self) -> str:
        host, port = self._server.server_address
        return f"http://{host}:{port}"

    def issue(self, component: str, method: str, params: Mapping[str, Any]) -> ProbeCommand:
        return self.state.issue(component, method, params)

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        return self.state.wait_result(command_id, timeout)

    def wait_registered(self, expected: Set[str], timeout: float) -> None:
        self.state.wait_registered(expected, timeout)

    def __enter__(self) -> "LoopbackBridge":
        state = self.state
        token = self.token
        allowed = self._allowed_origins

        class Handler(BaseHTTPRequestHandler):
            def log_message(self, format, *args):
                pass

            def _origin_ok(self) -> bool:
                origin = self.headers.get("Origin")
                if not origin:
                    return True
                if "*" in allowed:
                    return origin.startswith("http://127.0.0.1:")
                return origin in allowed

            def _auth_ok(self) -> bool:
                auth = self.headers.get("Authorization", "")
                if not auth.startswith("Bearer "):
                    return False
                return secrets.compare_digest(auth[7:], token)

            def _send_json(self, code: int, body: Any = None):
                payload = b"" if body is None else json.dumps(body).encode()
                self.send_response(code)
                origin = self.headers.get("Origin")
                if origin and (origin in allowed or ("*" in allowed and origin.startswith("http://127.0.0.1:"))):
                    self.send_header("Access-Control-Allow-Origin", origin)
                    self.send_header("Vary", "Origin")
                    self.send_header("Access-Control-Allow-Headers", "Authorization, Content-Type")
                    self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                if payload:
                    self.wfile.write(payload)

            def _send_error(self, code: int, err_code: str, message: str):
                self._send_json(code, {"error": {"code": err_code, "message": message}})

            def do_OPTIONS(self):
                if not self._origin_ok():
                    self._send_error(403, "forbidden_origin", "Origin not allowed")
                    return
                self._send_json(204)

            def do_GET(self):
                if not self._origin_ok():
                    self._send_error(403, "forbidden_origin", "Origin not allowed")
                    return
                if not self._auth_ok():
                    self._send_error(401, "unauthorized", "Missing or invalid token")
                    return
                parsed = urlparse(self.path)
                if parsed.path == "/health":
                    self._send_json(200, {"status": "ok"})
                    return
                if parsed.path == "/v1/next":
                    params = parse_qs(parsed.query)
                    component = params.get("component", [""])[0]
                    if component not in COMPONENTS:
                        self._send_error(400, "bad_component", f"Unknown component: {component}")
                        return
                    command = state.next(component, timeout=10.0)
                    if command is None:
                        self._send_json(204)
                    else:
                        self._send_json(200, command.to_dict())
                    return
                self._send_error(404, "not_found", "Unknown path")

            def do_POST(self):
                if not self._origin_ok():
                    self._send_error(403, "forbidden_origin", "Origin not allowed")
                    return
                if not self._auth_ok():
                    self._send_error(401, "unauthorized", "Missing or invalid token")
                    return
                parsed = urlparse(self.path)
                try:
                    length = int(self.headers.get("Content-Length", 0))
                except (ValueError, TypeError):
                    length = 0
                if length <= 0:
                    raw = b""
                else:
                    raw = self.rfile.read(min(length, 1048576))
                try:
                    body = json.loads(raw) if raw else {}
                except (json.JSONDecodeError, UnicodeDecodeError):
                    self._send_error(400, "bad_json", "Invalid JSON body")
                    return

                if parsed.path == "/v1/register":
                    component = body.get("component", "")
                    try:
                        state.register(component)
                    except ProtocolError as exc:
                        self._send_error(400, "bad_component", str(exc))
                        return
                    self._send_json(204)
                    return
                if parsed.path == "/v1/result":
                    try:
                        result = ProbeResult.from_dict(body)
                        state.complete(result)
                    except ProtocolError as exc:
                        self._send_error(400, "protocol_error", str(exc))
                        return
                    self._send_json(204)
                    return
                self._send_error(404, "not_found", "Unknown path")

        self._server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()
        return self

    def __exit__(self, *exc):
        if self._server:
            self._server.shutdown()
            self._server.server_close()
        if self._thread:
            self._thread.join(timeout=5)
