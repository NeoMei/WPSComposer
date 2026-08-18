from __future__ import annotations

import hashlib
import hmac
import json
import secrets
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from queue import Empty, Queue
from threading import Condition, Thread
from time import monotonic
from typing import Any, Mapping, Optional
from urllib.parse import parse_qs, urlsplit

from .models import (
    COMPONENTS,
    ProbeCommand,
    ProbeResult,
    ProtocolError,
)

MAX_BODY_BYTES = 1024 * 1024
CLEANUP_GRACE_SECONDS = 5.0
BOOTSTRAP_TTL_SECONDS = 60.0


def derive_client_credentials(root_token: str) -> dict[str, dict[str, str]]:
    """Derive per-component bootstrap claims from a non-served root secret."""
    secret = root_token.encode("utf-8")
    credentials: dict[str, dict[str, str]] = {}
    for component in COMPONENTS:
        client_id = hmac.new(
            secret, f"client:{component}".encode("utf-8"), hashlib.sha256
        ).hexdigest()
        capability = hmac.new(
            secret,
            f"bootstrap:{component}:{client_id}".encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()
        credentials[component] = {
            "clientId": client_id,
            "capability": capability,
        }
    return credentials


class BridgeState:
    """Thread-safe command queues and results for one probe session."""

    def __init__(
        self,
        *,
        session_nonce: Optional[str] = None,
        root_token: Optional[str] = None,
        bootstrap_ttl: float = BOOTSTRAP_TTL_SECONDS,
        lease_seconds: float = 15.0,
        clock=monotonic,
    ):
        if lease_seconds <= 0:
            raise ValueError("lease_seconds must be positive")
        self.session_nonce = session_nonce or secrets.token_urlsafe(24)
        self.lease_seconds = float(lease_seconds)
        self._clock = clock
        self._queues = {component: Queue() for component in COMPONENTS}
        self._registered: set[str] = set()
        self._issued: dict[str, str] = {}
        self._results: dict[str, ProbeResult] = {}
        self._completed: dict[str, ProbeResult] = {}
        self._canceled: set[str] = set()
        self._last_seen: dict[str, float] = {}
        self._condition = Condition()
        self._bootstrap_deadline = self._clock() + float(bootstrap_ttl)
        self._bootstrap = (
            derive_client_credentials(root_token) if root_token is not None else {}
        )
        self._client_tokens: dict[str, tuple[str, str]] = {}

    def claim_session(
        self, component: str, client_id: str, capability: str
    ) -> str:
        """Consume one short-lived component/client capability."""
        self._require_component(component)
        expected = self._bootstrap.get(component)
        if expected is None or self._clock() > self._bootstrap_deadline:
            raise ProtocolError("Bridge bootstrap capability is unavailable or expired")
        try:
            matched_client = secrets.compare_digest(
                client_id, expected["clientId"]
            )
            matched_capability = secrets.compare_digest(
                capability, expected["capability"]
            )
        except TypeError:
            matched_client = matched_capability = False
        if not matched_client or not matched_capability:
            raise ProtocolError("Invalid bridge bootstrap capability")
        del self._bootstrap[component]
        token = secrets.token_urlsafe(32)
        self._client_tokens[token] = (component, client_id)
        return token

    def authorize_client(self, token: str) -> tuple[str, str]:
        try:
            return self._client_tokens[token]
        except KeyError as exc:
            raise ProtocolError("A valid client bearer token is required") from exc

    @staticmethod
    def require_client_binding(
        binding: tuple[str, str], component: str, client_id: str
    ) -> None:
        try:
            component_matches = secrets.compare_digest(binding[0], component)
            client_matches = secrets.compare_digest(binding[1], client_id)
        except TypeError:
            component_matches = client_matches = False
        if not component_matches or not client_matches:
            raise ProtocolError("Bearer token is not bound to this component/client")

    def register(self, component: str, session_nonce: str) -> None:
        """Create or renew a lease after bearer-token authentication."""
        self._require_component(component)
        self._require_session(session_nonce)
        with self._condition:
            self._registered.add(component)
            self._last_seen[component] = self._clock()
            self._condition.notify_all()

    def issue(
        self, component: str, method: str, params: Mapping[str, Any]
    ) -> ProbeCommand:
        command = ProbeCommand.create(component, method, params)
        with self._condition:
            if component not in self._fresh_registered_locked():
                raise ProtocolError(
                    f"Component has no fresh client lease: {component}"
                )
            self._issued[command.id] = component
        self._queues[component].put(command)
        return command

    def next(
        self, component: str, session_nonce: str, timeout: float = 10.0
    ) -> Optional[ProbeCommand]:
        """Renew the authenticated polling client's lease and await work."""
        self._require_component(component)
        self._require_session(session_nonce)
        with self._condition:
            self._registered.add(component)
            self._last_seen[component] = self._clock()
            self._condition.notify_all()
        try:
            return self._queues[component].get(timeout=timeout)
        except Empty:
            return None

    def complete(self, result: ProbeResult, session_nonce: str) -> None:
        """Accept a known result and renew its component's client lease."""
        self._require_session(session_nonce)
        with self._condition:
            if result.id in self._canceled:
                return
            component = self._issued.get(result.id)
            if component is None:
                raise ProtocolError(f"Unknown result id: {result.id}")
            existing = self._completed.get(result.id)
            if existing is not None:
                if existing != result:
                    raise ProtocolError(
                        f"Conflicting duplicate result: {result.id}"
                    )
                self._renew_lease_locked(component)
                return
            self._completed[result.id] = result
            self._results[result.id] = result
            self._renew_lease_locked(component)
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
                if command_id in self._canceled:
                    raise TimeoutError(f"Command was canceled: {command_id}")
                remaining = deadline - monotonic()
                if remaining <= 0:
                    raise TimeoutError(
                        f"Timed out waiting for result {command_id}"
                    )
                self._condition.wait(remaining)
            return self._results.pop(command_id)

    def wait_registered(self, expected: set[str], timeout: float) -> None:
        deadline = self._clock() + timeout
        with self._condition:
            while not expected <= self._fresh_registered_locked():
                remaining = deadline - self._clock()
                if remaining <= 0:
                    missing = sorted(expected - self._fresh_registered_locked())
                    raise TimeoutError(
                        f"Timed out waiting for components: {missing}"
                    )
                self._condition.wait(remaining)

    def last_seen(self, component: str) -> Optional[float]:
        self._require_component(component)
        with self._condition:
            return self._last_seen.get(component)

    def registered_components(self) -> set[str]:
        with self._condition:
            return self._fresh_registered_locked()

    def _fresh_registered_locked(self) -> set[str]:
        now = self._clock()
        stale = {
            component
            for component in self._registered
            if now - self._last_seen.get(component, float("-inf"))
            > self.lease_seconds
        }
        self._registered.difference_update(stale)
        for component in stale:
            self._last_seen.pop(component, None)
        return set(self._registered)

    def _renew_lease_locked(self, component: str) -> None:
        self._registered.add(component)
        self._last_seen[component] = self._clock()

    def _require_session(self, session_nonce: str) -> None:
        try:
            matched = secrets.compare_digest(session_nonce, self.session_nonce)
        except TypeError:
            matched = False
        if not matched:
            raise ProtocolError("Stale or invalid bridge session nonce")

    @staticmethod
    def _require_component(component: str) -> None:
        if component not in COMPONENTS:
            raise ProtocolError(f"Unsupported component: {component}")


def _handler_class(state: BridgeState, allowed_origins: frozenset[str]):
    class BridgeHandler(BaseHTTPRequestHandler):
        server_version = "WPSComposerProbe/1"

        def do_OPTIONS(self) -> None:
            if not self._validate_origin():
                return
            self._send(204)

        def do_GET(self) -> None:
            binding = self._authorize()
            if binding is None:
                return
            if not self._validate_origin():
                return
            parsed = urlsplit(self.path)
            if parsed.path == "/health":
                self._send_json(200, {"status": "ok"})
                return
            if parsed.path != "/v1/next":
                self._send_error_json(404, "NOT_FOUND", "Unknown endpoint")
                return
            query = parse_qs(parsed.query)
            component = query.get("component", [""])[0]
            client_id = query.get("clientId", [""])[0]
            try:
                state.require_client_binding(binding, component, client_id)
                command = state.next(component, state.session_nonce)
            except ProtocolError as exc:
                self._send_error_json(400, "INVALID_COMPONENT", str(exc))
                return
            if command is None:
                self._send(204)
            else:
                self._send_json(200, command.to_dict())

        def do_POST(self) -> None:
            if self.path == "/v1/session":
                # Unauthenticated but loopback-only and origin-checked.
                if not self._validate_origin():
                    return
                try:
                    body = self._read_json()
                    client_token = state.claim_session(
                        str(body.get("component", "")),
                        str(body.get("clientId", "")),
                        str(body.get("capability", "")),
                    )
                except (ProtocolError, ValueError) as exc:
                    self._send_error_json(400, "INVALID_REQUEST", str(exc))
                    return
                self._send_json(
                    200,
                    {"token": client_token, "leaseSeconds": state.lease_seconds},
                )
                return
            binding = self._authorize()
            if binding is None:
                return
            if not self._validate_origin():
                return
            try:
                body = self._read_json()
                component = str(body.get("component", binding[0]))
                client_id = str(body.get("clientId", ""))
                state.require_client_binding(binding, component, client_id)
                if self.path == "/v1/register":
                    state.register(component, state.session_nonce)
                elif self.path == "/v1/result":
                    body.pop("component", None)
                    body.pop("clientId", None)
                    result = ProbeResult.from_dict(body)
                    issued_component = state._issued.get(result.id)
                    if issued_component != component:
                        raise ProtocolError("Result is not bound to this component/client")
                    state.complete(result, state.session_nonce)
                else:
                    self._send_error_json(
                        404, "NOT_FOUND", "Unknown endpoint"
                    )
                    return
            except (ProtocolError, ValueError) as exc:
                self._send_error_json(400, "INVALID_REQUEST", str(exc))
                return
            self._send(204)

        def _authorize(self) -> Optional[tuple[str, str]]:
            supplied = self.headers.get("Authorization", "")
            prefix = "Bearer "
            if not supplied.startswith(prefix):
                self._send_error_json(
                    401, "UNAUTHORIZED", "A valid bearer token is required"
                )
                return None
            try:
                return state.authorize_client(supplied[len(prefix):])
            except ProtocolError:
                self._send_error_json(
                    401, "UNAUTHORIZED", "A valid bearer token is required"
                )
                return None

        def _validate_origin(self) -> bool:
            origin = self.headers.get("Origin", "")
            if origin not in allowed_origins:
                self._send_error_json(
                    403, "ORIGIN_REJECTED", "Origin is not allowed"
                )
                return False
            return True

        def _read_json(self) -> dict[str, Any]:
            raw_length = self.headers.get("Content-Length", "0")
            try:
                length = int(raw_length)
            except ValueError as exc:
                raise ValueError("Invalid Content-Length") from exc
            if length < 0 or length > MAX_BODY_BYTES:
                raise ValueError("Request body is too large")
            raw = self.rfile.read(length)
            try:
                body = json.loads(raw or b"{}")
            except json.JSONDecodeError as exc:
                raise ValueError("Request body must be valid JSON") from exc
            if not isinstance(body, dict):
                raise ValueError("Request body must be a JSON object")
            return body

        def _send_error_json(
            self, status: int, code: str, message: str
        ) -> None:
            self._send_json(
                status, {"error": {"code": code, "message": message}}
            )

        def _send_json(self, status: int, payload: Mapping[str, Any]) -> None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self._send(status, data, "application/json; charset=utf-8")

        def _send(
            self,
            status: int,
            data: bytes = b"",
            content_type: Optional[str] = None,
        ) -> None:
            self.send_response(status)
            origin = self.headers.get("Origin", "")
            if origin in allowed_origins:
                self.send_header("Access-Control-Allow-Origin", origin)
                self.send_header("Vary", "Origin")
                self.send_header(
                    "Access-Control-Allow-Headers",
                    "Authorization, Content-Type",
                )
                self.send_header(
                    "Access-Control-Allow-Methods", "GET, POST, OPTIONS"
                )
            if content_type:
                self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            if data:
                self.wfile.write(data)

        def log_message(self, format: str, *args: Any) -> None:
            return

    return BridgeHandler


class LoopbackBridge:
    """Authenticated HTTP bridge bound to an ephemeral loopback port."""

    def __init__(self, allowed_origins: set[str]):
        if not allowed_origins:
            raise ValueError("At least one allowed origin is required")
        self.token = secrets.token_urlsafe(32)
        self.session_nonce = hashlib.sha256(
            self.token.encode("utf-8")
        ).hexdigest()
        self.state = BridgeState(
            session_nonce=self.session_nonce, root_token=self.token
        )
        origins = frozenset(allowed_origins)
        self._server = ThreadingHTTPServer(
            ("127.0.0.1", 0),
            _handler_class(self.state, origins),
        )
        self._server.daemon_threads = True
        host, port = self._server.server_address
        self.url = f"http://{host}:{port}"
        self._thread: Optional[Thread] = None

    def bootstrap_credentials(self, component: str) -> dict[str, str]:
        try:
            return dict(derive_client_credentials(self.token)[component])
        except KeyError as exc:
            raise ValueError(f"Unknown component: {component}") from exc

    def __enter__(self) -> "LoopbackBridge":
        self._thread = Thread(
            target=self._server.serve_forever,
            name="wpscomposer-probe-bridge",
            daemon=True,
        )
        self._thread.start()
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        shutdown = Thread(
            target=self._server.shutdown,
            name="wpscomposer-bridge-shutdown",
            daemon=True,
        )
        shutdown.start()
        shutdown.join(timeout=CLEANUP_GRACE_SECONDS)
        self._server.server_close()
        if self._thread is not None:
            self._thread.join(timeout=0)

    def issue(
        self, component: str, method: str, params: Mapping[str, Any]
    ) -> ProbeCommand:
        return self.state.issue(component, method, params)

    def wait_result(self, command_id: str, timeout: float) -> ProbeResult:
        return self.state.wait_result(command_id, timeout)

    def wait_registered(self, expected: set[str], timeout: float) -> None:
        self.state.wait_registered(expected, timeout)

    def registered_components(self) -> set[str]:
        return self.state.registered_components()
