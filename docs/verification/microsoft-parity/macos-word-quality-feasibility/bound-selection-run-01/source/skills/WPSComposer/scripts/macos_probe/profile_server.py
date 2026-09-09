from __future__ import annotations

from functools import partial
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import mimetypes
from pathlib import Path
import socket
import threading
from typing import Optional
from urllib.parse import unquote, urlsplit


_CONTENT_TYPES = {
    ".html": "text/html",
    ".js": "text/javascript",
    ".json": "application/json",
}


class _LoopbackHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True
    daemon_threads = False

    def __init__(self, *args, **kwargs) -> None:
        self._connection_lock = threading.Lock()
        self._connections: set[socket.socket] = set()
        super().__init__(*args, **kwargs)

    def process_request(self, request: socket.socket, client_address) -> None:
        with self._connection_lock:
            self._connections.add(request)
        try:
            super().process_request(request, client_address)
        except BaseException:
            with self._connection_lock:
                self._connections.discard(request)
            raise

    def shutdown_request(self, request: socket.socket) -> None:
        try:
            super().shutdown_request(request)
        finally:
            with self._connection_lock:
                self._connections.discard(request)

    def close_connections(self) -> None:
        with self._connection_lock:
            connections = tuple(self._connections)
        for connection in connections:
            try:
                connection.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass

    def handle_error(self, request, client_address) -> None:
        # Closing an active handler can interrupt a read or write.  Keep the
        # profile service's no-request-logging contract during cleanup too.
        return None


class _ProfileRequestHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def __init__(self, *args, profile_root: Path, **kwargs) -> None:
        self._profile_root = profile_root
        super().__init__(*args, **kwargs)

    def log_message(self, format: str, *args: object) -> None:
        # Profile URLs can include cache keys and the profile contains private
        # bridge credentials.  This service never emits request data.
        return None

    def _target(self) -> Optional[Path]:
        parsed = urlsplit(self.path)
        if parsed.scheme or parsed.netloc:
            return None
        try:
            decoded = unquote(parsed.path, errors="strict")
        except UnicodeDecodeError:
            return None
        if not decoded.startswith("/") or "\x00" in decoded or "\\" in decoded:
            return None
        relative = decoded[1:] or "index.html"
        parts = relative.split("/")
        if any(part in {"", ".", ".."} for part in parts):
            return None
        candidate = self._profile_root.joinpath(*parts)
        try:
            target = candidate.resolve(strict=True)
            target.relative_to(self._profile_root)
        except (FileNotFoundError, OSError, ValueError):
            return None
        if not target.is_file():
            return None
        return target

    def _reject(self, status: int, *, allow: Optional[str] = None) -> None:
        self.send_response(status)
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", "0")
        if allow is not None:
            self.send_header("Allow", allow)
        self.end_headers()

    def _serve(self, *, include_body: bool) -> None:
        target = self._target()
        if target is None:
            self._reject(404)
            return
        try:
            payload = target.read_bytes()
        except OSError:
            self._reject(404)
            return
        content_type = _CONTENT_TYPES.get(target.suffix.lower())
        if content_type is None:
            content_type = (
                mimetypes.guess_type(target.name)[0]
                or "application/octet-stream"
            )
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        if include_body:
            self.wfile.write(payload)

    def do_GET(self) -> None:
        self._serve(include_body=True)

    def do_HEAD(self) -> None:
        self._serve(include_body=False)

    def _reject_write(self) -> None:
        self._reject(405, allow="GET, HEAD")

    do_POST = _reject_write
    do_PUT = _reject_write
    do_PATCH = _reject_write
    do_DELETE = _reject_write


class ProfileServer:
    """Serve one generated WPS profile from an IPv4 loopback socket."""

    def __init__(self, profile_root: Path, port: int) -> None:
        if isinstance(port, bool) or not isinstance(port, int):
            raise TypeError("port must be an integer")
        if port < 0 or port > 65535:
            raise ValueError("port must be between 0 and 65535")
        self.profile_root = Path(profile_root).expanduser().resolve()
        self.port = port
        self._server: Optional[_LoopbackHTTPServer] = None
        self._thread: Optional[threading.Thread] = None

    @property
    def address(self) -> tuple[str, int]:
        return ("127.0.0.1", self.port)

    def start(self) -> "ProfileServer":
        if self._server is not None:
            return self
        if not self.profile_root.is_dir():
            raise RuntimeError(
                f"WPS profile directory is unavailable: {self.profile_root}"
            )
        handler = partial(_ProfileRequestHandler, profile_root=self.profile_root)
        server = _LoopbackHTTPServer(("127.0.0.1", self.port), handler)
        thread = threading.Thread(
            target=server.serve_forever,
            name=f"wpscomposer-profile-{server.server_port}",
            daemon=True,
        )
        try:
            thread.start()
        except BaseException:
            server.server_close()
            raise
        self._server = server
        self._thread = thread
        self.port = int(server.server_port)
        return self

    def close(self) -> None:
        server = self._server
        thread = self._thread
        self._server = None
        self._thread = None
        if server is None:
            return
        try:
            server.shutdown()
        finally:
            try:
                server.close_connections()
            finally:
                server.server_close()
                if thread is not None:
                    thread.join()


__all__ = ["ProfileServer"]
