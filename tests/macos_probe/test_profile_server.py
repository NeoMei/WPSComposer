from __future__ import annotations

import http.client
from pathlib import Path
import socket
import threading
import time
from urllib.error import HTTPError
from urllib.request import Request, urlopen

import pytest

from skills.WPSComposer.scripts.macos_probe.profile_server import ProfileServer


def _request(server: ProfileServer, path: str, *, method: str = "GET"):
    return urlopen(
        Request(f"http://127.0.0.1:{server.port}{path}", method=method),
        timeout=2,
    )


@pytest.fixture
def profile(tmp_path: Path):
    root = tmp_path / "profile"
    root.mkdir()
    (root / "index.html").write_text("<main>ready</main>", encoding="utf-8")
    (root / "component.js").write_text("globalThis.ready = true;", encoding="utf-8")
    (root / "session.json").write_text('{"component":"writer"}', encoding="utf-8")
    nested = root / "nested"
    nested.mkdir()
    (nested / "asset.js").write_text("nested", encoding="utf-8")
    server = ProfileServer(root, 0)
    server.start()
    try:
        yield server, root
    finally:
        server.close()


@pytest.mark.parametrize(
    ("path", "body", "content_type"),
    [
        ("/", b"<main>ready</main>", "text/html"),
        ("/component.js", b"globalThis.ready = true;", "text/javascript"),
        ("/session.json", b'{"component":"writer"}', "application/json"),
    ],
)
def test_profile_server_serves_files_without_caching(profile, path, body, content_type):
    server, _ = profile

    with _request(server, path) as response:
        assert response.status == 200
        assert response.read() == body
        assert response.headers.get_content_type() == content_type
        assert response.headers["Cache-Control"] == "no-store"


def test_profile_server_binds_only_ipv4_loopback(profile):
    server, _ = profile

    assert server.address == ("127.0.0.1", server.port)
    connection = http.client.HTTPConnection("127.0.0.1", server.port, timeout=2)
    connection.request("HEAD", "/index.html")
    response = connection.getresponse()
    try:
        assert response.status == 200
        assert response.read() == b""
    finally:
        connection.close()


@pytest.mark.parametrize(
    "path",
    [
        "/../outside.txt",
        "/%2e%2e/outside.txt",
        "/%2e%2e%2foutside.txt",
        "/nested",
        "/nested/",
        "/missing.txt",
    ],
)
def test_profile_server_rejects_traversal_directories_and_missing_files(profile, path):
    server, root = profile
    (root.parent / "outside.txt").write_text("secret", encoding="utf-8")

    with pytest.raises(HTTPError) as caught:
        _request(server, path)

    assert caught.value.code == 404


def test_profile_server_rejects_symlink_escape(profile):
    server, root = profile
    outside = root.parent / "outside.json"
    outside.write_text('{"secret":true}', encoding="utf-8")
    (root / "escape.json").symlink_to(outside)

    with pytest.raises(HTTPError) as caught:
        _request(server, "/escape.json")

    assert caught.value.code == 404


@pytest.mark.parametrize("method", ["POST", "PUT", "PATCH", "DELETE"])
def test_profile_server_rejects_writes_without_changing_profile(profile, method):
    server, root = profile
    before = (root / "session.json").read_bytes()
    connection = http.client.HTTPConnection("127.0.0.1", server.port, timeout=2)
    connection.request(method, "/session.json", body=b'{"component":"rogue"}')
    response = connection.getresponse()
    try:
        assert response.status == 405
        assert response.headers["Allow"] == "GET, HEAD"
    finally:
        response.read()
        connection.close()
    assert (root / "session.json").read_bytes() == before


def test_profile_server_does_not_log_request_data(profile, capsys):
    server, _ = profile

    with pytest.raises(HTTPError):
        _request(server, "/missing?capability=private-secret")

    captured = capsys.readouterr()
    assert "private-secret" not in captured.out
    assert "private-secret" not in captured.err


def test_profile_server_close_is_idempotent(profile):
    server, _ = profile

    server.close()
    server.close()


def test_profile_server_close_terminates_existing_keep_alive_connection(
    tmp_path: Path,
):
    root = tmp_path / "profile"
    root.mkdir()
    (root / "index.html").write_text("ready", encoding="utf-8")
    server = ProfileServer(root, 0).start()
    connection = http.client.HTTPConnection("127.0.0.1", server.port, timeout=2)
    try:
        connection.request("GET", "/index.html")
        response = connection.getresponse()
        assert response.status == 200
        assert response.read() == b"ready"

        server.close()

        with pytest.raises((OSError, http.client.HTTPException)):
            connection.request("GET", "/index.html")
            connection.getresponse()
    finally:
        connection.close()
        server.close()


@pytest.mark.parametrize(
    "request_prefix",
    [b"", b"GET /index.html HTTP/1.1\r\nHost: 127.0.0.1"],
    ids=["idle", "partial-request"],
)
def test_profile_server_close_releases_request_handler(
    tmp_path: Path, request_prefix: bytes
):
    root = tmp_path / "profile"
    root.mkdir()
    (root / "index.html").write_text("ready", encoding="utf-8")
    server = ProfileServer(root, 0).start()
    existing_threads = set(threading.enumerate())
    connection = socket.create_connection(server.address, timeout=2)
    try:
        if request_prefix:
            connection.sendall(request_prefix)
        deadline = time.monotonic() + 2
        handler = None
        while time.monotonic() < deadline:
            candidates = [
                thread
                for thread in threading.enumerate()
                if thread not in existing_threads
                and thread.is_alive()
            ]
            if candidates:
                handler = candidates[0]
                break
            time.sleep(0.01)
        assert handler is not None
        assert handler.is_alive()

        server.close()

        handler.join(timeout=0.1)
        assert not handler.is_alive()
        assert connection.recv(1) == b""
    finally:
        connection.close()
        server.close()
