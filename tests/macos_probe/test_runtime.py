import json
import signal
import stat
from pathlib import Path

import pytest

from skills.WPSComposer.scripts.macos_probe import runtime
from skills.WPSComposer.scripts.macos_probe.runtime import (
    COMPONENT_CONFIG,
    RegistrationSnapshot,
    build_profile,
    find_node,
)


def test_component_config_uses_distinct_ports_and_wps_types():
    assert COMPONENT_CONFIG == {
        "writer": {
            "addon_type": "wps",
            "port": 3889,
            "script": "writer.js",
        },
        "presentation": {
            "addon_type": "wpp",
            "port": 3890,
            "script": "presentation.js",
        },
        "spreadsheet": {
            "addon_type": "et",
            "port": 3891,
            "script": "spreadsheet.js",
        },
    }


def test_build_profile_writes_runtime_config(tmp_path: Path):
    assets = tmp_path / "assets"
    assets.mkdir()
    for name in (
        "index.html",
        "manifest.xml",
        "ribbon.xml",
        "bridge-client.js",
        "writer.js",
    ):
        (assets / name).write_text(name, encoding="utf-8")

    profile = build_profile(
        assets,
        tmp_path / "profiles",
        "writer",
        "http://127.0.0.1:45678",
        "secret-token",
    )

    package = json.loads((profile / "package.json").read_text())
    session = json.loads((profile / "session.json").read_text())
    assert package["addonType"] == "wps"
    assert session == {
        "bridgeUrl": "http://127.0.0.1:45678",
        "component": "writer",
        "token": "secret-token",
    }
    assert (profile / "component.js").read_text() == "writer.js"


def test_registration_snapshot_restores_existing_bytes(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    original = b"<jsplugins><original/></jsplugins>\n"
    publish.write_bytes(original)
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    publish.write_bytes(b"changed")
    snapshot.restore()
    assert publish.read_bytes() == original
    assert not (tmp_path / "recovery").exists()


def test_registration_snapshot_removes_probe_created_file(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    publish.write_bytes(b"probe")
    snapshot.restore()
    assert not publish.exists()
    assert not (tmp_path / "recovery").exists()


def test_find_node_honors_explicit_override(tmp_path: Path):
    node = tmp_path / "node"
    node.write_text("#!/bin/sh\nprintf 'v24.0.0\\n'\n")
    node.chmod(0o755)
    assert find_node(str(node)) == node.resolve()


def test_probe_toolchain_is_pinned_and_audit_override_is_reviewed():
    package = json.loads(
        Path("macos/wps-jsapi-probe/package.json").read_text(encoding="utf-8")
    )
    assert package["devDependencies"] == {
        "wps-jsapi-declare": "2.2.0",
        "wpsjs": "2.2.3",
    }
    assert package["overrides"] == {"tmp": "0.2.7"}


def test_component_activation_uses_a_fresh_wps_instance(tmp_path: Path):
    fixture = tmp_path / "fixture.pptx"
    command = runtime.activation_command(
        Path("/Applications/wpsoffice.app"), fixture
    )
    assert command == [
        "open",
        "-n",
        "-a",
        "/Applications/wpsoffice.app",
        str(fixture),
    ]


def test_owned_wps_pids_only_returns_processes_started_after_snapshot():
    assert runtime.owned_wps_pids({101, 102}, {102, 201, 202}) == {201, 202}


def test_default_staging_root_is_inside_wps_container():
    assert str(runtime.WPS_STAGING_ROOT).endswith(
        "Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer"
    )


def test_create_staging_session_is_private(tmp_path: Path):
    session = runtime.create_staging_session(tmp_path / "WPSComposer")
    try:
        assert session.parent == (tmp_path / "WPSComposer").resolve()
        assert stat.S_IMODE(session.stat().st_mode) == 0o700
        assert session.name.startswith("session-")
    finally:
        session.rmdir()


def test_runtime_removes_staging_session_after_success(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    monkeypatch.setattr(runtime.ProbeRuntime, "_preflight", lambda self: None)
    monkeypatch.setattr(runtime, "list_wps_pids", lambda app: set())
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
        staging_root=tmp_path / "container" / "WPSComposer",
    )

    with probe:
        session = probe.staging_dir
        assert session is not None
        (session / "artifact.pdf").write_bytes(b"data")

    assert not session.exists()


def test_runtime_removes_staging_session_after_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    monkeypatch.setattr(runtime.ProbeRuntime, "_preflight", lambda self: None)
    monkeypatch.setattr(runtime, "list_wps_pids", lambda app: set())
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
        staging_root=tmp_path / "container" / "WPSComposer",
    )

    with pytest.raises(RuntimeError, match="forced failure"):
        with probe:
            session = probe.staging_dir
            assert session is not None
            (session / "artifact.pdf").write_bytes(b"data")
            raise RuntimeError("forced failure")

    assert not session.exists()


def test_list_wps_pids_only_matches_exact_main_executable(monkeypatch):
    class Result:
        stdout = """\
 101 /Applications/wpsoffice.app/Contents/MacOS/wpsoffice
 102 /Applications/wpsoffice.app/Contents/MacOS/wpsoffice --flag
 103 /Applications/wpsoffice.app/Contents/Frameworks/wpsoffice helper
"""

    monkeypatch.setattr(runtime.subprocess, "run", lambda *args, **kwargs: Result())

    assert runtime.list_wps_pids(Path("/Applications/wpsoffice.app")) == {101}


def test_runtime_close_terminates_only_wps_processes_created_by_probe(
    monkeypatch, tmp_path: Path
):
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    probe._wps_pids_before = {101}
    observed = iter(({101, 201, 202}, {101}))
    monkeypatch.setattr(runtime, "list_wps_pids", lambda app: set(next(observed)))
    signals = []
    monkeypatch.setattr(
        runtime.os,
        "kill",
        lambda pid, action: signals.append((pid, action)),
    )

    probe.close()

    assert signals == [(201, signal.SIGTERM), (202, signal.SIGTERM)]


def test_activate_component_can_retry_one_missing_component(
    monkeypatch, tmp_path: Path
):
    probe_root = tmp_path / "probe"
    resource_dir = probe_root / "node_modules/wpsjs/src/lib/res"
    resource_dir.mkdir(parents=True)
    (resource_dir / "wpsDemo.docx").write_bytes(b"fixture")
    probe = runtime.ProbeRuntime(
        probe_root,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    probe.staging_dir = tmp_path / "container" / "session-1"
    probe.staging_dir.mkdir(parents=True)
    commands = []
    monkeypatch.setattr(
        runtime.subprocess,
        "run",
        lambda command, **kwargs: commands.append(command),
    )

    first = probe.activate_component("writer")
    second = probe.activate_component("writer")

    assert first == second
    assert first.parent == probe.staging_dir / "fixtures"
    assert first.read_bytes() == b"fixture"
    assert commands == [
        runtime.activation_command(probe.wps_app, first),
        runtime.activation_command(probe.wps_app, first),
    ]
