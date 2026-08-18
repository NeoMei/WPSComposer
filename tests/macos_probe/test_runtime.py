import json
import os
import signal
import stat
import xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import urlsplit

import pytest

from skills.WPSComposer.scripts.macos_probe import runtime
from skills.WPSComposer.scripts.macos_probe.runtime import (
    COMPONENT_CONFIG,
    ProcessIdentity,
    RegistrationSnapshot,
    build_profile,
    find_node,
    read_configured_node,
)

# ponytail: these tests exercise POSIX-only probe behaviour (fcntl, ps, POSIX perms)
posix_only = pytest.mark.skipif(
    os.name != "posix", reason="macOS probe uses POSIX-only runtime facilities"
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
        "client-a",
        "capability-a",
    )

    package = json.loads((profile / "package.json").read_text())
    session = json.loads((profile / "session.json").read_text())
    assert package["addonType"] == "wps"
    assert session == {
        "bridgeUrl": "http://127.0.0.1:45678",
        "component": "writer",
        "clientId": "client-a",
        "capability": "capability-a",
    }
    assert "nonce" not in (profile / "session.json").read_text().lower()
    assert (profile / "component.js").read_text() == "writer.js"


def test_registration_snapshot_restores_existing_bytes(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    original = b"<jsplugins><original/></jsplugins>\n"
    publish.write_bytes(original)
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-a",
        client_credentials=runtime.derive_client_credentials("private-a"),
    )
    snapshot.restore()
    assert publish.read_bytes() == original
    assert not (tmp_path / "recovery").exists()


def test_registration_uses_stable_authorized_origin_without_secret(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    publish.write_bytes(b"<jsplugins/>")
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    credentials = runtime.derive_client_credentials("private-root")

    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="session-id",
        client_credentials=credentials,
    )

    entry = next(iter(ET.parse(publish).getroot()))
    url = urlsplit(entry.attrib["url"])
    assert entry.attrib["name"] == "wpscomposer-phase0-writer"
    assert url.query == ""
    assert url.fragment == ""
    assert url.geturl() == "http://127.0.0.1:3889/"
    assert credentials["writer"]["capability"] not in publish.read_text()
    assert stat.S_IMODE(publish.stat().st_mode) == 0o600
    snapshot.restore()


def test_registration_reuses_authorized_profile_name_without_duplicates(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    publish.write_text(
        '<jsplugins><jspluginonline name="wpscomposer-phase0-writer" '
        'type="wps" url="http://stale/"/></jsplugins>',
        encoding="utf-8",
    )
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    credentials = runtime.derive_client_credentials("private-root")

    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="session-id",
        client_credentials=credentials,
    )

    entries = list(ET.parse(publish).getroot())
    assert [entry.attrib["name"] for entry in entries] == [
        "wpscomposer-phase0-writer"
    ]
    assert entries[0].attrib["url"] == "http://127.0.0.1:3889/"
    snapshot.restore()


def test_registration_replaces_namespaced_authorized_profile_entry(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    publish.write_text(
        '<jsplugins xmlns="urn:wps"><jspluginonline '
        'name="wpscomposer-phase0-writer" type="wps" '
        'url="http://stale/"/></jsplugins>',
        encoding="utf-8",
    )
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")

    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="session-id",
        client_credentials=runtime.derive_client_credentials("private-root"),
    )

    entries = list(ET.parse(publish).getroot())
    assert [entry.attrib["name"] for entry in entries] == [
        "wpscomposer-phase0-writer"
    ]
    snapshot.restore()


def test_capture_recovers_stale_recovery_directory(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    original = b"<jsplugins><original/></jsplugins>\n"
    publish.write_bytes(original)
    first = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    runtime.install_registration_entries(
        first,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-a",
        client_credentials=runtime.derive_client_credentials("private-a"),
    )
    # next run: stale recovery dir is restored instead of failing
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    assert publish.read_bytes() == original
    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-b",
        client_credentials=runtime.derive_client_credentials("private-b"),
    )
    snapshot.restore()
    assert publish.read_bytes() == original
    assert not (tmp_path / "recovery").exists()


def test_stale_prewrite_recovery_preserves_current_external_registration(tmp_path):
    publish = tmp_path / "publish.xml"
    original = b"<jsplugins><original/></jsplugins>"
    external = b'<jsplugins><original/><jspluginonline name="external"/></jsplugins>'
    publish.write_bytes(original)
    recovery = tmp_path / "recovery"
    snapshot = RegistrationSnapshot.capture(publish, recovery)
    publish.write_bytes(external)

    snapshot.record_prewrite(("wpscomposer-nonce-a-writer",), external)
    RegistrationSnapshot._recover_stale(recovery)

    assert publish.read_bytes() == external
    assert not recovery.exists()


def test_registration_snapshot_removes_probe_created_file(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-a",
        client_credentials=runtime.derive_client_credentials("private-a"),
    )
    snapshot.restore()
    assert not publish.exists()
    assert not (tmp_path / "recovery").exists()



@posix_only
def test_find_node_honors_explicit_override(tmp_path: Path):
    node = tmp_path / "node"
    node.write_text("#!/bin/sh\nprintf 'v24.0.0\\n'\n")
    node.chmod(0o755)
    assert find_node(str(node)) == node.resolve()


def test_probe_toolchain_is_pinned_and_audit_override_is_reviewed():
    package = json.loads(
        Path("macos/wps-jsapi-probe/package.json").read_text(encoding="utf-8")
    )
    assert package["dependencies"] == {"wpsjs": "2.2.3"}
    assert package["devDependencies"] == {"wps-jsapi-declare": "2.2.0"}
    assert package["overrides"] == {"tmp": "0.2.7"}


def test_read_configured_node_uses_installer_runtime_file(tmp_path: Path):
    node = tmp_path / "runtime" / "node"
    node.parent.mkdir()
    node.write_text("node", encoding="utf-8")
    (tmp_path / "runtime.json").write_text(
        json.dumps({"node": str(node)}), encoding="utf-8"
    )

    assert read_configured_node(tmp_path) == str(node.resolve())



@posix_only
def test_component_activation_uses_a_fresh_wps_instance(tmp_path: Path):
    fixture = tmp_path / "fixture.pptx"
    command = runtime.activation_command(
        Path("/Applications/wpsoffice.app"), fixture, reuse_running=False
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



@posix_only
def test_default_staging_root_is_inside_wps_container():
    assert str(runtime.WPS_STAGING_ROOT).endswith(
        "Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer"
    )



@posix_only
def test_create_staging_session_is_private(tmp_path: Path):
    session = runtime.create_staging_session(tmp_path / "WPSComposer")
    try:
        assert session.parent == (tmp_path / "WPSComposer").resolve()
        assert stat.S_IMODE(session.stat().st_mode) == 0o700
        assert session.name.startswith("session-")
    finally:
        session.rmdir()



@posix_only
def test_runtime_removes_staging_session_after_success(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    monkeypatch.setattr(runtime.ProbeRuntime, "_preflight", lambda self: None)
    monkeypatch.setattr(runtime, "list_wps_processes", lambda app, **kwargs: {})
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



@posix_only
def test_runtime_removes_staging_session_after_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    monkeypatch.setattr(runtime.ProbeRuntime, "_preflight", lambda self: None)
    monkeypatch.setattr(runtime, "list_wps_processes", lambda app, **kwargs: {})
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



@posix_only
def test_runtime_surfaces_staging_cleanup_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    monkeypatch.setattr(runtime.ProbeRuntime, "_preflight", lambda self: None)
    monkeypatch.setattr(runtime, "list_wps_processes", lambda app, **kwargs: {})
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
        staging_root=tmp_path / "container" / "WPSComposer",
    )
    original_rmtree = runtime.shutil.rmtree

    with pytest.raises(RuntimeError, match="remove WPS staging session"):
        with probe:
            session = probe.staging_dir
            assert session is not None
            monkeypatch.setattr(
                runtime.shutil,
                "rmtree",
                lambda path: (_ for _ in ()).throw(OSError("busy")),
            )

    monkeypatch.setattr(runtime.shutil, "rmtree", original_rmtree)
    original_rmtree(session)



@posix_only
def test_runtime_lock_serializes_overlapping_sessions(tmp_path: Path):
    lock_path = tmp_path / "wpscomposer.lock"

    with runtime.wps_runtime_lock(lock_path, timeout=1):
        with pytest.raises(TimeoutError, match="another WPSComposer session"):
            with runtime.wps_runtime_lock(lock_path, timeout=0.01):
                pass


@posix_only
@pytest.mark.parametrize("failure", [KeyboardInterrupt(), SystemExit(17)])
def test_interrupted_runtime_enter_releases_lock_and_rethrows_original(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    failure: BaseException,
):
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
        staging_root=tmp_path / "container" / "WPSComposer",
    )

    def interrupt_preflight():
        raise failure

    monkeypatch.setattr(probe, "_preflight", interrupt_preflight)

    with pytest.raises(type(failure)) as caught:
        probe.__enter__()

    assert caught.value is failure
    assert probe._runtime_lock is None
    assert probe.staging_dir is None
    with runtime.wps_runtime_lock(probe.state_dir / "runtime.lock", timeout=0.1):
        pass



@posix_only
def test_list_wps_pids_only_matches_exact_main_executable(monkeypatch):
    class Result:
        stdout = """\
 101 Mon Aug 18 01:00:00 2026 /Applications/wpsoffice.app/Contents/MacOS/wpsoffice
 102 Mon Aug 18 01:00:01 2026 /Applications/wpsoffice.app/Contents/MacOS/wpsoffice --flag
 103 Mon Aug 18 01:00:02 2026 /Applications/wpsoffice.app/Contents/Frameworks/wpsoffice helper
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
    first = _identity(201)
    second = _identity(202)
    probe._owned_wps_processes = {201: first, 202: second}
    alive = {201: first, 202: second}
    monkeypatch.setattr(
        runtime, "read_wps_process_identity", lambda pid, app, **kwargs: alive.get(pid)
    )
    signals = []
    monkeypatch.setattr(
        runtime.os,
        "kill",
        lambda pid, action: (signals.append((pid, action)), alive.pop(pid, None)),
    )

    probe.close()

    assert signals == [(201, signal.SIGTERM), (202, signal.SIGTERM)]


def test_runtime_close_attempts_every_resource_and_aggregates_errors(tmp_path):
    calls = []

    class Resource:
        def __init__(self, name, failure=None):
            self.name = name
            self.failure = failure

        def close(self, *args, **kwargs):
            calls.append(self.name)
            if self.failure is not None:
                raise self.failure

    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    probe._log_streams = [Resource("stream-bad", OSError("fd")), Resource("stream-good")]

    with pytest.raises(runtime.RuntimeCleanupError) as caught:
        probe.close()

    assert calls == ["stream-bad", "stream-good"]
    assert len(caught.value.errors) == 1
    assert probe._log_streams == []


def _identity(pid: int, started: str = "start-a") -> ProcessIdentity:
    return ProcessIdentity(
        pid=pid,
        start_time=started,
        executable="/Applications/wpsoffice.app/Contents/MacOS/wpsoffice",
    )


def test_runtime_never_signals_user_wps_launched_during_run(monkeypatch, tmp_path):
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    owned = _identity(201)
    user = _identity(202)
    probe._owned_wps_processes = {owned.pid: owned}
    current = {owned.pid: owned, user.pid: user}
    monkeypatch.setattr(
        runtime, "read_wps_process_identity", lambda pid, app, **kwargs: current.get(pid)
    )
    signals = []

    def fake_kill(pid, action):
        signals.append((pid, action))
        if action == signal.SIGTERM:
            current.pop(pid, None)

    monkeypatch.setattr(runtime.os, "kill", fake_kill)

    probe.close()

    assert signals == [(201, signal.SIGTERM)]


def test_runtime_kill_set_cannot_grow_during_term_grace(monkeypatch, tmp_path):
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    owned = _identity(201)
    late_user = _identity(303)
    probe._owned_wps_processes = {owned.pid: owned}
    now = [0.0]
    monkeypatch.setattr(runtime.time, "monotonic", lambda: now[0])
    monkeypatch.setattr(runtime.time, "sleep", lambda seconds: now.__setitem__(0, now[0] + seconds))
    monkeypatch.setattr(
        runtime,
        "read_wps_process_identity",
        lambda pid, app, **kwargs: {201: owned, 303: late_user}.get(pid),
    )
    signals = []
    monkeypatch.setattr(runtime.os, "kill", lambda pid, action: signals.append((pid, action)))

    probe.close()

    assert signals == [(201, signal.SIGTERM), (201, signal.SIGKILL)]


def test_runtime_does_not_signal_reused_pid(monkeypatch, tmp_path):
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    probe._owned_wps_processes = {201: _identity(201, "original-start")}
    monkeypatch.setattr(
        runtime,
        "read_wps_process_identity",
        lambda pid, app, **kwargs: _identity(pid, "reused-start"),
    )
    signals = []
    monkeypatch.setattr(runtime.os, "kill", lambda pid, action: signals.append((pid, action)))

    probe.close()

    assert signals == []


def test_preexisting_young_wps_is_never_claimed_by_elapsed_age(monkeypatch, tmp_path):
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "runtime",
        "http://127.0.0.1:45678",
        "token",
        wps_app=tmp_path / "wpsoffice.app",
    )
    preexisting = _identity(101)
    probe._wps_processes_before = {101: preexisting}
    probe._owned_wps_processes = {}
    signals = []
    monkeypatch.setattr(runtime.os, "kill", lambda pid, action: signals.append((pid, action)))

    probe.close()

    assert signals == []


def test_start_servers_launches_managed_wpsjs_processes(monkeypatch, tmp_path: Path):
    commands = []

    class Process:
        def poll(self):
            return 0

    def popen(command, **kwargs):
        commands.append((command, kwargs))
        return Process()

    monkeypatch.setattr(runtime.subprocess, "Popen", popen)
    monkeypatch.setattr(runtime, "find_node", lambda override=None: Path("/node"))
    monkeypatch.setattr(runtime, "find_wpsjs_cli", lambda root: Path("/wpsjs"))
    publish = tmp_path / "publish.xml"
    publish.write_text("<jsplugins/>", encoding="utf-8")
    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "random-runtime-a",
        "http://127.0.0.1:45678",
        "token",
        publish_xml=publish,
        staging_root=tmp_path / "stable" / "WPSComposer",
    )
    probe.runtime_dir.mkdir()
    for component in COMPONENT_CONFIG:
        profile = tmp_path / "profiles" / component
        profile.mkdir(parents=True)
        probe.profiles[component] = profile
    monkeypatch.setattr(probe, "_wait_for_server", lambda *args: None)

    try:
        probe.start_servers()
    finally:
        probe.close()

    assert [command for command, _ in commands] == [
        ["/node", "/wpsjs", "debug", "--server", "--port", "3889"],
        ["/node", "/wpsjs", "debug", "--server", "--port", "3890"],
        ["/node", "/wpsjs", "debug", "--server", "--port", "3891"],
    ]
    assert [kwargs["cwd"] for _, kwargs in commands] == [
        probe.profiles[component] for component in COMPONENT_CONFIG
    ]


def test_stale_registration_recovers_across_random_runtime_roots_before_preflight(
    monkeypatch, tmp_path: Path
):
    publish = tmp_path / "publish.xml"
    original = b"<jsplugins><original/></jsplugins>"
    publish.write_bytes(original)
    stable_root = tmp_path / "stable" / "WPSComposer"
    recovery = stable_root.parent / ".wpscomposer-runtime/registration-recovery"
    crashed = RegistrationSnapshot.capture(publish, recovery)
    runtime.install_registration_entries(
        crashed,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="crashed-session",
        client_credentials=runtime.derive_client_credentials("private-crashed"),
    )

    probe = runtime.ProbeRuntime(
        tmp_path,
        tmp_path / "different-random-runtime",
        "http://127.0.0.1:45678",
        "token",
        publish_xml=publish,
        staging_root=stable_root,
    )
    monkeypatch.setattr(
        probe,
        "_preflight",
        lambda: publish.read_bytes() == original
        or pytest.fail("recovery must happen before port preflight"),
    )
    monkeypatch.setattr(runtime, "list_wps_processes", lambda app, **kwargs: {})

    with probe:
        pass


def test_registration_restore_merges_external_concurrent_changes(tmp_path: Path):
    publish = tmp_path / "publish.xml"
    publish.write_text(
        '<jsplugins><jspluginonline name="original" url="http://old/"/></jsplugins>',
        encoding="utf-8",
    )
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-a",
        client_credentials=runtime.derive_client_credentials("private-a"),
    )
    tree = ET.parse(publish)
    ET.SubElement(
        tree.getroot(),
        "jspluginonline",
        {"name": "external", "url": "http://external/"},
    )
    tree.write(publish, encoding="utf-8", xml_declaration=True)

    snapshot.restore()

    names = {element.attrib.get("name") for element in ET.parse(publish).getroot()}
    assert names == {"original", "external"}


def test_registration_restore_uses_latest_prewrite_base_after_external_change(
    tmp_path: Path,
):
    publish = tmp_path / "publish.xml"
    original = b"<jsplugins><original/></jsplugins>"
    external = b'<jsplugins><original/><jspluginonline name="external"/></jsplugins>'
    publish.write_bytes(original)
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    publish.write_bytes(external)

    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-a",
        client_credentials=runtime.derive_client_credentials("private-a"),
    )
    snapshot.restore()

    assert publish.read_bytes() == external


@pytest.mark.parametrize("capture_existed", [True, False])
def test_installed_crash_recovery_uses_latest_prewrite_base(
    tmp_path: Path, capture_existed: bool
):
    publish = tmp_path / "publish.xml"
    if capture_existed:
        publish.write_bytes(b"<jsplugins><original/></jsplugins>")
    recovery = tmp_path / "recovery"
    snapshot = RegistrationSnapshot.capture(publish, recovery)
    external = b'<jsplugins><jspluginonline name="external"/></jsplugins>'
    publish.write_bytes(external)

    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-a",
        client_credentials=runtime.derive_client_credentials("private-a"),
    )
    RegistrationSnapshot._recover_stale(recovery)

    assert publish.read_bytes() == external
    assert not recovery.exists()


def test_registration_install_retries_concurrent_external_edit(
    monkeypatch, tmp_path: Path
):
    publish = tmp_path / "publish.xml"
    publish.write_text("<jsplugins><original/></jsplugins>", encoding="utf-8")
    snapshot = RegistrationSnapshot.capture(publish, tmp_path / "recovery")
    original_write = runtime._atomic_write
    raced = False

    def racing_write(path, content, mode, *, expected=None):
        nonlocal raced
        if not raced and path == publish:
            raced = True
            path.write_text(
                '<jsplugins><original/><jspluginonline name="external"/></jsplugins>',
                encoding="utf-8",
            )
        if expected is None:
            return original_write(path, content, mode)
        return original_write(path, content, mode, expected=expected)

    monkeypatch.setattr(runtime, "_atomic_write", racing_write)

    runtime.install_registration_entries(
        snapshot,
        {"writer": {"addon_type": "wps", "port": 3889}},
        session_nonce="nonce-a",
        client_credentials=runtime.derive_client_credentials("private-a"),
    )

    names = {element.attrib.get("name") for element in ET.parse(publish).getroot()}
    assert "external" in names


def test_activate_component_does_not_relaunch_one_component(
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
    assert commands == [["open", "-a", str(probe.wps_app), str(first)]]


def test_activate_component_uses_launchservices(monkeypatch, tmp_path: Path):
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

    target = probe.activate_component("writer")

    assert commands == [["open", "-a", str(probe.wps_app), str(target)]]
    assert probe._owned_wps_processes == {}


def _probe_with_writer_fixture(tmp_path: Path) -> runtime.ProbeRuntime:
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
    return probe


def test_activate_component_propagates_launchservices_failure(
    monkeypatch, tmp_path: Path
):
    probe = _probe_with_writer_fixture(tmp_path)
    failure = OSError("launch failed")
    commands = []

    def fail(command, **kwargs):
        commands.append(command)
        raise failure

    monkeypatch.setattr(runtime.subprocess, "run", fail)

    with pytest.raises(OSError) as caught:
        probe.activate_component("writer")

    assert caught.value is failure
    assert commands == [[
        "open",
        "-a",
        str(probe.wps_app),
        str(probe.staging_dir / "fixtures/wpsDemo.docx"),
    ]]
