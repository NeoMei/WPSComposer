import json
from pathlib import Path

from skills.WPSComposer.scripts.macos_probe.runtime import (
    COMPONENT_CONFIG,
    RegistrationSnapshot,
    build_profile,
    find_node,
)


def test_component_config_uses_distinct_ports_and_wps_types():
    assert COMPONENT_CONFIG == {
        "writer": {"addon_type": "wps", "port": 3895, "script": "writer.js"},
        "presentation": {"addon_type": "wpp", "port": 3896, "script": "presentation.js"},
        "spreadsheet": {"addon_type": "et", "port": 3897, "script": "spreadsheet.js"},
    }


def test_build_profile_writes_runtime_config(tmp_path: Path):
    assets = tmp_path / "assets"
    assets.mkdir()
    for name in ("index.html", "manifest.xml", "ribbon.xml", "bridge-client.js", "writer.js"):
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
