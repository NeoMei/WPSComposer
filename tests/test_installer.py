from __future__ import annotations

import json
from pathlib import Path
import subprocess

import pytest

import install
from install import InstallerError, install_plugin


ROOT = Path(__file__).resolve().parents[1]
INSTALL_MACOS_RUNTIME = install._install_macos_runtime


@pytest.fixture(autouse=True)
def stub_macos_runtime_install(monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setattr(install, "_install_macos_runtime", lambda destination: None)


def test_installer_copies_plugin_and_merges_personal_marketplace(tmp_path):
    codex_home = tmp_path / ".codex"
    marketplace_root = tmp_path
    marketplace_file = marketplace_root / ".agents" / "plugins" / "marketplace.json"
    marketplace_file.parent.mkdir(parents=True)
    marketplace_file.write_text(
        json.dumps(
            {
                "name": "personal",
                "interface": {"displayName": "Personal"},
                "plugins": [
                    {
                        "name": "existing-plugin",
                        "source": {"source": "local", "path": "./plugins/existing"},
                        "policy": {
                            "installation": "AVAILABLE",
                            "authentication": "ON_INSTALL",
                        },
                        "category": "Productivity",
                    }
                ],
            }
        ),
        encoding="utf-8",
    )

    result = install_plugin(ROOT, codex_home, marketplace_root)

    assert result.destination == codex_home / "plugins" / "wps-composer"
    assert (result.destination / ".codex-plugin" / "plugin.json").is_file()
    assert (result.destination / "skills" / "WPSComposer" / "SKILL.md").is_file()
    assert not (result.destination / ".git").exists()
    assert not (result.destination / "tests").exists()
    assert not (result.destination / "macos/wps-jsapi-probe/node_modules").exists()
    marketplace = json.loads(marketplace_file.read_text(encoding="utf-8"))
    assert [plugin["name"] for plugin in marketplace["plugins"]] == [
        "existing-plugin",
        "wps-composer",
    ]
    added = marketplace["plugins"][-1]
    assert added["source"] == {
        "source": "local",
        "path": "./.codex/plugins/wps-composer",
    }
    assert result.source_path == "./.codex/plugins/wps-composer"


def test_installer_refuses_existing_destination_without_force(tmp_path):
    codex_home = tmp_path / ".codex"
    destination = codex_home / "plugins" / "wps-composer"
    destination.mkdir(parents=True)

    with pytest.raises(InstallerError, match="already exists"):
        install_plugin(ROOT, codex_home, tmp_path)


def test_installer_force_replaces_existing_destination(tmp_path):
    codex_home = tmp_path / ".codex"
    destination = codex_home / "plugins" / "wps-composer"
    destination.mkdir(parents=True)
    (destination / "stale.txt").write_text("stale", encoding="utf-8")

    result = install_plugin(ROOT, codex_home, tmp_path, force=True)

    assert result.destination == destination
    assert not (destination / "stale.txt").exists()
    assert (destination / ".codex-plugin" / "plugin.json").is_file()


def test_installer_rejects_malformed_marketplace_before_copy(tmp_path):
    codex_home = tmp_path / ".codex"
    marketplace_file = tmp_path / ".agents" / "plugins" / "marketplace.json"
    marketplace_file.parent.mkdir(parents=True)
    marketplace_file.write_text("{broken", encoding="utf-8")

    with pytest.raises(InstallerError, match="invalid marketplace JSON"):
        install_plugin(ROOT, codex_home, tmp_path)

    assert not (codex_home / "plugins" / "wps-composer").exists()


def test_macos_runtime_install_uses_pinned_production_dependencies(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    monkeypatch.setattr(install.platform, "system", lambda: "Darwin")
    destination = tmp_path / "plugin"
    probe = destination / "macos" / "wps-jsapi-probe"
    probe.mkdir(parents=True)
    (probe / "package.json").write_text("{}", encoding="utf-8")
    node = tmp_path / "node"
    npm = tmp_path / "npm"
    node.write_text("node", encoding="utf-8")
    npm.write_text("npm", encoding="utf-8")
    commands = []

    def fake_which(name: str):
        return str({"node": node, "npm": npm}[name])

    def fake_run(command, **kwargs):
        commands.append((command, kwargs))
        if command[0] == str(node.resolve()):
            return subprocess.CompletedProcess(command, 0, "v20.20.2\n", "")
        return subprocess.CompletedProcess(command, 0, "", "")

    monkeypatch.setattr(install.shutil, "which", fake_which)
    monkeypatch.setattr(install.subprocess, "run", fake_run)

    INSTALL_MACOS_RUNTIME(destination)

    assert commands[1][0] == [
        str(npm.resolve()),
        "ci",
        "--omit=dev",
        "--ignore-scripts",
    ]
    assert commands[1][1]["cwd"] == probe
    assert json.loads((probe / "runtime.json").read_text(encoding="utf-8")) == {
        "node": str(node.resolve())
    }


def test_non_macos_install_does_not_require_node(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    destination = tmp_path / "plugin"
    probe = destination / "macos" / "wps-jsapi-probe"
    probe.mkdir(parents=True)
    (probe / "package.json").write_text("{}", encoding="utf-8")
    monkeypatch.setattr(install.platform, "system", lambda: "Windows")
    monkeypatch.setattr(
        install.shutil,
        "which",
        lambda name: (_ for _ in ()).throw(AssertionError(name)),
    )

    INSTALL_MACOS_RUNTIME(destination)

    assert not (probe / "runtime.json").exists()


def test_runtime_install_failure_leaves_destination_and_marketplace_unchanged(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    codex_home = tmp_path / ".codex"

    def fail_runtime_install(destination: Path):
        raise InstallerError("npm ci failed")

    monkeypatch.setattr(install, "_install_macos_runtime", fail_runtime_install)

    with pytest.raises(InstallerError, match="npm ci failed"):
        install_plugin(ROOT, codex_home, tmp_path)

    assert not (codex_home / "plugins" / "wps-composer").exists()
    assert not (tmp_path / ".agents/plugins/marketplace.json").exists()


def test_force_upgrade_restores_old_plugin_and_marketplace_after_commit_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
):
    codex_home = tmp_path / ".codex"
    destination = codex_home / "plugins" / "wps-composer"
    destination.mkdir(parents=True)
    (destination / "old.txt").write_text("old", encoding="utf-8")
    marketplace_file = tmp_path / ".agents/plugins/marketplace.json"
    marketplace_file.parent.mkdir(parents=True)
    original_marketplace = json.dumps(
        {
            "name": "personal",
            "plugins": [
                {
                    "name": "wps-composer",
                    "source": {"source": "local", "path": "./old"},
                }
            ],
        }
    ).encode()
    marketplace_file.write_bytes(original_marketplace)

    def fail_after_marketplace_replace(path: Path, data: dict):
        path.write_text('{"changed": true}\n', encoding="utf-8")
        raise OSError("marketplace commit failed")

    monkeypatch.setattr(install, "_write_json_atomically", fail_after_marketplace_replace)

    with pytest.raises(OSError, match="marketplace commit failed"):
        install_plugin(ROOT, codex_home, tmp_path, force=True)

    assert (destination / "old.txt").read_text(encoding="utf-8") == "old"
    assert not (destination / ".codex-plugin/plugin.json").exists()
    assert marketplace_file.read_bytes() == original_marketplace
