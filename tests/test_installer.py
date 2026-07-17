from __future__ import annotations

import json
from pathlib import Path

import pytest

from install import InstallerError, install_plugin


ROOT = Path(__file__).resolve().parents[1]


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
