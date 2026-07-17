from __future__ import annotations

import json
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_plugin_manifest_matches_bundle_layout():
    manifest = json.loads(
        (ROOT / ".codex-plugin" / "plugin.json").read_text(encoding="utf-8")
    )

    assert manifest["name"] == "wps-composer"
    assert manifest["version"] == "0.2.0"
    assert manifest["skills"] == "./skills/"
    assert (ROOT / manifest["skills"]).is_dir()
    assert (ROOT / "skills" / "WPSComposer" / "SKILL.md").is_file()
