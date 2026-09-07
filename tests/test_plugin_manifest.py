from __future__ import annotations

import json
from pathlib import Path
import re


ROOT = Path(__file__).resolve().parents[1]


def test_plugin_manifest_matches_bundle_layout():
    manifest = json.loads(
        (ROOT / ".codex-plugin" / "plugin.json").read_text(encoding="utf-8")
    )

    assert manifest["name"] == "wps-composer"
    assert manifest["version"] == "0.8.1"
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert re.search(r'^version = "0\.8\.1"$', pyproject, re.MULTILINE)
    assert manifest["skills"] == "./skills/"
    assert (ROOT / manifest["skills"]).is_dir()
    assert (ROOT / "skills" / "WPSComposer" / "SKILL.md").is_file()
