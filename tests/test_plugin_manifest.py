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
    assert manifest["version"] == "0.9.3"
    pyproject = (ROOT / "pyproject.toml").read_text(encoding="utf-8")
    assert re.search(r'^version = "0\.9\.3"$', pyproject, re.MULTILINE)
    assert manifest["skills"] == "./skills/"
    assert (ROOT / manifest["skills"]).is_dir()
    assert (ROOT / "skills" / "WPSComposer" / "SKILL.md").is_file()


def test_root_and_skill_version_declarations_stay_in_sync():
    versions = []
    for root in (ROOT, ROOT / "skills" / "WPSComposer"):
        manifest = json.loads(
            (root / ".codex-plugin" / "plugin.json").read_text(encoding="utf-8")
        )
        pyproject = (root / "pyproject.toml").read_text(encoding="utf-8")
        declared = re.findall(r'^version = "([^"]+)"$', pyproject, re.MULTILINE)
        assert len(declared) == 1
        versions.extend([manifest["version"], declared[0]])
    assert len(set(versions)) == 1, versions
