from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
README = (ROOT / "README.md").read_text(encoding="utf-8")
SKILL = (ROOT / "skills" / "WPSComposer" / "SKILL.md").read_text(encoding="utf-8")
AGENTS = (ROOT / "AGENTS.md").read_text(encoding="utf-8")


def test_docs_use_supported_public_import():
    for document in (README, SKILL):
        assert "from orchestrator import" not in document
        assert "from wps_engine import" not in document
        assert "from skills.WPSComposer import" in document


def test_docs_name_all_five_presets():
    for name in ("academic", "consultant", "business", "tech", "proposal"):
        assert name in README
        assert name in SKILL
    assert "4 colour+font presets" not in SKILL
    assert "4 套配色+字体" not in README


def test_docs_describe_current_install_and_output_contract():
    assert "install.py" in README
    assert "personal marketplace" in SKILL.lower()
    assert "only the requested artifact" in SKILL.lower()
    assert "~/.codex/skills/WPSComposer" not in AGENTS
    assert "do not split" not in AGENTS.lower()
