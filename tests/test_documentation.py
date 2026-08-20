from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
README = (ROOT / "README.md").read_text(encoding="utf-8")
SKILL = (ROOT / "skills" / "WPSComposer" / "SKILL.md").read_text(encoding="utf-8")
AGENTS = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
LONGFORM_M0 = (ROOT / "docs" / "longform-m0.md").read_text(encoding="utf-8")


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


def test_longform_m0_docs_cover_matrix_gate_and_recovery():
    for capability_id in range(1, 16):
        assert f"| {capability_id} |" in LONGFORM_M0
    for required in (
        "1–14",
        "no-go",
        "SVG",
        "--platform macos",
        "--platform windows",
        "--platform verify",
        "unrestricted local run",
        "pywin32",
        "separate user WPS document",
        "registration",
        "relative filename",
        "must not contain",
        "Windows native gate is pending",
    ):
        assert required in LONGFORM_M0
