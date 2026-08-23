from __future__ import annotations

from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
README = (ROOT / "README.md").read_text(encoding="utf-8")
SKILL = (ROOT / "skills" / "WPSComposer" / "SKILL.md").read_text(encoding="utf-8")
AGENTS = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
LONGFORM_M0 = (ROOT / "docs" / "longform-m0.md").read_text(encoding="utf-8")
API = (ROOT / "skills" / "WPSComposer" / "references" / "api.md").read_text(
    encoding="utf-8"
)
LONGFORM_MARKDOWN_PATH = ROOT / "docs" / "longform-markdown.md"
MACOS_M5_PATH = ROOT / "docs" / "macos-longform-m5-verification.md"
PROGRESS = (ROOT / ".superpowers" / "sdd" / "progress.md").read_text(
    encoding="utf-8"
)


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


def test_m5_public_docs_describe_default_route_and_closed_failure_boundary():
    assert LONGFORM_MARKDOWN_PATH.is_file()
    longform = LONGFORM_MARKDOWN_PATH.read_text(encoding="utf-8")
    required = (
        "layout_engine: legacy",
        "DOCX/PDF",
        "600",
        "Pillow>=10",
        "pypdf>=4",
        "pdfplumber>=0.11",
        "Unicode",
        "top-left",
        "FORMULA_FALLBACK_IMAGE_UNAVAILABLE",
        "ENGINE_LOST",
        "source paths",
        "0.8.0",
        "Windows",
    )
    for document in (README, SKILL, API, longform):
        for token in ("layout_engine: legacy", "DOCX/PDF", "600", "ENGINE_LOST"):
            assert token in document
    for token in required:
        assert token in longform
    assert "M5/final work is separate" not in SKILL
    assert "M4 does not change that routing" not in SKILL
    assert "M4 does not reroute" not in API


def test_m5_macos_evidence_is_explicit_and_windows_release_gate_remains_open():
    assert MACOS_M5_PATH.is_file()
    evidence = MACOS_M5_PATH.read_text(encoding="utf-8")
    for token in (
        "WPS 12.1.26055",
        "arm64",
        "63",
        "139",
        "three consecutive",
        "HEADING_ORPHAN",
        "FORMULA_MALFORMED",
        "Windows evidence is pending",
    ):
        assert token in evidence
    for token in ("M5 Task 8", "0d61345", "three consecutive"):
        assert token in PROGRESS
