# WPSComposer Engineering Baseline Remediation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Establish a tested Python API, current Codex plugin packaging, cross-platform installation, and internally consistent documentation before the macOS WPS JSAPI feasibility work begins.

**Architecture:** Keep the existing COM composers and internal relative imports unchanged. Add a thin public package facade at `skills.WPSComposer`, a standard-library installer that copies the complete plugin into a personal Codex marketplace, and platform-independent pytest coverage around these boundaries.

**Tech Stack:** Python 3.9+, pytest, standard-library `argparse`/`json`/`pathlib`/`shutil`, PowerShell and POSIX shell wrappers, Codex plugin manifest v current local-marketplace layout.

## Global Constraints

- Use `/Users/neomei/项目/WpsComposer` as the only writable repository.
- Do not modify `/Users/neomei/项目/codexprojects/WpsComposer`.
- Do not implement the macOS WPS add-in, JSON-RPC bridge, or `MacWpsJsBackend` in this plan.
- Do not change Writer, Sheet, Slide, or PDF runtime behavior.
- Keep the public skill name `WPSComposer`; use `wps-composer` only as the plugin identifier.
- Tests must never write to the live `~/.codex` or `~/.agents` directories.
- Every production behavior change follows RED-GREEN-REFACTOR and gets its own commit.

---

## Planned File Map

- Create `pyproject.toml`: project metadata, optional dependencies, pytest configuration.
- Create `skills/__init__.py`: mark the repository skill collection as a Python package.
- Create `skills/WPSComposer/__init__.py`: stable public API facade.
- Create `skills/WPSComposer/scripts/__init__.py`: mark engine scripts as a package.
- Modify `skills/WPSComposer/scripts/wps_engine.py`: complete the public `__all__` list.
- Create `tests/test_public_api.py`: public import and registry regression tests.
- Create `tests/test_markdown_parser.py`: representative platform-independent parser tests.
- Modify `skills/WPSComposer/scripts/md_parser.py`: remove one duplicate helper definition.
- Modify `.codex-plugin/plugin.json`: current plugin identifier and skill-directory path.
- Create `tests/test_plugin_manifest.py`: plugin layout and manifest regression tests.
- Create `install.py`: cross-platform plugin and personal-marketplace installer.
- Create `tests/test_installer.py`: isolated installer behavior and error tests.
- Replace `install.ps1`: thin PowerShell wrapper around `install.py`.
- Create `install.sh`: thin POSIX wrapper around `install.py`.
- Create `tests/test_documentation.py`: documented import, preset, install, and output-contract checks.
- Modify `README.md`, `AGENTS.md`, and `skills/WPSComposer/SKILL.md`: align user and contributor guidance with runtime behavior.

---

### Task 1: Add the Public Python Package and Test Harness

**Files:**
- Create: `pyproject.toml`
- Create: `tests/test_public_api.py`
- Create: `skills/__init__.py`
- Create: `skills/WPSComposer/__init__.py`
- Create: `skills/WPSComposer/scripts/__init__.py`
- Modify: `skills/WPSComposer/scripts/wps_engine.py`

**Interfaces:**
- Produces: `skills.WPSComposer.__all__: list[str]`
- Produces: `skills.WPSComposer.generate`, `inspect`, `edit`, `attach_active`, `parse`, `parse_file`, composer classes, model classes, and registry helpers.
- Consumes: existing exports imported by `skills/WPSComposer/scripts/wps_engine.py`.

- [ ] **Step 1: Add project/test configuration**

Create `pyproject.toml`:

```toml
[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[project]
name = "wps-composer"
version = "0.2.0"
description = "WPS-backed rich document composition for Codex"
readme = "README.md"
requires-python = ">=3.9"
license = "MIT"

[project.optional-dependencies]
windows = ["pywin32>=306; platform_system == 'Windows'"]
pdf = ["pypdf>=4", "pdfplumber>=0.11", "reportlab>=4"]
dev = ["pytest>=8"]

[tool.setuptools.packages.find]
include = ["skills*"]

[tool.pytest.ini_options]
addopts = "-ra"
testpaths = ["tests"]
pythonpath = ["."]
```

Create the local environment and install development dependencies:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/python -m pip install -e '.[dev]'
```

Expected: installation exits 0 and `.venv/bin/python -m pytest --version` reports pytest 8 or newer.

- [ ] **Step 2: Write the failing public-import test**

Create `tests/test_public_api.py`:

```python
from __future__ import annotations


def test_public_package_exports_supported_api():
    import skills.WPSComposer as api

    required = {
        "WriterComposer",
        "SheetComposer",
        "SlideComposer",
        "PdfComposer",
        "generate",
        "list_formats",
        "list_available_presets",
        "inspect",
        "edit",
        "attach_active",
        "parse",
        "parse_file",
        "StructuredDocument",
    }

    assert required <= set(api.__all__)
    for name in required:
        assert hasattr(api, name), name


def test_public_import_does_not_require_pywin32():
    import skills.WPSComposer as api

    assert api.list_formats() == ["docx", "pptx", "xlsx", "pdf"]
    assert api.list_available_presets() == [
        "academic",
        "consultant",
        "business",
        "tech",
        "proposal",
    ]
```

- [ ] **Step 3: Run the test and verify RED**

Run:

```bash
.venv/bin/python -m pytest tests/test_public_api.py -v
```

Expected: FAIL because `skills.WPSComposer` has no `__all__` or public attributes.

- [ ] **Step 4: Implement the minimal public package**

Create `skills/__init__.py`:

```python
"""Python package containing WPSComposer's bundled Codex skills."""
```

Create `skills/WPSComposer/scripts/__init__.py`:

```python
"""Implementation package for the WPSComposer skill."""
```

Create `skills/WPSComposer/__init__.py`:

```python
"""Stable Python API for WPSComposer."""

from .scripts.wps_engine import *  # noqa: F401,F403
from .scripts.wps_engine import __all__
```

Extend `skills/WPSComposer/scripts/wps_engine.py::__all__` with the already imported high-level names:

```python
    # High-level generation API
    "generate",
    "list_formats",
    "list_available_presets",
    # Document model and parser
    "StructuredDocument",
    "Section",
    "Paragraph",
    "Span",
    "ListBlock",
    "TableBlock",
    "CodeBlock",
    "ImageBlock",
    "BlockQuote",
    "parse",
    "parse_file",
```

- [ ] **Step 5: Run the focused and full tests and verify GREEN**

Run:

```bash
.venv/bin/python -m pytest tests/test_public_api.py -v
.venv/bin/python -m pytest -v
```

Expected: 2 tests pass; the full suite also exits 0.

- [ ] **Step 6: Commit the public package baseline**

```bash
git add pyproject.toml skills/__init__.py skills/WPSComposer/__init__.py \
  skills/WPSComposer/scripts/__init__.py \
  skills/WPSComposer/scripts/wps_engine.py tests/test_public_api.py
git commit -m "Add tested WPSComposer Python API"
```

---

### Task 2: Protect Markdown Parsing and Remove the Duplicate Helper

**Files:**
- Create: `tests/test_markdown_parser.py`
- Modify: `skills/WPSComposer/scripts/md_parser.py`

**Interfaces:**
- Consumes: `skills.WPSComposer.parse(md_text: str, base_dir: str = "") -> StructuredDocument`
- Produces: no new runtime interface; preserves current parser behavior.

- [ ] **Step 1: Write the representative parser regression test**

Create `tests/test_markdown_parser.py`:

```python
from __future__ import annotations

from skills.WPSComposer import parse
from skills.WPSComposer.scripts.document_model import Paragraph, TableBlock, TaskList


def test_parse_preserves_heading_task_list_and_table():
    document = parse(
        """# 项目状态

正文 **加粗**

- [x] 已完成
- [ ] 待处理

| 任务 | 状态 |
|:---|---:|
| 基线修复 | 进行中 |
"""
    )

    assert document.title == "项目状态"
    assert len(document.sections) == 1
    elements = document.sections[0].elements
    assert [type(element) for element in elements] == [
        Paragraph,
        TaskList,
        TableBlock,
    ]
    table = elements[-1]
    assert table.headers == ["任务", "状态"]
    assert table.rows == [["基线修复", "进行中"]]
    assert table.alignments == ["left", "right"]
```

- [ ] **Step 2: Run the parser test as a characterization test**

Run:

```bash
.venv/bin/python -m pytest tests/test_markdown_parser.py -v
```

Expected: PASS, establishing the behavior that must remain unchanged during cleanup.

- [ ] **Step 3: Remove only the second duplicate definition**

Delete the second copy at the end of `_parse_table_block`:

```python
def _is_separator_cell(cell: str) -> bool:
    return bool(re.match(r"^[\-\s:]+\+?$", cell))
```

Keep the first definition immediately after `_detect_alignments` unchanged.

- [ ] **Step 4: Verify parser behavior remains GREEN**

Run:

```bash
.venv/bin/python -m pytest tests/test_markdown_parser.py -v
.venv/bin/python -m pytest -v
```

Expected: all tests pass.

- [ ] **Step 5: Commit the deterministic cleanup**

```bash
git add tests/test_markdown_parser.py skills/WPSComposer/scripts/md_parser.py
git commit -m "Cover Markdown parsing behavior"
```

---

### Task 3: Correct the Codex Plugin Manifest

**Files:**
- Create: `tests/test_plugin_manifest.py`
- Modify: `.codex-plugin/plugin.json`

**Interfaces:**
- Produces: plugin identifier `wps-composer`.
- Produces: manifest skill directory `./skills/`.
- Preserves: bundled skill frontmatter name `WPSComposer`.

- [ ] **Step 1: Write the failing manifest test**

Create `tests/test_plugin_manifest.py`:

```python
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
```

- [ ] **Step 2: Run the test and verify RED**

Run:

```bash
.venv/bin/python -m pytest tests/test_plugin_manifest.py -v
```

Expected: FAIL because the current manifest name is `WPSComposer` and `skills` is an array.

- [ ] **Step 3: Update the manifest**

Replace `.codex-plugin/plugin.json` with:

```json
{
  "name": "wps-composer",
  "version": "0.2.0",
  "description": "Agent-friendly WPS document inspection and rich document composition for Writer, Spreadsheet, Presentation, and PDF workflows.",
  "license": "MIT",
  "repository": "https://github.com/NeoMei/WPSComposer",
  "keywords": ["wps", "documents", "docx", "pptx", "xlsx", "pdf"],
  "skills": "./skills/"
}
```

- [ ] **Step 4: Verify GREEN and commit**

Run:

```bash
.venv/bin/python -m pytest tests/test_plugin_manifest.py -v
.venv/bin/python -m pytest -v
```

Expected: all tests pass.

```bash
git add .codex-plugin/plugin.json tests/test_plugin_manifest.py
git commit -m "Align plugin manifest with Codex layout"
```

---

### Task 4: Implement the Isolated Cross-Platform Installer

**Files:**
- Create: `tests/test_installer.py`
- Create: `install.py`

**Interfaces:**
- Produces: `InstallerError(RuntimeError)`.
- Produces: `InstallResult(destination: Path, marketplace_file: Path, source_path: str)`.
- Produces: `install_plugin(source_root: Path, codex_home: Path, marketplace_root: Path, force: bool = False) -> InstallResult`.
- Produces CLI: `python install.py [--force] [--codex-home PATH] [--marketplace-root PATH]`.

- [ ] **Step 1: Write failing installer tests**

Create `tests/test_installer.py`:

```python
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
```

- [ ] **Step 2: Run installer tests and verify RED**

Run:

```bash
.venv/bin/python -m pytest tests/test_installer.py -v
```

Expected: collection ERROR with `ModuleNotFoundError: No module named 'install'`.

- [ ] **Step 3: Implement the minimal installer**

Create `install.py`:

```python
#!/usr/bin/env python3
"""Install WPSComposer into a personal Codex plugin marketplace."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
from pathlib import Path
import shutil
import tempfile


PLUGIN_NAME = "wps-composer"
IGNORED_NAMES = {
    ".git",
    ".pytest_cache",
    ".venv",
    "__pycache__",
    "build",
    "dist",
    "docs",
    "tests",
}


class InstallerError(RuntimeError):
    """Raised when the local plugin installation cannot be completed safely."""


@dataclass(frozen=True)
class InstallResult:
    destination: Path
    marketplace_file: Path
    source_path: str


def _validate_source(source_root: Path) -> None:
    required = [
        source_root / ".codex-plugin" / "plugin.json",
        source_root / "skills" / "WPSComposer" / "SKILL.md",
    ]
    missing = [str(path) for path in required if not path.is_file()]
    if missing:
        raise InstallerError(
            "invalid WPSComposer source bundle; missing: " + ", ".join(missing)
        )


def _read_marketplace(path: Path) -> dict:
    if not path.exists():
        return {
            "name": "personal",
            "interface": {"displayName": "Personal"},
            "plugins": [],
        }
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise InstallerError(f"invalid marketplace JSON at {path}: {exc}") from exc
    if not isinstance(data, dict) or not isinstance(data.get("plugins"), list):
        raise InstallerError(
            f"invalid marketplace JSON at {path}: 'plugins' must be a list"
        )
    return data


def _marketplace_source_path(destination: Path, marketplace_root: Path) -> str:
    try:
        relative = destination.resolve().relative_to(marketplace_root.resolve())
    except ValueError as exc:
        raise InstallerError(
            "plugin destination must be inside the marketplace root"
        ) from exc
    return "./" + relative.as_posix()


def _plugin_entry(source_path: str) -> dict:
    return {
        "name": PLUGIN_NAME,
        "source": {"source": "local", "path": source_path},
        "policy": {
            "installation": "AVAILABLE",
            "authentication": "ON_INSTALL",
        },
        "category": "Productivity",
    }


def _write_json_atomically(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=path.name + ".", suffix=".tmp", dir=path.parent
    )
    temporary_path = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(data, handle, ensure_ascii=False, indent=2)
            handle.write("\n")
        os.replace(temporary_path, path)
    finally:
        if temporary_path.exists():
            temporary_path.unlink()


def _copy_plugin(source_root: Path, destination: Path, force: bool) -> None:
    if destination.exists():
        if not force:
            raise InstallerError(
                f"plugin destination already exists: {destination}; use --force"
            )
        shutil.rmtree(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(
        source_root,
        destination,
        ignore=shutil.ignore_patterns(*sorted(IGNORED_NAMES), "*.pyc", "*.pyo"),
    )


def install_plugin(
    source_root: Path,
    codex_home: Path,
    marketplace_root: Path,
    force: bool = False,
) -> InstallResult:
    source_root = source_root.resolve()
    codex_home = codex_home.expanduser().resolve()
    marketplace_root = marketplace_root.expanduser().resolve()
    _validate_source(source_root)

    destination = codex_home / "plugins" / PLUGIN_NAME
    marketplace_file = (
        marketplace_root / ".agents" / "plugins" / "marketplace.json"
    )
    marketplace = _read_marketplace(marketplace_file)
    source_path = _marketplace_source_path(destination, marketplace_root)

    if destination.exists() and not force:
        raise InstallerError(
            f"plugin destination already exists: {destination}; use --force"
        )

    _copy_plugin(source_root, destination, force)
    plugins = [
        plugin
        for plugin in marketplace["plugins"]
        if not isinstance(plugin, dict) or plugin.get("name") != PLUGIN_NAME
    ]
    plugins.append(_plugin_entry(source_path))
    marketplace["plugins"] = plugins
    _write_json_atomically(marketplace_file, marketplace)
    return InstallResult(destination, marketplace_file, source_path)


def _parser() -> argparse.ArgumentParser:
    home = Path.home()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--force", action="store_true")
    parser.add_argument(
        "--codex-home", type=Path, default=Path(os.environ.get("CODEX_HOME", home / ".codex"))
    )
    parser.add_argument("--marketplace-root", type=Path, default=home)
    return parser


def main() -> int:
    args = _parser().parse_args()
    try:
        result = install_plugin(
            Path(__file__).resolve().parent,
            args.codex_home,
            args.marketplace_root,
            force=args.force,
        )
    except InstallerError as exc:
        print(f"error: {exc}")
        return 1
    print(f"Installed plugin: {result.destination}")
    print(f"Updated marketplace: {result.marketplace_file}")
    print("Restart the ChatGPT/Codex desktop app, then install or enable wps-composer from Plugins.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run installer tests and verify GREEN**

Run:

```bash
.venv/bin/python -m pytest tests/test_installer.py -v
.venv/bin/python -m pytest -v
```

Expected: 4 installer tests pass; full suite exits 0.

- [ ] **Step 5: Commit the installer core**

```bash
git add install.py tests/test_installer.py
git commit -m "Add tested cross-platform plugin installer"
```

---

### Task 5: Add Platform Wrappers and Align Documentation

**Files:**
- Create: `install.sh`
- Replace: `install.ps1`
- Create: `tests/test_documentation.py`
- Modify: `README.md`
- Modify: `AGENTS.md`
- Modify: `skills/WPSComposer/SKILL.md`

**Interfaces:**
- Produces POSIX command: `./install.sh [installer options]`.
- Produces PowerShell command: `pwsh ./install.ps1 [installer options]`.
- Documents public import: `from skills.WPSComposer import generate, inspect, edit, attach_active`.

- [ ] **Step 1: Write failing documentation contract tests**

Create `tests/test_documentation.py`:

```python
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
```

- [ ] **Step 2: Run documentation tests and verify RED**

Run:

```bash
.venv/bin/python -m pytest tests/test_documentation.py -v
```

Expected: all three tests fail against the current documentation.

- [ ] **Step 3: Add the thin platform wrappers**

Create `install.sh`:

```sh
#!/bin/sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
exec "${PYTHON:-python3}" "$SCRIPT_DIR/install.py" "$@"
```

Replace `install.ps1` with:

```powershell
$ErrorActionPreference = 'Stop'
$Python = if ($env:PYTHON) { $env:PYTHON } else { 'python' }
& $Python (Join-Path $PSScriptRoot 'install.py') @args
exit $LASTEXITCODE
```

Make the POSIX wrapper executable:

```bash
chmod +x install.sh
```

- [ ] **Step 4: Update user-facing documentation exactly to the supported contracts**

In `README.md`:

- Replace the Windows-only installation block with:

```markdown
### 安装为 Codex 本地插件

```bash
python3 install.py
# 覆盖已安装版本
python3 install.py --force
```

Windows 也可以使用 `pwsh ./install.ps1`，macOS/Linux 可以使用
`./install.sh`。安装后重启 ChatGPT/Codex Desktop，在 Plugins 目录中安装或启用
`wps-composer`。
```

- Replace Python examples with:

```python
from skills.WPSComposer import generate, inspect, edit, attach_active
```

- Describe five presets: `academic`, `consultant`, `business`, `tech`, and `proposal`.
- State that Windows COM generation requires WPS/MS Office and `pywin32`, while parser/model/PDF utilities remain importable cross-platform.

In `skills/WPSComposer/SKILL.md`:

- Replace direct module imports with `from skills.WPSComposer import generate, inspect, edit, attach_active, PdfComposer` as appropriate for each example.
- Change the preset description from four presets to five and include `proposal`.
- Replace the output discipline paragraph with:

```markdown
Generate and deliver only the artifact format requested by the user. During
development, create PDF evidence separately when a native WPS layout change
needs visual verification; do not make that PDF an automatic public companion
output.
```

- Add a short installation note stating that the plugin bundle is installed through a Codex personal marketplace by `install.py`.

In `AGENTS.md`:

- Replace the legacy `~/.codex/skills/WPSComposer` commands with `python install.py` and `python install.py --force`.
- Replace the stale `wps_engine.py` no-split rule with:

```markdown
- Keep `wps_engine.py` as the backward-compatible public facade; implementation
  belongs in focused composer and infrastructure modules.
```

- Clarify that PDF export is a development verification artifact, while public generation returns only the requested format.

- [ ] **Step 5: Run focused checks and verify GREEN**

Run:

```bash
sh -n install.sh
./install.sh --help
.venv/bin/python install.py --help
.venv/bin/python -m pytest tests/test_documentation.py -v
.venv/bin/python -m pytest -v
```

Expected: shell syntax exits 0, both help commands exit 0, 3 documentation tests pass, and the full suite passes.

- [ ] **Step 6: Commit wrappers and documentation**

```bash
git add install.sh install.ps1 README.md AGENTS.md \
  skills/WPSComposer/SKILL.md tests/test_documentation.py
git commit -m "Document and wrap local plugin installation"
```

---

### Task 6: Run the Complete Baseline Verification

**Files:**
- Verify all files changed by Tasks 1-5.
- Do not create or modify live installation files.

**Interfaces:**
- Consumes the complete test suite and installer CLI.
- Produces verification evidence for the baseline handoff.

- [ ] **Step 1: Run the complete automated suite**

```bash
.venv/bin/python -m pytest -v
```

Expected: every test passes with zero failures and zero errors.

- [ ] **Step 2: Parse every Python source file without writing bytecode**

```bash
PYTHONDONTWRITEBYTECODE=1 .venv/bin/python -B - <<'PY'
import ast
from pathlib import Path

files = sorted(Path('.').rglob('*.py'))
for path in files:
    if '.venv' in path.parts:
        continue
    ast.parse(path.read_text(encoding='utf-8'), filename=str(path))
print(f'parsed {sum(1 for p in files if ".venv" not in p.parts)} Python files')
PY
```

Expected: exits 0 and reports the number of parsed project/test Python files.

- [ ] **Step 3: Exercise the installer only against an isolated temporary home**

```bash
tmp=$(mktemp -d)
.venv/bin/python install.py --codex-home "$tmp/.codex" --marketplace-root "$tmp"
test -f "$tmp/.codex/plugins/wps-composer/.codex-plugin/plugin.json"
test -f "$tmp/.agents/plugins/marketplace.json"
rm -rf "$tmp"
```

Expected: install exits 0 and both asserted files exist before cleanup.

- [ ] **Step 4: Run repository integrity checks**

```bash
git diff --check
git status --short --branch
git -C /Users/neomei/项目/codexprojects/WpsComposer status --short --branch
```

Expected: no whitespace errors; the primary repository contains only intentional committed changes; the second repository still has only its pre-existing `.DS_Store` files.

- [ ] **Step 5: Review requirements line by line**

Confirm from fresh command output that:

- `from skills.WPSComposer import generate` succeeds without pywin32;
- runtime formats are DOCX/PPTX/XLSX/PDF;
- runtime presets are academic/consultant/business/tech/proposal;
- the plugin manifest uses `wps-composer` and `./skills/`;
- installer tests do not touch live home directories;
- docs contain no unsupported direct imports;
- no Mac backend code was introduced.

- [ ] **Step 6: Record the final implementation state**

```bash
git log --oneline -7
git status --short --branch
```

Expected: implementation commits are visible, the worktree is clean, and the branch is ahead of `origin/master` only by intentional local commits.
