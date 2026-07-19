# Repository Guidelines

Guidelines for contributing to **WPSComposer**, a Codex plugin that generates
rich-layout DOCX / PPTX / XLSX / PDF documents through WPS Office automation.
Windows uses COM; macOS uses WPS JSAPI add-ins over a loopback bridge.

## Project Structure & Module Organization

```text
WPSComposer/
├── .codex-plugin/plugin.json        # Plugin manifest
├── skills/WPSComposer/
│   ├── SKILL.md                     # Skill entry (name: WPSComposer)
│   ├── references/api.md            # Public composer API reference
│   └── scripts/
│       ├── wps_engine.py            # Backward-compatible public facade
│       ├── _dispatch.py             # COM dispatch and format constants
│       ├── _colors.py               # Unified colour model
│       ├── _base.py                 # Shared composer lifecycle
│       ├── writer.py                # WriterComposer
│       ├── sheet.py                 # SheetComposer
│       ├── slide.py                 # SlideComposer
│       ├── pdf.py                   # Cross-platform PDF editing
│       ├── design_presets.py        # Five design presets
│       ├── layout_templates.py      # Slide layouts
│       └── quality_checks.py        # Layout validation
├── tests/                          # Platform-independent pytest suite
├── install.py                     # Cross-platform marketplace installer
├── install.ps1                   # PowerShell wrapper
└── install.sh                    # macOS/Linux wrapper
```

- Import public APIs through `from skills.WPSComposer import ...`.
- Keep `wps_engine.py` as the backward-compatible public facade; implementation
  belongs in focused composer and infrastructure modules.
- `_dispatch.py`, `_colors.py`, and `_base.py` are internal modules.

## Build, Test, and Development Commands

Create the development environment and run the platform-independent suite:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
.venv/bin/python -m pytest -v
```

Install the local plugin through the Codex personal marketplace:

```bash
python3 install.py
python3 install.py --force
```

The PowerShell and POSIX wrappers call the same Python installer:

```bash
pwsh ./install.ps1
./install.sh
```

## Coding Style & Naming Conventions

- Python 3.9+, 4-space indentation, and `from __future__ import annotations`.
- Composer classes use `PascalCase`; methods use `snake_case`.
- Color arguments accept `#RRGGBB` strings or BGR integers.
- Preserve the public facade and update `SKILL.md` plus `references/api.md`
  whenever a public import or signature changes.

## Testing Guidelines

- Follow RED-GREEN-REFACTOR for behavior changes.
- Run `.venv/bin/python -m pytest -v` before committing.
- Keep platform-independent tests free of live COM and user-home writes.
- For native WPS layout changes, create native output plus separate PDF or
  screenshot evidence and inspect columns, wrapping, field results, and images.
- Public generation returns only the artifact format requested by the user;
  development PDF evidence is not an automatic public companion output.

## Commit & Pull Request Guidelines

- Use a short imperative commit subject.
- Reference the affected composer or subsystem in the body when useful.
- PRs for native layout changes include sample output or screenshots.

## Runtime Requirements

- Windows WPS/MS Office generation requires `pywin32`.
- PDF editing requires `pypdf` and `pdfplumber`; text watermarks also require
  `reportlab`.
- The parser, models, presets, and public package remain importable on macOS and
  Linux without `pywin32`.
