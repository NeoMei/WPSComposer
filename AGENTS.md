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
│       ├── wps_engine.py            # Backward-compatible public facade (re-exports)
│       ├── _dispatch.py             # COM dispatch and format constants
│       ├── _colors.py               # Unified colour model
│       ├── _base.py                 # Shared composer lifecycle
│       ├── writer.py                # WriterComposer (docx) + paraId helpers
│       ├── sheet.py                 # SheetComposer (xlsx)
│       ├── slide.py                 # SlideComposer (pptx)
│       ├── pdf.py                   # Cross-platform PDF editing
│       ├── document_api.py          # inspect/edit/apply_ops/validate_op orchestration
│       ├── orchestrator.py          # generate() — markdown → document
│       ├── md_parser.py             # Markdown → StructuredDocument
│       ├── document_model.py        # StructuredDocument data model
│       ├── conversion.py            # convert_to_pdf (Office→PDF)
│       ├── recording_composers.py   # COM-free recording doubles for generation plans
│       ├── generation_plan.py       # Closed generation plan validation
│       ├── design_presets.py        # Five design presets
│       ├── layout_templates.py      # Slide layouts
│       ├── quality_checks.py        # Layout validation
│       └── macos_probe/             # macOS WPS JSAPI bridge (generation/conversion only)
├── tests/                          # Platform-independent pytest suite
├── docs/windows-verification.md   # READ FIRST on Windows — COM verification handoff
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
- Preserve the public facade and update `skills/WPSComposer/SKILL.md` plus
  `skills/WPSComposer/references/api.md` whenever a public import or signature
  changes.

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

## Windows verification handoff (READ FIRST on Windows)

Several features were implemented on macOS but the COM-coupled parts could only
be **written blind** — they need a live WPS/Office host on Windows to verify
and fix. **Before any COM work on Windows, read `docs/windows-verification.md`.**

It covers:

- The full checklist (items A–G) with runnable verify scripts.
- The 33 `# WINDOWS-VERIFY:` markers in `skills/WPSComposer/scripts/` —
  `grep -rn WINDOWS-VERIFY skills/WPSComposer/scripts/` lists every site.
- Known COM-logic concerns flagged for Windows (e.g. Writer `doc.Range` vs
  `doc.Content`, heading-style early-return ordering, clipboard fidelity, shape
  paste targets, Sheet whole-sheet delete safety).

The macOS-verifiable layer (orchestration: `apply_ops`, `validate_op`,
`validate_target`, `snapshot_to_patches`, `_extract_paraids`, atomic `edit()`)
is fully tested (609 passing) and should not need rework — only the COM bodies
of `inspect_document` / `apply_format_patch` / `apply_structural_op` in
`writer.py`, `sheet.py`, `slide.py` need verification and fixes against a real
host. Update `docs/windows-verification.md` (and clear the `WINDOWS-VERIFY`
markers you resolve) as you go.
