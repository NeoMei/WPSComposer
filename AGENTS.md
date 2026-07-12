# Repository Guidelines

Guidelines for contributing to **WPSComposer**, a Codex plugin that generates
rich-layout DOCX / PPTX / XLSX / PDF documents via WPS Office COM automation.

## Project Structure & Module Organization

`	ext
WPSComposer/
├── .codex-plugin/plugin.json        # Plugin manifest (name, version, skill refs)
├── skills/WPSComposer/
│   ├── SKILL.md                     # Skill entry + frontmatter (name: WPSComposer)
│   ├── references/api.md            # Per-composer API reference
│   └── scripts/
│       ├── wps_engine.py            # Backward-compat re-export facade
│       ├── _dispatch.py             # COM dispatch, ProgIDs, format constants
│       ├── _colors.py               # Unified colour model (hex ↔ BGR, semantic names)
│       ├── _base.py                 # BaseComposer with shared save/export lifecycle
│       ├── writer.py                # WriterComposer (WPS Writer → docx)
│       ├── sheet.py                 # SheetComposer  (WPS Sheet  → xlsx)
│       ├── slide.py                 # SlideComposer  (WPS Present → pptx)
│       ├── pdf.py                   # PdfComposer    (pypdf/pdfplumber → pdf edit)
│       ├── design_presets.py        # DesignPreset + 4 presets + registration API
│       ├── layout_templates.py      # LayoutTemplate + 14 layouts + colour resolution
│       └── quality_checks.py        # Slide validation + WPS-aware checks
├── install.ps1                      # Publish plugin to ~/.codex/skills/
└── README.md
`

- `wps_engine.py` is now a thin re-export facade; real code lives in the
  individual composer modules and `_*` infrastructure files.
- `_dispatch.py`, `_colors.py`, `_base.py` are internal; public API is
  the four composer classes + companion modules.
- Old `from wps_engine import WriterComposer` imports continue to work.
- There is no `tests/` directory — validation is done by generating sample
  documents in a scratch subdirectory and visually checking the exported PDF.## Build, Test, and Development Commands

There is no build step. The workflow is **debug → publish**:

```powershell
pwsh ./install.ps1            # first-time publish to ~/.codex/skills/WPSComposer
pwsh ./install.ps1 -Force     # overwrite an already-published version
```

To test engine changes locally without publishing, run a small driver script
from the repo root so `wps_engine` resolves on `sys.path`:

```powershell
python -c "from skills.WPSComposer.scripts.wps_engine import WriterComposer; print('ok')"
```

Always export a PDF alongside the native format so layout can be visually
verified before declaring a change done.

## Coding Style & Naming Conventions

- Python 3, 4-space indentation, `from __future__ import annotations` at top of
  the module.
- Composers are class names in `PascalCase` (`WriterComposer`,
  `SlideComposer`, `SheetComposer`, `PdfComposer`); methods are `snake_case`.
- Keep `wps_engine.py` as a single cohesive module — do not split it without
  also updating the import paths documented in `SKILL.md` and `api.md`.
- Color arguments accept either `#RRGGBB` hex strings or BGR ints; reuse this
  convention for any new color-taking API.

## Testing Guidelines

There is no formal test framework. Conventions:

- Drop a throwaway driver script in any subdirectory and run it with `python`.
- Confirm every composer change by producing both the native file **and** a PDF,
  then inspect the PDF for correct columns, wrapping, and field results.
- Treat `SKILL.md`'s capability matrix as the regression checklist — every
  "yes" cell should still render correctly after your change.

## Commit & Pull Request Guidelines

This repo has no commit-message convention in history yet. When contributing:

- Use a short imperative subject line (e.g. `Add set_columns support to Sheet`).
- Reference the affected composer in the body when relevant.
- PRs must include sample output (native + PDF) or screenshots proving the
  layout capability that changed still works.

## Runtime Requirements

Windows with WPS Office (Writer/Spreadsheets/Presentation) installed; MS Office
is an automatic fallback via the ProgID chains in `wps_engine.py`. The runtime
needs `pywin32`; PDF editing additionally needs `pypdf` and `pdfplumber`.
