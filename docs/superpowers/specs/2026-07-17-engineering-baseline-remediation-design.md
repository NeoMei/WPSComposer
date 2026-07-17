# WPSComposer Engineering Baseline Remediation Design

Date: 2026-07-17
Status: Approved direction (Option B)

## Summary

Before implementing the macOS WPS JSAPI backend, WPSComposer needs a reliable
engineering baseline. The current Windows COM implementation is syntactically
healthy and its platform-independent parser and design modules import when the
repository is treated as a package, but the documented top-level imports fail,
the project has no automated test suite or dependency metadata, and the
installer uses the legacy skills directory rather than the current Codex local
plugin marketplace layout.

This change repairs those foundations without changing Windows document
generation behavior or beginning the macOS backend. It establishes one tested
Python import surface, current plugin metadata and cross-platform installation,
and a small regression suite that later backend work can build on.

## Goals

- Provide one stable Python API import rooted at `skills.WPSComposer`.
- Add reproducible project and optional dependency metadata.
- Add automated tests for imports, parsing, presets, plugin metadata, and local
  installation behavior.
- Install the complete plugin bundle using the current Codex personal
  marketplace layout on Windows and macOS.
- Correct deterministic code defects and documentation contradictions found
  during the repository audit.
- Leave the repository ready to begin the macOS feasibility gate.

## Non-goals

- Implementing `MacWpsJsBackend`, the WPS JS add-in, or the JSON-RPC bridge.
- Refactoring the existing Writer, Sheet, or Slide COM implementations.
- Replacing broad COM compatibility exception handling in this baseline pass.
- Deleting or rewriting the second repository copy under
  `/Users/neomei/项目/codexprojects/WpsComposer`.
- Publishing, pushing, or installing the plugin into the user's live Codex
  configuration during automated tests.

## Decisions

### Canonical repository

`/Users/neomei/项目/WpsComposer` is the implementation source for this work.
The other repository copy remains untouched. Its longer macOS design draft may
be reconciled separately after this baseline is complete.

### Python API surface

The public development import is:

```python
from skills.WPSComposer import generate, inspect, edit, attach_active
```

The repository will add package initializers at `skills/__init__.py`,
`skills/WPSComposer/__init__.py`, and `skills/WPSComposer/scripts/__init__.py`.
`skills.WPSComposer` re-exports the supported facade from
`scripts/wps_engine.py`. Existing internal relative imports remain unchanged,
and direct top-level imports such as `from orchestrator import generate` are no
longer documented as supported.

This avoids duplicating fallback-import logic across every module and keeps the
package boundary explicit.

### Test and dependency baseline

`pyproject.toml` records Python version support, pytest configuration, and
optional dependency groups:

- `windows`: `pywin32` for COM automation;
- `pdf`: `pypdf`, `pdfplumber`, and `reportlab`;
- `dev`: `pytest`.

The first test suite stays platform independent. It verifies:

- the public package import works on macOS without importing `pywin32`;
- Markdown parsing preserves representative headings, task lists, and tables;
- the documented format and preset registries match runtime values;
- plugin metadata follows the current bundle layout;
- the installer copies a plugin into an isolated temporary Codex home and
  merges its marketplace entry without removing unrelated entries.

Native Windows COM and macOS WPS integration tests remain separate future
suites because neither can be meaningfully executed on every development host.

### Codex plugin structure

The plugin remains rooted at the repository top level:

```text
.codex-plugin/plugin.json
skills/WPSComposer/SKILL.md
skills/WPSComposer/scripts/
```

The manifest uses a kebab-case plugin identifier (`wps-composer`) and points
`skills` at `./skills/`, while the skill name remains `WPSComposer` so existing
`$WPSComposer` prompts remain valid.

### Cross-platform installation

`install.py` becomes the source of truth for installation. It uses only the
Python standard library and supports `--force`, `--codex-home`, and
`--marketplace-root` so tests can target temporary directories.

The installer:

1. copies the plugin bundle to `<codex-home>/plugins/wps-composer`;
2. excludes repository-only and generated files such as `.git`, caches,
   virtual environments, tests, and local artifacts;
3. creates or merges `<marketplace-root>/.agents/plugins/marketplace.json`;
4. adds a local source path beginning with `./` and relative to the marketplace
   root, normally `./.codex/plugins/wps-composer`;
5. preserves unrelated marketplace entries;
6. refuses to replace an existing installation unless `--force` is supplied;
7. reports that the ChatGPT/Codex desktop app must be restarted and the plugin
   installed or enabled from the Plugins Directory.

`install.ps1` and a new `install.sh` are thin wrappers around `install.py` so
platform behavior cannot drift.

### Deterministic cleanup

This baseline removes the duplicate `_is_separator_cell` definition in
`md_parser.py`. It does not broadly rewrite defensive COM exception handling;
that requires backend-specific tests and belongs to later hardening.

Documentation is updated so that:

- examples use the supported package import;
- the preset count and names match runtime values;
- development verification may generate PDF evidence, but public generation
  returns only the requested artifact, matching the macOS design contract;
- `AGENTS.md` describes the already-split module architecture accurately;
- installation instructions describe the local marketplace flow rather than
  copying a nested plugin into `~/.codex/skills`.

## Error Handling

Installer failures use non-zero exit status and actionable messages for an
existing destination, malformed marketplace JSON, an invalid source bundle, or
an unwritable destination. Marketplace updates are written to a temporary file
and atomically replaced so interruption cannot leave partial JSON.

Tests never use the live home directory. Every filesystem test supplies
temporary `--codex-home` and `--marketplace-root` paths.

## Validation

The baseline is accepted when:

- the new regression suite passes from the repository root;
- the public import and parser smoke tests pass under Python 3 without pywin32;
- an isolated installer test produces the required plugin bundle and valid
  merged marketplace JSON;
- the manifest and documentation agree on plugin name, skill path, presets,
  imports, and output contract;
- Python source parses without syntax errors;
- `git diff --check` reports no whitespace errors;
- no live Codex installation, marketplace, or second repository copy changed.

## Follow-up

After this baseline is implemented and verified, work proceeds to Phase 0 of
the approved macOS backend design: align the Node/npm toolchain, install
`wpsjs`, create the smallest add-in and loopback bridge, and prove independent
DOCX, PPTX, XLSX, and PDF output paths before introducing the full RenderPlan
architecture.
