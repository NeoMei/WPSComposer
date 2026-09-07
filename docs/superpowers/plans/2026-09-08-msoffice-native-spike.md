# MS Office Native Spike Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task.

**Goal:** Obtain native Word/PDF evidence for the proposed dual-platform Office backend and an honest capability matrix.

**Architecture:** Standalone explicit-MS-Office probes; no production routes change. A Mac AppleScript runner and Windows COM runner emit comparable observations; verification reads the resulting DOCX/PDF.

**Tech Stack:** Python 3.9+, macOS AppleScript, Windows pywin32, pdfplumber.

**Spec:** docs/superpowers/specs/2026-09-08-msoffice-native-spike-design.md

## Global Constraints

Only task-created documents; no global template/security mutation, no broad quit/kill, no WPS substitution. Native generation/save/reopen/export and caller document preservation are required evidence. Keep original exceptions. A missing Windows environment is a documented unrun gate, never a pass. No public API, install, release or production implementation changes.

### Task 1: Native Word probe and evidence

**Files:** Create fixtures/msoffice_spike/mac_word.py, fixtures/msoffice_spike/windows_word.py, fixtures/msoffice_spike/README.md and docs/msoffice-native-spike.md. Scratch artifacts live in build/msoffice-spike.

**Interfaces:** Each runner takes --output-dir PATH (new directory) and writes result.json plus probe.docx/probe.pdf on success. Results distinguish attempted/succeeded/failed/unrun operations, actual engine/version, existing-document snapshots, owned document identity, save/reopen/page count, failures and cleanup. Windows explicitly uses DispatchEx("Word.Application"); macOS uses the installed Word.sdef dictionary and exact task document references.

- [x] Inspect Word dictionaries/current COM implementation and compile an initial Mac script before native changes.
- [x] Implement a minimal native document/save/reopen/PDF probe, then expand it one capability at a time to headings/native numbering, TOC refresh and multipage tables. Unsupported operations are failures or gaps, never silently skipped successes.
- [x] Execute Mac probe in a fresh output directory; retain errors from every attempt. Verify Word-generated DOCX semantics and PDF text/pages/layout, and compare pre-existing documents before/after.
- [x] Prepare and syntax-check the Windows equivalent. Run natively only when an authorized Windows environment is available; otherwise record the exact handoff command and unrun state.
- [x] Review scripts for document ownership and stale-artifact hazards; run appropriate platform-independent checks and write the capability matrix/recommendation.

Verification commands:
```sh
python3 -m py_compile fixtures/msoffice_spike/mac_word.py fixtures/msoffice_spike/windows_word.py
python3 fixtures/msoffice_spike/mac_word.py --output-dir build/msoffice-spike/mac-run-01
# Windows desktop only:
python fixtures/msoffice_spike/windows_word.py --output-dir build/msoffice-spike/windows-run-01
```

Prototype scripts are investigation tools, not production features. Native output checks are the acceptance tests; no tests that merely mirror the script text. Do not commit until the repository-required full suite has passed, or leave the bounded probe uncommitted with its validation status stated.

Status: bounded probe implementation and Mac evidence complete; Windows native execution remains pending the user-provided connection. Task review and whole-branch review approved with no Critical/Important findings.
