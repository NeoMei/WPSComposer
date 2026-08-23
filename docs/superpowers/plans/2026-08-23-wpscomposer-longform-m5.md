# WPSComposer Longform M5 Implementation Plan

> **For Codex:** REQUIRED SUB-SKILL: Use `test-driven-development` for every behavior change, `systematic-debugging` for failures, and `verification-before-completion` before each completion claim. Agentic workers may review independent completed tasks, but the primary implementation remains serial because the runtime lifecycle and public routing share state.

**Goal:** Complete the PDF-driven long-form quality gate, bounded recovery lifecycle, default DOCX/PDF migration, performance evidence, and release-facing documentation without changing PPTX/XLSX behavior.

**Architecture:** Add a platform-pure quality model and PDF analyzer beneath a private long-form lifecycle coordinator. The coordinator owns one absolute deadline, at most two native WPS generation passes, one notice-only patch, and at most three PDF exports. Platform adapters perform native generation/export/patch actions; the public orchestrator selects the long-form route only for DOCX/PDF unless frontmatter explicitly requests deprecated `legacy`. Quality findings map through the executor pagination map, and only high-confidence findings in the closed repair matrix may cause one full relayout.

**Tech Stack:** Python 3.9+, frozen dataclasses and Protocols, `pypdf>=4`, `pdfplumber>=0.11`, `Pillow>=10`, pytest, WPS Windows COM, macOS WPS JSAPI loopback bridge.

**Global Constraints:** Preserve `generate()` signature and absolute-path return. Publish only the requested artifact. Missing core analyzers fail before WPS starts. No PDF text guessing for node identity. Never exceed two full native generations, one notice patch, or three exports. Warning/info never mutate content. A notice patch never changes body layout. Engine/protocol/save/export/validation/publish failures abort with no partial public artifact. PPTX/XLSX recordings and routes remain byte-for-byte compatible. Do not bump or release 0.8.0 until Windows evidence passes.

---

### Task 1: Define typed M5 quality and lifecycle contracts

**Files:**
- Create: `skills/WPSComposer/scripts/longform/quality.py`
- Modify: `skills/WPSComposer/scripts/longform/executor.py`
- Modify: `skills/WPSComposer/scripts/longform/__init__.py`
- Create: `tests/longform_m5/test_quality_contract.py`

**Interfaces:**
- `QualitySeverity`, `QualityConfidence`, `PageRole`, `QualityEvidence`, `QualityFinding`, `QualityReport`
- `GenerationOutcome(path, degraded, issues)` as private outcome
- Pagination bounds become a validated point tuple `[x0, y0, x1, y1]`; pagination version becomes `M5-v1` for real maps while accepting legacy M2 snapshots on input.

**RED:** Add round-trip, closed-enum, finite-bounds, severity/confidence, deterministic ordering, redaction, and degraded aggregation tests. Run `../../.venv/bin/python -m pytest tests/longform_m5/test_quality_contract.py -q`; expect failures because contracts do not exist.

**GREEN:** Implement only the typed contracts and serializers. Reject non-positive pages, inverted/non-finite bounds, uncontrolled evidence keys, and path-like evidence values. Preserve old `PaginationMap.from_dict()` compatibility.

**VERIFY:** Run the new tests plus `tests/longform/test_executor.py`, `tests/longform_m2/test_acceptance_m2.py`, and `tests/longform_m4/test_task8_pipeline_integration.py`.

**Commit:** `Define M5 quality contracts`

### Task 2: Add core dependency preflight and normalized PDF page model

**Files:**
- Create: `skills/WPSComposer/scripts/longform/pdf_quality.py`
- Modify: `pyproject.toml`
- Create: `tests/longform_m5/test_pdf_quality_dependencies.py`
- Create: `tests/longform_m5/test_pdf_page_model.py`

**Interfaces:**
- `require_quality_dependencies()` imports and version-checks Pillow, pypdf, and pdfplumber without starting WPS.
- `load_pdf_pages(path) -> tuple[PdfPage, ...]` validates PDF structure, normalizes rotation/MediaBox/CropBox into top-left point coordinates, and records bounded character/object geometry.

**RED:** Test missing modules, too-old versions, corrupt PDFs, rotated/cropped pages, parser page-count disagreement, non-finite geometry, and no path leakage.

**GREEN:** Promote pypdf/pdfplumber into base project dependencies; retain reportlab in the optional PDF-edit extra. Reuse `artifact_transport.validate_pdf` and M0 page-box logic where possible.

**VERIFY:** New tests plus `tests/test_artifact_transport.py`, `tests/test_pdf_composer.py`, and a clean editable install dependency resolution check.

**Commit:** `Add longform PDF analysis foundation`

### Task 3: Implement role-aware deterministic PDF checks

**Files:**
- Modify: `skills/WPSComposer/scripts/longform/pdf_quality.py`
- Modify: `skills/WPSComposer/scripts/longform/quality.py`
- Create: `tests/longform_m5/fixtures/pdf_cases.py`
- Create: `tests/longform_m5/test_pdf_quality_rules.py`

**Interfaces:**
- `analyze_pdf(pdf_path, pagination_map, page_roles, policy) -> QualityReport`
- Closed checks: unexpected blank page, body-boundary overflow, header/footer/page-number sequence, orphan heading, caption separation, image width/DPI, table width/header/split, TOC typography/spacing/utilization, caption/reference field errors, bibliography spacing, and final-page utilization.

**RED:** Generate synthetic vector PDFs for every rule and negative control. Assert confidence ceilings: structural/bbox may be high, extraction heuristics at most medium, missing vector geometry low.

**GREEN:** Implement deterministic checks. Map findings only through pagination fragments/page spans; never infer node IDs from extracted text. Exempt cover, explicit chapter starts, and explicit page breaks from utilization findings.

**VERIFY:** Rule suite twice with stable serialized reports; validate six acceptance fixture policies offline.

**Commit:** `Implement role-aware PDF quality checks`

### Task 4: Implement the closed one-pass relayout planner

**Files:**
- Create: `skills/WPSComposer/scripts/longform/relayout.py`
- Modify: `skills/WPSComposer/scripts/longform/plan.py`
- Create: `tests/longform_m5/test_relayout_planner.py`

**Interfaces:**
- `build_relayout_directives(report, plan) -> RelayoutDecision`
- Closed directives: compact TOC style, shorten overflowing header to 32 display units, fit image to body width, low-DPI notice, compress table once, forced table-row split, heading keep-with-next, and caption keep-with-object.

**RED:** Test every matrix row, duplicate coalescing, deterministic order, no action for medium/low/warning/info, no action outside the matrix, and no second relayout.

**GREEN:** Emit a validated immutable directive set keyed by stable `node_id`. Extend the plan builder/executors only with closed M5 relayout arguments; keep the original semantic operation graph unchanged for identity comparisons.

**VERIFY:** M1-M4 snapshots unchanged when no directives are supplied; add Windows/macOS executor contract tests for each directive.

**Commit:** `Add bounded longform relayout directives`

### Task 5: Implement notice-only patch contracts on both executors

**Files:**
- Modify: `skills/WPSComposer/scripts/longform/macos_executor.py`
- Modify: `skills/WPSComposer/scripts/longform/windows_executor.py`
- Modify: `macos/wps-jsapi-probe/writer-addon/js/writer-longform-v2.js`
- Modify: `skills/WPSComposer/scripts/macos_probe/models.py`
- Create: `tests/longform_m5/test_notice_patch.py`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Executor capability `patch_quality_notices(staged_docx, notices, deadline) -> ExecutionOutcome`
- Each notice contains only stable code, controlled reason/fallback, placement, mapped node ID, and final page number. Patch calls `finalize_fields` once, saves, and returns a fresh pagination map.

**RED:** Test mapped insertion, de-duplication, exact one patch, no layout directive acceptance, field refresh, rollback, engine-loss abort, no public artifact, and no path/body/hash leakage.

**GREEN:** Add the closed protocol command and platform implementations. Reuse the M4 degradation writers and fatal boundaries.

**VERIFY:** JS syntax/assets, focused macOS/Windows executor suites, and a real macOS WPS patch/reopen probe.

**Commit:** `Add notice-only quality patching`

### Task 6: Build the bounded long-form lifecycle coordinator

**Files:**
- Create: `skills/WPSComposer/scripts/longform/lifecycle.py`
- Modify: `skills/WPSComposer/scripts/longform/pipeline.py`
- Create: `tests/longform_m5/test_lifecycle.py`

**Interfaces:**
- `run_longform_lifecycle(build, adapter, output, format_name, timeout, overwrite) -> GenerationOutcome`
- Adapter methods: `execute(plan, resources, directives, deadline)`, `export_pdf(docx, deadline)`, `patch_quality_notices(docx, notices, deadline)`, `validate_docx`, and `publish`.

**RED:** Cover clean first pass, one relayout, final degraded notice patch, warning/info no mutation, export/validation failures, timeout at every stage, cleanup on every exit, publication rollback, final issue page refresh, requested-PDF reuse, requested-DOCX internal-PDF deletion, and hard operation/export count caps.

**GREEN:** Implement a single monotonic deadline and explicit state machine. Reuse the original prepared resource payload across passes. Analyze initial and final PDFs; after a patch perform boundary/marker validation only, not another automatic-quality loop.

**VERIFY:** Lifecycle fault matrix and privacy hooks; ensure no temporary artifact remains after success/failure.

**Commit:** `Orchestrate the M5 quality lifecycle`

### Task 7: Add production platform adapters and public DOCX/PDF routing

**Files:**
- Create: `skills/WPSComposer/scripts/longform/platform_runtime.py`
- Modify: `skills/WPSComposer/scripts/orchestrator.py`
- Modify: `skills/WPSComposer/scripts/longform/macos_executor.py`
- Modify: `skills/WPSComposer/scripts/longform/windows_executor.py`
- Modify: `tests/test_generation.py`
- Create: `tests/longform_m5/test_public_routing.py`

**Interfaces:**
- Private `_generate_outcome(...)` chooses longform/legacy from normalized frontmatter.
- Public `generate()` returns only `.path`.
- DOCX/PDF default to M5 longform; `layout_engine: legacy` explicitly uses the old Writer route; protocol/capability mismatch never falls back. PPTX/XLSX remain on their existing routes.

**RED:** Test default route, explicit legacy, invalid layout value normalization, source file/raw text/base directory, plugins, timeout/overwrite, absolute return, only requested artifact, fatal no-output, and unchanged PPTX/XLSX snapshots.

**GREEN:** Delay all platform imports until after dependency/build validation. On macOS, own one `ProbeRuntime`/bridge across generation/export/patch passes. On Windows, use the dedicated worker/executor contract and never generic `Dispatch` fallback.

**VERIFY:** Full `tests/test_generation.py`, public import tests, installer tests, and clean-subprocess platform import tests.

**Commit:** `Route DOCX and PDF through longform by default`

### Task 8: Add real fixtures, performance accounting, and macOS visual acceptance

**Files:**
- Create: `tests/longform_m5/fixtures/` with six non-customer Markdown fixtures and controlled media
- Create: `skills/WPSComposer/scripts/longform_m5_evidence.py`
- Create: `tests/longform_m5/test_acceptance_m5.py`
- Create: `tests/longform_m5/test_performance_budget.py`
- Create: `tests/longform_m5/test_macos_real_wps_m5.py`

**Evidence:** Record hardware, OS, WPS/add-in/protocol/semantic versions, page count, operation count, each stage duration, generation/export/patch counts, issue summaries, hashes, and representative screenshots without customer content or paths.

**RED/GREEN:** First make deterministic offline fixture acceptance pass, then generate a synthetic 50-100-page performance fixture. Run the real macOS WPS path for all six fixtures, export PDF, reopen/refresh, inspect representative cover/TOC/body/image/table/formula/code/degradation/bibliography pages, and confirm total lifecycle under 600 seconds.

**VERIFY:** Repeat the real gate three times; hashes may differ but semantic/quality metrics and pass counts must remain stable. Confirm no unexpected blank pages, clipping, or notice/body overlap.

**Commit:** `Verify M5 on macOS WPS`

### Task 9: Update public documentation and migration metadata

**Files:**
- Modify: `README.md`
- Modify: `skills/WPSComposer/SKILL.md`
- Modify: `skills/WPSComposer/references/api.md`
- Create: `docs/longform-markdown.md`
- Modify: `.superpowers/sdd/progress.md`
- Create: `docs/macos-longform-m5-verification.md`
- Modify package/lock metadata only as required by the core dependency change; do not bump version yet.

**RED:** Extend documentation tests to require default route, explicit legacy escape, degradation/fatal boundaries, dependency/runtime requirements, timeout budget, grammar, Unicode/range/coordinates, resource privacy, and macOS evidence truth.

**GREEN:** Document only verified behavior. State that 0.8.0 remains unreleased until Windows passes; preserve the known macOS equation degradation truth.

**VERIFY:** Documentation/link/import checks, package build content inspection, install into a fresh temporary venv, and plugin manifest consistency.

**Commit:** `Document the M5 longform migration`

### Task 10: Multi-round full-system audit and Windows handoff

**Files:**
- Modify discovered bug locations and corresponding tests only.
- Create/update: `docs/windows-verification.md` with M5 commands and evidence schema.

**Round A:** Run the entire pytest suite, JS syntax/assets, package build/install, public DOCX/PDF/PPTX/XLSX route tests, privacy/security/failure matrices, and macOS real WPS/UI interaction suite. Fix every critical, important, or material issue via RED-GREEN-REFACTOR.

**Round B:** Repeat the entire suite from a clean worktree and fresh temporary install. Exercise real WPS menu/add-in registration, bridge registration/restoration, generation, PDF export, reopen, field refresh, notice patch, overwrite conflict, cancel/timeout, and no-engine failure. Visually inspect representative screenshots and PDFs.

**Round C:** Run an independent diff/architecture review and all affected tests again. A further round is required if any code or test changes after Round B, or if reviewers find any critical/important/material issue. Stop local cycling only when a fresh round makes no material change.

**Windows handoff:** Push the reviewed branch only after local closure. On Windows run the same six fixtures, native COM semantic/visual gate, 50-100-page performance budget, dedicated-worker timeout/ownership tests, default/legacy public routing, and PPTX/XLSX regression. Fix and repeat on both platforms if Windows changes shared code.

**Final verification commands:**
- `../../.venv/bin/python -m pytest -v`
- `git diff --check && git status --short`
- clean temporary `pip install .` and import/public smoke
- three successful real macOS evidence runs
- Windows evidence JSON/DOCX/PDF/screenshots before declaring 0.8.0 complete

**Commit:** `Complete M5 full-system verification`

---

## Plan self-review

- Coverage: Every M5 design requirement is assigned to a concrete task, test, interface, verification command, and commit boundary.
- Bounds: The lifecycle counts are explicit and testable; notice patch cannot silently become a third relayout.
- Compatibility: Public signature/return and PPTX/XLSX paths are protected; legacy is explicit and deprecated.
- Safety: Dependency failure precedes WPS startup; staging, timeout, publication, privacy, and engine-loss paths are tested.
- Evidence: Offline tests, real macOS WPS/UI/visual evidence, performance evidence, and a Windows gate are separate surfaces.
- Placeholders: None. Real filenames, interfaces, commands, and expected outcomes are specified.
