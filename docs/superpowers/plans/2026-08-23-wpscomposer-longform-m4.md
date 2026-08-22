# WPSComposer Long-form M4 Formulas, Bibliography & Degradation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox syntax.

**Goal:** Complete M4 with editable native Office Math/WPS formula content, deterministic numeric citations and bibliography layout, and one shared degradation system that always returns a marked artifact for recoverable local failures while stopping immediately for engine-level failures.

**Architecture:** Keep restricted-LaTeX validation/conversion, citation ordering, fallback selection, and operation schemas platform-independent. Emit a closed protocol-v2 plan containing trusted native-math descriptors, ordered citation runs, structured bibliography entries, and deterministic degradation descriptors. Windows COM and macOS JSAPI map those descriptors to native `OMaths`, paragraph/range, table, bookmark, and field APIs without reinterpreting policy. A shared recovery controller owns the whitelist, one simplified retry, visible notice placement, issue redaction, and fatal-error boundary.

**Tech Stack:** Python 3.9+, pytest, Pillow-backed private resources from M3, existing generation protocol v2, pywin32 COM mocks, WPS JSAPI loopback add-in, real macOS WPS evidence, read-only OOXML/PDF inspection.

## Global Constraints

- Implement only M4. PDF-driven quality checks, final pagination/bbox mapping, automatic full re-layout, notice-only patching, performance gate, public `generate()` default migration, release/version changes, and final Windows real-WPS verification remain M5/final-gate work.
- Keep `protocolVersion: 2`, `semanticVersion: longform-1`, and `resourceManifestVersion: 1`. Every nested M4 object is closed, bounded, JSON-only, rejects unknown fields, and validates before WPS starts.
- A valid formula must be inserted as an editable native `OMath`/WPS math object. A borderless table is only the center/number layout container. Plain source text or an image is a marked runtime fallback and never counts as native-formula success.
- Formula input remains the existing restricted LaTeX subset: at most 10,000 Unicode code points and 64 brace levels; no custom macros, packages, file/shell operations, URLs, or external resources. Unknown/forbidden commands are normalized to a planned degradation before WPS starts and are never sent to `OMaths.Add`.
- `fallback_image` is optional and local-only. It uses M3's signature, size, pixel, hash, normalization, staging, and privacy rules. It is used only after a named native-math creation failure; the system never invokes a non-WPS formula renderer.
- Numeric citations are assigned by first semantic-source occurrence across abstract, lists, page-break content, and body paragraphs. Repeated citations reuse the number. Declared but uncited entries follow cited entries in declaration order unless `bibliography_include_uncited: false`.
- Missing citations stay in the same paragraph as a short styled inline degradation span. Bibliography entries use native paragraphs with a consistent hanging indent, left alignment, no artificial blank paragraphs, and identical spacing for Chinese and Latin entries.
- Every recoverable runtime path is: native attempt -> at most one declared simplified fallback -> visible inline/block/document notice at the same stable node -> continue. Unknown exceptions, rollback failure, engine acquisition, protocol/capability, staging/hash, required field/repagination, save/export/validation/publication, and cleanup failures remain fatal.
- A document that simply has no formula, citation, bibliography, image, or other optional object receives normal layout with no issue or notice.
- Visible notices contain only a stable code, controlled object label, readable reason, and actual fallback. They never contain source paths, payloads, field results/hashes, bookmark maps, exception reprs, or private resource IDs.
- Inline notices must remain in the owning paragraph. Block notices use a restrained native one-cell box and stay near their node. Document notices are deduplicated in one fixed `生成质量提示` region outside headings, TOC/index fields, headers, and footers.
- Windows implementation remains mock-tested on macOS. Real Windows M4 verification is deferred until M5 and all local work are complete. M4 itself requires a non-skipped real macOS WPS DOCX/PDF/reopen test and manual screenshot inspection.
- Public generation still returns only the format requested by the caller; development PDFs/screenshots remain internal evidence.
- `uv.lock` remains untouched. Follow RED-GREEN-REFACTOR and commit each task only after focused tests pass.

## Task List

### Task 1: Unified restricted-math grammar and trusted WPS-linear conversion

Files: modify `skills/WPSComposer/scripts/longform/formula.py`; create `skills/WPSComposer/scripts/longform/native_math.py` and `tests/longform_m4/test_native_math.py`.

Interfaces: add immutable `NativeMathDescriptor(syntax: str, linear_text: str, source_hash: str)` and `convert_restricted_latex(source) -> NativeMathDescriptor`. `syntax` is fixed to `wps-linear-v1`. Validation and conversion share one bounded tokenizer/parser: a source is native-valid only when conversion succeeds, so the validator can never accept a construct the converter later rejects.

- [ ] Write failing conversion tests for plain expressions, superscript/subscript, fractions, square/nth roots, sums/products/integrals with limits, scalable delimiters, Greek letters, matrices/cases, common relations/operators, nested constructs, Unicode input, and whitespace normalization.
- [ ] Write failing safety tests proving forbidden/unknown commands, malformed environments/groups, excess token/depth/row/column counts, external references, and control characters are rejected before a descriptor exists.
- [ ] Write failing determinism tests for NFC normalization, stable `source_hash`, JSON/repr privacy, and identical output on repeated calls.
- [ ] Implement a small recursive parser and controlled WPS-linear serializer; do not invoke LaTeX, MathML, a browser, or any non-WPS renderer.
- [ ] Rerun focused tests plus `tests/longform/test_formula.py`.
- [ ] Commit: Convert restricted formulas to WPS linear math.

### Task 2: Formula resources, citation runs, and bibliography page semantics

Files: modify `skills/WPSComposer/scripts/document_model.py`, `skills/WPSComposer/scripts/longform/md_parser_longform.py`, `skills/WPSComposer/scripts/longform/semantic.py`, `skills/WPSComposer/scripts/longform/resources.py`, `skills/WPSComposer/scripts/longform/pipeline.py`, `skills/WPSComposer/scripts/longform/page_policy.py`, and `skills/WPSComposer/scripts/longform/policy.py`; create `tests/longform_m4/test_semantic_formulas.py`, `tests/longform_m4/test_semantic_citations.py`, and `tests/longform_m4/test_formula_resources.py`; extend parser/resource/page-policy tests.

Interfaces: `FormulaBlock` gains `fallback_image: Optional[str]`, a validated `NativeMathDescriptor`, and an optional resolved fallback-resource binding that stays private. Add immutable `CitationRun(node_id, target_id, target_node_id, number, fallback_text)` and `InlineDegradationRun`; `Span` can own a citation or same-paragraph degradation run. Bibliography references gain deterministic `number`, `declaration_index`, and `cited` metadata. Add a `bibliography`/back-matter page role after body with continued Arabic page numbering and isolated header/footer linkage.

- [ ] Write failing parser tests for canonical `:::equation` and legacy `:::formula`, optional `fallback_image`, empty/escaped values, and preservation of raw source.
- [ ] Write failing semantic tests using the Task 1 grammar: conversion success remains a formula; forbidden, unknown, overlong, malformed, over-deep, or valid-but-unconvertible input becomes a same-node planned block degradation with readable source preserved and no native descriptor.
- [ ] Write failing resource tests proving only an explicitly declared formula fallback image is scanned, signature/preflighted, manifested, bound exactly once, staged privately on both executors, and rejected on missing/hash mismatch without leaking its path. A valid resource not referenced by its formula and one resource referenced by two formula nodes both fail before WPS starts.
- [ ] Write failing citation tests for first-occurrence numbering across abstract/list/page-break/body content, repeated citations, multiple citations in one paragraph, code/math literalness, missing IDs, duplicate IDs, deterministic occurrence IDs, and same-paragraph inline degradation runs.
- [ ] Write failing bibliography tests for cited-first ordering, uncited declaration order, `bibliography_include_uncited: false`, empty/malformed entries, no-reference documents, and byte-stable canonical JSON.
- [ ] Write failing page-policy tests for bibliography after all body sections, continued Arabic numbering, correct link-to-previous behavior, no empty bibliography section, and unchanged cover/front-matter/body/landscape ordering.
- [ ] Implement parsing, shared-grammar integration, citation/degradation span splitting, numbering, bibliography role, formula-resource scanning/binding, and canonical serialization without WPS/PDF imports.
- [ ] Rerun focused tests plus all M1-M3 parser/semantic/resource/page-policy tests.
- [ ] Commit: Resolve M4 formula, citation, and bibliography semantics.

### Task 3: Closed M4 plan schemas and bibliography/citation emission

Files: modify `skills/WPSComposer/scripts/generation_plan.py`, `skills/WPSComposer/scripts/longform/plan.py`, and `skills/WPSComposer/scripts/recording_composers.py`; create `tests/longform_m4/test_plan_schema.py` and `tests/longform_m4/test_plan.py`.

Interfaces: `writer.add_equation` gains required `nativeMath`, optional `fallbackResourceId`, and `fallbackMode`; its failure policy becomes closed degrade mode for the existing stable `EQUATION_INSERT_FAILED` code with `explicit-image-then-source-notice`. `writer.add_cross_reference.runs` accepts strict `citation` and `degradation` variants so unresolved references/citations remain in their owning paragraph. `writer.add_bibliography.entries` becomes a bounded list of `{id, nodeId, number, text, cited}` with `hangingIndentPt`, `leftIndentPt`, `spaceAfterPt`, and fixed numeric style.

- [ ] Write failing schema tests for strict native-math syntax, length/hash shape, optional formula resource binding, citation numbers, bibliography entry bounds/uniqueness/order, hanging-indent geometry, and unknown-field rejection at every nested level.
- [ ] Write failing schema tests rejecting raw paths, bytes/base64, arbitrary field codes, unvalidated formula source as native math, non-positive/duplicate citation numbers, citations without a matching bibliography target, and invalid recoverable-code/fallback pairs.
- [ ] Write failing plan tests for literal/reference/citation ordering in one paragraph, cited-first bibliography order, uncited omission/configuration, no empty bibliography operation, native formula descriptors, explicit-image fallback binding, and preflight-invalid formula degradation.
- [ ] Extend the state-machine validator so one semantic node has one owner, each citation exactly matches one final bibliography entry/number, bibliography is emitted in its back-matter section after body content, and M1-M3 plans stay valid.
- [ ] Add one fixed document-quality anchor immediately after the title/cover transition and before TOC/index/body content. Validate that document-level notices can only upsert at this anchor; they may not be appended at plan end or inserted in headings, fields, headers, or footers.
- [ ] Implement closed validators, plan emission, and recording mirrors.
- [ ] Rerun focused tests plus all generation-plan and M1-M3 plan tests.
- [ ] Commit: Emit closed M4 math and bibliography plans.

### Task 4: Shared whitelist recovery and styled degradation primitives

Files: create `skills/WPSComposer/scripts/longform/degradation.py`; modify `skills/WPSComposer/scripts/longform/executor.py`, `skills/WPSComposer/scripts/longform/field_contract.py`, `skills/WPSComposer/scripts/longform/windows_executor.py`, `skills/WPSComposer/scripts/longform/macos_executor.py`, `skills/WPSComposer/scripts/writer.py`, and `macos/wps-jsapi-probe/addin/writer-longform-v2.js`; create `tests/longform_m4/test_degradation_contract.py` and extend M3 field-convergence tests.

Interfaces: add immutable `DegradationDescriptor(code, placement, object_label, reason, fallback_text, fallback_kind)` with redacted serialization and `RecoveryDecision`. `run_recoverable_operation(...)` accepts an exact whitelist and one fallback callback, records one issue, inserts/deduplicates the matching visible notice, and rethrows everything outside the contract.

- [ ] Write failing matrix tests for every recoverable code already declared by M1-M4 and every fatal category; assert unknown errors, engine loss, save/field/index/configuration/cleanup errors, fallback failure, and rollback failure terminate.
- [ ] Write failing tests proving one and only one simplified attempt, stable issue order, same-node placement, nested child locality, deduplication, and no catch-all continuation.
- [ ] Write failing privacy tests for paths, payload/resource identifiers, hashes, exception reprs, field results, and bookmark maps in issue/notice/evidence serialization.
- [ ] Write failing Writer-mock tests for inline styling in the owning paragraph, restrained one-cell block notice styling, and one document-quality region outside outline/TOC/header/footer stories.
- [ ] Write failing field-convergence tests proving three unstable mutation rounds upsert exactly one `FIELD_REFRESH_UNSTABLE` notice at the fixed quality anchor, run the required fourth round, refresh again after notice insertion, freeze the fourth snapshot, and return the issue without duplicating the notice on any pass. Missing/throwing field APIs remain fatal.
- [ ] Implement the shared controller and native notice primitives; route existing figure/table/reference and new formula/bibliography recovery through it without changing fatal boundaries.
- [ ] Rerun focused tests plus all M2/M3 executor, rollback, cleanup, and privacy tests.
- [ ] Commit: Centralize marked degradation recovery.

### Task 5: Windows COM native formulas, citations, and bibliography

Files: modify `skills/WPSComposer/scripts/writer.py` and `skills/WPSComposer/scripts/longform/windows_executor.py`; create `tests/longform_m4/test_windows_executor_m4.py` and extend COM fakes.

Interfaces: add `WriterComposer.add_equation_native(...)`, `add_citation_paragraph(...)` through the existing run paragraph primitive, and `add_bibliography_native(...)`. Remove bibliography from `_M2_DEFERRED_OPERATIONS`. Native formula insertion writes trusted linear text to the center cell, calls `document.OMaths.Add(range)` and `BuildUp()`, verifies the OMath count/range, then adds the unchanged M3 number/bookmark shell in the right cell.

- [ ] Write failing COM-mock tests for a borderless formula/number layout table, centered editable OMath content, right-aligned native number, keep-together behavior, bookmark/reference refresh, and native-object verification.
- [ ] Write failing formula fallback tests: named OMath failure uses a valid explicit image then a visible mark; absent/failed image preserves source text and a visible mark; later operations continue. Unknown COM, number/bookmark, field, rollback, or container failures abort.
- [ ] Write failing citation/bibliography tests for inline `[n]` in the owning paragraph, repeated number reuse, left-aligned hanging-indent native paragraphs, fixed spacing, no blank separator paragraphs, and exact final order.
- [ ] Write failing lifecycle/privacy tests for reopen persistence, private resource cleanup, hash mismatch, and no path/payload leakage.
- [ ] Implement COM mappings with `DispatchEx` ownership unchanged and no OOXML mutation.
- [ ] Rerun focused tests plus all Windows M2/M3 tests.
- [ ] Commit: Render M4 content through Windows COM.

### Task 6: macOS JSAPI native formulas, citations, and bibliography

Files: modify `skills/WPSComposer/scripts/longform/macos_executor.py`, `skills/WPSComposer/scripts/macos_probe/models.py`, `skills/WPSComposer/scripts/macos_probe/runtime.py`, `skills/WPSComposer/scripts/macos_probe/templates.py`, `macos/wps-jsapi-probe/addin/writer-longform-v2.js`, and generated assets; create `tests/longform_m4/test_macos_executor_m4.py`.

Interfaces: add-in handlers consume exactly the Task 3 descriptors and return only applied-operation status, stable issue codes, closed child results, field snapshots, and pagination metadata. Bibliography leaves `LONGFORM_DEFERRED`; formula handler upgrades M3's text shell to native `OMaths.Add`/`BuildUp` with the shared fallback ladder.

- [ ] Write failing Python bridge tests for request shape, staged fallback-image binding, response closure, cleanup on every exit, and fatal capability/protocol mismatch.
- [ ] Write failing JavaScript asset/fake tests for native OMath creation/build-up/verification, number-shell preservation, inline citations, hanging bibliography, and the three notice placements.
- [ ] Write failing recovery tests for formula image/source fallback, bibliography paragraph-local fallback, exact whitelist handling, one fallback attempt, later-operation continuation, and fatal engine/save/field/rollback failures.
- [ ] Implement the add-in/Python mappings and regenerate/verify bundled assets through the existing template path.
- [ ] Rerun focused tests plus all macOS M2/M3 and add-in asset tests.
- [ ] Commit: Render M4 content through macOS JSAPI.

### Task 7: M4 acceptance fixtures and real macOS WPS evidence

Files: create `tests/longform_m4/fixtures/`, `tests/longform_m4/snapshots/`, `tests/longform_m4/test_acceptance_m4.py`, `tests/longform_m4/test_macos_real_wps_m4.py`, and `skills/WPSComposer/scripts/longform_m4_evidence.py`; produce internal artifacts under `build/longform-m4/macos-native-<run-id>/`.

Interfaces: fixtures cover every supported formula family, chapter/global formula numbering and references, cited/uncited bibliography order, repeated/missing citations, valid/invalid fallback images, all notice placements, named recoverable failures, and fatal engine scenarios. Evidence JSON contains only WPS version, relative artifact names, artifact digests, native-object/count/order metrics, issue codes, refresh rounds, and screenshot names.

- [ ] Add offline snapshots proving exact native-math descriptors, citation/bibliography order, closed recovery plans, and deterministic privacy-safe failures.
- [ ] Add a bridge-gated real-WPS test that creates, saves, closes, reopens, refreshes, and verifies editable native formulas remain `OMaths`; formula numbers/references remain correct after moving a formula; citations and bibliography remain ordered/styled.
- [ ] Add real recoverable-failure fixtures proving source/image/bibliography/reference degradation remains visible at the correct location and a valid DOCX is still returned; separately assert an injected engine-level failure returns no public artifact.
- [ ] Export internal PDF/screenshots and visually inspect formula centering/number alignment, citation paragraph continuity, bibliography hanging indent/spacing, inline/block/document notices, absence of clipping/overlap/blank pages, and absence of diagnostic placeholders.
- [ ] Require the real macOS test to execute without skip before M4 completion; validate evidence privacy and artifact hashes.
- [ ] Commit: Add M4 native WPS acceptance evidence.

### Task 8: M4 pipeline integration, documentation, ledger, and final regression

Files: modify `skills/WPSComposer/scripts/longform/pipeline.py`, `skills/WPSComposer/scripts/longform/__init__.py`, `skills/WPSComposer/SKILL.md`, `skills/WPSComposer/references/api.md`, and `.superpowers/sdd/progress.md`; extend import-purity, cleanup, compatibility, and privacy tests.

Interfaces: retain `build_longform_generation(...)` and `execute_longform_plan(...)`. Private formula fallback payloads use the M3 resource lifecycle and are released on success/error/timeout. No public `generate()` routing change occurs in M4.

- [ ] Write failing integration tests proving platform-pure offline build, no WPS launch during parsing/planning, private payload cleanup, no-path diagnostics, M1-M3 snapshot compatibility, and no optional-content notices.
- [ ] Document equation syntax/fallbacks, citation ordering/configuration, notice meanings/placement, recoverable versus fatal boundaries, and the M4/M5 boundary.
- [ ] Update the progress ledger with each task commit/review/evidence path and carry forward only PDF quality/re-layout, performance, default migration, final cross-platform gate, and release work.
- [ ] Run all M4 tests, all longform tests, generation-plan/recording/add-in/privacy tests, then `.venv/bin/python -m pytest -v`.
- [ ] Perform an independent whole-M4 review against this plan and the design spec; fix every critical, important, or minor discrepancy and rerun fresh tests.
- [ ] Commit: Complete M4 pipeline and documentation.

## Scope Boundaries

### In Scope (M4)

- Validated restricted-LaTeX conversion to editable native WPS/Office Math.
- Formula numbering/reference shell preservation and exactly-bound explicit image/source fallback with visible marking.
- First-occurrence numeric citations, repeated-number reuse, configured uncited handling, and native bibliography styling.
- One shared, closed, privacy-safe degradation controller and native inline/block/document notices.
- Windows COM mock implementation and real macOS WPS acceptance.

### Out of Scope (M5/final gate/release)

- PDF DPI/label OCR, page-role geometry checks, final bbox issue mapping, deterministic full re-layout, notice-only patch, performance fixture, pagination-map completion, and public default migration.
- Real Windows WPS execution before all local milestones finish.
- CSL, BibTeX import, external bibliography managers, arbitrary TeX/macros/packages, automatic formula-to-image rendering, MathML/OOXML post-processing, or network resources.
- Version bump, marketplace publication, release notes, and 0.8.0 release.

## Verification

- Semantic/math/plan: `.venv/bin/python -m pytest tests/longform_m4/test_semantic_formulas.py tests/longform_m4/test_semantic_citations.py tests/longform_m4/test_native_math.py tests/longform_m4/test_plan_schema.py tests/longform_m4/test_plan.py -v`.
- Recovery/executors: `.venv/bin/python -m pytest tests/longform_m4/test_degradation_contract.py tests/longform_m4/test_windows_executor_m4.py tests/longform_m4/test_macos_executor_m4.py -v`.
- Offline acceptance: `.venv/bin/python -m pytest tests/longform_m4/test_acceptance_m4.py -v`.
- Real macOS acceptance: `.venv/bin/python -m pytest tests/longform_m4/test_macos_real_wps_m4.py -v`; a skip is not acceptable for completion on this machine.
- Regression: `.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/longform_m3 tests/longform_m4 tests/test_generation_plan.py tests/test_recording_composers.py tests/longform_m0/test_addin_assets.py tests/macos_probe -v`.
- Full suite: `.venv/bin/python -m pytest -v`.

## Key Design Decisions

1. **Formula conversion is bounded and platform-independent.** Both executors receive the same trusted WPS-linear descriptor; neither parses raw LaTeX or chooses a fallback policy.
2. **Native math and formula numbering stay separate.** M4 replaces only M3's readable content shell. The existing native number/bookmark contract remains stable.
3. **Explicit images are fallback resources, never formula success.** WPS is always tried first. A fallback image is used only after the named native-math failure and is visibly marked.
4. **Citation numbers are semantic, not pagination-derived.** First source occurrence fixes the number before execution, so both platforms and later re-layouts agree.
5. **Bibliography entries are structured plan objects.** Executors receive final order and geometry, not loose strings that require platform-specific parsing.
6. **One recovery controller owns all local degradation.** This prevents divergent whitelist, retry, issue, notice, and privacy behavior across Windows/macOS and across figures/tables/formulas/references.
7. **Engine loss never degrades.** If the engine or required document APIs disappear, returning a plausible-looking partial file would violate the user's explicit boundary; execution stops.

## Open Questions

1. No blocking M4 product question remains. M0 already proved `OMaths.Add`/`BuildUp` and reopen on both platforms; M4 turns that primitive into the production descriptor and recovery path.
2. WPS linear-math serialization may expose platform quirks for a specific supported construct. The shared converter must use the proven common subset; any construct that cannot remain editable and equivalent on real macOS is removed from the accepted subset and deterministically degraded rather than implemented with static OOXML or a hidden renderer.
3. Real Windows parity remains intentionally unconfirmed until M5 is complete; mocks must still assert the exact COM calls and fatal/recoverable boundaries now.
