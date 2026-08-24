# WPSComposer Long-form M4 Formulas, Bibliography & Degradation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox syntax.

**Goal:** Complete M4 with editable native Office Math/WPS formula content, deterministic numeric citations and bibliography layout, and one shared degradation system that always returns a marked artifact for recoverable local failures while stopping immediately for engine-level failures.

**Architecture:** Keep restricted-LaTeX validation/conversion, citation ordering, fallback selection, and operation schemas platform-independent. Emit a closed protocol-v2 plan containing trusted native-math descriptors, ordered citation runs, structured bibliography entries, and deterministic degradation descriptors. Windows COM and macOS JSAPI map those descriptors to native `OMaths`, paragraph/range, table, bookmark, and field APIs without reinterpreting policy. Python is the canonical source for the recovery matrix and generates the frozen JS table; Windows Python and the macOS add-in each run the same closed local state machine because checkpoint/rollback cannot cross the bridge. Cross-platform parity tests lock code, fallback, placement, attempt count, redaction, and fatal boundaries.

**Tech Stack:** Python 3.9+, pytest, Pillow-backed private resources from M3, existing generation protocol v2, pywin32 COM mocks, WPS JSAPI loopback add-in, real macOS WPS evidence, read-only OOXML/PDF inspection.

## Global Constraints

- Implement only M4. PDF-driven quality checks, final pagination/bbox mapping, automatic full re-layout, notice-only patching, performance gate, public `generate()` default migration, release/version changes, and final Windows real-WPS verification remain M5/final-gate work.
- Keep `protocolVersion: 2`, `semanticVersion: longform-1`, and `resourceManifestVersion: 1`. Backward compatibility is mandatory: the closed M3 equation-shell shape and legacy string bibliography shape remain valid and executable. M4 adds separately discriminated closed shapes (`renderMode: native-m4` for equations and `schemaVersion: 1` for structured bibliography); only those shapes require native math/structured entries and the M4 recovery policy. Every nested M4 object is closed, bounded, JSON-only, rejects unknown fields, and validates before WPS starts.
- A valid formula must be inserted as an editable native `OMath`/WPS math object. A borderless table is only the center/number layout container. Plain source text or an image is a marked runtime fallback and never counts as native-formula success.
- Formula input remains the existing restricted LaTeX subset: at most 10,000 Unicode code points and 64 brace levels; no custom macros, packages, file/shell operations, URLs, or external resources. Unknown/forbidden commands are normalized to a planned degradation before WPS starts and are never sent to `OMaths.Add`.
- `fallback_image` is optional and local-only. It uses M3's signature, size, pixel, hash, normalization, staging, and privacy rules. It is used only after a named native-math creation failure; the system never invokes a non-WPS formula renderer.
- Numeric citations are assigned by first semantic-source occurrence across every M4-defined visible inline container: abstract/body paragraphs, list items, block-quote paragraphs, explicit page-break paragraphs, and table-cell static citation text. Repeated citations reuse the number. Declared but uncited entries follow cited entries in declaration order unless `bibliography_include_uncited: false`.
- Missing citations stay in the same paragraph as a short styled inline degradation span. Bibliography entries use native paragraphs with a consistent hanging indent, left alignment, no artificial blank paragraphs, and identical spacing for Chinese and Latin entries.
- Every recoverable runtime path is: native attempt -> at most one declared simplified fallback -> visible inline/block/document notice at the same stable node -> continue. Unknown exceptions, rollback failure, engine acquisition, protocol/capability, staging/hash, required field/repagination, save/export/validation/publication, and cleanup failures remain fatal.
- A document that simply has no formula, citation, bibliography, image, or other optional object receives normal layout with no issue or notice.
- Visible notices contain only a stable code, controlled object label, readable reason, and actual fallback. They never contain source paths, payloads, field results/hashes, bookmark maps, exception reprs, or private resource IDs.
- Inline notices must remain in the owning paragraph. Block notices use a restrained native one-cell box and stay near their node. Document notices are deduplicated in one fixed `生成质量提示` region outside headings, TOC/index fields, headers, and footers.
- Windows implementation remains mock-tested on macOS. Real Windows M4 verification is deferred until M5 and all local work are complete. M4 itself requires a non-skipped real macOS WPS DOCX/PDF/reopen test and manual screenshot inspection.
- Public generation still returns only the format requested by the caller; development PDFs/screenshots remain internal evidence.
- `uv.lock` remains untouched. Follow RED-GREEN-REFACTOR and commit each task only after focused tests pass.
- This worktree uses the repository virtualenv at `../../.venv/bin/python`; all commands below use that path rather than assuming a worktree-local `.venv`.

## Task List

### Task 1: Unified restricted-math grammar and trusted WPS-linear conversion

Files: modify `skills/WPSComposer/scripts/longform/formula.py`; create `skills/WPSComposer/scripts/longform/native_math.py` and `tests/longform_m4/test_native_math.py`.

Interfaces: add immutable `NativeMathDescriptor(syntax: str, linear_text: str, source_hash: str)` and `convert_restricted_latex(source) -> NativeMathDescriptor`. `syntax` is fixed to `wps-linear-v1`. `source_hash` is a controlled digest of the normalized formula text, is allowed in semantic/plan JSON alongside the already-visible linear formula, and is distinct from forbidden resource `sourceSha256`/`payloadSha256`; it is not copied into ordinary issues or evidence. Validation and conversion share one bounded tokenizer/parser: a source is native-valid only when conversion succeeds, so the validator can never accept a construct the converter later rejects.

- [ ] Write failing conversion tests for plain expressions, superscript/subscript, fractions, square/nth roots, sums/products/integrals with limits, scalable delimiters, Greek letters, matrices/cases, common relations/operators, nested constructs, Unicode input, and whitespace normalization.
- [ ] Write failing safety tests proving forbidden/unknown commands, malformed environments/groups, excess token/depth/row/column counts, external references, and control characters are rejected before a descriptor exists.
- [ ] Write failing determinism tests for NFC normalization, stable `source_hash`, JSON/repr privacy, and identical output on repeated calls.
- [ ] Implement a small recursive parser and controlled WPS-linear serializer; do not invoke LaTeX, MathML, a browser, or any non-WPS renderer.
- [ ] Rerun focused tests plus `tests/longform/test_formula.py`.
- [ ] Commit: Convert restricted formulas to WPS linear math.

### Task 2: Formula resources, citation runs, and bibliography page semantics

Files: modify `skills/WPSComposer/scripts/document_model.py`, `skills/WPSComposer/scripts/longform/md_parser_longform.py`, `skills/WPSComposer/scripts/longform/semantic.py`, `skills/WPSComposer/scripts/longform/resources.py`, `skills/WPSComposer/scripts/longform/pipeline.py`, `skills/WPSComposer/scripts/longform/page_policy.py`, and `skills/WPSComposer/scripts/longform/policy.py`; create `tests/longform_m4/test_semantic_formulas.py`, `tests/longform_m4/test_semantic_citations.py`, and `tests/longform_m4/test_formula_resources.py`; extend parser/resource/page-policy tests.

Interfaces: `FormulaBlock` gains only the user-authored relative `fallback_image` declaration and a validated `NativeMathDescriptor`; it never stores source paths, bytes, hashes, or prepared bindings. Opaque IDs and private bindings live only in `ResourcePreflight`/`PreparedLongformResource`. Add immutable `CitationRun(node_id, target_id, target_node_id, number, fallback_text)` and `InlineDegradationRun`; `Span` can own a citation or same-paragraph degradation run. Native run rendering covers abstract/body paragraphs, list items, block-quote paragraphs, and paragraphs preserved inside explicit page-break content. Table cells use a narrower explicit contract: resolved citations become deterministic static `[n]` text, while unresolved citations remain in the same cell as `[REFERENCE_UNRESOLVED 引用目标未解析]` plus cell-local degradation metadata/shading; table cells do not claim arbitrary rich-run support in M4. Bibliography references gain deterministic `number`, `declaration_index`, and `cited` metadata. Add a `bibliography`/back-matter page role after body with continued Arabic page numbering and isolated header/footer linkage.

- [ ] Write failing parser tests for canonical `:::equation` and legacy `:::formula`, canonical `:::bibliography` and legacy `:::references`, optional `fallback_image`, empty/escaped values, and preservation of raw source.
- [ ] Write failing semantic tests using the Task 1 grammar: conversion success remains a formula; forbidden, unknown, overlong, malformed, over-deep, or valid-but-unconvertible input becomes a same-node planned block degradation with readable source preserved and no native descriptor.
- [ ] Write failing pure resource tests proving only an explicitly declared formula fallback image is scanned, signature/preflighted, manifested, and bound exactly once as a private prepared resource. A missing/unreadable/invalid source image becomes a formula-local fallback-resource planned degradation while native math is still attempted. A valid resource not referenced by its formula and one resource referenced by two formula nodes both fail before WPS starts. Semantic JSON, plan JSON, repr, issues, and evidence must not expose relative/absolute paths, bytes, source/payload hashes, or prepared bindings. Platform staging/hash/cleanup assertions belong to Tasks 5 and 6.
- [ ] Write failing citation tests for first-occurrence numbering across every supported visible inline container, including abstract, paragraphs, lists, block quotes, page-break content, and table cells; repeated citations; multiple citations in one paragraph; code/math/alt literalness; missing IDs; duplicate IDs; deterministic occurrence IDs; and same-paragraph inline degradation runs. Preserve block-quote/page-break child paragraph semantics for Task 3 to emit recursively. Require table cells to compile resolved citations to static `[n]` and unresolved citations to cell-local coded fallback metadata; no scanned citation may be silently lost.
- [ ] Write failing bibliography tests for cited-first ordering, uncited declaration order, `bibliography_include_uncited: false`, empty/malformed entries, no-reference documents, and byte-stable canonical JSON. Malformed entry source text must remain visible at the original bibliography position with one block degradation; it may not be silently discarded.
- [ ] Write failing global-ID collision tests across bibliography, figure, table, formula, and section IDs: first definition wins; later definitions retain visible content but lose target capability and produce the configured original-location degradation.
- [ ] Write failing page-policy tests for bibliography after all body sections, continued Arabic numbering, correct link-to-previous behavior, no empty bibliography section, and unchanged cover/front-matter/body/landscape ordering.
- [ ] Implement parsing, shared-grammar integration, citation/degradation span splitting, numbering, bibliography role, formula-resource scanning/binding, and canonical serialization without WPS/PDF imports.
- [ ] Rerun focused tests plus all M1-M3 parser/semantic/resource/page-policy tests.
- [ ] Commit: Resolve M4 formula, citation, and bibliography semantics.

### Task 3: Closed M4 plan schemas and bibliography/citation emission

Files: modify `skills/WPSComposer/scripts/generation_plan.py`, `skills/WPSComposer/scripts/longform/plan.py`, and `skills/WPSComposer/scripts/recording_composers.py`; create `tests/longform_m4/test_plan_schema.py` and `tests/longform_m4/test_plan.py`.

Interfaces: preserve the exact M3 `writer.add_equation` shape as the `equation-shell-m3` branch. Add an M4 branch discriminated by `renderMode: native-m4` with two orthogonal closed choices: `content` is exactly one of `{nativeMath}` or `{plannedDegradation}`, and optional `fallbackResource` is exactly one of `{fallbackResourceId}` or `{fallbackResourcePlannedDegradation}`. Only content degradation suppresses OMath; a fallback-resource degradation never suppresses valid native math and never supplies a locator. The M4 failure policy uses the existing stable `EQUATION_INSERT_FAILED` code with `explicit-image-then-source-notice`. Both branches preserve the M3 number/bookmark shell, so preflight-invalid formula source is never passed to OMath but keeps its formula ID, numbering, bookmark, and formula reference target. `writer.add_cross_reference.runs` accepts strict `citation` and `degradation` variants so unresolved references/citations remain in their owning paragraph. `writer.add_semantic_table` gains bounded `cellDegradations[{row,column,code,fallbackText}]` for the explicit cell-local M4 path, while resolved cell citations are already static `[n]` strings. Preserve legacy `writer.add_bibliography.entries: string[]`; add the M4 branch discriminated by `schemaVersion: 1` whose entries are bounded `{id, nodeId, number, text, cited}` objects with `hangingIndentPt`, `leftIndentPt`, `spaceAfterPt`, and fixed numeric style.

- [ ] Write failing one-of schema tests for untouched M3 equation shell, untouched legacy bibliography strings, strict M4 discriminators, orthogonal content and fallback-resource exclusivity, length/hash shape, cell-degradation coordinates/codes, citation numbers, bibliography entry bounds/uniqueness/order, hanging-indent geometry, and unknown-field rejection at every nested level.
- [ ] Write failing schema tests rejecting raw paths, bytes/base64, arbitrary field codes, unvalidated formula source as native math, non-positive/duplicate citation numbers, citations without a matching bibliography target, and invalid recoverable-code/fallback pairs.
- [ ] Write failing plan tests for literal/reference/citation ordering in one paragraph, recursive emission of block-quote/page-break child paragraphs, cited-first bibliography order, uncited omission/configuration, no empty bibliography operation, native formula descriptors, explicit-image fallback binding, and preflight-invalid formula degradation.
- [ ] Extend the state-machine validator so one semantic node has one owner, each citation exactly matches one final bibliography entry/number, bibliography is emitted in its back-matter section after body content, and M1-M3 plans stay valid.
- [ ] Add one always-emitted `writer.reserve_document_quality_anchor` operation with unique owner `doc:quality`, fixed `failurePolicy: fail`, immediately after the title/cover transition and before TOC/index/body content. It creates no visible paragraph, spacing, or page when empty. Replace the existing late `_build_quality_notices` emission: all initial semantic document issues and later runtime issues upsert through this one anchor. Validate one owner/one operation, placement before fields/body, no late notice operation, and fatal behavior when anchor/upsert APIs are absent or even the minimal notice text cannot be written.
- [ ] Implement closed validators, plan emission, and recording mirrors.
- [ ] Rerun focused tests plus all generation-plan and M1-M3 plan tests.
- [ ] Commit: Emit closed M4 math and bibliography plans.

### Task 4: Shared whitelist recovery and styled degradation primitives

Files: create `skills/WPSComposer/scripts/longform/degradation.py` and `skills/WPSComposer/scripts/update_longform_recovery_matrix.py`; modify `skills/WPSComposer/scripts/longform/executor.py`, `skills/WPSComposer/scripts/longform/field_contract.py`, `skills/WPSComposer/scripts/longform/windows_executor.py`, `skills/WPSComposer/scripts/longform/macos_executor.py`, `skills/WPSComposer/scripts/writer.py`, and `macos/wps-jsapi-probe/addin/writer-longform-v2.js`; create `tests/longform_m4/test_degradation_contract.py` and extend M3 field-convergence tests.

Interfaces: add immutable `DegradationDescriptor(code, placement, object_label, reason, fallback_text, fallback_kind)` with redacted serialization and `RecoveryDecision`. Extend `ExecutionIssue` compatibly with `stage`, `fallback`, and `recoverable` fields plus deterministic round-trip serialization. `degradation.py::render_js_recovery_matrix()` is the canonical generator; the tracked add-in contains explicit generated-block markers, and `skills/WPSComposer/scripts/update_longform_recovery_matrix.py` updates only that block. Windows calls the Python state machine; macOS runs the generated frozen table with its equivalent local checkpoint state machine. Both allow exactly one declared fallback, record one issue, insert/deduplicate the matching visible notice, and rethrow everything outside the contract.

- [ ] Write failing matrix tests for every recoverable code already declared by M1-M4 and every fatal category; assert unknown errors, engine loss, save/field/index/configuration/cleanup errors, fallback failure, and rollback failure terminate.
- [ ] Write failing tests proving one and only one simplified attempt, stable issue order, same-node placement, nested child locality, deduplication, and no catch-all continuation.
- [ ] Write failing privacy tests for paths, payload/resource identifiers, hashes, exception reprs, field results, and bookmark maps in issue/notice/evidence serialization.
- [ ] Write failing `ExecutionIssue` compatibility/round-trip tests for `stage`, `fallback`, and `recoverable`, including old payloads with no new fields and stable redacted serialization.
- [ ] Write failing Writer-mock tests for inline styling in the owning paragraph, restrained one-cell block notice styling, and one document-quality region outside outline/TOC/header/footer stories.
- [ ] Write failing field-convergence tests proving the empty runtime anchor is present even when the initial plan has no document issue; three unstable mutation rounds upsert exactly one `FIELD_REFRESH_UNSTABLE` notice at that anchor, run the required fourth round, refresh again after notice insertion, freeze the fourth snapshot, and return the issue without duplicating the notice on any pass. Missing/throwing field APIs remain fatal.
- [ ] Write failing parity tests that compare `render_js_recovery_matrix()` byte-for-byte with the tracked generated block, prove regeneration is idempotent and touches no surrounding JS, and exercise identical code/fallback/placement/attempt-count/fatal outcomes on Windows and macOS fakes.
- [ ] Run `../../.venv/bin/python skills/WPSComposer/scripts/update_longform_recovery_matrix.py --check` as the drift gate; run without `--check` only to regenerate the marked block after changing the canonical Python matrix.
- [ ] Implement the shared controller and native notice primitives; route existing figure/table/reference and new formula/bibliography recovery through it without changing fatal boundaries.
- [ ] Rerun focused tests plus all M2/M3 executor, rollback, cleanup, and privacy tests.
- [ ] Commit: Centralize marked degradation recovery.

### Task 5: Windows COM native formulas, citations, and bibliography

Files: modify `skills/WPSComposer/scripts/writer.py` and `skills/WPSComposer/scripts/longform/windows_executor.py`; create `tests/longform_m4/test_windows_executor_m4.py` and extend COM fakes.

Interfaces: add `WriterComposer.add_equation_native(...)`, `add_citation_paragraph(...)` through the existing run paragraph primitive, and `add_bibliography_native(...)`. Remove bibliography from `_M2_DEFERRED_OPERATIONS`. Native formula insertion writes trusted linear text to the center cell, calls `document.OMaths.Add(range)` and `BuildUp()`, verifies the OMath count/range, then adds the unchanged M3 number/bookmark shell in the right cell.

- [ ] Write failing COM-mock tests for a borderless formula/number layout table, centered editable OMath content, right-aligned native number, keep-together behavior, bookmark/reference refresh, and native-object verification.
- [ ] Write failing formula fallback tests for the exact ladder: OMath -> one already-validated explicit image when present, otherwise readable source -> visible notice. If image insertion fails, roll back it and place source inside the terminal notice; that terminal notice is not a second simplified retry. When OMath succeeds but `fallbackResourcePlannedDegradation` exists, keep the native formula and insert exactly one formula-node `FORMULA_FALLBACK_IMAGE_UNAVAILABLE` issue/notice without attempting a locator. Later operations continue. Unknown COM, number/bookmark, field, rollback, or container failures abort.
- [ ] Write failing citation/bibliography tests for inline `[n]` in the owning paragraph, repeated number reuse, left-aligned hanging-indent native paragraphs, fixed spacing, no blank separator paragraphs, and exact final order.
- [ ] Write failing table-cell tests for static resolved `[n]`, same-cell coded unresolved text, exact cell-local shading metadata, and no table/data loss. Write failing compatibility tests proving an M3 equation with no `renderMode` still uses readable source plus number shell without OMath and legacy bibliography strings still execute through the legacy path.
- [ ] Write failing lifecycle/privacy tests for reopen persistence, private resource cleanup, hash mismatch, and no path/payload leakage.
- [ ] Implement COM mappings with `DispatchEx` ownership unchanged and no OOXML mutation.
- [ ] Rerun focused tests plus all Windows M2/M3 tests.
- [ ] Commit: Render M4 content through Windows COM.

### Task 6: macOS JSAPI native formulas, citations, and bibliography

Files: modify `skills/WPSComposer/scripts/longform/macos_executor.py`, `skills/WPSComposer/scripts/macos_probe/models.py`, `skills/WPSComposer/scripts/macos_probe/runtime.py`, `skills/WPSComposer/scripts/macos_probe/templates.py`, `macos/wps-jsapi-probe/addin/writer-longform-v2.js`, and generated assets; create `tests/longform_m4/test_macos_executor_m4.py`.

Interfaces: add-in handlers consume exactly the Task 3 descriptors and return only applied-operation status, stable issue codes, closed child results, field snapshots, and pagination metadata. Bibliography leaves `LONGFORM_DEFERRED`; formula handler upgrades M3's text shell to native `OMaths.Add`/`BuildUp` with the shared fallback ladder.

- [ ] Write failing Python bridge tests for request shape, staged fallback-image binding, response closure, cleanup on every exit, and fatal capability/protocol mismatch.
- [ ] Write failing JavaScript asset/fake tests for native OMath creation/build-up/verification, number-shell preservation, inline citations, hanging bibliography, and the three notice placements.
- [ ] Write failing table-cell tests for static resolved `[n]`, same-cell coded unresolved text/shading, and no table/data loss. Write failing compatibility tests proving an M3 equation with no `renderMode` never enters OMath and legacy bibliography strings still execute through the legacy handler.
- [ ] Write failing recovery tests for the same exact OMath/image-or-source/terminal-notice ladder, including successful OMath plus exactly one formula-node `FORMULA_FALLBACK_IMAGE_UNAVAILABLE` issue/notice and no locator attempt; bibliography paragraph-local fallback; exact whitelist handling; one fallback attempt; later-operation continuation; and fatal engine/save/field/rollback failures.
- [ ] Write failing staging tests for private fallback-image copies, manifest/hash verification, changed staged payload, create/write/cleanup failures, and redacted bridge responses.
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
- [ ] Run all M4 tests, all longform tests, generation-plan/recording/add-in/privacy tests, then `../../.venv/bin/python -m pytest -v`.
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

- Semantic/math/plan: `../../.venv/bin/python -m pytest tests/longform_m4/test_semantic_formulas.py tests/longform_m4/test_semantic_citations.py tests/longform_m4/test_native_math.py tests/longform_m4/test_plan_schema.py tests/longform_m4/test_plan.py -v`.
- Recovery/executors: `../../.venv/bin/python -m pytest tests/longform_m4/test_degradation_contract.py tests/longform_m4/test_windows_executor_m4.py tests/longform_m4/test_macos_executor_m4.py -v`.
- Offline acceptance: `../../.venv/bin/python -m pytest tests/longform_m4/test_acceptance_m4.py -v`.
- Real macOS acceptance: `../../.venv/bin/python -m pytest tests/longform_m4/test_macos_real_wps_m4.py -v`; a skip is not acceptable for completion on this machine.
- Regression: `../../.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/longform_m3 tests/longform_m4 tests/test_generation_plan.py tests/test_recording_composers.py tests/longform_m0/test_addin_assets.py tests/macos_probe -v`.
- Full suite: `../../.venv/bin/python -m pytest -v`.

## Key Design Decisions

1. **Formula conversion is bounded and platform-independent.** Both executors receive the same trusted WPS-linear descriptor; neither parses raw LaTeX or chooses a fallback policy.
2. **Native math and formula numbering stay separate.** M4 replaces only M3's readable content shell. The existing native number/bookmark contract remains stable.
3. **Explicit images are fallback resources, never formula success.** WPS is always tried first. A fallback image is used only after the named native-math failure and is visibly marked.
4. **Citation numbers are semantic, not pagination-derived.** First source occurrence fixes the number before execution, so both platforms and later re-layouts agree.
5. **Bibliography entries are structured plan objects.** Executors receive final order and geometry, not loose strings that require platform-specific parsing.
6. **One generated recovery specification governs two local state machines.** Python owns the matrix and Windows implementation; the macOS add-in receives a generated frozen table and must perform checkpoint/rollback locally. Parity tests prevent divergent whitelist, retry, issue, notice, and privacy behavior.
7. **Engine loss never degrades.** If the engine or required document APIs disappear, returning a plausible-looking partial file would violate the user's explicit boundary; execution stops.

## Open Questions

1. No blocking M4 product question remains. M0 already proved `OMaths.Add`/`BuildUp` and reopen on both platforms; M4 turns that primitive into the production descriptor and recovery path.
2. WPS linear-math serialization may expose platform quirks for a specific supported construct. The shared converter must use the proven common subset; any construct that cannot remain editable and equivalent on real macOS is removed from the accepted subset and deterministically degraded rather than implemented with static OOXML or a hidden renderer.
3. Real Windows parity remains intentionally unconfirmed until M5 is complete; mocks must still assert the exact COM calls and fatal/recoverable boundaries now.
