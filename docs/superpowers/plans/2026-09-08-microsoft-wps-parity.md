# Microsoft / WPS Parity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use subagent-driven-development to implement task-by-task, with independent task review and a final branch review. Execute continuously in this approved session.

**Goal:** Match existing WPSComposer business capabilities with Microsoft Word, Excel and PowerPoint on Windows/macOS, backed by native evidence.
**Architecture:** Engine-pinned public operations and shared semantic plans, app-specific native hosts and per-platform adapters. Extend the existing Word backend without weakening WPS behavior or ownership/atomicity safeguards.
**Tech Stack:** Python 3.9+, Windows pywin32, macOS native AppleScript dictionaries, existing JSAPI WPS runtime, pytest and native Office evidence.
**Spec:** docs/superpowers/specs/2026-09-08-microsoft-wps-parity-design.md

## Global Constraints

- Base v0.9.0, commit 6dd3a00. Existing isolated worktree office-description now branch codex/microsoft-parity; preserve its documentation changes and all other worktrees.
- Default engine=wps; msoffice and auto explicit, selected before mutation. No substitution with static Office-file generators or GUI automation as the product backend.
- Public actions compare against the union of implemented WPS business operations. Record current platform differences; unsupported or untested Microsoft rows remain incomplete rather than removed.
- Python 3.9+, preserve public positional signatures and lazy imports. Native dependencies imported only on appropriate platforms.
- Preserve unrelated files, unsaved documents, global templates/security and user processes. Owned and attached sessions have distinct close/rollback contracts. Never infer ownership merely from DispatchEx or a window title.
- Never lower macro security or run macros contained in user documents. Preserve current live atomic-composite rejection when reliable rollback cannot be proved.
- Native output must remain editable. Formula values, shape identity, saved TOC/numbering and source hashes are independently inspected after reopen.
- Read docs/regression-guardrails.md before Word typography changes.
- Full pytest before commit; individual workers run targeted RED/GREEN and leave edits for coordinated full-suite commits. Candidate pushes and draft PR are authorized for remote validation; no new release/merge is authorized by this implementation step.
- No native GUI automation from concurrent agents; AppleScript/COM runs use different apps or are serialized. UI acceptance is coordinated by the parent.

## Task 1: Frozen capability baseline and native feasibility evidence

Files: new scripts/capability_catalog.py, tests/msoffice/test_capability_catalog.py, docs/verification/microsoft-parity/baseline.json; fixtures/microsoft_parity/macos_excel.py, macos_powerpoint.py, windows_native.py; tests/msoffice/test_parity_probes.py.

Interfaces: catalog exposes capability_records() -> list[dict] with component, operation, source, WPS per-platform support, Microsoft per-platform implementation and native-verification states. Pure data must not start apps. Probe runners require --output-dir NEW_PATH, reject existing directories and write report.json, native outputs, logs, source hashes and checksums.

- [ ] Enumerate public Composer methods, public pipeline operations, patch targets/verbs and format lists from the frozen source. Distinguish raw COM object access from semantic operations.
- [ ] Write tests that fail for missing catalog coverage, accidental Mac WPS structural-edit claims and premature MS Excel/PPT support; run to RED. Example assertion: `assert not by_key[("spreadsheet", "generate")]["msoffice"]["darwin"]["verified"]` on the initial baseline.
- [ ] Implement data-only baseline and operation coverage checks, with explicit notes for legacy direct COM fallback and shared independent PDF editing.
- [ ] Native Excel probes create two sheets, formulas, styled/merged ranges, native chart, save/reopen and PDF; inspect formula and calculated value plus actual PDF. PowerPoint probes create native slide/textbox/shape/image/table/notes, save/reopen/PDF and structural edits. Use only owned documents and preserve a synthetic unsaved sentinel.
- [ ] Windows remote probe obtains actual EXCEL.EXE/POWERPNT.EXE identity and verifies document/instance ownership before mutation; do not copy Word-specific HWND assumptions.
- [ ] Review baseline completeness and raw failures. Dictionary presence/COM activation is not a successful native gate. Record verified primitives needed by later tasks.

## Task 2: Native spreadsheet and presentation generation/conversion

Files: new msoffice/macos_excel_script.py, macos_powerpoint_script.py, macos_office_runtime.py, windows_office_host.py, windows_office_runtime.py; extend office_engines.py, orchestrator.py, conversion.py, presentation.py; tests/msoffice/test_office_generation.py, test_macos_office_runtime.py, test_windows_office_host.py.

Interfaces: `compile_plan(plan, resources, native_path, pdf_path=None, timeout=...) -> str` per Mac app consumes GenerationPlan and owned normalized resources; app workers return validated native artifact paths and semantic result metadata. Shared dispatcher `generate_office(plan, request, deadline)` uses the existing artifact publication contracts. `convert(request, timeout=...)` mirrors Word conversion; native format constants remain app-specific.

- [ ] RED tests explicitly select msoffice XLSX/PPTX and assert no WPS ProgID/app can be selected, invalid operations fail before launch, requested output stays atomic, and conversion source bytes are preserved.
- [ ] Extend read-only discovery for installed Excel/PowerPoint with same-view registry lookup on Windows and exact app locations on Mac. Apply engine filtering for msoffice as well as wps; auto chooses only a backend capable of the complete request.
- [ ] Implement Windows hosts using app-specific identity bindings established in Task 1, inherited SheetComposer/SlideComposer business methods and component-specific save/PDF operations. Shared process attachment cannot be quit as owned.
- [ ] Implement Mac native compilers for the existing spreadsheet/presentation GenerationPlan operations. Reuse resource validation, deadline and atomic-publication semantics. Closed grammar prevents script injection; per-app locks prevent concurrent jobs colliding; uncertain AppleEvents retain recovery state.
- [ ] Add missing business plan operations from the baseline, with RED tests for each real operation before enabling it. Formula strings and native objects remain native; do not pre-render entire worksheets/slides.
- [ ] Execute public native generate/convert/reopen fixtures on both systems, including exact app provenance, source preservation and chosen-engine presentation. Review code and results independently before enabling support in capability records.

## Task 3: Word advanced generation parity

Files: msoffice/macos_script.py, macos_runtime.py, tests/msoffice/test_macos_script.py and new test_word_parity.py; fixtures/microsoft_parity/word_advanced.md; docs/verification/microsoft-parity/word/.

Interfaces: existing MacWordAdapter and compile_plan signatures remain stable; newly supported GenerationPlan args emit actual Word objects and share M5 fields, quality and semantic contracts.

- [ ] Read regression guardrails and baseline, create minimal fixtures for every current capability rejection: native equations, merged/semantic tables, landscape media, multiple-column figures, supported explicit styles/list glyphs and page combinations.
- [ ] RED compilation/validation tests for each primitive, followed by bounded native probes showing the real dictionary/object operation works. Keep unsupported cases closed until confirmed.
- [ ] Implement each native operation with bookmark/field refresh and section/list continuity. Roll back failed owned-document changes and preserve accurate capability errors.
- [ ] Run Word public M5 representative and structural insertion/deletion, TOC/page-number, long table, image/reference and recovery regression in both Word and WPS. Review each newly enabled family and update the capability matrix.

## Task 4: Explicit file inspection and formatting edits across three apps

Files: document_api.py, new msoffice/document_sessions.py, macos_document_api.py, windows_document_api.py; per-app Mac scripting modules; tests/msoffice/test_document_engine_routing.py and test_document_sessions.py.

Interfaces: add keyword-only engine="wps" to open_document, attach_active, inspect and edit. Bound sessions implement inspect_document, inspect_selection, apply_format_patch, apply_structural_op, save, save_copy, export_pdf, close and is_bound_to with existing snapshot/patch/report schemas. apply_ops/apply_patches use the bound engine and never redispatch.

- [ ] RED routing tests: `inspect(path, engine="msoffice")` uses a Microsoft session, `edit(..., engine="wps")` remains WPS, invalid engine fails before opening, and unsupported mixed format fails without mutation.
- [ ] Implement explicit session factory dispatch without passing engine through **options to legacy inspect methods. Preserve existing caller/test injection contracts for default WPS.
- [ ] Implement native file snapshots for Word paragraphs/tables/sections/shapes, Excel cells/ranges/sheets/charts and PowerPoint slides/shapes/tables/notes, with stable identifiers and content/type/property assertions.
- [ ] Implement existing PATCH_GRAMMAR dimensions by family using owned staging files. Validate final file and publish only after successful batch; failed atomic operations never save partial results.
- [ ] Native acceptance checks formatting after reopen, outside-range preservation, unchanged source and PDF output where the baseline supports it. Validate boundary, Unicode, merged-cell and grouped-shape cases rather than only empty/basic files.

## Task 5: Structural editing and active documents

Files: the same document session adapters plus focused app structural modules where needed; tests/msoffice/test_structural_ops.py and test_active_documents.py; fixtures/microsoft_parity/active_sessions.py.

Interfaces: existing set/insert/remove/move/clone and INSERT_TYPES remain the semantic contract. Active sessions bind an exact app/document identity once and preserve original saved-state/path when saving a copy.

- [ ] RED tests cover same-kind ambiguous documents, wrong-app attachment, synthetic unsaved sentinel preservation, no global Quit, stale IDs after structural change and rejected unsupported atomic live batches before mutation.
- [ ] Implement Word paragraph/heading/table/image/textbox operations, Excel row/column/sheet operations and PowerPoint slide/textbox/image operations, based on verified native primitives. Inspection identities survive operations or produce explicit stale-target errors.
- [ ] Implement attach_active for each app with exact binding and native selection snapshots. An attached document cannot be silently replaced by a newly opened file.
- [ ] Verify actual UI edit -> Undo -> explicit save/discard -> close -> reopen on owned synthetic documents. Test save-copy binding and unrelated unsaved document preservation.
- [ ] Review recovery/cleanup and compare semantic outcomes with WPS reference fixtures. No GUI macro or static Office-file replacement can satisfy the product gate.

## Task 6: Full parity verification, installation and documentation

Files: docs/verification/microsoft-parity/, fixtures/verify_microsoft_parity.py, skill/API/README metadata and tests.

Interfaces: final acceptance runner compares capability records to native evidence: every required implemented operation has source-bound passing evidence for each target platform, or remains a release blocker. Reports distinguish implementation tests, native execution, native artifacts, UI and cleanup.

- [ ] Run complete local suite and complete Windows suite, followed by cross-engine fixtures and fault/recovery cases. Retain raw failures with separate corrected reruns.
- [ ] Independent review of each implementation task and entire branch; fix worthwhile findings and rerun affected native gates.
- [ ] Install exact candidate in isolated destinations and execute public generation, conversion and editing from installed modules. Check source/manifest/version hashes and lazy import without native dependencies.
- [ ] Update description/README/API/capability registry only for verified behavior. Require consistent engine selection, scope, examples and installation requirements. Do not modify published v0.9.0.
- [ ] Commit/push reviewed candidate, prepare draft PR and report exact parity status. No merge/tag/new release until separately authorized.

## Execution ledger

2026-09-08: Design approved by user with “可以”. Ruling: execute in the existing isolated office-description worktree, renamed to codex/microsoft-parity, preserving documentation changes. Default to subagent-driven development per required skill and prior project preference; no additional execution-choice question is needed. Native probing and candidate evidence pushes are within the approved cross-platform task. Runtime code unchanged at start.
