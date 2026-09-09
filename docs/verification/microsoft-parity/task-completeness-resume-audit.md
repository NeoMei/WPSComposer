# Microsoft/WPS parity task-completeness resume audit

Date: 2026-09-09  
Scope: read-only reconciliation of the approved design and plan, frozen capability baseline, current source declarations, `progress.md`, `word-direct-remaining.md`, `task-audit-20260909.md`, `audit-summary-20260909.md`, and the evidence gate. This audit did not edit production or tests, run the whole suite, start Office/WPS, use UI automation, or delegate work.

## Verdict

**All six approved plan tasks remain partial.** The branch has substantial reviewed Mac implementation and representative evidence, but it has not met the approved definition of Microsoft/WPS parity. The remaining work is not one final regression run: it includes missing Word methods, incomplete argument/format/target coverage, controller rollback, current-source native evidence, all Windows execution gates, and per-capability certification.

The frozen baseline remains **628 required capability rows / 1,256 Darwin and Windows gates**. `coverage-checkpoint-03.json` records `ready: false` and `verified_count: 0`. That zero is a formal certification count, not an implementation count: existing native reports have not been converted into reviewed, current-source, capability-to-check claims.

At this audit point:

- HEAD is `cb9898f28e42927248a42421858d549edc54bc3b`, equal to `origin/codex/microsoft-parity`.
- A Word field-topology repair is in flight in `macos_word_session.py`, `macos_word_fields.py`, and `test_macos_word_fields.py`. It is not a frozen completion result and is intentionally not re-reviewed here.
- Static current-source declaration counts remain Word **59/80**, Excel **24/24**, and PowerPoint **21/21** for the frozen direct-method names. Declaration does not establish full semantics, arguments, or native verification.
- Round 19 (`4,112 passed / 12 skipped`) and `installed-audit-07` are valid evidence for their recorded pre-topology-repair source. They are not final-current-source regression or installation evidence after the in-flight production change.

## Remaining implementation requirements

- [ ] Finish and independently freeze the field-topology repair already in flight. Preserve the exact external-drift and TOC-growth checks while accepting acknowledged same-session structural changes. This item records status only; it does not reopen that scoped review.
- [ ] Implement the **21 absent Mac Word direct methods** listed in `word-direct-remaining.md`. The currently probed but unimplemented controller pair, `degradation_checkpoint` and `rollback_degradation_checkpoint`, remains part of those 21 until production, review, and native acceptance finish. The other 19 names remain required; native probe success or generation-pipeline equivalence does not declare their direct contracts.
- [ ] Complete declared Word method argument behavior. Known open branches include floating-image alt text, unproved wrap values, floating image blocks, raster variants without native evidence, character-style/property combinations, full list-glyph/resource behavior, attached-session variants, and the multiline/page-boundary reference cases identified by the existing reports.
- [ ] Audit and implement every frozen argument/behavior of the **seven spreadsheet** and **nine presentation** generation-plan operations. Their operation IDs and all direct method names exist, but arbitrary combinations are not thereby supported. Existing-content PowerPoint resizing and untested size combinations remain explicit examples.
- [ ] Close the full public inspection/edit grammar by target, verb, and property, including boundary/Unicode/merged-cell/grouped-shape cases. Current representative edits do not cover every frozen `PATCH_GRAMMAR`, `ALL_OPS`, or `INSERT_TYPES` row.
- [ ] Close source, save, conversion, and generation format rows against the frozen WPS union. Legacy input/output formats and native SaveAs targets remain open and must either be implemented or retained as explicit blockers; they cannot be removed from the baseline.
- [ ] Finish active-session semantics that remain narrower than the baseline: trustworthy unsaved/active binding where native IDs are absent, selection coverage, attached save-copy/PDF without rebinding, and remaining copy/structural variants. Existing Excel worksheet deletion that can trigger a confirmation dialog also remains unresolved.
- [ ] Make and record the supplemental-design decision for Mac native API gaps only where the native-adapter path cannot meet a specific baseline row. The proposed Office.js supplement is not installed or approved and cannot be counted as implementation.

## Remaining native-evidence requirements

- [ ] Run affected Mac Word native acceptance after the topology and checkpoint implementations freeze. Prove snapshot after legitimate owned REF/TOC/content mutations, exact rollback of text, CR, tables and fields, old tracked handles after rollback, new handles after acknowledged mutations, failure retention, deadline behavior, source preservation, reopen, and unrelated-document preservation.
- [ ] Add source-bound native checks for the remaining 21 Word direct methods as they are implemented. Existing object, section, field/index, reference/bibliography, logical-save, and cancellation reports remain representative evidence for their exact source snapshots; they do not certify later methods or all argument branches.
- [ ] Complete Mac native coverage for every argument of the seven spreadsheet and nine presentation plan operations, richer direct Composer methods, full patch targets/properties, active attachment, save/copy/export, and failure recovery. A six-check public smoke per app is not this coverage.
- [ ] Complete the plan's same-input cross-engine regression. WPS Writer inspection and WPS presentation inspect/edit evidence exist; WPS Spreadsheet inspection remains blocked before ET profile registration. After an authorized trust/enable decision, rerun the official isolated ET path and preserve the existing failure separately.
- [ ] Build explicit claims for all 1,256 capability/platform gates. Each accepted row needs implementation state, exact current shipped-source digest, native report hash, platform/engine/component, and a capability-to-named-check mapping. Each unresolved row needs an explicit blocker; a generic passing report, method presence, portable test, or skip is insufficient.
- [ ] Run a final current-source full regression only after all production repairs freeze. The six passive bridge skips and six opt-in native skips remain skips, not acceptance.

## Remaining UI requirements

- [ ] Preserve the completed Mac Word, Excel, and PowerPoint edit -> Undo -> explicit save -> close -> reopen evidence, but repeat affected UI/editability checks for newly implemented document objects or lifecycle paths whose source or output semantics change.
- [ ] Exercise Mac active attachment and selection behavior for the remaining supported target/verb families. Existing file-based UI flows do not prove every attached-session or selection contract.
- [ ] Complete all three Windows UI flows on exact candidate outputs: edit existing semantic content, Undo, explicit save or discard, close, reopen the exact file, inspect formulas/fields/shapes/layout, and confirm unrelated sentinels remain unchanged.
- [ ] Do not reinterpret the failed Word open-panel attempt as success. The accepted Word references UI evidence used the exact recent-file route and must remain described that way.
- [ ] Resolve the WPS ET add-in enable/trust gate only with the already requested user authorization; then capture the actual UI/runtime result. No current document claims Spreadsheet inspection succeeded.

## Remaining Windows requirements

- [ ] Restore fresh Remote Control command execution. Accepted dispatches and the old interrupted turn are not candidate execution proof.
- [ ] Check out one exact reviewed candidate on Windows and record its commit and shipped-source digest. Run the complete Windows pytest suite there; portable Mac runs do not substitute for it.
- [ ] Run `fixtures/microsoft_parity/windows_native.py` in a new evidence directory for Word, Excel, and PowerPoint. Verify executable and app provenance, exact document/instance ownership, generation, conversion, save/reopen, PDF, inspection, formatting, typed structural handles, and preserved saved/unsaved sentinels.
- [ ] Natively validate COM-dependent values and target lookup that portable preflight cannot decide, including deferred host values and unclassified structural properties. Explicit `auto` routing tests establish selection policy, not host acceptance.
- [ ] Exercise Windows recovery paths that currently have portable proof only: acknowledged save conflicts retain the same live session/handles, post-save native-close failure quarantines, timeout/worker uncertainty preserves recovery evidence, retry/discard obeys ownership, and exact binding is rechecked within the original deadline.
- [ ] Run the optional fault-timeout phase only after the normal Windows run and with its retained-document recovery procedure ready. Do not delete quarantine evidence or kill Office to obtain a clean rerun.
- [ ] Complete Windows UI and exact installed-bundle acceptance after the native run. The older v0.9.0 Word evidence and current portable worker tests do not validate this three-application candidate.

## Remaining installation, documentation, and delivery requirements

- [ ] After all production changes freeze, create a fresh isolated Mac installation and compare every shipped file with the live candidate. Repeat the three public native smoke flows and retain post-run source equality. `installed-audit-07` remains valid for its snapshot but is reopened by the current Word production diff.
- [ ] Perform the equivalent isolated Windows installation from the same exact reviewed candidate, followed by public generation/conversion/edit/reopen/PDF checks and installed-vs-source hashes.
- [ ] Refresh the static inventory after the final Word slices. `task-inventory-round18.json` still reports the correct 59/80 name count at this audit point, but its recorded source hashes predate current changes.
- [ ] Replace the empty certification checkpoint with reviewed per-row claims; keep representative native reports, UI evidence, installation evidence, and certification as separate gates.
- [ ] Reconcile `status.md`, `audit-summary-20260909.md`, `word-direct-remaining.md`, SKILL/API/README metadata, and the capability registry against the final source. Preserve dated historical failures and avoid leaving earlier “in progress” paragraphs as the only current summary.
- [ ] Run the final independent task/whole-branch review after implementation and evidence freeze, then update draft PR #9 to the exact candidate. The pushed branch is a review candidate, not a release.
- [ ] Obtain separate authorization before merge, tag, release, or personal marketplace installation. The approved parity implementation did not authorize publishing a new release or replacing the personal installation.

## Closed work that this audit does not reopen

The following scoped repairs have reviewed evidence and should not be treated as remaining bugs merely because the whole task is open: shared publication conflict/rollback protection, Excel logical save and failed-saving-close ownership, PowerPoint logical save/background append, Word logical save, the run-12 field/index repairs, run-04 references/bibliography behavior, Word interruption quarantine, and the portable Windows saving-close/handle/binding repair. Their native/source scope still matters when constructing final per-row claims.

This checklist is intentionally narrower than a claim of “Office parity.” It covers the approved WPSComposer baseline and does not require implementing arbitrary Microsoft Office menus. Conversely, it does not allow the remaining baseline rows to be dropped because representative Mac outputs and portable tests pass.
