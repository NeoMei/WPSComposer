# MS Office native spike repeated audit — 2026-09-08

Scope: approved standalone native Microsoft Word/PDF feasibility spike on macOS and Windows. No production WPS routing, public API, installers, website or release changes. The plan tasks were cross-checked against the implementation and retained native evidence. UI means the actual desktop Word document interaction; this spike adds no frontend application.

## Review rounds

1. Task/evidence audit and native-runner review found sentinel setup cleanup leakage, an unbounded child runner wait, overlapping evidence/output directories, false-positive disabled table headers/direct numbering overrides, and stale successful validation after malformed input. Fixed with failure-path and real saved-artifact mutation regressions.
2. Independent review found effective numbering format/start values were unchecked: `numFmt=none` and `startOverride=99` passed against an unchanged PDF. Fixed current and referenced parent levels, including instance override precedence; added eight real-artifact tests.
3. Independent closure review reran the original two reproductions (both rejected), reviewed cleanup/timeout/validation and ran 56 focused tests successfully. No further worthwhile P1/P2 findings in the reviewed scope.

## Native and UI evidence

Mac Word 16.112.3, macOS 26.6.2 arm64: a fresh native run completed all 14 operation gates, saved/reopened DOCX, exported a 9-page PDF and preserved the initial empty document set. Raw native output is in `mac-round1/`. Read-only PDF checks verified 80 records, headers on every page, end marker, numbered headings and TOC page values; pages 1, 2 and 9 were visually inspected. Mac PDF text uses compatibility radical glyphs, so extraction was normalized with NFKC and visually checked radical mappings; the initial unnormalized comparison did fail and was not mistaken for a native generation defect.

Actual Mac Word UI: opened an independent synthetic copy, typed `UI-AUDIT-MAC-20260908` at its end, observed the marker, pressed Undo and Save, closed, reopened through Word's local Open dialog, observed correct heading/TOC layout and the absent marker, then closed. Saved XML text exactly matches the original. The copy required a per-file access grant; this manual UI check is separate from the successful automatic native run. Original generated artifacts were not edited.

Previous final Windows native runs 09/10 and the unsaved sentinel gate remain retained under `../windows-word-spike/final-cleanup-review/`. The stronger validator passed temporary copies of both with all 12 checks; original 65 + 25 evidence hashes were unchanged. Fresh Windows evidence f7e3ed7 now closes the remaining gate: run11 on exact d4884a5 code passed all 14 native operations and 12 artifact checks; sentinel04 preserved the 62-character unsaved document and closed only that sentinel. Task PID 3936 exited after 2.797 seconds and registered Word returned to zero documents. The controller verified all 146 root-manifest entries plus nested manifests, matched exact executed sources, reran all 12 artifact checks on an independent copy, and inspected the contact sheet. See `../windows-word-spike/audit-native-round3/report.md`.

## Validation and boundaries

First full suite after round1: 2672 passed, 12 explicitly skipped native gates. Final full suite after round2 fixes: 2680 passed, 12 explicitly skipped, 184.22 seconds. Focused round3: 56 passed. Native-gated skips are not claimed as executed WPS acceptance; existing WPS frontend/bridge/composer regression coverage is portable, with no changes to those modules in this spike.

The child Python deadline cannot bound the sentinel wrapper's own COM calls. A timeout reports child Word cleanup unverified, preserves partial logs and never kills Word. Prior failed-run PID ownership is not reconstructed from elapsed time and no unowned process is closed. Synthetic table page splits and sparse last page reflect probe formatting, not a polished production template.

Production WPS/MS Office backend selection, Excel/PowerPoint parity and release admission remain later work outside this approved investigation. The architecture recommendation and complete original task mapping are in `../../msoffice-native-spike.md` and `../../superpowers/plans/2026-09-08-msoffice-native-spike.md`.

## Final task closure

All approved spike plan tasks and this bounded repeated audit are complete. Windows also executed all56 focused regressions (18.58s) and the actual CUA mouse/keyboard edit/Undo/save/close/reopen flow on a synthetic copy. Its full persisted DOCX matched the original byte for byte. Screenshots and independent COM readbacks are in `../windows-word-spike/audit-ui-round1/`; the controller inspected insertion, Undo and reopen captures. Windows round2's registration experiment failed with E_INVALIDARG and never launched the sentinel; this historical failure was superseded by the authorized normal registered-empty-instance workflow in round3, without a registration workaround.

Final snapshot review found no remaining worthwhile defects in the approved spike scope. This conclusion is bounded by the tested scenarios and is not a guarantee about every WPS/Office application function. The 12 skipped platform gates, unbounded native COM calls in the wrapper, and production integration remain explicitly outside this completion claim.
