# Native Word production acceptance — 2026-09-08

Release status: **PENDING final Windows native acceptance after section-numbering correction**. No release tag or package publication has been made. Existing 0.8.1 metadata is unchanged.

| Gate | Evidence | Result |
|---|---|---|
| macOS Microsoft Word public DOCX / conversion / direct PDF | macos-word/report.json and original artifacts | PASS, 5 pages |
| macOS WPS existing public flow | macos-wps/report.json and original artifacts | PASS, 4 pages |
| Real edit → Undo → explicit save → close → reopen | each platform ui-report.json | PASS |
| Structural heading edits, automatic renumbering and refreshed TOC pagination | additional task-copy native verification | Mac PASS for all four levels and TOC after insert/delete/save/reopen; Windows correction and final acceptance pending |
| Word unsaved document preservation | macos-word/preservation.json | PASS |
| Word native images, captions, REF / SEQ | macos-word/media-report.json, transparent-report.json | PASS |
| Word timeout, quarantine, refused unsafe recovery and successful explicit recovery | macos-word/timeout-recovery.json | PASS |
| Legacy DOC conversion with source preservation | macos-word/legacy-doc-report.json | PASS |
| Portable integrated suite | pytest-legal-outline.log | 2859 passed, 12 native-gated skips |
| Windows production public API, native UI and ownership | remote production evidence required | PENDING |
| WPS bootstrap tabs | six old task tabs closed; full representative rerun from WPS home returned to home without document tabs | PASS for current successful flow |

The runner records the Git base and source SHA256 values for the actual dirty working tree used; these source digests, not only the base commit, bind each native report to its implementation. Artifacts are unmodified native outputs. macOS Word theme colors and pagination may differ from WPS; the shared typography, numbering, content and A4 guards all pass.

The UI reports describe observed CUA interactions. WPS records one undo step per typed character, so all inserted characters were undone before saving. The saved WPS copy changes package metadata, but its document text and native typography checks match the original. Word saved copy is byte-identical.

Reproduce with `python fixtures/verify_msoffice_production.py --output-root NEW_DIRECTORY --fixture representative --engine msoffice --timeout 600`; use `--engine wps` for the regression. The runner does not certify native UI or Windows by itself. Operator support boundaries and recovery instructions: [native-word.md](../../../skills/WPSComposer/references/native-word.md).

Final local rerun: Mac Word public-representative-final-2 and WPS wps-regression-05 pass the strengthened title-occurrence gate. The WPS Title outline regression is fixed; its earlier PASS used an insufficient exact-paragraph-only title test. Current report/artifacts are the corrected native rerun. `ui-source.docx` binds the previously observed edit/undo/reopen report to its exact source; the UI report is not relabelled as an edit test of the new bytes. Native unsupported-equation rejection also passed through the public API with the explicit safe error code. Windows first production API/UI/sentinel round passed its initial checks with the separately committed identity fix. Parent visual and OOXML review subsequently found the cover PAGE field and missing front-matter/body numbering restarts; these are release blockers under native correction, so initial PASS does not admit release. Final merged-source admission remains pending.

Candidate deb67ba portable verification: 2836 passed, 12 native-gated skips, exit 0 in 189.62 seconds. Isolated installation, public signatures, both native adapter imports and ten installed-source hash comparisons pass (install-deb67ba.json). Independent final review accepted the cleanup compatibility correction, cancellation handling, Title mapping and strengthened representative page-number gates. Windows section numbering and final native admission remain open.

Candidate ead655a integrates the native Windows section fix: unlink headers/footers before PAGE insertion, use COM boolean restart values and verify section policy readback. Independent review and all ten new section tests pass. Complete portable rerun: 2846 passed, 12 native-gated skips, exit 0 in 221.96 seconds. Clean isolated installation and both adapter imports/source comparisons pass (install-ead655a.json). The complete final Windows representative and WPS COM regression remain pending.

Structural acceptance correction: generic Mac field updates left stale TOC entries, and style-side list linking created separate numbering instances for Heading 1–4. Dedicated native collection updates and level-side localized style binding correct both defects. Current public-representative-shared-outline passes all three routes and the new shared-list gate; structural-shared-outline verifies native descendant list strings, cached/visible TOC, actual PDF pages, insertion/deletion and save/close/reopen. The earlier structural-edit-final report is explicitly superseded because its Heading-1-only checks missed descendants. Mac portable snapshot: 2855 passed, 12 skips in 186.39 seconds. Windows round3 independently found the same structural defect and remains under correction.

Current Mac hybrid-bid native rerun preserves Chinese chapter numbers and Arabic dotted descendants through Word legal numbering, with both chapters verified; chinese-formal and sequence-transparent figure/reference native regressions also pass. Integrated suite after WPS first-section compatibility and legal numbering: 2859 passed, 12 skipped, exit 0 in 199.34 seconds. The independent shared-outline-recheck.json accepts both existing WPS artifacts and corrected Mac Word, while rejecting the pre-correction Windows Word list instances. Windows final structural and auto-detection correction remains pending.
