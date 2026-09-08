# Microsoft parity candidate status

This is an implementation checkpoint, not a release approval. The frozen baseline is `baseline.json` (628 required capability rows at v0.9.0, commit 6dd3a00). Raw COM access is recorded separately and is not a semantic parity requirement. No baseline row has been removed to make the candidate appear complete.

| Area | macOS evidence available | Windows evidence for this candidate |
|---|---|---|
| Word advanced generation | `macos-word-advanced/final-normal-02` and `final-source-fallback-02`: native objects, fields/numbering, PDF/reopen, controlled fallback, sentinel | Pending new candidate execution |
| Excel generation/conversion | `macos-excel/public-run-01` and `public-run-02`: public route, formulas/caches, PDF, unchanged input | COM-free host/runtime tests only |
| PowerPoint generation/conversion | `macos-powerpoint/public-native-02`: public route, native shapes/images/tables, PDF/reopen | COM-free host/runtime tests only |
| Word file edits | `macos-word-sessions/native-10` and `new-public-05-reviewed`: public new document, formatting/structure, native save/PDF/reopen | Session/proxy tests; native execution pending |
| Excel business/session APIs | `macos-excel-sessions/business-04`, `structural-04`, `active-03`, `new-03`: public factory, formulas/charts/conditions, copy/move boundaries, saved active binding | Session/proxy tests; native execution pending |
| PowerPoint business/session APIs | `macos-powerpoint-sessions/semantic-03-public`, `save-current-01`: all current semantic method signatures, native save/current/reopen | Session/proxy tests; native execution pending |
| UI editability | `ui-word-01`: actual typing, Undo, explicit save, close and UI reopen; Excel/PPT pending because Mac locked | Pending |

Additional code checkpoint: 17 direct Mac Word business methods have matching signatures and focused tests, but their native typography/formatting acceptance is pending. Independent review fixed an invalid header border enum (actual AppleScript compilation failure) and rich-paragraph inherited indentation; 48 business tests and pure compilation regressions pass. `fixtures/microsoft_parity/windows_native.py` prepares bounded three-app acceptance; its portable tests are not Windows execution. Full suite round 6 passed **3,574 tests with 12 skips** in 210.02 seconds; see `unit-round6.txt`. The skips include unavailable WPS registration and opt-in native gates, so they are not native passes.

The final runner-only fault-injection fix explicitly passes its factory into the real worker protocol (avoiding a definition-time default-argument binding). The final focused supplement passed **292 tests**, including the two new protocol/factory regressions; see `unit-final-focused.txt`. Independent scoped re-reviews accepted the Word fixes, preflight fixes, evidence gate and Windows runner fixes. These reviews are not a final whole-parity approval. Raw native logs retain their original whitespace; source files pass the Git whitespace check separately.

Representative evidence does not validate every argument combination, every baseline row, the exact final branch, or installed-plugin behavior. Older failed runs and corrected reruns are retained separately. Word `new-public-02` passed its original narrow assertions but visual review exposed concatenated paragraphs; `new-public-03-boundary-red` proves the defect, and `new-public-04-boundary`/`new-public-05-reviewed` verify the strengthened paragraph assertion.

Outstanding work:

- Close all remaining required semantic/format rows: Word direct business methods, advanced styles/list glyphs/resources, full patch dimensions and structural targets, legacy formats and native SaveAs targets, and arbitrary PowerPoint size.
- Verify every argument and behavior of the frozen seven spreadsheet and nine presentation plan operations. Their operation IDs are implemented; limitations such as arbitrary slide height remain open. The richer direct Composer methods are separate required rows; adding new plan operation IDs beyond the frozen grammar is not necessary for parity.
- Complete request-wide capability preflight for `engine="auto"`, including property/value-specific support, before native mutation.
- Resolve macOS native gaps: trustworthy unsaved/active binding where native IDs are absent, non-rebinding attached save-copy/PDF, and existing Excel worksheet deletion without a confirmation dialog. The proposed local Office.js supplement still requires the user's design choice; it has not been installed.
- Finish independent reviews, native fault/recovery tests and source-bound per-row evidence. Windows persistent child deadlines and typed semantic handles have platform-independent coverage only.
- Run Windows native acceptance on the review candidate, all-app UI edit/Undo/save/reopen, native WPS regressions, isolated installation/public API smoke tests and final full-suite verification.
- Commit and push a reviewed candidate/draft PR after its checks stabilize. No new merge, tag, release or installed-skill update has been performed for this parity work.

Only the previous v0.9.0 release has release acceptance. This candidate must not be described as fully matching WPS until the outstanding gates are closed.

`fixtures/microsoft_parity/evidence_gate.py` now expands the unchanged 628 capability rows into 1,256 platform gates. `coverage-checkpoint-03.json` intentionally has no certified claims: existing representative evidence has not been retroactively bound to the evolving source tree. This is an evidence-certification count, not a count of implemented features or passing native examples. A closing claim must identify an implemented capability, native verification, exact source digest (including pinned runtime dependencies), report hash, platform/engine/component, and an explicit capability-to-check mapping with passing named checks. Source changes reopen certification. Independent runner review, UI acceptance, and installation/release gates remain separate.
