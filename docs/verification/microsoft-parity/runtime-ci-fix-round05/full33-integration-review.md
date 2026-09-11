# Full33 staged integration review

Scope: `595268e` to the staged 10-file candidate. All ten staged files matched their working copies at review time; the Word/Office one-shot and bridge hashes matched their previously approved snapshots. `git diff --cached --check` passed. No source, permanent tests, index, Git state, or native Office state was changed. Existing scoped suites were not repeated.

## Finding

**P2 — session proxy evidence metadata can replace a selected timeout classification.** `skills/WPSComposer/scripts/msoffice/windows_session_proxy.py:158-169`, `_error`'s nested `evidenced` helper, catches only `OSError` during recovery/quarantine `is_file()` calls. `_call` invokes this after catching timeout and completing `_abort`. A secondary `KeyboardInterrupt`/`SystemExit` during either lookup escapes instead of returning `NATIVE_WORD_TIMEOUT` / `NATIVE_OFFICE_TIMEOUT`. Thus the interactive path still violates the metadata-failure preservation contract now satisfied by both one-shot runtimes.

New bounded probes in `checkpoint-integration-scratch/test_proxy_stat_cancel_integration.py`: **12 failed in 0.92 seconds**, covering three application proxies × both secondary signal types × recovery/quarantine paths. They use the existing real Python-only transport fixture. All cases first confirm in-memory uncertainty, lock closure and exact-child termination, then fail only on the missing/incorrect timeout code. No Office process is involved. Reviewed proxy source SHA-256: `ced8874be05101c8fcf5afd0fdb0ece1379e10b2d6e6726840bdeedd83ca7653` in both index and worktree.

Suggested narrow repair: make the optional metadata lookup tolerate secondary BaseException while retaining the intended typed error, and record safe diagnostic failure categories as in the one-shot helpers. Keep the frozen full33 snapshot intact as evidence; any repair needs its own scoped validation and subsequent source-bound integration evidence. The prior one-shot and bridge approvals remain valid; this is a distinct uncovered proxy path.

## Other integration observations

- One-shot runtime flags consistently distinguish known prelaunch/acknowledged-close cleanup from uncertainty. Shared OfficeJobLock quarantine and typed Word/Office error paths remain compatible. No additional actionable runtime finding was identified.
- The semantic-table fixture imports shipped helper/fixture modules and adds no production route or public method. It is a source-checkout acceptance fixture: `SOURCES` explicitly includes its permanent test file, while `install.py` excludes `tests`. Therefore it must not be represented as an installed-plugin native fixture without retaining that development source. Public runtime import/install closure is unaffected; this is a verification-scope limitation, not a newly claimed product parity bug.
- Current status top does **not** claim completed parity, release approval, or fresh Windows native execution. It explicitly retains the Windows 3.12 failure, nine missing Mac Word method names and open quality acceptance.
- Status freshness needs editorial cleanup separately: line 7 still says Word one-shot work is in progress and quotes the earlier 43-test Office review, although the scoped reviews now pass 49 and 57 tests; line 9 still says hosted run04 has not run despite line 7 recording its completion. Lower text is expressly labeled historical, but line 9 precedes that history disclaimer. Do not infer current-source acceptance from older “still match live” wording or old absent-method counts.

## Verdict

**One open P2; no full33 integration PASS yet.** This report does not invalidate the running snapshot's test evidence or claim that full33 has finished. Root owns documentation refresh, repair sequencing, new full-suite/Windows CI results, installation, native/UI acceptance and remaining overall parity work.

## Post-full33 resolution

The sole integration P2 was independently closed on the later proxy source `dfa9abfba9962db621a63bccf4058d923ad293a2c4f93888aa6ec507d5294b77` and tests `5f8916ce8eaf6e83d25dcd88939fd472b594f34cffbb546a1aee5dfd040020f3`: original 12 reviewer probes plus 12 permanent cases all pass (24 passed, 19 deselected, 2.29 seconds). See `proxy-stat-cancel-fix-report.md` for acceptance. This approves the corrected files for full34, without rewriting full33's original source-bound result or historical RED evidence.
