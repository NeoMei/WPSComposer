# Office runtime independent final review

Initial reviewed candidate: worktree `office-description`, HEAD `1686155`, runtime SHA-256 `c0a76254c4dee8a73030aa3c5e728d87e3f95efc850a2772a935788fa584f357`. Review date: 2026-09-09. Scope: the separately preserved Windows lock/path patch and diagnostic I/O patch described by `office-runtime-windows-fix-report.md` and `office-runtime-diagnostics-fix-report.md`. No native application, production-source change, index mutation, or commit by reviewer.

## Initial decision

The Windows backend and diagnostic-write corrections satisfy their narrow contracts. Independent combined verification: **159 passed in 1.19s**. Additional pure edge checks: **6 passed**. Two pre-existing cleanup-error masking cases were independently reproduced and reported to root for correction, because root confirmed worthwhile existing bugs are in scope. They are not attributed as regressions introduced by the two candidate patches. Final approval remains pending their correction/review.

## P2: final lock release replaces the primary native error or cancellation

`_execute` calls `lock.close()` unguarded in `finally`. If unlock reports an I/O error while a native timeout/error/cancellation is already propagating, callers receive the cleanup exception instead. The updated `OfficeJobLock.close` correctly releases its descriptor in `finally`, but the caller still loses the primary exception identity and code. Independent bounded reproduction `test_office_runtime_review_edges.py::test_unlock_failure_does_not_replace_primary_cancellation` injects a final close error after real descriptor release: expected original `KeyboardInterrupt`; actual `OSError(EIO)`.

Preserve an already-active primary exception and attach safe cleanup failure evidence. A close failure on an otherwise successful execution must still propagate; this is not a request to suppress all release errors.

## P2: prelaunch staging cleanup replaces the preparation error

`_execute` also calls `shutil.rmtree(job)` unguarded for prelaunch failures. A permission/I/O failure removing that private directory replaces the error that prevented transport from launching. Independent reproduction `test_office_runtime_review_edges.py::test_prelaunch_cleanup_failure_does_not_replace_prepare_error`: expected original `ValueError`; actual `PermissionError(EACCES)`. This is the same cleanup-versus-primary contract as lock close. Preserve the primary exception, report the retained directory truthfully when useful, and continue propagating cleanup errors after otherwise successful work.

## Behaviors independently verified

- Office lock uses actual host `os.name`, byte offset 0 and length 1 for both CRT lock/unlock, and makes no initialization write. POSIX retains nonblocking exclusive flock. Contention errno classification is bounded; unrelated errors propagate. The host-native contender test is not skipped on Windows. Explicit unlock failure closes and resets its descriptor.
- Persistent quarantine checks happen after exclusion. Normal jobs refuse the marker; recovery obtains the lock and preserves the marker. Separate component roots keep independent component jobs isolated.
- Native return code and exact success/close marker are classified before stdout/stderr persistence. Timeout, nonzero return, and transport cancellation attempt quarantine before recovery/log writes. Each diagnostic failure is independent and cannot replace the primary during that diagnostic phase.
- The 12-case native/timeout/cancel × single diagnostic-path matrix passed. All three primary outcomes with **all four diagnostic paths failing** also passed independent checks. Existing output stayed unchanged; cancellation retained identity.
- A failed quarantine creation reports `quarantine_path=None`; independent follow-up demonstrated that another job can then launch, so no cross-process block is falsely claimed. This is a reported storage-failure limitation, not a hidden fallback guarantee.
- If the quarantine file is created but its `fsync` fails, the initial error does not claim a successfully persisted quarantine path, while the existing marker still blocks the next normal job. Independently verified.
- Confirmed close with failed logs retains staging, refuses publication, and does not invent uncertain-close quarantine. Confirmed close with a missing staged artifact likewise reports execution failure without inventing an uncertain close. Clean success still publishes and removes private staging; expiry before subprocess invocation does not create quarantine or claim deleted paths.
- The three path-only tests use complete escaped AppleScript literals or POSIX manifest spelling. Exact private-document binding, recovery job prefix including its separator, and no-close/no-quit checks remain intact.

## Evidence and limitations

Combined independent run used root's complete `clean-dev-venv` and exactly the eight paths listed in `office-runtime-diagnostics-fix-report.md`: **159 passed**. Additional probes are retained in `test_office_runtime_review_edges.py` next to this report. At initial review these contain **6 passing edge probes and 2 RED cleanup contract probes**. The two failures above were sent to root immediately.

No AppleEvent or native Office transport was executed. This is source review and macOS-host pure/mock evidence. Real Windows CI and native Office acceptance remain separate integration gates.

## Final closure after cleanup correction — 2026-09-09

**Both P2 findings above are closed; no open actionable findings remain in this bounded candidate. Approved to freeze into full32.** The earlier initial decision and RED evidence are retained as historical review findings, superseded by this closure.

Independent re-review verified runtime SHA-256 `8836a0932e699ef72c3589ff9fe597f7cc3beb29d4f4f36677e0ed756b4a7a14` and diagnostic-test SHA-256 `d9381bd097aed5120d6bfd7737d326fbe410f239fe2d31b3d90cc4d3c6bdb235`. The original eight edge probes, including both unmodified primary-exception identity assertions, all pass. The full combined bounded regression was independently run: **173 passed in 1.35s**; scoped diff whitespace checks passed.

Primary errors/cancellation now survive secondary lock and eligible staging cleanup failures, with safe cleanup evidence attached. With no primary failure, a cleanup failure still propagates. The preceding Windows lock and diagnostic/quarantine behaviors remain covered. Details are also recorded under independent acceptance in `office-runtime-cleanup-fix-report.md`. Full32 integration, actual Windows CI, and native acceptance remain separate gates.
