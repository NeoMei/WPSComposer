# Independent review: Mac Word interrupted transport recovery

Date: 2026-09-09  
Scope: the frozen `MacWordSession._execute` cancellation repair and its focused tests. This review made no production/test edits and started no Word, Office, WPS, UI, or subprocess cancellation job. The reviewer-owned reproduction and this report are the only reviewer files.

## Verdict

**PASS for the scoped repair.** The observed `KeyboardInterrupt`/`SystemExit` gap is closed, the original exception semantics are preserved, and no new P1/P2 was found in the bounded diff.

## What was verified

- The new `except BaseException` applies only around `subprocess.run`. It does not reclassify a completed nonzero AppleScript response, field-identity error, malformed completion envelope, timeout, or `OSError`.
- On an interrupted bound operation, `_retain("Native AppleEvent completion uncertain")` creates the existing quarantine record before control returns to the context manager. `close()` sees `_quarantined` and rejects locally before another `_execute`, so no second AppleEvent is sent against a document with uncertain mutation completion.
- Bare `raise` preserves `KeyboardInterrupt` or `SystemExit`, its value, and traceback. The diagnostic records the exception type without serializing arbitrary exception text.
- `_execute_structural` still marks structural targets stale and sets `_retain_evidence=True` on the escaping cancellation. A quarantined nonstructural call also retains its stage because `close()` cannot enter normal cleanup.
- Interrupted `open_document` retains its private script/stage and gains the missing per-script diagnostic. Its existing outer `BaseException` cleanup releases the advisory lock while preserving quarantine. An interrupted attached/open path likewise respects `_quarantined` before deleting its stage.
- Lock release plus persistent quarantine matches the existing timeout ownership model: the OS lock is not leaked, but a later native Word job is blocked until explicit recovery clears the quarantine.

## Evidence

The owner retained the original focused RED: `3 failed, 56 deselected`. Both bound cases observed two `subprocess.run` calls, and the interrupted-open case lacked its transport log.

After the repair:

- owner cancellation/neighbouring paths: **11 passed, 48 deselected**;
- complete `test_macos_word_session.py`: **59 passed**;
- independent reviewer reproduction: **4 passed**;
- independent Mac Word session/logical-save/fields/references plus cancellation set: **181 passed in 9.10s**.

The independent reproduction covers bound structural `KeyboardInterrupt` and `SystemExit`, interrupted open for both exception families, one-and-only-one transport invocation, retained script/diagnostic/stage, exact quarantine ownership, released lock, and the absence of context-exit close traffic.

## Frozen hashes

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `532550b1fa7dbb10de1e9025307f406bfc69e6409cf5a056da1266808488a467` |
| `tests/msoffice/test_macos_word_session.py` | `468d692cd2f73c09d2f4ec0bb1583b1308bb19268304e19e199fa08f2b9f54d8` |
| `word-cancel-recovery-review-repros.py` | `21377cce42ed58df880fed5ae5e8f01bcad84ea4f6e3793d15febfb44e393fc5` |
| owner report | `ccb107769e65d53299016eb46255fcdc5c11e329884fec69d8ef3cd10214af42` |

## Limits

Pure fault injection proves Python-side exception, quarantine, diagnostic, lock, and context-manager behavior. It does not recreate the exact timing of the native run-05 user interruption or prove clean native Word save/reopen after this source change. The retained run-05 incident and root-owned full19/installed07/native checks remain separate evidence. This scoped pass does not certify checkpoint/rollback semantics, other direct Word methods, Windows behavior, the 1,256 platform gates, or full Microsoft/WPS parity.
