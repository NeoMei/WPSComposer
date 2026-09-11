# Word interrupted transport recovery report

## Scope

- Worktree: `/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description`
- Production ownership: `skills/WPSComposer/scripts/msoffice/macos_word_session.py`, limited to `MacWordSession._execute` cancellation handling.
- Test ownership: `tests/msoffice/test_macos_word_session.py`, limited to focused cancellation cases.
- No native Word call, cancellation injection, process kill, commit, or push was performed.

## Root cause

`MacWordSession._execute` retained and quarantined uncertain completion only when `subprocess.run` raised `subprocess.TimeoutExpired` or `OSError`. A `KeyboardInterrupt`, `SystemExit`, or other `BaseException` escaped before `_retain` and before a diagnostic log was written. When this happened inside a bound session context, `MacWordSession.__exit__` called `close`; because the session was not yet quarantined, `close` issued a second AppleEvent against a document whose prior mutation completion was unknown.

The opening path had an outer `BaseException` guard that quarantined the session, but it did not produce the transport diagnostic that `_execute` owns.

## RED evidence

Command:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest -q tests/msoffice/test_macos_word_session.py -k 'cancelled_bound_transport or cancelled_open_transport'
```

Observed before the production change: `3 failed, 56 deselected`.

- Both bound cases (`KeyboardInterrupt` and `SystemExit`) observed two `subprocess.run` calls rather than one, proving context exit sent a close attempt.
- The interrupted-open case retained the staging directory and script but had no `.log` diagnostic.

## Minimal repair

After the existing timeout/OSError translation branch, `_execute` now catches remaining `BaseException` values, quarantines with the existing `Native AppleEvent completion uncertain` reason, writes the exception type to the per-script diagnostic, and uses bare `raise` to preserve the original exception type, value, and traceback.

This leaves the established behavior unchanged for timeout, OSError, acknowledged nonzero native errors, and invalid completion envelopes. Quarantine makes the subsequent context-manager `close` fail locally before another subprocess call.

## Verification

- Focused cancellation plus neighboring uncertainty/error paths: `11 passed, 48 deselected`.
- Complete `tests/msoffice/test_macos_word_session.py`: `59 passed in 1.01s`.
- Independent reviewer reproduction `word-cancel-recovery-review-repros.py`: `4 passed in 0.79s`; it covers bound/open `KeyboardInterrupt` and `SystemExit`, quarantine contents, evidence retention, released lock, and absence of a close retry.
- Independent reviewer scoped suite: reported by root as `181 passed`, with no new P1/P2 finding.
- `git diff --check` on the owned production and test files: clean.
- A broad repository run was intentionally stopped after `3746 passed, 12 skipped` when it entered an unrelated long wait; this is not recorded as a completed full-suite pass. Root owns the planned full19/install07/native Word smoke.

## Source hashes

Before this task:

- `macos_word_session.py`: `39129d22315eb5701dde8f1f5cfe305de404d07ce538d10d44b2fa1f562d502a`
- `test_macos_word_session.py`: `b4008b5c5a0a9493cf729cecbe2a712b1736a1dc209bc97121106fc57901dacd`

Frozen after repair:

- `macos_word_session.py`: `532550b1fa7dbb10de1e9025307f406bfc69e6409cf5a056da1266808488a467`
- `test_macos_word_session.py`: `468d692cd2f73c09d2f4ec0bb1583b1308bb19268304e19e199fa08f2b9f54d8`
- Independent reviewer reproduction: `21377cce42ed58df880fed5ae5e8f01bcad84ea4f6e3793d15febfb44e393fc5`

## Residuals

- Pure fault injection proves Python-side quarantine, diagnostics, lock release, exception preservation, and no second transport call. It does not prove the timing behavior of a real user cancellation inside `osascript`; the retained run-05 evidence remains the native incident record.
- Quarantine intentionally blocks automated cleanup after uncertain completion. Recovery of the native document and retained staging evidence remains an explicit operator action.
- Native clean save/reopen regression and installed/full validation remain with root under the Word lease.
