# Windows Word one-shot worker diagnostic and cleanup fix

Scope frozen for independent review. Only windows_runtime.py and the new tests/msoffice/test_windows_word_runtime_diagnostics.py changed. No other Office runtime, session proxy, worker protocol, COM executor, native Office, or index changes. A real Python-child cancellation test also verifies an unrelated Python sentinel remains alive; neither is a native Office process.

## Confirmed defects

- Shared quarantine write failures and final lock.close errors replaced the original typed timeout/cancellation.
- Popen failures followed by diagnostic-write errors lost cleanup_verified=True and could quarantine Word before any child had started.
- Request-staging failures before Popen similarly lacked the known-clean marker.
- Timeout diagnostic write/stat failures replaced the timeout; failed child kill/wait could do the same.
- Cancellation during child.wait left the owned Python process alive.
- Non-object diagnostics (list/null/bool) caused a second data.get exception, and fallback errors could advertise nonexistent worker.log evidence.
- A worker-log close I/O failure could replace the original timeout.

## Change

_run_worker preserves original exception identity while attempting shared quarantine and releasing the lock independently. Successfully created quarantine evidence is reported only if a regular file can be observed; quarantine-write failure records a safe category and leaves quarantine_path=None. Successful operations continue to propagate lock-close failures.

_run_worker_unlocked marks every pre-Popen exception cleanup_verified=True, including diagnostic persistence failure after Popen itself failed. Expected launch OSError still becomes NATIVE_WORD_EXECUTION_FAILED, while launch cancellation identity is preserved. Any wait exception first attempts kill and bounded wait on the exact owned Python handle, then attempts diagnostic persistence. Timeout remains NativeWordTimeoutError; cancellation remains the original object. Cleanup and diagnostic failures attach safe label/type tuples instead of overriding that primary exception.

Diagnostic existence/stat probes are best effort and occur after the child cleanup attempt in wait failure paths. A failed or partially completed diagnostic write is not advertised as successful evidence; the existing worker log is used when verifiable, otherwise diagnostic_path=None. _worker_failure tolerates non-object JSON and never invents an absent fallback log. Final log close preserves a propagating primary exception while still raising a standalone close failure after otherwise successful work.

Uncertain Word state remains quarantined; no Word PID or process tree is terminated. If the OS refuses durable quarantine creation, the error reports that persistence failure rather than claiming that other processes are blocked. Existing post-success deadline/quarantine policy and the native worker protocol were not broadened or weakened.

## RED / GREEN evidence

- windows-word-runtime-diagnostics-red.log: 18 failed / 1 passed before production edits.
- windows-word-runtime-log-close-red.log: 1 failed / 1 passed against the frozen original _run_worker_unlocked loaded only into the test process. This proves the added log-close preservation regression without rewriting the source or old evidence snapshot.
- windows-word-runtime-diagnostics-green.log: 41 passed in 1.00s, covering 21 dedicated regressions and all 20 existing Word runtime tests.
- git diff --check passed.

Command:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_windows_word_runtime_diagnostics.py tests/msoffice/test_windows_runtime.py -q --tb=short
```

The independent patch is windows-word-runtime-diagnostics-fix.patch and before snapshot is windows-word-runtime-before-diagnostic-fix.py. Per root's bounded scope, no whole-suite run was performed. Actual Windows CI and native acceptance remain separate gates.

## Frozen SHA-256

- `4777325edcfbbb895654776c537a5e6cf3bc0d51c2b461060fa41443a70182e0`  `skills/WPSComposer/scripts/msoffice/windows_runtime.py`
- `49a6da3e1187b91a8995fce1015be4f95dc6f33542a9eada16582bc5e9919118`  `tests/msoffice/test_windows_word_runtime_diagnostics.py`

## Independent review — checkpoint_integration_review

Both reviewed files matched the frozen SHA-256 values above. Review was read-only for production/test/index/application state and covered `_run_worker`, `_run_worker_unlocked`, diagnostic helpers, and the new tests. The prelaunch cleanup marker, independent lock/log close handling, exact Python-handle kill, bounded recovery wait, non-object JSON handling, and existing deadline policy are coherent. No Word process or process tree is killed by the changed path.

**P2: secondary cancellation during evidence stat still replaces the primary failure** — `skills/WPSComposer/scripts/msoffice/windows_runtime.py:78-85`, `_evidenced_path`.

The helper catches only `OSError`, although its callers already hold a primary timeout or cancellation and the neighboring diagnostic/quarantine/cleanup guards deliberately catch `BaseException`. If `Path.is_file()` is interrupted by `KeyboardInterrupt` or `SystemExit`, that secondary exception escapes `_evidenced_path`. On the worker-timeout branch the exact child has been stopped correctly, but the caller receives the secondary cancellation instead of `NativeWordTimeoutError`; `_run_worker` then records `NATIVE_WORD_EXECUTION_FAILED` in quarantine rather than the timeout classification. An existing primary cancellation's object identity can likewise be lost. The same helper is used after successful quarantine persistence, so that evidence-location check has the same problem.

Pure reproduction used a fake timed-out Python child, an isolated lock, and `Path.is_file` patched to raise a distinct `KeyboardInterrupt` only for `diagnostics.json`. Running the real `_run_worker` produced:

```json
{"caught":"KeyboardInterrupt","is_secondary":true,"code":null,"child_returncode":-9,"quarantine_code":"NATIVE_WORD_EXECUTION_FAILED"}
```

Recommended narrow fix: capture `BaseException` from this best-effort evidence-location helper and attach its safe label/type via `_record_io_failure`, preserving the already-selected primary error. Add regressions for timeout plus secondary stat cancellation and primary cancellation plus secondary stat cancellation (including the quarantine-path stat branch).

Independent execution: the 21 new diagnostic tests passed in 0.38 seconds, and the existing 20 Word runtime tests passed in 0.93 seconds, with the provided clean Python, `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.`, pytest cache disabled, and scratch-only basetemp. Total: 41 passed across two independent invocations. No native execution, native compilation, index modification, or source edit was performed.

**Freeze assessment:** do not treat these two hashes as the final diagnostic fix for full33 until the P2 above is addressed and reviewed. This finding is scoped to the requested primary-identity/classification contract; it does not reopen the native worker protocol or Windows UI acceptance.
