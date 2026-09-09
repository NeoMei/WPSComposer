# Office runtime Windows CI fix report

Scope frozen for independent review. No commit or index mutation performed. Production change is restricted to imports and OfficeJobLock in macos_office_runtime.py; shared Word runtime, sessions, _execute, native UI, and other agents' changes were not edited.

## Root cause and changes

Actual Windows CI run 34339669535 (run-03, Python 3.9 and 3.12) failed eight Office runtime tests at unconditional `import fcntl`. OfficeJobLock now selects the OS backend from os.name, using nonblocking msvcrt locking on byte zero on Windows and fcntl flock on POSIX. The lockfile is opened in binary read/write mode without any initialization write. Lock/unlock target the same byte. Only EACCES/EAGAIN contention (plus Windows EDEADLK) retries within the original deadline; unrelated I/O errors propagate immediately. Failed acquisition closes its local stream without a potentially masking unlock. Explicit close releases the descriptor in a finally block even if unlocking fails. Existing lock path, quarantine path, error code, recovery behavior, and quarantine persistence are preserved.

Two other observed Windows failures were test assumptions: source manifest paths now normalize with relative_to(ROOT).as_posix(); new-document exact private binding checks the AppleScript-escaped private path using apple_string. A third path assertion hidden behind fcntl failure was corrected with root's explicit approval: recovery verification checks the escaped exact job prefix including trailing slash. Native close/quit prohibitions remain asserted.

## RED / GREEN evidence

New bounded tests cover real host lock exclusion/reacquisition, native and mocked Windows quarantine/recovery, same-byte lock/unlock with no initialization write, each Windows contention errno, deadline expiry, unrelated I/O errors (including a non-contention BlockingIOError), and descriptor closure after unlock failure. Windows backend is mocked because this run is on macOS; the native test will use the real CRT on Windows CI. These tests do not skip the functional path on Windows.

Before production edits: `office-runtime-windows-fix-red.log` records 9 failed, 2 passed, all nine failing because fcntl is unavailable in the modeled Windows backend.

After production edits: `office-runtime-windows-fix-green.log` records 141 passed. Command:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_office_job_lock_portable.py tests/msoffice/test_macos_office_runtime.py tests/msoffice/test_macos_word_quality_fixture.py tests/msoffice/test_macos_word_session.py tests/msoffice/test_word_job_lock_portable.py tests/msoffice/test_macos_runtime.py -q
```

`git diff --check` passes. Actual Windows rerun and whole-branch validation remain root-owned gates; local modeled backend success does not claim Windows CI or native Office acceptance.

## Separate diagnostic I/O defect, intentionally not fixed

Scratch reproduction: `test_office_runtime_diagnostic_red.py`; failing output: `office-runtime-diagnostic-red.log` (1 failed). After osascript returns nonzero, inject OSError(ENOSPC) only for job/recovery.json. Expected NativeOfficeError(NATIVE_OFFICE_EXECUTION_FAILED) plus quarantine; actual tuple is ('OSError', None, False). The diagnostic write precedes lock.quarantine in the outer exception handler, so an underlying OS diagnostic-write failure bypasses the quarantine call and replaces the primary native error. The job directory remains retained, but a later job is not blocked by the absent quarantine marker. The fault injection is deliberately limited to the diagnostic path to prove quarantine can still be persisted if attempted independently; truly exhausted storage can independently fail quarantine persistence too and requires explicit handling.

Additional code inspection shows timeout stdout/stderr writes can replace NATIVE_OFFICE_TIMEOUT, and successful result logs are written before the closed marker is evaluated. These are separate from portability, and _execute has not been changed pending root review.

## Frozen SHA-256

- `ca82ca86e84d1f968ffb51d22a121159302f02251d5d6126b87ca3ab9c23efcb`  `skills/WPSComposer/scripts/msoffice/macos_office_runtime.py`
- `d641f8ed96402e8381edbacfc8e345a89d70635aba16e3dcfb7acc9e251a0c77`  `tests/msoffice/test_office_job_lock_portable.py`
- `a29d606b13c56e11ac12ae9903127e5c48c9fc6ab8799b969193c1dfc3e4f9a4`  `tests/msoffice/test_macos_office_runtime.py`
- `57606b00f6197ed0f343952f1411eaab43463e4035f569f0660b1f34f01b321d`  `tests/msoffice/test_macos_word_quality_fixture.py`
- `d539bee64f245e2936020fde1d40040539f348d1912ecf8a32b8b55fd93b0429`  `tests/msoffice/test_macos_word_session.py`
