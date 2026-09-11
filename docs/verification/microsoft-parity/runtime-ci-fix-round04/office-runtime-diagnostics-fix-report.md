# Office runtime diagnostic I/O fix report

Scope frozen for independent review. This is a separately authorized follow-up to the CI lock and path correction. It changes only _execute in macos_office_runtime.py and adds tests/msoffice/test_macos_office_runtime_diagnostics.py. No commit, index mutation, shared Word runtime, session, or native UI changes.

## Independent patch evidence

- Earlier narrow patch: office-runtime-windows-fix.patch; its new tests: office-runtime-windows-fix-test-snapshot.py; frozen hashes and 141-pass result: office-runtime-windows-fix-report.md.
- Runtime snapshot after the CI correction and before this follow-up: office-runtime-before-diagnostic-fix.py.
- This follow-up only: office-runtime-diagnostics-only.patch.

The earlier report accurately describes its historical frozen state; this report supersedes its statement that the diagnostic defect remains unfixed.

## Change

Native result classification and confirmed close marker are established before writing logs. On native failure or cancellation, uncertain completion attempts quarantine first, then recovery metadata and each log independently. Diagnostic I/O faults cannot replace the primary native error or cancellation; safe (label, exception class) tuples are attached as diagnostic_io_failures. NativeOfficeError reports only successfully completed recovery/quarantine writes. Failed or partially written evidence does not gain a claimed path merely because a file exists.

After an acknowledged close, failed logs retain the private job and propagate the I/O error without publishing or inventing an uncertain-close quarantine. Before invoking the subprocess, deadline evaluation now precedes setting launched, so a pre-transport expiry causes neither quarantine nor stale deleted staging/diagnostic paths. Ordinary success still publishes and removes private staging.

No in-process uncertainty cache was introduced. If the OS refuses quarantine-file creation itself, cross-process exclusion after the current lock closes cannot be guaranteed; the error explicitly has quarantine_path=None and a quarantine diagnostic failure. If quarantine succeeds and later diagnostics fail, the next job is rejected before launch. This distinction is covered by the test matrix. Same-object repeated lock.acquire semantics remain unchanged because internal entry points acquire once.

## RED / GREEN

- Original minimal scratch regression: test_office_runtime_diagnostic_red.py failed with ('OSError', None, False) for recovery.json ENOSPC and nonzero native completion; now passes.
- New matrix RED: office-runtime-diagnostics-matrix-red.log records 16 failed, 1 passed before the implementation. It covers timeout, nonzero native result, cancellation identity, each of stdout/stderr/recovery/quarantine faults, all four concurrent faults, confirmed-close log faults, success publication, and prelaunch deadline expiry.
- Combined GREEN: office-runtime-diagnostics-green.log records 159 passed (including the original scratch reproduction). git diff --check passes.

Command:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_office_job_lock_portable.py tests/msoffice/test_macos_office_runtime.py tests/msoffice/test_macos_office_runtime_diagnostics.py tests/msoffice/test_macos_word_quality_fixture.py tests/msoffice/test_macos_word_session.py tests/msoffice/test_word_job_lock_portable.py tests/msoffice/test_macos_runtime.py .superpowers/sdd/2026-09-08-microsoft-wps-parity/test_office_runtime_diagnostic_red.py -q
```

This is macOS portable-unit evidence with the AppleEvent transport replaced. Actual Windows CI, full-suite integration, and native Office acceptance are separate root-owned gates.

## Frozen SHA-256

- `c0a76254c4dee8a73030aa3c5e728d87e3f95efc850a2772a935788fa584f357`  `skills/WPSComposer/scripts/msoffice/macos_office_runtime.py`
- `6c0f1aeacf92098edb48a46f4735ee614de58c3e105c8110b2a21f3b3513d174`  `tests/msoffice/test_macos_office_runtime_diagnostics.py`
- `d641f8ed96402e8381edbacfc8e345a89d70635aba16e3dcfb7acc9e251a0c77`  `tests/msoffice/test_office_job_lock_portable.py`
- `a29d606b13c56e11ac12ae9903127e5c48c9fc6ab8799b969193c1dfc3e4f9a4`  `tests/msoffice/test_macos_office_runtime.py`
- `57606b00f6197ed0f343952f1411eaab43463e4035f569f0660b1f34f01b321d`  `tests/msoffice/test_macos_word_quality_fixture.py`
- `d539bee64f245e2936020fde1d40040539f348d1912ecf8a32b8b55fd93b0429`  `tests/msoffice/test_macos_word_session.py`
