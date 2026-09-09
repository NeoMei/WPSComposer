# Office runtime cleanup review follow-up

Scope frozen for independent review. Only _execute and its dedicated diagnostic test file changed in this wave. No AppleEvents, native UI, commits, or index mutations. Earlier CI and diagnostic patches/reports were preserved unchanged.

## Confirmed review defects and fix

Independent review found two pre-existing exception-masking paths: final lock.close EIO replaced an original KeyboardInterrupt, and final staging rmtree EACCES replaced a prelaunch prepare ValueError. Both reproduced locally without altering the review assertions.

_execute now retains a reference to its primary exception. Final lock and eligible staging cleanup still execute. If an error is already propagating, secondary cleanup failures are recorded on that same exception as safe (label, exception class) tuples in cleanup_io_failures. The primary exception identity and native error code remain intact. Without a primary exception, cleanup failure still propagates unchanged; successful native publication is not silently reported as an entirely clean success.

No new retry behavior, quarantine state, or cleanup eligibility rules were introduced. Uncertain launched jobs stay retained; prelaunch failures remain eligible for removal. Existing successful-publication cleanup failure behavior (output may already be published when cleanup fails) remains explicit in tests.

## RED / GREEN

- Independent review repro: office-runtime-cleanup-review-red.log — 2 failed, 6 passed before the fix.
- Dedicated regression RED: office-runtime-cleanup-tests-red.log — 4 failed, 19 passed. The four failing cases cover native error, timeout, cancellation, and prepare-error masking. Two success-path cleanup-failure propagation tests already passed and guard against over-broad suppression.
- Dedicated tests plus unchanged reviewer probes after fix: 31 passed.
- Combined regression: office-runtime-cleanup-fix-green.log — 173 passed. This includes earlier lock/runtime/Word regression modules, the original diagnostic scratch reproduction, and all eight reviewer edge probes.
- git diff --check passes.

The isolated wave patch is office-runtime-cleanup-fix-only.patch. Before-wave snapshots are office-runtime-before-cleanup-fix.py and office-runtime-diagnostics-before-cleanup-tests.py. The old office-runtime-windows-fix.patch and office-runtime-diagnostics-only.patch remain untouched for independent historical review.

## Frozen SHA-256

- `8836a0932e699ef72c3589ff9fe597f7cc3beb29d4f4f36677e0ed756b4a7a14`  `skills/WPSComposer/scripts/msoffice/macos_office_runtime.py`
- `d9381bd097aed5120d6bfd7737d326fbe410f239fe2d31b3d90cc4d3c6bdb235`  `tests/msoffice/test_macos_office_runtime_diagnostics.py`

## Independent review acceptance — 2026-09-09

The independent reviewer verified both frozen hashes above against the live candidate and reviewed the isolated cleanup patch plus its interaction with the preceding Windows lock and diagnostic persistence fixes. **PASS: no open actionable findings in this bounded scope; this exact candidate may be frozen into full32.**

Both previously reported P2 findings are closed. `_execute` retains the original primary exception, preserves native error codes and cancellation identity, attempts lock cleanup and eligible prelaunch staging removal, and adds safe cleanup failure categories without substituting the cleanup exception. If there is no primary failure, cleanup errors still propagate unchanged; publication may already have completed, as the tests explicitly verify. Quarantine ordering, uncertainty retention, recovery/log path truth, and success/close classification remain unchanged by the cleanup patch.

Independent validation, with the complete `clean-dev-venv`, re-ran all nine combined test paths (seven repository modules, original diagnostic scratch regression, and independent edge probes): **173 passed in 1.35s**. Scoped `git diff --check` passed. The original independent scratch identity assertions were read and run unchanged, including the two formerly RED cleanup cases. All eight scratch probes now pass; scratch SHA-256 is `97948401578526a27e8ba33d9a505046a5971e5d61650721ab6f08bc25ad4fd3` for `test_office_runtime_review_edges.py`.

No production source, test assertion, shared index, or native process was modified by this independent re-review. This approves entry into full32 integration; it does not claim full32, a corrected Windows CI run, or native Office acceptance has already passed.
