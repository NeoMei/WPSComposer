# Windows CI repair independent review

Reviewed 2026-09-09 in `office-description`, limited to the working changes in `msoffice/macos_runtime.py`, `msoffice/input_validation.py`, their portable tests, the `CheckpointSnapshot` fixture constructor, and the five additional path-portability changes requested by root. Other Word/Excel/PowerPoint changes and staged work were not reviewed. No production source, shared index, native application, or commit was changed by this reviewer.

**Final verdict: no open actionable findings in this bounded candidate.** One P2 initialization race was identified, reported, and corrected during review. This conclusion covers inspected source and pure tests; it does not establish actual Windows lock contention or native Office acceptance.

## Closed finding: P2 Windows lock initialization can bypass contention handling

The initial candidate performed `fstat(fd).st_size == 0`, then wrote/flushed a NUL byte before entering the retry loop. Two processes could observe the empty file; one could initialize and acquire byte 0 before the other wrote that byte. Windows then rejects the second write, and the error escapes without waiting for the job deadline. The final candidate removes initialization and locks offset 0, length 1 directly.

Microsoft documents that `_locking` prevents other-process access to locked bytes and supports ranges beyond EOF. Therefore an empty lock file needs no initialization. `EACCES` denotes contention and `EBADF` denotes an invalid descriptor. The final implementation retries contention while allowing unrelated I/O errors to propagate. [Microsoft CRT `_locking`](https://learn.microsoft.com/en-us/cpp/c-runtime-library/reference/locking?view=msvc-170)

Pure model: `windows-ci-review/repro_lock_init_race.py`. Run from the worktree with `PYTHONPATH=.` and the project Python environment. Root had already removed initialization while this model was being prepared, so the script explicitly reconstructs the exact previously reviewed three-line initialization block from the current class; it is not presented as an immutable captured production snapshot or a real Windows run. Observed output: `REPRODUCED: initialization EACCES before contender lock attempt; elapsed 0.0002 seconds of 3-second deadline`. The initial attempt using the already-corrected snapshot did not reproduce; that temporary snapshot was removed.

## Review conclusions

- Lock backend selection uses actual `os.name`; test simulation of `sys.platform` cannot incorrectly import `fcntl` on Windows. Both lock and unlock target offset 0 / length 1. POSIX retains nonblocking exclusive `flock`. Repeated acquire/close remains idempotent. Normal acquisition checks persistent quarantine after obtaining exclusivity, while recovery acquisition leaves the quarantine marker intact. Acquisition errors close the descriptor; explicit unlock failures close/reset the stream in `finally`.
- Windows model checks independently confirmed repeated contention reaches the original deadline and closes the descriptor, and an injected unlock failure propagates after closing/resetting ownership. Existing POSIX exclusion/deadline coverage passed. Real Windows contention is not established by the backend double; existing `test_macos_runtime` contention tests still skip without `fcntl`.
- ZIP validation now inspects `ZipInfo.orig_filename` before normalized `filename` or directory detection can hide backslashes or NUL suffixes. Removing only one terminal slash permits ordinary directory entries while rejecting malformed repeated separators. Existing canonical names, content-type checks, case-folded duplicate rejection, limits, and readonly behavior remain intact. The raw ZIP regression construction asserts both local and central raw names are present, exercises both reader separator settings, and checks source bytes are unchanged.
- The snapshot constructor now names all six actual fields, including the empty numbering prefix. The injected failure location, original command sequence, ordering, and single-injection assertions are unchanged.
- `retain_sources` uses `relative.as_posix()` only for manifest keys; source copies and SHA-256 verification are unchanged. Three script assertions now require the complete escaped AppleScript path literal (the Excel assertion also includes `native.xlsx`). The rollback snapshot uses the host's absolute temporary path while retaining exact-document binding, native guard, and no-global-mutation assertions. These adjustments preserve or strengthen the prior intent on Windows.

## Verification

- Root evidence inspected: `portable-ci/run-02/lock-green.log` reported 28 passed and `input-name-green.log` reported 57 passed. These are root's earlier test records, distinct from the independent checks below.
- Independent current candidate: `test_word_job_lock_portable.py`, `test_macos_runtime.py`, `test_input_validation.py`, and `test_macos_word_rollback_failure.py`: **91 passed** with the main checkout `.venv/bin/python`.
- Independent extended path-related files: `test_macos_word_business_objects.py`, `test_macos_word_objects_probe.py`, and `test_parity_excel_probe.py`: **80 passed, 2 failed initially**; both failures were missing `fitz` in that environment. Re-ran only the two failed artifact-verifier tests using root's complete `clean-dev-venv`: **2 passed**, with PyMuPDF/SWIG deprecation warnings. No dependency or production edits were made to resolve this environment mismatch.
- No native Office or Windows process was launched. Final Windows CI confirmation remains the integration owner's next gate.

Reviewed production SHA-256 values:

```text
01c4acbfda5dfd9294517d681c8673bba7c55e617fedda14f1670e14277decc5  skills/WPSComposer/scripts/msoffice/macos_runtime.py
1c3ae7028447db3fad35c6c397d7ff88180601b41f320f17f5ef8bf5d1cbb0ed  skills/WPSComposer/scripts/msoffice/input_validation.py
c3a5e5936ec3f278d2338984f0bfd135434b2b5b8150992e8a63dc1eead36a77  fixtures/microsoft_parity/macos_word_fields.py
8642c43547e33281b558861d61f890ec3fdd63d54f6997b772352538dd7c0e17  tests/msoffice/test_word_job_lock_portable.py
306214855bd7862065ba8a63b112665649892b293ad74e7b514ac11c825c510e  tests/msoffice/test_input_validation.py
64afe7dfa776716e5e4302e6547aaa3b092427a1f0648d7b5e424002b402ef12  tests/msoffice/test_macos_word_rollback_failure.py
6fdef19207400f9b61813c27d6260cb50d9a74bce8ff55f3c7cdefda3fa3dff6  tests/msoffice/test_macos_word_business_objects.py
2af367472080d6d08b8990d2db2fcb5440ba983a6e6062b8a3646ad9779894ef  tests/msoffice/test_macos_word_objects_probe.py
956d9ebbaa1870b58e2cce13250c6f59e8dad011e521feff07aca351c2a60326  tests/msoffice/test_parity_excel_probe.py
```

## Follow-up: four Windows Python 3.9 failures

Root subsequently requested independent review of `orchestrator.py`, its added `test_generation.py` regression, and the minimal environment changes in `tests/longform/test_pipeline.py`, `tests/longform/test_executor.py`, and `tests/longform_m2/test_task9_pipeline.py`. **No actionable findings in this additional scope.**

The actual Windows 3.9.13 CI log at `portable-ci/run-02/artifacts/pytest-windows-2022-python-3.9/pytest.log` shows three child-interpreter failures in `_Py_HashRandomization_Init` before application imports (lines 4516, 4559, 4604), and the default generation output remaining `document.pptx` rather than an absolute path (line 6507). This distinguishes the import-test failures from forbidden-module failures.

The path root cause is directly supported by CPython v3.9.13: `_WindowsFlavour.resolve` repeatedly splits missing components and returns the input path if it stops progressing; a missing relative filename can reach that branch before finding an existing absolute ancestor. `Path.absolute` anchors an ordinary relative filename at the current directory without removing `..` or resolving symbolic links, and the retained final `resolve` still performs canonicalization. The narrow `expanduser().absolute().resolve()` change therefore fixes the observed default-output case while retaining home expansion, suffix enforcement, overwrite refusal, and downstream absolute output selection. [CPython v3.9.13 pathlib source](https://github.com/python/cpython/blob/v3.9.13/Lib/pathlib.py#L190-L212)

The legacy-path regression emulates precisely the missing-relative failure branch; the supplied `python39-output-red.log` demonstrates it fails before the production change. It is a host-independent model, not a real Windows execution. Already-absolute paths and symlink resolution continue through `resolve`; no public signature, renderer choice, or artifact format was changed.

The child environment now passes only project `PYTHONPATH` plus any case-insensitively named `SystemRoot`, rather than inheriting unrelated parent variables. Python's subprocess documentation explicitly requires callers supplying `env` to provide necessary startup variables and notes the Windows `SystemRoot` requirement. Together with the actual preinitialization failure, this supports the fix; final proof of Windows randomness initialization remains the corrected CI run. No fixed hash seed, skip, broad environment inheritance, or forbidden-module removal was introduced. All three scripts retain their complete forbidden-prefix lists, nonzero-error path, and `pure` success check. [Python 3.9 subprocess environment contract](https://docs.python.org/3.9/library/subprocess.html#subprocess.Popen)

Independent verification with the complete `clean-dev-venv`: **7 passed** covering the new legacy-path regression, existing resolved default output, overwrite refusal, extension mismatch, and all three child import-purity tests. Root's completed full-file record `portable-ci/run-02/python39-fixes-green.log` separately reports **101 passed in 95.77s**. No native application or Windows process was launched by this reviewer, and no code/index changes were made.

Additional reviewed SHA-256 values:

```text
43652f6ef0be1c0210e300a52c5999f3af2e2a0d46be772df649886970a64e23  skills/WPSComposer/scripts/orchestrator.py
c9b545acf485989debe94aeb90f0c80b3128b378841f7bd08bb56e571adf8c74  tests/test_generation.py
7474db995a1d84d20be13a9730d1f6c1fb585c6b66df8234e9db72a108cfcfea  tests/longform/test_pipeline.py
d30e126a27f3720c4ddf67ae2c3b3df12ffb37f3511ef1f99325f61330163b02  tests/longform/test_executor.py
45a9c4323d717488b384b8a8c64ec589ec4dcfe31efaff78fd27dd9983353830  tests/longform_m2/test_task9_pipeline.py
```
