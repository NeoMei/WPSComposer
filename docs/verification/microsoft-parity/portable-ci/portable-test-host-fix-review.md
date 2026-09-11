# Portable test host fix scoped review

Decision: **PASS**. No actionable findings in the four-file test-only patch.

## Scope

Reviewed the diffs for `tests/longform_m5/test_public_routing.py`, `tests/macos_probe/test_longform_evidence.py`, `tests/msoffice/test_macos_word_rules.py`, and `tests/test_generation.py`, plus the worker report and retained actual CI evidence. Reviewed file SHA-256 values:

- `tests/longform_m5/test_public_routing.py`: `5f7834e14073643dcdc9864d522ec8c238a40345906a9365883264f9294a4435`
- `tests/macos_probe/test_longform_evidence.py`: `080044bc340ae16409d3e78c9b1be4f896a45934ac37a5b2845a2aff90af5e0c`
- `tests/msoffice/test_macos_word_rules.py`: `b0433229222d1cafaeb16f139932dc202b979148939c837d0f72605351172ab5`
- `tests/test_generation.py`: `a84800813b30c2352ce6f382af7f5eea349fb7805b23e65d352e689daf187cd0`

Concurrent production/Word features and the separate filename-portability fix are excluded. No production changes are required by this patch.

## Assessment

- Both original Linux JUnit reports from CI run `34332171915` contain the same ten failing test cases. They map exactly to the six mocked-generation tests, two mocked-evidence tests, and two macOS native compilation cases addressed here.
- The six generation tests already replace the native outcome; explicitly replacing only `orchestrator.sys` with a Darwin namespace selects their intended supported route without altering the global Python/pytest host. Their existing artifact, private outcome, preset, output-opening, failure, and warning assertions remain intact. The patch changes no production platform guard.
- The existing `test_generate_rejects_unsupported_platform_before_renderer_import` remains unchanged, active, and independently passes in both scoped runs. It explicitly selects Linux, checks the unsupported-platform error code, forbids renderer imports, and verifies no output file is created. Linux rejection is not hidden by the new per-test supported-host mocks.
- The two recording-runtime evidence tests now mock the remaining external WPS version lookup. Removing their previous POSIX skips increases portable coverage; the version assertion now verifies an exact controlled return value. No mock is substituted for the artifact, report, operation-count, or runtime behavior assertions.
- Only the two actual `/usr/bin/osacompile` cases receive Darwin-only skips. The mixed readback test is split so all three pure generated-script assertions remain active on every host. The separate compiler test retains its exit-code assertion. No whole-file skip, xfail, or deleted semantic assertion is introduced.

## Independent checks

Read the simulation helper before execution: it changes the orchestrator's module-local sys reference and the rules test module's imported sys, raises if unmocked WPS version access occurs, and rejects attempted native compiler execution. Python, pytest, pathlib, and multiprocessing remain on the actual macOS host; this is a dependency simulation, not native Linux/Windows acceptance.

Using the clean Python 3.9 venv with `PYTHONPATH=.` pointing to the current worktree, independently ran the worker's exact scoped selection for both `linux` and `win32`, including the retained unsupported-platform test and separated readback assertions:

- Simulated Linux: **10 passed, 2 skipped, 118 deselected**, exit 0, 0.12 seconds.
- Simulated Windows: **10 passed, 2 skipped, 118 deselected**, exit 0, 0.15 seconds.
- Both skips are precisely the two native Word dictionary compilation tests.
- Read the worker's broader four-file log: **128 passed, 2 skipped**. This broader run was not repeated by this reviewer.
- Scoped `git diff --check` passed.

## Boundary

The test-host patch is ready for integration with the committed candidate. Actual hosted Linux/Windows execution remains necessary to verify the combined checkpoint. No native application or compiler was invoked. Only this review report was written; no test/production edits, Git index changes, commits, or pushes were performed.
