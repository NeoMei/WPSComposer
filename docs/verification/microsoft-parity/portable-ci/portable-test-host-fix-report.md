# Portable test host fix

Date: 2026-09-09
Scope: test-host assumptions from CI run 34332171915, Linux Python 3.9 and 3.12. No production changes, staging, commits, pushes, native app actions, or actual AppleScript compiler execution were performed by this worker.

## Root causes and patch

- `tests/longform_m5/test_public_routing.py`: three public long-form tests mock the native outcome but leave `orchestrator.sys.platform` on the real Linux host. The public unsupported-platform check correctly rejects the call before the mock. Pin the intended supported Darwin host using a module-local `SimpleNamespace`, scoped to each test.
- `tests/test_generation.py`: the same omission affects three presentation/default-output tests. Add the same narrow host pin; retain the original output, failure, and warning assertions.
- `tests/macos_probe/test_longform_evidence.py`: two recording-runtime tests still invoke real `read_wps_version()`, which opens `/Applications/wpsoffice.app/Contents/Info.plist`. Mock that external lookup alongside the existing bridge/runtime doubles and assert the returned version. Remove the POSIX-only skips so these pure tests also execute on Windows.
- `tests/msoffice/test_macos_word_rules.py`: two tests invoke `/usr/bin/osacompile` without a host requirement. Mark only those native compilation cases as requiring Darwin. Split the mixed readback test so its exact-name/script assertions remain a portable test.

Production Linux rejection is unchanged. `test_generate_rejects_unsupported_platform_before_renderer_import` still runs and passes under both simulated host runs. No whole-file skip, xfail, deleted assertion, or runtime relaxation was introduced.

## Evidence

Actual CI failure evidence remains in `docs/verification/microsoft-parity/portable-ci/run-01/artifacts/pytest-ubuntu-24.04-python-{3.9,3.12}/pytest.{log,xml}`: each Linux leg had 10 failed, 4396 passed, 32 skipped.

Local evidence is in `docs/verification/microsoft-parity/portable-ci/test-host-fix/`:

| Evidence | Result |
| --- | --- |
| `red-linux-scoped.log` | Before the patch: the exact same ten failing test cases, 10 failed / 119 deselected. |
| `green-linux-scoped.log` | After the patch: 10 passed / 2 skipped / 118 deselected, including retained Linux rejection and the newly separated portable readback assertions. |
| `green-win32-scoped.log` | Same scoped checks with a simulated Windows host: 10 passed / 2 skipped / 118 deselected. |
| `green-linux.log` | All four affected files with simulated Linux: 128 passed / 2 skipped / 5 PyMuPDF SWIG deprecation warnings, 100.75 seconds. |
| `reproduce_host.py` | Reusable local host-dependency simulation; it does not claim to emulate a native OS or its Python version. |

The simulator changes only the orchestrator's `sys` reference and the rules test module's imported `sys`. It models an absent WPS bundle, and makes any attempted native compiler call fail. It leaves Python, pytest, pathlib and multiprocessing on the actual macOS host. The two explicit compiler skips are the only skips in the broad green run.

Initial broad-run failure traces are retained in `red-linux.log` and `failed-harness-linux.log`. They are **not** product regression evidence: the first used an incomplete root venv (missing PyMuPDF), and the initial helper also lacked a multiprocessing main guard. The exact ten-case RED run involved no validator subprocesses and reproduced the intended failures independently. The helper now has the main guard, and the final broad run used the fully provisioned clean environment.

## Commands

Run from the worktree root. Final verification used:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python docs/verification/microsoft-parity/portable-ci/test-host-fix/reproduce_host.py linux -r s
```

For scoped Linux or Windows simulation, append the following expression, and change the positional host to `win32` for Windows:

```sh
-k 'defaults_to_longform or public_preset_override or run_longform_m2_evidence or generate_default_does_not_present or generate_failure_never_presents or generate_opener_failure or generated_rule_script_compiles or native_style_readback or rejects_unsupported_platform'
```

The pre-patch exact RED command used `/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python`, host `linux`, and the same selection expression without `or rejects_unsupported_platform`.

`git diff --check` passed. Existing dirty work by other workers was left untouched. Root owns review, full merged-suite validation and the next actual Linux/Windows CI run; no new actual remote-CI acceptance is claimed here.
