# Final review fix report

## Status

Both final scoped review findings are fixed from base `da866d6`.

- `ProfileServer.close()` now shuts down every accepted socket, including idle,
  partial-request, and HTTP/1.1 keep-alive connections. Request handlers are
  non-daemon threads reclaimed by `server_close()` after socket shutdown, and
  cleanup-time socket errors remain silent under the existing no-request-log
  contract.
- macOS DOCX, PDF, XLSX, and PPTX generation now activates the selected
  component with `isolated=True`. Registration retries preserve the same
  isolation choice instead of falling back to arbitrary running WPS state.
- Loopback binding, static-profile path restrictions, no-store responses,
  write rejection, selected-component routing, and public return behavior are
  unchanged.

## TDD evidence

Initial RED command:

```text
.venv/bin/python -m pytest -q tests/macos_probe/test_profile_server.py tests/macos_probe/test_generation.py::test_generation_uses_only_staged_path_and_publishes_valid_package tests/macos_probe/test_generation.py::test_generation_registration_retry_preserves_isolated_activation
```

Actual RED result: `17 passed, 6 failed in 12.62s`.

- A real `HTTPConnection` completed a second GET with HTTP 200 after
  `ProfileServer.close()`.
- Writer, spreadsheet, and presentation generation recorded
  `isolated=False`.
- The generation registration retry recorded `isolated=False`.
- The first partial-request handler test also exposed that its initial thread
  name filter was implementation-specific. The corrected test observes the
  newly created live handler itself and failed because it remained alive after
  close.

First GREEN command after the production fixes:

```text
.venv/bin/python -m pytest -q tests/macos_probe/test_generation.py::test_generation_uses_only_staged_path_and_publishes_valid_package tests/macos_probe/test_generation.py::test_generation_registration_retry_preserves_isolated_activation tests/macos_probe/test_profile_server.py
```

Actual result: `23 passed in 10.56s`.

The final socket test covers both an entirely idle accepted connection and a
partial HTTP request. Shared public-generation coverage also asserts that the
runtime receives the selected component together with `isolated=True`.

## Verification

Affected profile/runtime/generation/deadline command:

```text
.venv/bin/python -m pytest -q tests/macos_probe/test_profile_server.py tests/macos_probe/test_runtime.py tests/macos_probe/test_generation.py tests/test_generation.py tests/macos_probe/test_deadline.py
```

Actual result: `194 passed in 145.49s`.

Fresh full repository command, run once after the final test changes:

```text
.venv/bin/python -m pytest -q
```

Actual result: `2624 passed, 12 skipped in 169.59s`. The skips are the existing
native-gated macOS/Windows checks and bounded M2 registration skips; this
worker did not enable or run native WPS.

`git diff --check` exited 0 before commit preparation.

## Controller native evidence

After the generation isolation change became GREEN, the controller separately
ran the bounded legacy Writer acceptance. It exited 0 in 9.632 seconds and
observed `isolated=True`, the native blank activation document, exactly one
owned WPS host, and a completed artifact. Evidence is retained under
`build/task-session-startup/native-legacy-result.json` and the corresponding
log. This worker did not launch WPS.

## Limitations

- This fix does not alter conversion or M5 activation behavior; those routes
  already passed `isolated=True` and retain their existing tests.
- No merge, push, installation, release, or broad process termination is part
  of this fix wave.

Final fix commit: the commit containing this report, titled
`Close profile connections and isolate generation`.
