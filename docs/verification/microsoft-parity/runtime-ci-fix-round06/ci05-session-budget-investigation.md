# CI05 session-budget boundary investigation

Result: confirmed P2 in the shared persistent Windows session proxy. A valid new session can emit `remaining_seconds > 600` by floating-point rounding when both monotonic readings observe the same clock tick. The unchanged real worker correctly rejects that frame, so the public session fails during initial binding. This is production request generation, not a worker-limit or assertion problem.

## Source and raw evidence

Reviewed source SHA-256:

- `skills/WPSComposer/scripts/msoffice/windows_session_proxy.py`: `dfa9abfba9962db621a63bccf4058d923ad293a2c4f93888aa6ec507d5294b77`
- `skills/WPSComposer/scripts/msoffice/windows_session_worker.py`: `0afc0de5c8f2a86c3d48bfaa6fbee4446ec82a9496c927e192699462117d88da`
- `tests/msoffice/test_windows_session_proxy.py`: `a49eea7a88899df4fb72c85e77828d4b5f35507364eee0d06361085069abcfc9`

Proxy line 115 initializes `self._deadline = time.monotonic() + 600`. Lines 194–198 subtract a later monotonic reading and return the raw positive result. Line 318 serializes it into the request. Worker lines 353–355 require a finite numeric budget satisfying `0 < budget <= 600`. The proxy's base class supplies Word, Excel, and PowerPoint session factories. On initial failure its existing abort path runs; no successful native binding should be claimed.

Read CI05 XML with ElementTree from `docs/verification/microsoft-parity/portable-ci/run-05/artifacts/*/pytest.xml`. Both Ubuntu jobs: 4992 tests, 0 failures, 44 skips. Windows 3.9: 4992 tests, 1 failure, 75 skips. Windows 3.12: 4992 tests, 2 failures, 75 skips. All three failures are `ValueError: Invalid remaining session budget` in the initial real-worker binding. The Windows 3.9 handle-roundtrip failure includes the actual serialized request:

```json
{"protocol":1,"id":1,"kind":"sheet","method":"new_document","args":[],"kwargs":{"visible":false},"deadline":2158.5460000000003,"remaining_seconds":600.0000000000002}
```

Its actual worker response is `status: error`, `type: ValueError`, `message: Invalid remaining session budget`. The Windows 3.12 XML establishes the same rejection/classification, but does not expose the numeric request frame in its traceback; do not invent its exact tick. The observed Windows 3.9 values are exactly reproduced by `(1558.546 + 600) - 1558.546`. Parent's independent candidate tick `424.005` similarly yields `600.0000000000001`. A repeated coarse clock tick is sufficient; no claim about the CI host's clock implementation is needed to establish the invalid frame bug.

## Deterministic real-worker reproduction

Scratch: `checkpoint-integration-scratch/test_session_budget_float_review.py`, SHA-256 `ed408cc8219c11a925f7ccae03bb04284f7e05fcd2cedf0dfb8f10692fbfe6a0`.

The scratch imports and invokes both existing tests without changing their assertions:

- `test_actual_worker_serve_subprocess_preserves_error_metadata_and_close`
- `test_handles_roundtrip_real_worker_with_com_object_return_doubles`

It replaces only the proxy module's `time` binding with a constant-clock namespace, avoiding global clock changes in threading/queues and leaving each real Python worker's clock and strict protocol validator untouched. Both tests run at both ticks above, spawning actual `windows_session_worker.serve` processes with COM-free session doubles.

Unmodified production: **4 failed, 1 passed in 7.13s**. All four original-contract tests fail on initial binding with the same `ValueError`; the separate expiry/non-extension arithmetic guard passes. Log: `checkpoint-integration-scratch/session-budget-red.log`.

Process-local prototype only (`REVIEW_CAP_PROTOTYPE=1`): wrap `_remaining` as `min(600, original(self))`. Identical test assertions: **5 passed in 3.13s**. Log: `checkpoint-integration-scratch/session-budget-prototype.log`. This is diagnostic proof, not an implemented or accepted production repair.

Commands use the full clean-dev Python, `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.:tests/msoffice`, `pytest -q -p no:cacheprovider`, and separate scratch basetemps. No source, permanent tests, Git/index, or native Office state was changed.

## Narrow repair and acceptance requirements

Propose changing only the positive return in `_remaining` to `return min(600, value)`, keeping the existing `value <= 0` timeout check first. This can only reduce a computed positive budget, never increases it, never modifies the proxy's absolute deadline, and preserves expiry. Do not add an epsilon to the worker's maximum, extend deadlines, retry session creation, or relax the existing real-worker assertions. Worker deadline shortening remains unchanged.

Permanent RED coverage should retain a fixed same-tick real-worker regression for both existing public scenarios, plus direct checks for smaller remaining budgets and expired/zero budgets. Preserve a worker-side rejection check for a frame just above 600; the repair belongs to the producer. The shared base-class placement covers all three components, while native Office acceptance remains a separate gate.

The current source is not approved for this boundary. A production repair needs independent re-review and the parent's next full36/CI validation; the concurrently frozen full35 bridge run cannot establish acceptance of a later session-proxy change. No overall Office parity completion claim follows from this investigation.
