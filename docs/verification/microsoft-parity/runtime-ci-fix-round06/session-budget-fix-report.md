# Session remaining-budget float boundary repair

Implementation frozen for original reviewer recheck. No native Office/WPS operation, shared-index operation, commit, or existing full35 snapshot modification occurred. This correction requires full36.

## Defect and exact change

The recorded CI05 Windows 3.9 case computes deadline `2158.5460000000003` from clock `1558.546` plus 600. At the same clock value subtraction produces `600.0000000000002`, outside the worker's strict `0 < remaining_seconds <= 600` protocol. A second clock value, `424.005`, reproduces the same category of rounding overflow.

Production change is exactly `_remaining()` returning `min(600, value)` after its unchanged `value <= 0` expiry guard. `_deadline` is not modified. Positive budgets below the cap remain unchanged, expired sessions still raise the same TimeoutError, and no epsilon, retry, deadline extension, or worker-limit change was added.

## RED/GREEN evidence

- Original reviewer scratch unchanged, without its optional prototype environment switch: **4 failed, 1 passed in 2.22s**. `session-budget-implementation-red.log`.
- Six new permanent cases before production edit: **5 failed, 1 passed, 43 deselected in 3.59s**. `session-budget-permanent-red.log`.
- After the one-line repair: **85 passed in 12.44s**. `session-budget-implementation-green.log`.

Command:

```sh
PYTHONPATH=.:tests/msoffice /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest -q tests/msoffice/test_windows_session_proxy.py tests/msoffice/test_windows_session_proxy_diagnostics.py .superpowers/sdd/2026-09-08-microsoft-wps-parity/checkpoint-integration-scratch/test_session_budget_float_review.py
```

The four permanent real-worker cases call the two original subprocess scenario tests unchanged at both offending clock values. Their native session objects are doubles; actual Python transport/worker protocol, error metadata, close behavior, returned handles and stale-handle rejection remain exercised. Only the proxy clock is replaced, not the worker clock.

The budget invariant case verifies the exact publication deadline, unchanged positive sub-600 values, cap without increasing available time, and failure at/after expiry. The separate real `serve` guard supplies `600.0000000000002` directly and requires ValueError before the session factory is invoked, proving the worker maximum was not relaxed. The original proxy/diagnostic suites were included.

## Freeze

- Proxy SHA-256: `96bebb087a83fc2fda164cad6d98a43a46e131b0d896f3aace037626026a2dbf`
- Proxy test SHA-256: `30175bbaafbfdc674fd135dbf748fbcf7ddf7908aa107f584af5a2d64fc5064a`
- Unchanged worker SHA-256: `0afc0de5c8f2a86c3d48bfaa6fbee4446ec82a9496c927e192699462117d88da`
- Patch: `session-budget-frozen.patch`
- Original scratch retained: `checkpoint-integration-scratch/test_session_budget_float_review.py`
- Prior reviewer/CI failure and prototype logs were not replaced.

Changed-file `git diff --check` passed; worker source has no working diff. These are local pure/Python-worker checks, not fresh Windows native or hosted CI acceptance. Independent recheck and full36 remain pending.

## Independent reviewer acceptance

Original investigator independently reviewed the exact one-line production delta and all six permanent cases. No new actionable finding in this bounded repair. The positive budget cap only reduces the computed remaining time; the expiry guard, publication deadline, protocol classification and worker's strict maximum remain unchanged. Both original real-worker scenarios retain their existing assertions and use isolated proxy clocks. The direct over-limit worker frame still fails before factory invocation.

Independent verification, without `REVIEW_CAP_PROTOTYPE`: **11 passed, 43 deselected in 2.74s**. This includes the six new permanent cases and all five unchanged original scratch cases; real subprocess binding, handles/error metadata and close assertions pass against actual repaired production code. Log: `checkpoint-integration-scratch/session-budget-final-review.log`. Used the complete clean-dev interpreter, `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.:tests/msoffice`, `-p no:cacheprovider`, and an owned scratch basetemp. Selection was `floating_point_budget_boundary or budget_cap_preserves or strict_maximum or original_real_worker_contract_at_same_tick or cap_preserves_expiry` across the permanent proxy test and original scratch.

Hashes checked before and after this run:

- Proxy: `96bebb087a83fc2fda164cad6d98a43a46e131b0d896f3aace037626026a2dbf`
- Permanent test: `30175bbaafbfdc674fd135dbf748fbcf7ddf7908aa107f584af5a2d64fc5064a`
- Worker remains `0afc0de5c8f2a86c3d48bfaa6fbee4446ec82a9496c927e192699462117d88da`.
- Original scratch remains `ed408cc8219c11a925f7ccae03bb04284f7e05fcd2cedf0dfb8f10692fbfe6a0`.

Accepted for full36/CI06 freeze. This acceptance closes the reproduced producer-side budget defect at the reviewed hashes; full36 and hosted Windows results remain separate pending gates. No overall parity completion claim. Reviewer changed only this report and owned scratch test artifacts; no source, permanent test, Git/index, or native state mutation.
