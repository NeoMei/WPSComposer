# Excel / PowerPoint diagnostic I/O fix candidate

2026-09-09. Worktree `.worktrees/office-description`. Parent-authorized bounded repair; only the two session modules and one related test module were changed for this implementation. The Word transport is independently owned and was not edited here. No native Office execution, UI, staging, commit, installation or publication occurred. This candidate awaits the parent's independent review.

## Defect and resulting behavior

The parent reproduced Excel's successful submitted result followed by failed `finally` log writing, leaving `_failed=False` (`docs/verification/microsoft-parity/office-diagnostic-io/excel-parent-red.json`). Independent RED tests reproduced this in both Excel and PowerPoint and showed auxiliary log/recovery/lock writes replacing the primary native error, timeout, or cancellation.

- Excel now finishes native execution/result classification before writing successful-result logs. If those logs cannot be persisted, it sets `_failed` before recovery-file/lock attempts and raises the diagnostic error. Failed native execution also sets `_failed` before any evidence attempt. Each diagnostic write is isolated so a later failure cannot replace the primary result/error.
- Excel retains its public `RuntimeError` wrapper for ordinary native failures and native timeouts; the wrapper carries the existing native timeout/execution classification as `code`, reports `diagnostic_path` only when recovery JSON was written, and exposes failed evidence steps as `diagnostic_io_failures`. `KeyboardInterrupt`/other non-Exception cancellation now propagates the original object after memory quarantine instead of being converted into a generic RuntimeError. Clipboard uncertainty metadata is preserved on recovery records and ordinary failure wrappers.
- PowerPoint classifies native exit/envelope failures before writing logs, retains `NativeOfficeError` execution/timeout codes, and propagates the original cancellation. Native failure diagnostic paths are populated only for successfully written logs. Failed diagnostics on an otherwise successful call quarantine even when `mutation=False`. Existing ordinary read-error behavior remains unchanged when diagnostic persistence succeeds.
- Both transports set their in-memory failed/uncertain flag before fallible quarantine persistence. Log/recovery/lock failures are recorded in memory by step and exception type, without private diagnostic text. Persisting one file is not reported as proof that all evidence/lock files were persisted. Where evidence writing failed, public messages avoid asserting successful diagnostic retention.
- Existing close gates skip native cleanup on failed/uncertain sessions; tests assert that a second operation and explicit close cause no second process submission. Script writing remains outside the submission boundary: injected script-write failure produces zero calls and does not set the failed/uncertain flag.

## RED / GREEN evidence

New regression file: `tests/msoffice/test_macos_office_diagnostic_io.py`. It runs the real `_run` implementations and replaces only process launch plus selected file writes/lock persistence. It covers both applications: successful result then lost logs; native failure/timeout/cancellation crossed with log, recovery, lock, and all diagnostic failures; and pre-submission script writing.

1. RED before production edits: **25 failed, 3 passed in 1.48s**. The 3 baseline PASS cases are preserved pre-submission script behavior for each application and an already-correct PowerPoint cancellation/log path. Retained output: `docs/verification/microsoft-parity/office-diagnostic-io/excel-ppt-red.log`.
2. First GREEN with session regressions: **148 passed in 1.35s**. Retained output: `docs/verification/microsoft-parity/office-diagnostic-io/excel-ppt-green.log`.
3. Final affected regression suite: **253 passed in 15.86s**. Retained output: `docs/verification/microsoft-parity/office-diagnostic-io/excel-ppt-green-final.log`. This includes the new matrix, both session suites, Excel/PPT script suites, PowerPoint logical-save and size-probe tests, and Excel/PPT parity-probe tests. `git diff --check` passed for the changed code/test paths.

Final command from the worktree:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_macos_office_diagnostic_io.py tests/msoffice/test_macos_excel_session.py tests/msoffice/test_macos_excel_script.py tests/msoffice/test_macos_powerpoint_session.py tests/msoffice/test_macos_powerpoint_script.py tests/msoffice/test_macos_powerpoint_logical_save.py tests/msoffice/test_powerpoint_size_probe.py tests/msoffice/test_parity_excel_probe.py tests/msoffice/test_parity_powerpoint_probe.py -q
```

## Frozen candidate hashes

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_excel_session.py` | `832f9517a140ec20ae69cdbf059cc7dfba48080b0a07cdb24356191686b1e797` |
| `skills/WPSComposer/scripts/msoffice/macos_powerpoint_session.py` | `7b531b31bb8094e02e159eef5072a5d079b6ac42a970299e7f4cb966f0d92cea` |
| `tests/msoffice/test_macos_office_diagnostic_io.py` | `cf848d3d13ac704afd6442dc2f8e1e8d605e07dd93e8def6c47897fe788de472` |

The existing session files contain prior parent-owned work; these are complete working-file hashes. The index was not changed. The implementation's diff from the current index is confined to diagnostic persistence helpers and `_run` failure handling in these two modules.

## Limits and review focus

This is a transport/diagnostic fault-injection result, not an actual disk-full incident, native Excel/PowerPoint round trip, cross-process quarantine guarantee under total filesystem failure, or repository-wide test result. A failed on-disk quarantine cannot guarantee exclusion of another process; the current session is blocked in memory and its close method emits no native request. Review the intentionally narrow cancellation compatibility change in Excel, successful-log failure behavior for nonmutating PPT reads, diagnostic-path truthfulness, and before-launch/after-submission separation. Native client acceptance and the complete branch suite remain parent-owned gates.
