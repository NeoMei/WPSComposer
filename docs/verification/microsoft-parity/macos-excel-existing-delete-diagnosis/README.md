# Existing Excel worksheet deletion: isolated-process diagnosis

September 10, 2026; production candidate `55b27fa40835059ffad08a1d8d5a286cc3f4da52`.
**Diagnostic only: existing-sheet deletion remains blocked. No production code or capability gate changed.**

## Observations

- The original Excel PID was 9212. Launching Excel with `open -n -g -a` created PID 99513 at `Thu Sep 10 12:03:54 2026`, initially containing only saved blank 工作簿1. PID-addressed ScriptingBridge reads distinguished both inventories.
- The Excel `smXL/1169` open command returned an empty reply without opening the file through ScriptingBridge, direct Apple events, and PID-retargeted AppleScript. A zero process exit status alone was not accepted as success. POSIX, HFS and container-path trials are retained separately.
- A standard `aevt/odoc` open event successfully opened the exact synthetic container copy in PID 99513, replacing its initial blank workbook. This particular input was a byte-verified macro-free XLSX fixture with no external-link parts; this does not establish a general secure-open alternative to the existing editable-open command.
- Writing `PID-WRITE-PROBE-99513` into Data!Z40 read back correctly and changed saved to false. This proves only in-memory editing in the isolated process.
- Setting display alerts to false read back as **true**. Both instances continued to report true. No deletion was attempted, since prompt suppression was unverified.
- Native save returned without an error, but saved remained false. Independent final file hashing proves the marker was not persisted. Native close similarly returned without removing the workbook. The roundtrip probe therefore failed before reopen. These observations do not establish their underlying cause or claim all PID-addressed Excel commands fail.

## Cleanup and preservation

After rechecking process start time, bundle identity, the sole workbook's exact synthetic path, sheet sequence, and original-process inventory, the parent sent SIGTERM **only to owned PID 99513**. The test process exited. This manual diagnostic cleanup is not an accepted production recovery implementation.

The final exact original-workbook read confirms 工作簿5 and 工作簿7 remain unsaved with their original markers `sentinel-58d9ceb7df8e` and `sentinel-e9a7d3dce1ad`; display alerts remains true. The first iterator-based postcheck failed with native error -50; its log is retained, and the exact-name retry passed. Source and disposable on-disk copy both retain SHA-256 `afa1135ef9ed62cd07e18cb337a8383f3f5af894cc7c833db32ea1a059012156`; ZIP integrity and original three-sheet order pass. No certificate, macro/security preference, template, or installed plugin changed.

## Scope and next step

Do not remove the existing-sheet deletion guard based on this experiment. A usable isolated transport still needs secure open, verified alert suppression or a confirmation-free delete primitive, durable save, close, timeout handling, and user-process preservation. Resolve the ineffective setters/commands before integrating it.

The archived Objective-C/Swift and AppleScript files contain run-specific PIDs and owned paths. They are historical diagnostic source, **not reusable scripts or production helpers**; do not run them unchanged. Compiled executables are excluded.

Current hosted CI metadata independently confirms all four Linux/Windows Python 3.9/3.12 jobs succeeded for the exact candidate in run 34403359923. This is portable CI only. A fresh remote Windows snapshot remains `notLoaded` on old interrupted turn 01a07cdd-8e86-7970-b9dd-6318bc2598a6 (cursor 2598579f-1dd6-4480-a080-218100d4aff5:5), with no current native command evidence. All six parity tasks remain partial; no release approval is given.

Local follow-up: `tests/msoffice/test_macos_excel_session.py` passes **69 tests in 0.16s**. All 72 archived diagnostic file hashes verified before this documentation addition; production, tests and fixtures remain unchanged. The full suite was not rerun for this diagnostic-only change.
