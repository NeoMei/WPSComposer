# Windows session recovery / ownership independent review

Candidate HEAD: `595268e`. Scope: Windows session proxy/worker, document-session lifecycle, and the Word/Excel/PowerPoint ownership guards reached by close/save. Review used code inspection and Python-only test transports; no native Office application was operated. No production source or shared index was changed.

**Decision: one new P2 defect boundary; no new P1 identified.** This is the Windows persistent-session proxy path, separate from the already-closed `macos_office_runtime._execute` diagnostic/cleanup fixes. Existing save-close acknowledgement and binding-reverification fixes are not re-reported.

## P2: abort diagnostics can bypass durable quarantine and mask primary failure

`skills/WPSComposer/scripts/msoffice/windows_session_proxy.py`, `_abort`, lines 209–234, writes session `recovery.json` before calling `self._lock.quarantine(detail)`. An I/O failure limited to recovery metadata skips quarantine even when the independent quarantine path is writable. The `finally` branch still kills the exact Python child and releases its lock. Office is intentionally not killed, so its command/document may remain live; a later same-component session can now launch without encountering an uncertainty marker.

At the caller boundary this also substitutes the diagnostic `OSError` for the intended typed timeout or cancellation. `_start` has the same problem: a failed binding acknowledgement is replaced by `_abort`'s recovery-write failure. `_error` unconditionally reports `recovery.json`, and reports the quarantine path based only on the in-memory `_uncertain` flag; a subsequent call can therefore advertise both paths despite neither file having been persisted.

These manifestations share the same abort-persistence boundary and should be handled as one narrow repair. Attempt quarantine independently before optional recovery metadata; preserve typed timeout, initial binding error, and cancellation identity while recording safe secondary failures; return only evidenced diagnostic/quarantine paths. If quarantine creation itself fails, state that explicitly without claiming a cross-process block. Aborted Python-child cleanup and lock release must still be attempted, without touching unrelated native Office processes. Successful-close behavior and verified recoverable save-close errors must remain unchanged.

`_call` currently converts a caught cancellation to a quarantine error even without a diagnostic fault; the cancellation probes intentionally require the original cancellation to survive abort cleanup. The broader cleanup recommendation is limited to this abort/start boundary rather than arbitrary fault combinations elsewhere.

## Frozen pure repro

`test_windows_proxy_abort_review.py`, retained beside this report, uses the existing `test_windows_session_proxy.transport` fixture. This fixture launches only small Python line-protocol doubles, plus a lock double whose normal acquisition rejects an existing quarantine file. It does not load COM or launch Office.

Command:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest -q .superpowers/sdd/2026-09-08-microsoft-wps-parity/test_windows_proxy_abort_review.py --tb=short
```

Observed: **9 failed in 0.56s** on the frozen candidate:

- Writer, Excel, PowerPoint timeout × recovery-only ENOSPC: observed `(code=None, quarantine_exists=False, lock_closed=True, python_exited=True)` instead of the component timeout plus quarantine.
- Writer, Excel, PowerPoint injected `KeyboardInterrupt` while receiving × recovery-only ENOSPC: observed `(original_cancel=False, quarantine_exists=False, lock_closed=True, python_exited=True)`.
- Next Excel session: no quarantine exception; a second Python worker starts and completes close.
- Subsequent `_error('QUARANTINED')`: claims two nonexistent recovery/quarantine paths.
- Initial wrong-kind acknowledgement: expected the original binding `ValueError`; received recovery `OSError(ENOSPC)`.

The injected recovery failure matches only the first session's exact `recovery.json`, leaving the quarantine path writable; this isolates ordering from a hypothetical total-volume failure. The next-worker test verifies the resulting failure of the gate, rather than merely checking whether the quarantine method was called.

## Existing behavior checked

Independent existing suites: `test_windows_session_proxy.py` **43 passed in 1.53s**; `test_windows_document_api.py`, `test_windows_office_host.py`, `test_windows_host.py` **136 passed in 1.92s**. Thus **179 existing tests passed**, separately from the nine RED probes.

The inspected worker keeps failed saving-close requests retryable only after read-only binding verification and deadline validation; attached close is not split into a save. Session methods verify exact captured application/document tokens, and native owned close checks exact native document identity before closing. No additional actionable ownership or stale-binding defect was confirmed in this bounded pass. Passing these tests does not claim Windows/native acceptance.

Frozen SHA-256:

```text
34fb8fa05deed25a8b050ee232e55b5d82566510ba53c05f9dcf247d28d61201  skills/WPSComposer/scripts/msoffice/windows_session_proxy.py
0afc0de5c8f2a86c3d48bfaa6fbee4446ec82a9496c927e192699462117d88da  skills/WPSComposer/scripts/msoffice/windows_session_worker.py
6ae35a3de8289437ba23e36c5bd2850502f0986d951d3b626d6005995f2a9b32  skills/WPSComposer/scripts/msoffice/windows_document_api.py
90a30b4b4b76c416256009a070e2a125cc7cd6980cd57e194b5e457e05cca573  .superpowers/sdd/2026-09-08-microsoft-wps-parity/test_windows_proxy_abort_review.py
```

## Final closure — 2026-09-09

The P2 above is **closed** after independent review of proxy SHA-256 `ced8874be05101c8fcf5afd0fdb0ece1379e10b2d6e6726840bdeedd83ca7653` and dedicated tests SHA-256 `b07c57a1435c73ea4307df720193c0c25ea48b99a29f334576ac1fd1c9df3c8d`. All original nine scratch assertions remained unchanged and now pass. Combined independent re-run: **207 passed in 4.22s**. No additional actionable P1/P2 was identified; this exact proxy candidate can be frozen for integration. Detailed independent acceptance is appended to `windows-session-abort-fix-report.md`. Historical RED observations above are retained as failure evidence, not current candidate status.
