# Windows session proxy abort persistence fix

Scope frozen for independent review. This wave changes only windows_session_proxy.py and adds test_windows_session_proxy_diagnostics.py. No Windows runtime/generation/worker/document API files were edited. No COM, native Office, index mutation, or commit occurred. Tests launch only Python protocol doubles and one unrelated Python sentinel.

## Verified root cause

The independent nine-case review suite reproduced 9 failures before this wave. _abort wrote recovery.json before quarantine, so recovery-only ENOSPC bypassed a writable quarantine location and replaced typed timeout, cancellation, or initial binding failure. _error reported both evidence paths from intentions rather than successful writes. Independent next-session tests proved a second worker could start after that failure.

## Fix

_abort marks the instance uncertain immediately, attempts quarantine before recovery.json, and records safe diagnostic failure categories independently. It then attempts exact child kill, bounded wait, I/O thread stop, and lock close independently, recording safe cleanup failure categories. No Office process is targeted. Errors in these secondary steps cannot replace the caller's primary timeout, cancellation, or binding error.

_error exposes recovery/quarantine paths only when persistence succeeded (or a pre-existing regular quarantine marker was observed) and the file remains present. If quarantine creation itself fails, quarantine_path=None and diagnostic_io_failures identifies the failure. The same proxy remains blocked by _uncertain; cross-process blocking is not claimed when durable quarantine was not created. If quarantine succeeds and only recovery fails, the next same-component session is rejected before worker launch.

_start preserves original binding/creation exception identity and annotates secondary failures. _call preserves non-Exception cancellations such as KeyboardInterrupt and SystemExit, and close preserves those cancellations too. Existing saving-close timeout-to-QUARANTINED classification, retryable acknowledged save-close errors, and ordinary successful-close behavior remain unchanged.

## Evidence

- windows-session-abort-review-red.log: unchanged independent review suite, 9 failed before production edits.
- windows-session-abort-tests-red.log: new diagnostic tests, 18 failed / 1 passed before production edits.
- windows-session-abort-fix-green.log: 207 passed in 4.14s. Composition: 179 existing session/document/host tests, 9 unchanged independent review probes, and 19 new tests.
- git diff --check passed.

New tests cover writer/sheet/slide timeout and cancellation against recovery/quarantine failures; truthfulness of current and subsequent error paths; rejection before next worker launch; initial binding exception identity; individual child-kill/I/O-stop/lock-close failures; an unrelated Python sentinel; absent pre-abort diagnostic paths; and SystemExit during close.

Reproduction command:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_windows_session_proxy_diagnostics.py .superpowers/sdd/2026-09-08-microsoft-wps-parity/test_windows_proxy_abort_review.py tests/msoffice/test_windows_session_proxy.py tests/msoffice/test_windows_document_api.py tests/msoffice/test_windows_office_host.py tests/msoffice/test_windows_host.py -q --tb=short
```

The isolated wave is windows-session-abort-fix.patch; its before snapshot is windows-session-proxy-before-abort-fix.py. The review scratch assertions were not modified. Actual Windows CI and native acceptance remain separate root-owned gates.

## Frozen SHA-256

- `ced8874be05101c8fcf5afd0fdb0ece1379e10b2d6e6726840bdeedd83ca7653`  `skills/WPSComposer/scripts/msoffice/windows_session_proxy.py`
- `b07c57a1435c73ea4307df720193c0c25ea48b99a29f334576ac1fd1c9df3c8d`  `tests/msoffice/test_windows_session_proxy_diagnostics.py`

## Independent review acceptance — 2026-09-09

**PASS: the reported Windows proxy abort P2 is closed; no new P1/P2 identified in this bounded patch. The exact candidate may be frozen for integration.** Both frozen source/test hashes above were independently matched against the live files. The original nine-case reviewer scratch remained byte-identical (SHA-256 `90a30b4b4b76c416256009a070e2a125cc7cd6980cd57e194b5e457e05cca573`), and its assertions were run unchanged.

Independent re-run of the exact six-path command above: **207 passed in 4.22s**. Review confirmed that recovery-only I/O failure no longer bypasses quarantine; typed timeouts, non-Exception cancellation identity, and initial binding failure survive secondary diagnostics/cleanup errors. Same-instance uncertainty blocks immediately, while subsequent-session blocking is claimed only when a quarantine marker was evidenced. Reported recovery/quarantine paths require both successful persistence/observed prior marker and current file presence.

Verified saving-close error acknowledgement/retry behavior and successful close remain covered by the existing worker/session tests. Exact Python-child cleanup, unrelated sentinel survival, independent cleanup attempts, and SystemExit during close also passed. No worker, document API, production tests, index, or native Office state was changed by this re-review. Actual Windows CI and native acceptance remain separate gates.
