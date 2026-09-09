# Windows saving-close proxy fixwave

Date: 2026-09-09  
Status: portable implementation frozen; Windows native verification unavailable.

## Reproduced failure

`_WindowsSession.close(save_changes=True)` keeps its owned COM session open when
`save_current()` rejects a concurrent destination change. The worker returned a
complete, sequence-matched error frame and remained ready for another request.
`_SessionProxy.close()`, however, treated every exception as an unverified close:
it killed the exact worker, quarantined the component, released the lock, and
replaced the acknowledged source-conflict error with `*_QUARANTINED`.

The independent reproduction failed because the proxy was uncertain and closed
after the acknowledged error rather than remaining available for
`close(save_changes=False)`. A second reproduction showed that the worker also
cleared `NativeHandleRegistry.records` before invoking close, so the acknowledged
save failure left a live session whose existing handles had expired.

## Repair

The worker now splits an owned saving-close request into two explicit phases:

1. `save_current()`;
2. `close(save_changes=False)` only after the save succeeds.

Only an exception from phase 1 is eligible for the strict
`session_state: "open"` error acknowledgement. Before emitting it, the worker
reruns the session's read-only `_verify()` identity guard and checks the same
request deadline again. `_closed=False` alone is never accepted as proof of a
live native binding. A failed identity check, an expired deadline, or an
exception after a successful save while performing native close therefore has
no open-session acknowledgement. The proxy requires that marker on every
`close(save_changes=True)` error response before treating the error as
recoverable. Missing, unexpected, or incorrect state metadata is a malformed or
ambiguous response and still kills the exact Python worker, records recovery,
quarantines the component, and releases the lock. A transport timeout retains
the same quarantine behavior.

Attached sessions retain their prior contract: their close is not split through
`save_current()`, so `save_changes=True` cannot save an attached user document.

The worker clears native handles only after native close returns successfully.
An acknowledged save-phase error therefore leaves the registry, worker and
exclusive lock live. The original error reaches the caller, who may retry,
save elsewhere, or explicitly discard. A successful later discard clears the
handles, returns `{\"closed\": true}`, exits the worker, and releases the lock.

Production changes are limited to:

- `skills/WPSComposer/scripts/msoffice/windows_session_proxy.py`
- `skills/WPSComposer/scripts/msoffice/windows_session_worker.py`

The dedicated regressions are in
`tests/msoffice/test_windows_session_proxy.py` and cover the real subprocess
protocol, handle reuse, worker-loop continuation, real `_WindowsSession`
binding-check success and failure, same-deadline enforcement, strict state
metadata, timeout quarantine, post-save native-close failure, attached close,
explicit discard, and exact worker exit.

## Verification

Independent round-3 reproductions, including the real `_WindowsSession` failure
evidence path and handle preservation:

```text
5 passed in 0.10s
df200920edf63ab0c5afa5d523579fb4e8a7fd134814057e4d135f1056311f08  whole-integration-review-round3-repros.py
```

Focused proxy, worker, Windows business protocol, Windows document API, and the
complete independent reproduction file:

```text
158 passed in 5.92s
be714975d474250d8ff4a6c75682120b62aef99e617e478625aef7bb5bb6c8de  /tmp/windows-save-close-focused-final2.log
```

Windows engine-routing and parity-runner regressions:

```text
60 passed in 0.55s
```

`py_compile` and `git diff --check` pass for the two production modules and the
dedicated test module.

## Frozen hashes

```text
34fb8fa05deed25a8b050ee232e55b5d82566510ba53c05f9dcf247d28d61201  skills/WPSComposer/scripts/msoffice/windows_session_proxy.py
0afc0de5c8f2a86c3d48bfaa6fbee4446ec82a9496c927e192699462117d88da  skills/WPSComposer/scripts/msoffice/windows_session_worker.py
a49eea7a88899df4fb72c85e77828d4b5f35507364eee0d06361085069abcfc9  tests/msoffice/test_windows_session_proxy.py
```

No Windows Office host is available in this environment. These results prove
the closed protocol, proxy lifecycle, worker-loop, retained failure evidence,
and COM-free `_WindowsSession` integration contract; they are not native Windows
Office acceptance.
