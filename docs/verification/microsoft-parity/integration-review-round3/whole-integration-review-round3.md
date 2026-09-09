# Microsoft parity whole-integration review, round 3

Date: 2026-09-09  
Reviewer scope: read-only review of the `428283e` to working-tree candidate after the previously closed transport, Mac logical-save, Word field/index, and Word reference reviews. I did not edit production or repository tests, start an Office/WPS process, use UI automation, or delegate review work. Reviewer-owned reproductions and this report are the only files I changed.

## Verdict

**SCOPED PASS after three related P2 repairs.** I found two Windows saving-close lifecycle defects that the owner repaired while production was frozen. A final real-session reproduction then showed that the worker still labeled a failed identity check as an open, recoverable session; the owner repaired that state proof as a bounded follow-up. All three behaviors now have independent GREEN reproductions, and I found no further worthwhile P1/P2 in the reviewed cross-component paths.

This verdict is not full parity approval. It does not complete the frozen 628-row/1,256-platform-gate audit, certify every direct-method argument, execute native Windows Office, or replace final release/UI/installation acceptance.

## Findings and disposition

### R3-1 — P2 — ADDRESSED: acknowledged Windows save failure destroyed the recoverable session

The original `_SessionProxy.close(save_changes=True)` caught every `_call("close")` exception, quarantined and killed the exact worker, released its lock, and replaced an acknowledged source/publication conflict with `NATIVE_*_QUARANTINED`. The child `_WindowsSession` contract intentionally keeps an owned document open after a failed saving close so the caller can retry, save elsewhere, or explicitly discard. The proxy erased that recovery path.

The repair does not broadly treat every close error as recoverable. The worker now separates the save phase from `close(save_changes=False)`, emits `session_state: "open"` only when `save_current()` failed before native close and the session is still open, and the proxy accepts that exact response shape. A timeout, malformed response, missing/open-state acknowledgement, quarantine error, or failure after the save phase still quarantines and releases the proxy lock. A verified save conflict preserves the original typed error, worker, session, lock, and subsequent discard path.

The reviewer reproduction initially failed because the proxy returned `QUARANTINED` instead of `Source changed since open`. It is now GREEN in `whole-integration-review-round3-repros.py`.

### R3-2 — P2 — ADDRESSED: failed saving close invalidated every live Windows native handle

The worker originally cleared `NativeHandleRegistry.records` before calling the session's `close`. When saving close returned an acknowledged recoverable error, the session stayed open but all table/range/field/worksheet/shape handles already held by the caller became unusable. That made the advertised recovery state materially incomplete and could block a corrective edit that reused a live handle.

The repair clears the registry only after close returns successfully. The independent worker-stream reproduction creates a real table handle, receives an acknowledged saving-close failure, successfully resolves the same handle to `table:1` for a later formatting call, then discards and closes.

### Recovery-evidence invariant — VERIFIED

Splitting `save_current()` from `close(False)` could have bypassed `_WindowsSession.close`'s catch block. I exercised the actual `WindowsWordSession` class through `windows_session_worker.serve` with a fake Composer and a source-change failure from the real `save_current` path. The worker invokes the real `_retain_error`, sets `_failed=True`, and a later successful discard leaves the private staging directory and `failure-*.json` intact with `native_cleanup_verified: false`. This rules out the stateless fake-session false positive for this boundary.

### R3-3 — P2 — ADDRESSED: `_closed is False` was accepted as proof of a live native binding

The split worker currently emits `session_state: "open"` for every exception in the save phase when `session._closed` remains false. That Boolean is not an identity check. The first operation in the real `_WindowsSession.save_current()` is `_verify(mutation=True)`, which can raise `DocumentIdentityError` when the native document has already closed or the original application/document identity has changed. In that case `_closed` is still false, so the worker reports the lost binding as open. The proxy then preserves the worker and lock and returns an ordinary `RuntimeError`; subsequent operations can only fail identity checks, and native discard/close has not been proved safe.

The reviewer reproduction uses an actual `WindowsWordSession` instance and the real `save_current`/worker path, with only its native Composer replaced. It receives the expected identity error plus the incorrect `session_state: "open"`; the test requires that state claim to be absent and fails. This is P2 because it breaks the promised retry/discard recovery state and can leave the caller holding an unusable exclusive session, but no data loss or unrelated-document close has been demonstrated.

The repair now reruns the session's exact read-only `_verify()` guard under the same request deadline after a save-phase exception. It emits the open acknowledgement only when that verification succeeds before the deadline. Verification failure or deadline exhaustion omits the state, causing the proxy to quarantine. The production test uses a real `WindowsWordSession` with both live and lost bindings; the independent RED is also GREEN.

## Cross-component review

- Public `document_api.edit` still stages native document/PDF artifacts before closing the owned session and publishes only after verified close. The previously reviewed destination snapshots, group rollback, deadline sharing, foreign-file preservation, and recovery-path reporting remain in place; this round did not reopen those closed findings.
- Mac Word fields and references share the same exact bound-document session, structural invalidation, one deadline, source/logical-save publication path, and conservative field topology. Reference handles are registered only after full native acknowledgement and then participate in field snapshots. No new lifecycle conflict was found between the final reference helper and field/index repair.
- Windows DTO decoding remains closed to the declared value classes and typed session handles. Nested tuple preservation, stale/foreign handle rejection, capacity preflight, and resource staging occur before the corresponding native mutation. The saving-close repair now preserves those handles only while the same verified child session remains open.
- The current status and linked direct-method inventory agree on **59 declared / 21 absent** Mac Word direct names. The docs continue to distinguish name presence, representative evidence, and certified baseline coverage.

## Independent verification

Current-source focused command:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest \
  tests/msoffice/test_windows_session_proxy.py \
  tests/msoffice/test_windows_business_protocol.py \
  tests/msoffice/test_windows_document_api.py \
  tests/test_document_api.py \
  tests/msoffice/test_macos_word_fields.py \
  tests/msoffice/test_macos_word_references.py \
  tests/msoffice/test_macos_word_logical_save.py \
  tests/msoffice/test_macos_word_session.py \
  .superpowers/sdd/2026-09-08-microsoft-wps-parity/whole-integration-review-round3-repros.py -q
```

Before the last identity-state repair, the focused set passed **448 tests in 17.96s** and the final reviewer reproduction was **4 passed / 1 failed**, with the sole failure R3-3; those immutable outputs are `whole-integration-review-round3-green.log` and `whole-integration-review-round3-red.log`. On the final repaired source, the same expanded focused set passes **452 tests in 14.04s**; see `whole-integration-review-round3-final-green.log`.

Root's immutable `unit-round17.txt` reports **4,098 passed / 12 skipped in 211.73s**. It predates the final Windows proxy/worker repairs, so it is broad pre-repair evidence rather than an exact final-source run.

I independently parsed `installed-audit-05`: its manifest has 118 present shipped files; Word, Excel, and PowerPoint each report all six bounded native checks true; and its saved post-run equality report is true. The later round-3 repair means that installation no longer matches exactly two current live source files: `windows_session_proxy.py` and `windows_session_worker.py`. Therefore installed-audit-05 is valid evidence for its recorded pre-repair snapshot, not an installed-current-source claim. This review started no replacement native or installed-plugin job.

## Frozen source hashes

| File | SHA-256 |
|---|---|
| `artifact_transport.py` | `4a5740a2943eae5ccb875da6eae31ddeeddb3b5d9f9b815b1e68250215329ef7` |
| `document_api.py` | `d8b64d325cc8906776765bab61570099e9e977e482c82f8df1ecf20e03a39a05` |
| `macos_word_session.py` | `39129d22315eb5701dde8f1f5cfe305de404d07ce538d10d44b2fa1f562d502a` |
| `macos_word_fields.py` | `c390342553fa4f9c30f77dc751da99e97ffd0abc1ff0e20822f9c5c974fa776d` |
| `macos_word_references.py` | `f2243a6ce644f6a196db087756b8dba51d6d9bb713c586a9f0175819511200f6` |
| `windows_business_protocol.py` | `e2336b13876922a4a5046ba00194795a9ea2a4cf03eb73a974ac6a8a16350d97` |
| `windows_document_api.py` | `6ae35a3de8289437ba23e36c5bd2850502f0986d951d3b626d6005995f2a9b32` |
| `windows_session_proxy.py` | `34fb8fa05deed25a8b050ee232e55b5d82566510ba53c05f9dcf247d28d61201` |
| `windows_session_worker.py` | `0afc0de5c8f2a86c3d48bfaa6fbee4446ec82a9496c927e192699462117d88da` |
| `test_windows_session_proxy.py` | `a49eea7a88899df4fb72c85e77828d4b5f35507364eee0d06361085069abcfc9` |
| reviewer reproductions | `df200920edf63ab0c5afa5d523579fb4e8a7fd134814057e4d135f1056311f08` |
| focused GREEN log | `a5a2cfdb6450d9a7554a44edf5df2b7411e0d591abf831c14fa318f97ac9e414` |
| identity-state RED log | `42cb426d421957eb7d9aba9a430d55fe04791e44be2cbe7d5ebc73c1c31e1310` |
| final focused GREEN log | `4f42f568f388eab5d67e8cd83dfdb61d2fb2c3e455270826a823c528a701e79d` |
| `unit-round17.txt` | `8a4bc3ba1a17a8c87ffab8f57984f9cb76778d6e7f95b276bbf709dda7741cb4` |

The Windows repair has portable proof only. Native Windows saving-close conflict, post-save native-close failure, timeout cleanup, live-handle recovery, and desktop/UI behavior remain pending Windows execution. The broader status correctly retains missing direct methods, unsupported argument branches, native Windows gates, and source-bound certification work rather than treating this scoped pass as full Microsoft/WPS parity.
