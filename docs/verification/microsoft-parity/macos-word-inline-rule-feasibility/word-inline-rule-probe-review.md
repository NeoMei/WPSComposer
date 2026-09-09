# Inline rule runner scope review

Current verdict: **SCOPED PASS after correction**, for runner `ebf8a0480a25efa504311ae4d513c0e341ced6b36e264655e1788935e714ad82` and tests `8edefe47530909336f90c34983bca844222e00fabccd4c7ddba8dd330b7fe1ca`. The original P1 below is resolved; its evidence is retained as review history. This is fixture readiness only, not native or production acceptance.

Reviewed only the prepared inline fixture and matching pure tests, plus the imported helpers needed to follow their behavior. No production changes, native Word execution, AppleEvents, osascript, osacompile, UI actions, or lease acquisition were performed. Other workers' current changes were left untouched.

Reviewed runner SHA-256: `3077559de0062195c5a86472e743ecf64a9d8f6e4e28e258a7ef7ad9c9a78e32`.
Frozen constructor SHA-256 independently matches: `386aa4dc70de7eac48742535c37a2e97e36aa77ca40a21626e4124efa50a2ba3`.

## P1 — Stop all native follow-ups after uncertain independent sentinel cleanup

Location: `fixtures/microsoft_parity/macos_word_inline_rule_feasibility.py:138-141`, `:253`, and `:278-280`.

The sentinel close runs through raw `subprocess.run`, after the owned session has closed. If this call times out, loses its connection, or returns an invalid completion acknowledgement, the runner retains an error but neither closed session becomes quarantined. The finalizer computes quarantine exclusively from those sessions and then calls native `inventory(output, 'final')`. This violates the required zero-native-follow-up behavior after uncertain completion and can send another AppleEvent into an unresponsive or disconnected Word instance. The same gap applies if the independent post-owned-close inventory fails uncertainly. The sentinel close is not retried, but the final inventory is still a native follow-up.

Pure reproduction reused the test's fully mocked complete-success flow and replaced only the sentinel-close subprocess with `TimeoutExpired('synthetic-sentinel-close', 30)`. Observed:

```json
{
  "calls": ["before", "sentinel-before", "sentinel-after", "after-owned-close", "final"],
  "status": "FAIL",
  "quarantined": false,
  "error": {"type": "TimeoutExpired"},
  "remaining_sentinel": "sentinel"
}
```

Fix within this fixture: retain a runner-level native-uncertainty/quarantine state for independent inventory/cleanup transport, preserve diagnostic output and the sentinel recovery identity, and skip subsequent native calls once completion becomes uncertain. Keep verified sentinel-identity refusal distinguishable from transport uncertainty. Add pure failure-path coverage for sentinel-close timeout/connection loss/invalid acknowledgement and post-owned-close inventory uncertainty; assert no final inventory and no retry. Preserve an existing constructor error if cleanup subsequently fails.

## Other reviewed gates

- Explicit `execute is True` guard and fresh output directory precede native work.
- The constructor remains the single exact document-container inline candidate; the old multi-case runner and WordArt constructor are not invoked.
- Native count delta, inline type, finite positive geometry, repeated reopen readback, source DOCX hash, persisted inline XML, complete XML/relationship extraction, and PDF vector geometry have distinct checks.
- Owned close precedes independent inventory comparison and exact unsaved sentinel name/path/text/state guards.
- Ordinary constructor failure preserves its original error and attempts one partial checkpoint; existing session quarantine prevents partial native reads/saves and blocks session cleanup/final inventory.
- Initial/live/retained source and dictionary hashes are bound together; dictionary preflight failure retains a report without native inventory.
- Actual native success, PDF visual inspection, production contract implementation, and UI edit/undo/save/close/reopen remain unfulfilled gates.

## Pure validation

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_inline_rule_feasibility.py -q
```

Result: **22 passed**, 5 PyMuPDF SWIG deprecation warnings, 0.30 seconds. Additional fault injection above ran against mocked owners, inventory, and subprocess only. No native operation was executed.

## Correction re-review

Independently read the corrected runner and failure-path tests and reran the dedicated pure suite: **43 passed**, 5 upstream PyMuPDF SWIG deprecation warnings, 0.42 seconds. No native Word, osascript, osacompile, UI, or shared-helper mutation was performed.

- The runner-level uncertainty state now survives final quarantine calculation and blocks independent cleanup, reopen, and final inventory. The reviewed timeout, connection loss (`-609`), AppleEvent timeout (`-1712`), malformed JSON, incorrect ACK, and post-owned-close inventory error paths have no subsequent native calls or close retries.
- File-not-found/permission failures before child launch remain `not-submitted`; the exact script's explicit identity guard refusal remains `identity-refused-before-close`. Neither is silently converted to success. Other unverified outcomes fail closed. The imported inventory wrapper conservatively treats any error as uncertain because it does not expose its submission boundary.
- Uncertainty while the owner is live quarantines that owner through its existing local `_retain` operation, so native context cleanup is blocked. Uncertainty after verified owned close preserves `_closed=True` and does not falsely mark the closed session open or session-quarantined.
- Recovery markers retain the stage, owned path/staging root/real closed state, and known exact sentinel name/preimage. Partial stdout/stderr are retained on timeout. A valid sentinel-close ACK remains recorded if a later final inventory fails.
- The original constructor error stays primary when cleanup also fails, with partial artifacts and separate cleanup diagnostics retained. PASS remains impossible on uncertainty.
- Frozen constructor hash independently remains `386aa4dc70de7eac48742535c37a2e97e36aa77ca40a21626e4124efa50a2ba3`. Single-candidate/no-WordArt scope and the original native/XML/PDF/reopen/source/dictionary/execute gates remain intact.

No further concrete blocking finding in this scoped re-review. Root may schedule the one guarded inline native attempt after granting the exclusive lease and confirming source stability. Native result, PDF visual review, production implementation, and UI acceptance remain separate pending gates.
