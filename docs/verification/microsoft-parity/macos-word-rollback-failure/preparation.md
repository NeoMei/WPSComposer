# Ordinary-error rollback native gate preparation

Date: 2026-09-09. Status: **PREPARED FOR REVIEW; NATIVE NOT EXECUTED**.

Only two new formal files were added:

- `fixtures/microsoft_parity/macos_word_rollback_failure.py`
- `tests/msoffice/test_macos_word_rollback_failure.py`

Preparation first stayed in `rollback-failure-prep/` while full23 ran; formal files were added after root explicitly released that restriction. Every file in `continuation-fixwave-freeze.json` was rehashed and remains unchanged. No production, existing test, installed copy, native Office, UI, or Git publication changes were made.

## Fixture behavior for review

The guarded CLI requires `--execute` before making an output directory or invoking native work. Imports do not launch Office or create outputs. Execution creates a fresh owned Word document and separate synthetic unsaved sentinel, records the original unrelated inventory, exact owned path/window/staging/marker identity, sentinel token and full state/hash, and copies/hash-binds the exact fixture/runtime source dependencies.

The owned document receives a nonempty original paragraph and bookmark, then a public checkpoint. It appends one native REF inside one native table carrying a distinctive retained marker; this uses the construction primitives exercised by the existing recovery fixture. A real `snapshot_fields()` observation seeds the topology cache. Independent read-only native snapshots use exact Foundation path comparison and the original window ID before reading counts, full-body/prefix hashes and table facts.

Only `recovery.rollback_commands` is temporarily wrapped. The real compiler output must contain exactly one `delete recoveryField` and no previous injection. A single `error "WPSC_INJECTED_ROLLBACK_FAILURE"` is inserted immediately after it; every other command remains byte-for-byte equal. The real public rollback and real session `_execute` run unchanged. The wrapper is restored in a `finally` block. No transport is patched, general mutation hook added, fallback invoked, or destructive retry attempted.

Expected native proof requires:

- Exactly one compiler injection and the two expected session scripts for current-state preflight and rollback submission.
- Public `LOCAL_MUTATION_ROLLBACK_FAILED`, raw `WPSC_INJECTED_ROLLBACK_FAILURE (-2700)` log, quarantine marker/evidence retention, discarded observed topology and unchanged tracked prefixes.
- Later public write and actual context-exit close both raise `NATIVE_WORD_QUARANTINED`; the complete session-script hash inventory remains unchanged across both calls, and the owned session remains unclosed.
- A separately authorized **read-only diagnostic AppleEvent** after quarantine shows fields 1→0 while one table and its marker remain, with original prefix hash unchanged. This diagnostic is deliberately separate from the blocked session write/close calls, and does not reset quarantine or rebind the session.
- Unsaved sentinel exact state/hash and all preexisting unrelated inventory remain unchanged; the owned task document is still present.
- Exact source copies match their original hashes and `fixtures.microsoft_parity.evidence_gate.source_digest(ROOT)` before/after values are identical. Unknown ACK rows, malformed digests/types and boolean counts are rejected. No capability ID mappings or certification claims are generated.

Expected success status is `PASS_QUARANTINED_RETAINED`, not automatic recovery. Native logs, scripts, quarantine marker and exact recovery identities are copied to the output report. The fixture never closes either task document after the injected failure. It releases only the local process lock; root must perform exact UI discard of the owned target, independently verify owned absence and sentinel preservation, close the exact sentinel, then run guarded quarantine recovery. Unexpected fixture failures also retain the owned session for coordinated cleanup.

## Pure verification

Initial staged RED: 9 failures because the guarded fixture did not exist, retained in `rollback-failure-prep/red.txt`.

Final command:

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_rollback_failure.py -q
```

Result: **12 passed in 0.60s**, `rollback-failure-prep/green.txt`. No native or osacompile calls were made. Tests cover import/CLI guards, the actual production compiler's unique insertion point and byte-preserved remaining payload, missing/duplicate/preinjected rejection, wrapper restoration, exact read-only binding source, actual session signatures/lazy tracking, strict diagnostic ACK types, and full shipped-source digest capture even when execution is deliberately stopped before the first native call.

## Frozen identity

| File / scope | SHA-256 |
| --- | --- |
| Fixture | `5f2e3621470c874639cbc767d01005f9c7460627409bc3dc546442afb3536f0a` |
| New tests | `a12b2e1a75840b4ad12d870de71068c2363d9360218fc72dd763592ddde3714d` |
| Current complete shipped-source digest | `e3161f3ce0b894d1b1d00edfac84d63199fd8f74e85bfc1ae4532b1c3eb69361` |

Machine-readable preparation freeze: `rollback-failure-prep/freeze.json`. Runtime will recompute and retain the complete shipped digest before/after rather than assume the preparation-time value.

## Pending authorized native command

Do not run until root reviews these files and explicitly grants the exclusive Word lease. The expected output directory must not already exist.

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_rollback_failure.py --execute --output docs/verification/microsoft-parity/macos-word-rollback-failure/run-01
```

Root separately owns the existing 13-check recovery regression and subsequent native/UI cleanup acceptance. This preparation report does not claim those future results.
