# Word rollback ordinary-failure native fixture final review

Date: 2026-09-09

Scope: independent read-only review of the frozen guarded fault fixture and its
run-02 evidence for an ordinary native error injected during checkpoint
rollback. This review does not reopen the already closed production fix and
does not treat root's later UI cleanup as automatic recovery. I changed no
production, test or fixture file, launched no Office/native/UI process, and made
no commit or installation.

## Verdict

**SCOPED PASS — no actionable P1/P2 found in the requested 14-check native
failure evidence and cleanup ledger.** Run 02 proves the bounded failure shape:
the real rollback deleted the appended REF field, then returned an ordinary
`-2700` before deleting the table or appended body; the public call failed and
quarantined the session; later public write and context-close attempts sent no
AppleEvent. Recovery remained explicitly manual.

Run 01 remains an immutable `FAIL_RETAINED` setup failure caused by requiring a
window ID that Word did not provide. It stopped before the fault injection and
is not used as evidence for the production behavior. Run 02 uses the corrected
exact file-owned document identity while retaining missing window metadata.

## Frozen source identity

- `fixtures/microsoft_parity/macos_word_rollback_failure.py`:
  `80af0630d881199dbf89fbf136bb50cac63906cbffb93df3b9f7a7abd8303a44`
- `tests/msoffice/test_macos_word_rollback_failure.py`:
  `e0c2d0b87f6f605b913d5dfc71edc3efa151160b77bfbb4760769ae5d8963a1a`

Both match the requested freeze. All ten source dependencies recorded in the
run-02 report independently match the current files, including recovery
`20721fb2...`, session `3ab747a0...` and fields `446e8c9f...`. The report's full
shipped-source digest is unchanged before/after:
`e3161f3ce0b894d1b1d00edfac84d63199fd8f74e85bfc1ae4532b1c3eb69361`.

## Fault placement and real partial mutation

The fixture wraps the real `rollback_commands()` only for the bounded call and
requires exactly one original `delete recoveryField`. It inserts exactly one
`error "WPSC_INJECTED_ROLLBACK_FAILURE"` immediately after that deletion and
restores the compiler in `finally`.

Independent inspection of the retained native scripts found **9 total session
scripts** and exactly **1 injected script**,
`0552bc86050748f29212cbecf6a76db8.applescript`. Its relevant order is:

1. line 159: `delete recoveryField`
2. line 160: injected ordinary error
3. line 181: `delete recoveryTable`
4. line 186: clear `rollbackRange`
5. line 320: final checkpoint postcondition

The retained matching log is the only native log containing the injection and
records `WPSC_INJECTED_ROLLBACK_FAILURE (-2700)`. Therefore field deletion was
submitted before the ordinary error, while the table/body deletion commands
were unreachable.

Independent exact-path readback before versus after the failure confirms the
actual partial state:

- field count: **1 -> 0**;
- table count: **1 -> 1**;
- table marker remains present;
- the table's story extent/hash changes as the REF is deleted, so the fixture
  does not falsely demand the entire post-append body hash remain unchanged;
- the checkpoint-prefix SHA-256 remains exactly
  `d4fcc7f16249392bbd72bc79723ded8a2dbf2b0bbdc4ed02dd8f0ba3ad68550c`.

This is direct native readback of the exact private owned path, not an OOXML
inference or active-document selection.

## Public failure, quarantine and ownership evidence

The public call returned `NativeWriterObjectError` with code
`LOCAL_MUTATION_ROLLBACK_FAILED`. The observed field cache was invalidated, but
tracked state was not committed to the checkpoint. The session reports
`quarantined=true`, `closed=false`, and retains the ordinary-error log plus
quarantine marker.

Native script counts are **7 before rollback, 9 after rollback, and 9 after the
blocked write and context-close attempts**. Both later operations returned
`NATIVE_WORD_QUARANTINED`. This proves the public session emitted no cleanup,
retry or later mutation after the uncertain partial rollback.

The separately created unsaved sentinel retains the exact tuple and body hash:

```text
["文档68", "文档68", false,
 "19f68849c036ec4886e64f790373976d0beab54ab59aae2b583f07c66222b550"]
```

The post-quarantine inventory contains exactly that sentinel plus the retained
private owned document. The initial unrelated inventory was empty. The source
artifact digest also remains unchanged across the run.

All **14 named checks** in `run-02/report.json` are true, and the report status
is narrowly `PASS_QUARANTINED_RETAINED`. Report SHA-256:
`17c17d565d7afd3fabcb7e15f27db3606bd9f0148acb7b67f1341b23c5938b38`.

## Manual cleanup disposition

`parent-review-cleanup.json` is a separate parent/UI recovery ledger. Root
visually observed the exact private target with the original prefix and one
table while the REF was absent, then closed that exact document and explicitly
chose **Do Not Save**. Only the unchanged sentinel remained. Root then used the
independent exact name/path-empty/saved-false/token guard to close the sentinel
and separately cleared the quarantine; final inventory is empty.

The cleanup report explicitly sets
`cleanup_is_manual_not_automatic_rollback=true`. Its SHA-256 is
`165c5c2d8e34650d5dbcc3b91f6184d335b40dca64d210d80c3b76d73d7b03c9`.
I independently cross-checked that its post-owned-discard row equals the
run-02 sentinel preimage and that both post-sentinel and final inventories are
empty. This establishes controlled manual disposal and host recovery only. It
does not establish automatic rollback, automatic target closure or retry.

## Fixture quality and verification

The native fixture refuses execution without `--execute`, refuses an existing
output directory, snapshots its shipped sources, and deliberately leaves the
quarantined target for external review rather than attempting cleanup through
the failed session. Its after-failure diagnostics are independent, exact-path,
read-only AppleScripts and never clear/rebind the quarantined session.

Independent pure command:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest -q \
  tests/msoffice/test_macos_word_rollback_failure.py
```

Result: **21 passed in 0.84s**. These tests cover import/CLI guards, unique
injection placement, compiler restoration, exact diagnostic binding and
malformed/identity-drift rejection. Read-only consistency checks additionally
confirmed all source hashes, all 14 booleans, the script/log injection count
and ordering, field/table/prefix transitions, sentinel equality, original
source digest and final cleanup inventories.

This acceptance is intentionally limited to one ordinary error after a real
field deletion and before later destructive commands. It does not claim that a
partially mutated document was automatically restored, saved or safe to resume;
the required behavior is quarantine, retained evidence and explicit recovery.
