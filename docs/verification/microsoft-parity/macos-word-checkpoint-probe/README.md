# Mac Word checkpoint native probe report

Date: 2026-09-09  
Status: `BLOCKED_QUARANTINE` (partial native evidence; production implementation must not start from this report alone)

## Scope and source state

Only the new probe and its evidence paths were changed. No production source was modified. No commit or push was made.

- Probe: `fixtures/microsoft_parity/macos_word_checkpoint_probe.py`
- Evidence: `docs/verification/microsoft-parity/macos-word-checkpoint-probe/run-01` through `run-05`
- Word version: Microsoft Word 16.112.3 on macOS
- `writer.py`: `b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205`
- `macos_word_session.py`: `39129d22315eb5701dde8f1f5cfe305de404d07ce538d10d44b2fa1f562d502a`
- `macos_runtime.py`: `fad6d9d8c5f9fb6bef50d8bca60144918818a4fc7fb116fb48be71326bfca0bb`
- Latest probe source before pause: obtain from `shasum -a 256 fixtures/microsoft_parity/macos_word_checkpoint_probe.py`; run-05 was interrupted before a final accepted evidence bind.

## Native findings

The first four geometry cases completed with acknowledged AppleEvents in runs 01-04:

| Case | Text before | Checkpoint | Native content end | Exact restored text |
|---|---:|---:|---:|---|
| Empty | `\r` | 0 | 1 | `\r` |
| Nonempty terminal paragraph | `prefix\r` | 6 | 7 | `prefix\r` |
| CJK and emoji | `前缀中文😀\r` | 6 UTF-16 units | 7 | exact |
| Appended paragraph | `base\r` | 4 | 5 | `base\r` |

For these cases the correct public integer coordinate is Word's terminal insertion bound, `(end of content of text object of boundDoc) - 1`. The final mandatory CR occupies the following native coordinate. Deleting `[checkpoint, current terminal insertion bound)` restored the prefix exactly. A second rollback at the same coordinate was an acknowledged no-op. Python/AppleScript character counts are unsuitable for bounds with surrogate pairs; CJK/emoji must use Word's native UTF-16 coordinates.

The partial-table case proved table creation requires `activate object boundWindow`. Once activated, creation succeeded, but the first deletion assertion did not converge to the original terminal bound. The actual intermediate acknowledgement was not retained before the assertion, so the number of cleanup passes and residual paragraph geometry remain unproven. Production must not implement a guessed repeated-delete loop. A fresh isolated probe must record each post-delete body text, paragraph/table counts, terminal bound, and content end before deciding whether rollback needs a table-aware cleanup step.

Runs 01 and 02 retained ordinary acknowledged Word `-2710` failures from invalid table construction. These set `retain_evidence`; they did not set the production session's `_quarantined` state. Runs 03 and 04 retained probe assertion failures after successful table construction. They remain separate failure evidence and are not acceptance runs. The probe's JSON key named `quarantine` in runs 01-04 was misleading: it listed possible recovery targets but did not reflect the production lock state. The fixture now names that key `recovery_targets`; the historical reports are intentionally retained unchanged.

Runs 03-04 removed the session staging directory during acknowledged close before the exception copier ran, so their individual AppleScript ACK logs are unavailable. Their report shows that ordinary text cases restored exactly and that the table case reached the rollback assertion, but it does not contain the first post-delete bound. The likely Word behavior is that inserting a table into the empty terminal paragraph consumes that paragraph; deleting the table recreates only Word's mandatory final CR, losing the additional explicit CR that existed at the checkpoint. For `table-prefix\r`, that would move the terminal bound from 13 to 12. This explains why a loop that only retries while the observed bound is greater than the checkpoint did not converge. It is a static inference, not accepted native evidence.

## Quarantine and exact recovery state

Run-05 was manually interrupted while `osascript` was executing. This is an uncertain native completion and therefore quarantined under the production recovery rule. No more Word mutation was issued afterward.

The post-interrupt read-only inventory was:

```text
[['文档191', '文档191', False,
  '8f2ba9aca30fbdda05bb127f45c9a5a93d8562be438cbe034913630913774ef7  -']]
```

- Exact synthetic sentinel: `文档191`, unsaved, expected content prefix `Checkpoint sentinel 中文😀 `.
- Run-05 owned private document path: `/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-8cdkxg1c/document-ef648ffdc23d401da40afde00466654f.docx`.
- The owned private path was absent from the read-only inventory; only the sentinel remained open.
- Root CUA verified the exact sentinel content, used Cmd+W and Word's Delete/Don't Save path for only `文档191`, and obtained a subsequent empty read-only inventory. Recovery evidence is `docs/verification/microsoft-parity/macos-word-checkpoint-probe/run-05/parent-recovery.json`.

## Field and tracking requirements

The pre-existing REF plus newly appended REF rollback case did not run to completion before quarantine. The production design still needs this exact native proof:

1. Record the newly appended field's semantic/native identity and result range before deleting the extent.
2. Delete only after validating `0 <= target <= current terminal insertion bound` on the exact bound document.
3. After acknowledged rollback, remove only tracking entries recorded inside the reverted extent.
4. Never inspect a deleted native field to recover its identity.
5. Preserve every pre-checkpoint REF/index identity and bookmark.
6. Invalidate changed positional targets after rollback.

## Recommended minimal implementation after a clean follow-up probe

- `degradation_checkpoint()` should return the plain native terminal insertion integer. Preserve `int` coercion in rollback. Do not add nonce, LIFO, one-use, or session-token semantics.
- Bind checkpoint and rollback AppleEvents through the existing exact `boundDoc`/path verification and the session's single deadline/lock.
- Before mutation, obtain the current terminal insertion bound and reject targets below zero or above it with `LOCAL_MUTATION_ROLLBACK_FAILED` while leaving the document untouched.
- Delete the native range beginning at the checkpoint and ending at the proven rollback bound. Table cleanup must follow a new table-only probe. A minimal candidate is to compare the acknowledged post-delete terminal bound with the integer target and insert exactly enough CR characters at that bound when Word normalized away an owned terminal empty paragraph. Do not adopt this until the probe proves exact prefix/text/table/paragraph counts and repeated rollback.
- Restore the session insertion point to the checkpoint only after acknowledged deletion and exact postcondition readback.
- Update field tracking only after rollback acknowledgement, using identities captured before deletion.
- Any timeout, malformed acknowledgement, interrupted transport, stale binding, or uncertain completion must retain evidence and lock quarantine; controller fallback must not proceed.

## Acceptance remaining

Required before production implementation or native-support claim:

- Fresh table-only rollback probe with every intermediate acknowledgement retained. It should start from `table-prefix\r`, record body text and native bounds before insertion, after insertion, after the first range delete, after any CR restoration, and after repeated rollback. It should not rerun the four already-proven text cases.
- Pre-existing REF/index plus appended-field rollback and tracking cleanup proof.
- Invalid negative and beyond-terminal bounds, zero-mutation failure, repeated rollback, saved DOCX/PDF, reopen, and sentinel hash preservation in one clean accepted run.

## Prepared follow-up fixture

`fixtures/microsoft_parity/macos_word_checkpoint_followup.py` is prepared but has not been run. Its current SHA-256 is `e956c820fec72bbc1962fdfe1623f7acf9bc13c5c1ff1001bb03b44957709a7c`.

- `--mode table-only` records and atomically persists `before`, table insertion ACK, `after-table`, first deletion ACK, and `after-first-delete`. It writes observations before deciding whether a second delete is legal. If the current bound fell below the checkpoint it records `second-delete-skipped` as a contract miss; otherwise it records the second deletion ACK and geometry. Each snapshot includes exact body text, terminal bound, content end, character/paragraph/table/field counts, paragraph boundaries, and table boundaries. It performs no CR repair.
- `--mode fields-only` creates bookmark `wpsc_checkpoint_preexisting` over native range `0..5`, reads back its exact text and bounds to prove it contains only `target`, and creates the pre-existing REF after that range. It checkpoints, appends a second REF to the same explicit bookmark, inventories both fields' exact code/result ranges before deletion, then records the deletion ACK and post-delete bookmark/field identities.
- Every successful `_execute` result is written to `report.json` before any dependent assertion. A failure therefore retains all prior acknowledged geometry.
- The modes are intended for distinct evidence directories beginning at `run-06`; neither mode reruns the four completed text cases.

Native execution remains paused until the Word cancellation fix is frozen and root explicitly clears the lease.

## Follow-up native results

Root granted the Word lease after freezing `macos_word_session.py` at `532550b1fa7dbb10de1e9025307f406bfc69e6409cf5a056da1266808488a467`. The fresh inventory before each run was empty. Word was 16.112.3.

### Run 06: table-only

Evidence: `docs/verification/microsoft-parity/macos-word-checkpoint-probe/run-06/report.json`.

- Before: body `table-prefix\r\r`, checkpoint 13, terminal bound 13, two paragraphs, no table.
- Inserted table: native table boundary `[13,24)`, 2 rows by 2 columns; terminal bound 24.
- First `Range(13,24).Delete()`: ACK returned, but the table remained 2x2 at `[13,19)`. Only visible cell content was removed. Six Word row/cell terminators remained; body was `table-prefix\r\r\x07\r\x07\r\x07\r\x07\r\x07\r\x07\r` and terminal bound 19.
- A second identical range deletion was an acknowledged no-op with the same geometry.
- Owned document close was acknowledged and final inventory was empty.

Exact cause: Word range deletion cannot remove an appended table object. Minimal rollback must snapshot pre-checkpoint table identities/bounds, delete only tables proven wholly new (`start >= checkpoint`) as native table objects, then delete the remaining appended range and require exact prefix/bounds postconditions. A table that crosses the checkpoint or cannot be distinguished from pre-existing topology must fail rollback instead of guessing.

### Runs 07-09: field-only

Run 07 retained a compile-time AppleScript `-2741` caused by an unescaped `\h`; the ordinary acknowledged failure closed the owned document, and its source/logs are retained under `run-07/failed-runtime`. Run 08 corrected escaping and proved Word range end is exclusive: `0..5` over `target` reads `targe`. It also retained all field rollback ACKs before a probe-only assertion failure. Run 09 used the correct native range `0..6` and completed the field observations; owned close returned inventory to empty.

Run 09 evidence: `docs/verification/microsoft-parity/macos-word-checkpoint-probe/run-09/report.json`.

- Bookmark `wpsc_checkpoint_preexisting` read back exactly `target`, boundary `[0,6)`, excluding mandatory CR and both REF fields.
- Pre-existing REF code was `REF wpsc_checkpoint_preexisting \\h \\* MERGEFORMAT`, result `target`, result boundary `[60,66)`; checkpoint was 67.
- After appending ` tail ` and a second REF, the appended REF had the same exact code/result and result boundary `[127,133)`. Both fields were inventoried before deletion.
- `Range(67,134).Delete()` removed the second native REF while preserving the pre-existing bookmark and REF identity. It left one visible space at the checkpoint: before body `targettarget\r`, post-rollback body `targettarget \r`, terminal bound 68 instead of 67.

Exact cause: deleting a range containing a field removes the appended field object but does not reliably restore the exact surrounding insertion text/bound. Minimal rollback must inventory newly appended field identities before deletion, delete the new native fields explicitly while they are live, then delete the remaining appended plain range and require exact checkpoint text/bound. Tracking entries may be removed only after this native sequence is acknowledged. Pre-existing bookmark/REF identities must remain byte-for-byte equivalent in the native readback.

The follow-up fixture now calculates bookmark bounds from `len(text.encode("utf-16-le")) // 2` and verifies the bookmark in a separate ACK before inserting any REF. This final fixture revision was not rerun; the native 0..6/result behavior is already captured in run 09.

## Object-first rollback verification

The frozen native runtime remained `532550b1fa7dbb10de1e9025307f406bfc69e6409cf5a056da1266808488a467`. Fresh inventories before runs 10 and 11 were empty, and both owned documents closed with final inventories empty. Each mode was attempted once.

### Run 10: table object first — PASS

Evidence: `docs/verification/microsoft-parity/macos-word-checkpoint-probe/run-10/report.json`.

- Pre-checkpoint topology: no tables; body `table-prefix\r\r`, checkpoint 13, content end 14.
- Appended table identity: table 1, native text boundary `[13,24)`, 2x2.
- Reverse object deletion selected exactly table 1 because its start was at checkpoint. Native ACK reported the deleted identity `[table, 1, 13, 24]`, zero tables, terminal bound 13, content end 14.
- Immediate readback was byte-for-byte and geometry-for-geometry equal to the pre-checkpoint snapshot, including both paragraphs and both CRs.
- Residual range deletion was a no-op. Repeating object selection returned an empty deletion list; repeating range deletion was also a no-op; final snapshot remained exact.

This validates the minimal table rollback sequence: snapshot table topology, delete in reverse order only wholly appended native tables whose starts are at or after the checkpoint, then delete the residual appended range and require exact snapshot equality. The probe did not validate a table crossing the checkpoint; that must fail closed.

### Run 11: field object first — precise failure

Evidence: `docs/verification/microsoft-parity/macos-word-checkpoint-probe/run-11/report.json`.

- Pre-checkpoint bookmark was exactly `wpsc_checkpoint_preexisting`, text `target`, `[0,6)`. The pre-existing REF code/result was preserved at `[60,66)`. Body was `targettarget\r`, checkpoint 67, content end 68.
- The appended REF was inventoried at `[127,133)` with the same exact code. Before rollback the body was `targettarget tail target\r`, bound 134.
- Reverse object deletion selected only field 2 and ACKed `[field, 2, 127, 133]`. Field count returned to one and the original bookmark/REF remained exact. The remaining body was `targettarget tail \r`, bound 73.
- Residual `Range(67,73).Delete()` left `targettarget \r`, bound 68 rather than the checkpoint 67. Repeating the full sequence selected no field and produced no change.
- Owned close was clean and inventory returned empty.

The one-round object-first field scheme therefore does not satisfy exact rollback. Explicitly deleting the appended field fixes field identity/tracking but Word retains the trailing space at the checkpoint when the remaining appended plain range is deleted. A production implementation must not claim field rollback from this sequence. A separate reviewed probe would need to distinguish the checkpoint boundary character from appended text, then verify one exact deletion/set-content primitive against plain text following an existing REF. This report does not recommend widening the deletion one coordinate left because that risks modifying the pre-checkpoint field boundary.

### Run 12: field object then range content clear — PASS

Evidence: `docs/verification/microsoft-parity/macos-word-checkpoint-probe/run-12/report.json`. The final probe SHA-256 is `d34707596a17961c2c7150ca24fdaaa1e2bba2f2843974ba25ee9a0ca25c61f9`.

- Before checkpoint: body `targettarget\r`, bound 67/content end 68; bookmark `wpsc_checkpoint_preexisting` was text `target` at `[0,6)`; field 1 code was ` REF wpsc_checkpoint_preexisting \\h \\* MERGEFORMAT ` with result `target` at `[60,66)`.
- Before rollback: body `targettarget tail target\r`, bound 134; field 2 carried the same exact code and result `target` at `[127,133)`.
- Object deletion ACK selected only field 2 (`[127,133)`) and returned one remaining field. Readback preserved bookmark `[0,6)` and field 1 code/result `[60,66)` exactly; residual body was `targettarget tail \r`, bound 73.
- The residual range was read back before mutation as exact text ` tail ` at `[67,73)`. `set content of rollbackRange to ""` returned bound 67/content end 68 and restored the entire pre-checkpoint snapshot exactly.
- Repeated object deletion selected no fields. Repeating range content clear at the empty `[67,67)` boundary kept bound 67/content end 68. The final snapshot again equaled the complete pre-checkpoint snapshot.
- Owned close completed and `after_owned_close` was `[]`.

This validates the minimal field rollback sequence without changing Word's global smart-cut settings: inventory field identities before deletion; delete only fields proven wholly appended; create the residual range from the checkpoint to the current terminal bound; set its content to the empty string; verify exact document text/bounds plus all pre-existing bookmark/field identities; only then remove reverted tracking entries. Repetition is an acknowledged no-op.

The frozen evidence set for review is:

- `run-06/report.json`: range-delete table failure, with every ACK.
- `run-07/report.json`, `run-07/failure.txt`, `run-07/failed-runtime/`: retained ordinary AppleScript compile failure and exact scripts/logs.
- `run-08/report.json`: `0..5` exclusive-end and field-range evidence retained before assertion.
- `run-09/report.json`: range-delete field smart-space failure.
- `run-10/report.json`: table-object-first exact PASS and repeated no-op.
- `run-11/report.json`: field-object-first plus `delete` residual failure.
- `run-12/report.json`: field-object-first plus local range content clear exact PASS and repeated no-op.
- `fresh-before-run10.*`, `fresh-before-run11.*`, `fresh-before-run12.*`, and each run's `starting-inventory.*`/`after-owned-close.*`: guarded inventory evidence.

No more native runs are authorized from this probe report. Production source remains unchanged.
