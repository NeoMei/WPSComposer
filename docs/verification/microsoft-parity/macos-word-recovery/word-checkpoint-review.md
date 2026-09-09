# Independent scoped review: Mac Word checkpoint and rollback

Date: 2026-09-09  
Scope: the frozen two-method checkpoint implementation, its session forwards,
focused tests, and native fixture. This review started no Office process and
made no production or product-test changes.

## Verdict

**SCOPED PASS — no remaining P1/P2 in the reviewed checkpoint/rollback
slice.** The original floating-shape rollback false success and the follow-up
generic-inline-object gap are both addressed. The current implementation is
ready for its final source-bound native acceptance run.

This verdict covers append rollback of the native text/table/field path and
safe refusal when floating or inline object topology changes. It does not
claim shape/figure/equation deletion rollback, all Word object parity, or the
full Microsoft/WPS parity baseline.

## Source freeze reviewed

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_recovery.py` | `5ee6b2e65e4e13c6a963bd68d0d576fc24ab31cf6fe18fc586ff13356da9090f` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `ee86d43761c7e503b999bddc09b25d583cd4ca494a04e83b2c6722f5c3c6fa7c` |
| `tests/msoffice/test_macos_word_recovery.py` | `975336f85dc46c070540969866914916f343e96a0c917a44d481ac717b66b0e6` |
| `fixtures/microsoft_parity/macos_word_recovery.py` | `5efad969d6acb04875c5bf86a2ffd9a6e36e0d1ba30bfe34d4f632a5d80eb7f7` |

The session delta from the independently frozen topology source
`a96f96f99667e04ec64ed3eb018ba45de308db9a308d37b3df9c98bf2537bc6c`
is bounded to the checkpoint mapping, two public forwards, and exact stale
document / AppleEvent `-609` quarantine mappings. The separately reviewed
field-topology submission-boundary logic is unchanged.

## Original floating-shape finding: ADDRESSED

The immutable checkpoint now inventories every native floating `shape`,
including ordinal, name/type, `anchorID`, `editID`, anchor range, position,
dimensions, rotation, z-order, visibility and a SHA-256 digest of text-frame
content. The mandatory summary row independently counts shapes.

Rollback obtains a fresh complete inventory. `_preflight()` requires the
entire floating-shape fact sequence to equal the checkpoint sequence before it
generates or submits the destructive rollback batch. A new shape can keep the
same main-story terminal coordinate and prefix hash, but it still changes the
shape facts and therefore raises `LOCAL_MUTATION_ROLLBACK_FAILED`. The recovery
controller treats that rollback failure as fatal: fallback is not invoked and
no degradation issue is recorded. An unchanged pre-existing floating shape
continues through the acknowledged no-op path.

The old immutable RED remains at
`word-checkpoint-review-shape-repro.py` / `word-checkpoint-review-shape-red.log`
(`920e7365...` / `43dd0558...`). It deliberately models the former ACK grammar,
so the independent current-protocol reproduction is kept separately rather
than rewriting that evidence.

## Generic inline-object finding: ADDRESSED

The guard uses Word's full `inline shapes` collection, not the narrower
`inline pictures` collection. Each row includes ordinal, native IDs, native
text-object range, dimensions, inline-shape type and alternate-text digest;
the mandatory summary independently counts the same collection.

This covers pictures plus generic inline charts, OLE/embedded objects, WordArt,
horizontal lines and other Word inline-shape kinds at the inventory boundary.
As with floating shapes, any added, removed, reordered or changed fact fails
before the rollback mutation batch or controller fallback. The implementation
does not delete these objects; typed refusal is the intentional bounded
behavior. A summary count without the corresponding exact rows is malformed,
quarantines the session, and cannot create a checkpoint.

## ACK, identity, mutation, and lifecycle review

- The parser requires exactly one valid header, one terminal marker, contiguous
  object ordinals, finite geometry, fixed-size lowercase digests, and a summary
  matching the enumerated shape, inline-shape, table, field and bookmark facts.
- Checkpoint construction cross-checks exact tracked bookmark rows against the
  Python tracked index/reference `(bookmark, code)` map. Missing identities
  cannot become a usable preimage.
- Prefix fidelity uses a full `[0, checkpoint)` SHA-256 plus a bounded readable
  guard. Tables, fields, all bookmarks, floating shapes and inline shapes are
  checked separately, so objects that do not contribute visible text cannot
  hide behind an unchanged text hash.
- Rollback re-runs the full state producer in the destructive AppleScript and
  uses Foundation `NSArray.isEqualToArray` / `NSString.isEqualToString` for
  exact comparison before its first delete. The Python ACK then requires the
  exact planned deleted-field/deleted-table rows and the complete saved state.
- Appended fields are selected while live and deleted before tables; adjusted
  live table bounds and surviving fields are checked before table deletion.
  Residual content uses the native-proven `set content ... to ""` primitive.
- Python tracking, observed field topology, pending heading and positional
  invalidation change only after the complete native postcondition ACK.
  Malformed ACK or failed rollback leaves those values untouched and retains
  evidence. Timeout, cancellation, stale document and `-609` connection loss
  use the session quarantine path, which blocks later AppleEvents and fallback.
- A checkpoint may read a read-only document, while rollback runs
  `_mutation_preflight()` before native mutation and therefore rejects it.
  Plain `int` return/coercion and repeated/no-op rollback match the frozen
  direct-method contract; no nonce, LIFO or one-use restriction was added.

## Independent pure verification

The exact frozen focused suite passed:

```text
46 passed in 2.13s
```

It covers the object guards, count/row mismatch, tracked identities, malformed
ACKs, crossing objects, mixed field/table deletion order, retry/no-op,
read-only and stale-binding lifecycle, real recovery-controller routing, script
compilation, and Foundation exact-comparison/hash helpers. Log SHA-256:
`ae4dd833e206e3eaec6dcf09c5de0b9f91c6435edaaab05a2beb69c73fb08cd1`.

An independent five-case current-protocol reproduction also passed:

```text
5 passed in 0.14s
```

It proves that a newly appended floating shape or generic inline shape at an
unchanged story end prevents fallback and prevents submission of any
`rollbackRange` command; unchanged pre-existing instances allow an acknowledged
no-op; and a count without its identity row quarantines. Reproduction SHA-256:
`52c3daf113a85883a063fc78803a0c44a3f99832a281d3f4866922ccb58dbdea`;
log SHA-256:
`dd16cbdbb2cff39a7ff41edae14420833d1d368adba1f7523723a58e363ea28d`.

`py_compile` and `git diff --check` passed for the reviewed files.

## Native evidence boundary

Run 03 remains valid evidence for its older source: its report and retained
source copies agree for all nine recorded sources, all ten named checks are
true, DOCX/PDF hashes match, the partial native REF/table mutation was rolled
back, fallback appeared once, original REF/TOC survived, the unrelated unsaved
sentinel hash stayed unchanged, and final recorded inventories were empty.

Run 03 is explicitly **not current-source acceptance**. Its recovery/session
hashes predate exact Foundation comparison, expanded ACK counts, object guards
and stale-connection quarantine. The final fixture now includes an unchanged
pre-existing inline picture plus floating textbox no-op, followed by an
appended floating textbox rejection and explicit discard. It does not create a
generic non-picture inline object; that generic collection boundary is proven
only by pure protocol tests until a later capability needs native support.

The pending final native run should bind the four frozen hashes above and
retain the existing source-copy, unsaved-sentinel, exact inventory, saved
DOCX/PDF, reopen, and visual inspection gates. A passing run closes current
source native acceptance for this checkpoint slice only.
