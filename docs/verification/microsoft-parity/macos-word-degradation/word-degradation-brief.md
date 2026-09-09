# Mac Word degradation and quality follow-up preparation

Date: 2026-09-09. Preparation only; no implementation or native-support claim.
The approved parity baseline remains 6dd3a00. This follows the five-method
reference/bibliography slice; it must not widen that worker's current scope.

## Frozen sources inspected

- writer.py: b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205
- macos_word_session.py: b18905ec2165e5e7b0fae2eee5b52707b752501d9050bee396d1c5165ad4c96f
- macos_word_fields.py: 9b84521faa0524583f05bb81402052c168151a7e280db09ce356de8c20d72146
- longform/windows_executor.py: 9dc03d2b59271c434156b8ee096a652f5ec654ef96a8e2d7b731eb36af457148

The session and fields files will evolve in the reference slice. Rebase this
brief's implementation review on its eventual frozen code, preserving semantic
REF handles and the tracked-index growth regression.

## Eight exact public signatures and behavior

| Method after self | Direct behavior |
|---|---|
| degradation_checkpoint() | Return a local document-end integer token. Failure code LOCAL_MUTATION_CHECKPOINT_FAILED. |
| rollback_degradation_checkpoint(checkpoint) | Convert token to int, delete only the appended local mutation, restore insertion point; return None. Failure LOCAL_MUTATION_ROLLBACK_FAILED. |
| add_degradation_notice(code, message, fallback_text, placement="block") | placement="inline" delegates to inline; otherwise inserts one styled single-cell notice box, with baseline range fallback on failed table creation. Returns a semantic table/range handle, not issue records. |
| add_inline_degradation(code, message, fallback_text) | Insert and style display text at the current semantic insertion point without a new paragraph; return semantic range handle. |
| reserve_document_quality_anchor(title="生成质量提示", notices=()) | Create the fixed wpsc_document_quality_anchor bookmark only once, even if empty, remember title and dedup set, insert given notices; return None. |
| upsert_document_quality_notice(issue) | Require reserved anchor; append one previously unseen issue at the reserved location; return None. |
| add_document_quality_notice(notices) | Backward-compatible reserve-if-missing then upsert each; return None. |
| add_quality_notice_at_bookmark(*, code, message, fallback, node_id, page, bookmark_name=None) | Explicit bookmark inserts after its paragraph; default uses fixed anchor start. Format page as int, return inserted semantic table/range handle. node_id is inert metadata. |

Native COM access is outside semantic parity, but returned range/table identity
must support the consumer operations required by the frozen API. Do not claim
that a plain string meets an uninspected handle consumer. Audit consumers before
choosing a Mac semantic handle; raw AppleEvent references cannot outlive a batch.

## Required invariants and dependencies

- The executor checks both checkpoint and rollback callability before a node;
  missing methods abort at local-checkpoint. The reference slice cannot claim
  the actual controller recovery chain until these methods exist and run.
- A Mac document's mandatory final CR means the COM document-end token cannot
  be copied blindly to an AppleScript deletion bound. Native probes must prove
  exact before/after text, final CR, and prefix preservation for a partially
  appended reference, paragraph and table. Include empty doc, nonempty terminal
  paragraph, CJK/emoji and an operation that failed before making any change.
- Checkpoint design needs an explicit compatibility ruling before production.
  The baseline returns a document-end int and rollback applies int coercion;
  existing consumers pass it opaquely. A session-issued token map may protect
  the native document, but two equal plain integers cannot prove their source.
  Do not claim a map alone distinguishes cross-session equal values. One-use
  consumption, LIFO restrictions or nonce-encoded ints are candidate design
  choices, not existing baseline behavior. Probe valid zero-mutation/retry
  semantics and reject out-of-document deletion bounds without inventing a
  direct-call limitation. Keep the checkpoint token separate from native range
  bounds if a reviewed session-token design is eventually selected.
- Rollback must invalidate changed positional targets and remove only tracked
  fields created in the reverted extent. Record field identities before deletion
  and remove reverted tracking only after rollback acknowledgement; do not read
  deleted native objects to discover their former identity. Preserve preexisting
  REF/index semantic identities. It must never erase a pre-checkpoint bookmark or unsaved sentinel.
- A timeout or uncertain native completion is quarantined and cannot be repaired
  by blindly issuing more commands. Preserve evidence and the exclusive lock;
  controller fallback only follows acknowledged local rollback.
- Shared _degradation_display validates public issue code, redacts private text,
  avoids double wrapping, and uses distinct inline/block wrappers. Reuse this
  contract; message is not separately interpolated by the inline/block helpers.
- Styling is italic, dark red #9C0006, light red #FCE8E6; use UTF-16 native ranges.
  A block box has 0/3 spacing, keep-together, body outline, and no row-page split.
  If native table construction partially fails, rollback must be acknowledged
  before the plain-range fallback; never layer fallback on uncertain remnants.
- Only quality upsert notices deduplicate by (normalized code, placement,
  redacted node id). Upsert normalizes an invalid code to QUALITY_NOTICE;
  bookmark notices and degradation helpers use DEGRADATION through the shared
  display function. Do not unify those distinct defaults. The bookmark helper
  does not deduplicate repeated calls.
  Mapping and object-form issues are consumed in WriterComposer. Title appears
  only before the first actually inserted notice. An empty reservation produces
  no visible quality block. The first title and first notice share one box.
  Update dedup/anchor position only after native ack.
- Preserve arbitrary safely bounded direct literals and scalar coercions from
  WriterComposer. Generation-plan keys, fixed labels, or geometry restrictions
  must not silently narrow direct API calls. Preserve each consumed scalar contract: placement equals exactly
  "inline" selects inline; all other direct placement values select block.
  Quality placement is a dedup value and is not automatically converted to str.
  notices accepts the baseline mapping/object entries and falsy iterable forms;
  page/checkpoint use int coercion. Native-unsafe or unrepresentable values must
  be identified before mutation, but that does not authorize a single new strict
  schema over all direct calls. Unrelated metadata remains inert where ignored.
- Repeated reserve must keep the first anchor/title; notices still upsert.
  Explicit bookmark-not-found, missing reserved anchor and malformed ack produce
  typed failures with retained diagnostic evidence, not guessed insertion sites.
- Bind every AppleEvent to the exact private document and the single deadline.
  No UI selection, clipboard, global template/settings changes or process quit.

## Implementation and proof order

1. Probe checkpoint/rollback geometry on task-owned Word documents without
   enabling production support. Validate ownership and zero-mutation errors.
2. RED/GREEN two checkpoint methods plus pure executor recovery tests, then real
   reference controller failure -> exact rollback -> static fallback once ->
   issue once. Keep normal and fault reports separately hash-bound.
3. RED/GREEN inline and block helpers with shared display/UTF-16 tests and native
   formatting/readback/return-handle proof.
4. RED/GREEN quality reservation/upsert/bookmark family, native idempotence,
   title order, anchor stability after other edits and distinct issue identities.
5. Independent family review, affected field/index/source-retention regression,
   full suite, public save/PDF/close/reopen, real editability and nonempty sentinel
   preservation. Only then reconcile method inventory and required evidence rows.

The existing generation compiler's reserve_document_quality_anchor only appends
its planned notices and does not establish direct fixed-anchor/upsert behavior.
Its generic degradation branches add a paragraph, which is incompatible with the
direct inline method. Do not copy those branches as proof of direct parity.

## Independent preparation review

word_objects_probe checked the eight signatures and consumer routes. Feedback
on tracking cleanup, distinct code defaults, dedup scope and scalar coercions is
incorporated. Token ownership/consumption remains a design decision, not a new
implemented restriction. Production and native acceptance remain absent.
