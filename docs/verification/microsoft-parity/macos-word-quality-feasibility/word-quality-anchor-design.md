# Word direct quality-anchor family — bounded design brief

2026-09-09, Task 3 of the approved Microsoft/WPS parity plan. Read-only design work; this brief is the sole authored file. No production/test/fixture edits, native/UI execution, commit, or capability enablement occurred. The parent owns the native lease and the concurrently changing rollback acceptance. Existing approved plan/SDD authorization applies.

## Decision

Add one focused `msoffice/macos_word_quality.py` helper and four thin `MacWordSession` forwards, preserving the frozen signatures and behavior below. Reuse display/style/immutable range DTOs and the transport submission boundary. Introduce a **bounded arbitrary-position insertion transaction**, with its own geometry and rollback proof; do not send middle insertion through the append-only recovery API. This is a required native feasibility dependency, not an already established primitive. Table construction plus failed-construction range fallback must work at a real middle anchor before this family is complete.

Do not modify the shared generation compiler, WPS/COM implementation, generic editing subsystem, or public signatures as part of this slice. A future need for a focused rollback helper file is preferable to growing `macos_word_session.py`; do not build a generic transaction framework. Read `docs/regression-guardrails.md` before implementation.

## Source contract and evidence boundary

The frozen source was read with `git show 6dd3a00:skills/WPSComposer/scripts/writer.py`; its four methods match the live inspected implementations at `writer.py:2900–3019` (frozen positions begin slightly earlier). Supporting display/box behavior is at live lines 2807–2877. The implementation plan, progress ledger, `word-degradation-brief.md`, longform consumer and existing portable tests were also inspected.

Read-time SHA-256 values (not a freeze of the parent's later work):

| Source | SHA-256 |
| --- | --- |
| `writer.py` | `b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205` |
| `msoffice/macos_word_degradation.py` | `8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7` |
| `msoffice/macos_word_recovery.py` | `20721fb271df8aa7baed94036251cef1787ec96ae1372439cb95239602b7b756` |
| `msoffice/macos_word_session.py` | `3ab747a06c63deabc9e0dbf1de51857d8b7dfa843d6ea8e0822915e91e44f205` |

Existing native evidence, read from `.superpowers/sdd/2026-09-08-microsoft-wps-parity/word-degradation-final-native-report.md`, is run-03 / Word 16.112.3: six append cases, 57 checks, two real partial-table failure → acknowledged append rollback → range fallback cases. Its scope explicitly excludes these four methods. It proves reusable primitive behavior at the terminal body range only. This design does not independently rerun or re-certify that evidence.

## Frozen signatures, state, errors, and returns

### `reserve_document_quality_anchor(self, title="生成质量提示", notices=())`

1. If absent, create `_quality_notice_seen = set()` before any title/native work.
2. Only if `_quality_notice_anchor_position` is absent, evaluate `redact_private_text(str(title or "生成质量提示"))`. That evaluation is outside the native `try`; exceptions from user coercion propagate. Repeated reservation must not consume/coerce a new title.
3. Read `int(self.selection.End)`, create a collapsed native range, and require `Bookmarks.Add("wpsc_document_quality_anchor", range)`. Commit numeric `_quality_notice_anchor_position` and first safe `_quality_notice_title` only after successful native reservation. Native failures become `NativeWriterObjectError("DEGRADATION_INSERT_FAILED", "quality anchor insertion failed")`, without chained diagnostics. A failed reservation may leave the initialized empty seen set.
4. Iterate `notices or ()` serially through the shared upsert routine. An empty reservation emits no title, paragraph, or table. Repeated reservation retains first anchor/title but still consumes new notices.
5. Return `None`. The iterable is not a new strict list schema: falsy inputs are empty; errors obtaining/advancing an iterable propagate; an earlier successful notice is not undone when a later item fails.

The frozen selection is `self._app.Selection`. The Mac equivalent must read the selection belonging to **boundWindow**, verify the exact bound document and story/range, and use its End. Substituting document end changes this direct contract. Inactive-bound-window selection access requires a native probe; global `selection`, clipboard, or the unrelated active document cannot supply it. If bounded activation is necessary, activate only the verified `boundWindow`, then reverify the selection's ownership. Existing append APIs do not prove this behavior.

### Shared upsert mapping behavior

For a `dict`, read `code`, `placement` (default `"document"`), `nodeId`; otherwise use attributes `code`, `placement` (default `"document"`), `node_id`. It is an actual `dict` check, not all Mapping implementations.

Normalize code against `^[A-Z][A-Z0-9_]{0,63}$`; invalid/non-string code becomes `QUALITY_NOTICE`. Identity is exactly `(code, placement, redact_private_text(str(node_id or "")))`. Do not coerce placement, include message in identity, or reinterpret placement as layout. Explicit `None` placement differs from omitted placement. Unhashable placement or failing getter/coercion fails before the native try; do not silently normalize it. Duplicate lookup returns immediately, before obtaining fallback message or making a native call.

For a dictionary, display source is `notice.get("fallbackText") or notice.get("message", "")`. For an object, only `getattr(notice, "message", "")` is read; an object's `fallbackText`/`fallback` is ignored. These reads occur outside the native try.

Inside the native try: coerce saved numeric position using `int`, make a collapsed range at that position, call the frozen `_degradation_display(code, fallback_text)`, and prepend `first_safe_title + "\r"` only when seen is empty. Insert one block box. Update position to the returned `Range.End` (the frozen fallback expression consults Selection.End if needed); then add identity to seen **after** the insertion block succeeds. Mac's acknowledged DTO always supplies End, so no unrelated selection fallback is necessary.

An insertion/display/position failure is `DEGRADATION_INSERT_FAILED`, message `quality notice upsert failed`. Do not commit dedup/cursor on failure, malformed ACK, or uncertain completion. An `upsert` is append-once, not replacement of a prior message. Each new identity inserts a new one-cell box; the title and the first notice share the first box.

### `upsert_document_quality_notice(self, issue)`

Require the in-session `_quality_notice_anchor_position` before consuming issue properties; absence yields `DEGRADATION_INSERT_FAILED`, message `quality anchor is unavailable`. Otherwise run the shared mapping behavior. Return `None`. Finding a persisted bookmark on a reopened document does **not** automatically reconstruct the in-memory dedup set, title, or cursor.

### `add_document_quality_notice(self, notices)`

If no in-session numeric anchor exists, call `reserve_document_quality_anchor(notices=notices)` with the default title, then return `None`. Otherwise iterate `notices or ()` through shared upsert and return `None`. Preserve serial effects and errors; do not pre-validate/materialize the entire iterable before reservation.

### `add_quality_notice_at_bookmark(self, *, code, message, fallback, node_id, page, bookmark_name=None)`

All six named inputs except bookmark_name are required and keyword-only. `node_id` is unused; do not inspect, stringify, or derive a bookmark from it. Every call inserts anew; no document-upsert dedup or cursor advance occurs.

The whole method body is in one `try`. Resolve `name = bookmark_name or "wpsc_document_quality_anchor"`; check existence when available, then native collection lookup (COM callable, then Item fallback). For truthy explicit bookmark_name, use the **End of the first paragraph of bookmark.Range**; for a falsy/default name use **bookmark.Range.Start**, not the saved upsert cursor. Missing bookmark must fail; never append to the end as a substitute.

After resolving the native target, form `redact_private_text(str(message)) + " (page " + str(int(page)) + "; " + redact_private_text(str(fallback)) + ")"`, then `_degradation_display(code, text)`. `int(page)` semantics include numeric strings, booleans, truncatable floats, and negative values; do not replace this with a positive-integer schema. `str(None)` is `"None"` here. Display invalid-code default is `DEGRADATION`, different from document upsert's `QUALITY_NOTICE`. Exceptions from lookup, coercion, display, or insertion become `DEGRADATION_INSERT_FAILED`, message `quality notice insertion failed`.

Return the inserted table or fallback `.Range` wrapper. Mac should reuse `NativeDegradationBox`/`NativeDegradationRange`, with UTF-16 Start/End/Text and session id, never a string or live COM proxy. `table_index` must be the actual resulting ordinal; adding a middle table does not make it the last table. `longform/windows_executor.py:357` ignores this return; pagination code consumes `.Start/.End` of extracted ranges elsewhere. Snapshot handles are not persistent edit targets.

Named bookmark strings and native collection ordinals must retain the native baseline's accepted semantics; do not turn an ordinal into a quoted name or silently stringify arbitrary objects. Falsy bookmark_name still selects the default branch. Truthy numeric and boundary cases need targeted compatibility/native evidence rather than a guessed stronger schema. The frozen code does not impose its own case-sensitive bookmark-name lookup: preserve the native lookup rules, then compare the resolved canonical identity and subsequent preimage facts exactly. A Foundation identity guard does not authorize a new input-name restriction.

## Two different anchor mechanisms must remain distinct

The frozen implementation creates one public fixed bookmark but does not move/recreate it on every upsert. Upsert subsequently uses the saved numeric cursor and advances it to each inserted box's End. Default bookmark insertion independently reads the persisted bookmark Start. Repeated reserve does not rebind either. Do not replace both with a newly introduced moving bookmark or automatically rebase the numeric cursor after arbitrary external edits; that is not frozen behavior. Verify the native public bookmark's expansion/affinity on insertion and save/reopen. Fail stale/unverifiable targets safely; do not claim arbitrary-edit cursor stability already exists in the baseline.

`_insert_degradation_box` tries a native 1×1 table, applying italic, #9C0006 foreground, #FCE8E6 shading, 0/3 paragraph spacing, keep-together, body outline, and row non-splitting. Ordinary failure tries a styled plain range at the same supplied target without adding a paragraph. Fallback End uses UTF-16 units. The Mac safety requirement is acknowledged restoration before fallback if table construction already changed the document; reproducing the COM helper's unguarded partial-failure behavior would weaken the approved plan.

## Reuse versus new work

| Existing code | Reuse allowed | Limit requiring new work |
| --- | --- | --- |
| `_degradation_display`, privacy redaction, `_units`, literal encoding | Same code normalization/wrapping/redaction and UTF-16 display calculation | Combined title+display and bookmark message need the same native-safe literal checks; preserve their method error boundaries |
| `_style_commands`, `_style_ack`, `_style_valid` | Same style commands and strict style ACK values | Native verification that a middle table/fallback does not leak formatting into surrounding content |
| `NativeDegradationBox` / `NativeDegradationRange` | Immutable session-scoped return facts | ACK must report actual middle-table ordinal and geometry |
| `range_commands`, `table_commands` | Factor the common insertion/style/readback body with append default unchanged | Both call `_position('end')`; `position` asserts end rather than selecting it. Passing a middle integer cannot work |
| `_range_ack`, `_table_ack` | Common row types, exact display, style, cell/row checks | Both assert `beforeEnd == position + 1`; table ACK assumes total table count is returned object's ordinal. Add separate arbitrary-position ACK; never loosen append ACK globally |
| `macos_word_recovery` | Hash commands, privacy-safe snapshot concepts, Foundation exact array comparison, quarantine rules | Existing preflight expects unchanged prefix and old object coordinates; rollback deletes target→terminal range. Inserting before existing objects shifts those coordinates and leaves a suffix that must be preserved |
| `_execute_topology_mutation` / `_execute` | Preflight, shared deadline, staging/log retention and actual-submission invalidation | Use this boundary for new content mutations; degradation `_execute` currently invalidates only on successful `_committed_range`, which is insufficient for this new slice |

The inspected `_binding()` uses AppleScript `is not` for path/name equality. The new helper must add Foundation `NSString.isEqualToString` comparison of the exact expected full path before its first native read/mutation and again after any required bound-window activation. Do not treat basename lookup as identity. For unsaved attached documents, verify window id, exact name, and empty path. If the core binding is repaired concurrently, reuse its reviewed exact guard instead of maintaining two divergent policies. All target names, document paths, text, bookmark identity and structured preimage comparisons use Foundation exact comparison; AppleScript's default case-insensitive equality is not sufficient.

## Minimal new transaction contract

1. Resolve the operation-specific position and exact bound document in the same native batch that validates the target. Capture before end, prefix and suffix text hashes, the boundary paragraphs' format facts, public/other bookmarks, tables, fields with codes/results/coordinates, and the existing tracked semantic identities. Capture section/list and shape facts sufficient to reject unrelated changes. The current append snapshot is a starting point, not proof of formatting or suffix preservation. Return hashes/facts, not the entire private body.
2. Insert a collapsed range or one table at that verified position, retaining the actual created native table reference **inside the batch**. The new success ACK must prove before/after document delta, insertion range/table/cell UTF-16 geometry, exact display, actual table ordinal, required styles, and unchanged prefix/suffix after translating the suffix by the measured insertion delta. Middle-table terminators, automatically inserted paragraph marks, adjacent-table merging, bookmark affinity and terminal paragraph handling require native measurement; do not hardcode append geometry for them.
3. Execute content-changing batches through `_execute_topology_mutation`. Local coercion, read-only/deadline or script-I/O rejection before submission leaves observed field topology intact. Every actual submission invalidates it even on an ordinary native error or bad ACK. Commit `_structural_changed`, pending-heading cleanup, quality cursor/seen and return DTO only after a valid success ACK. Empty reservation and duplicates must not be marked as text mutations merely to simplify bookkeeping.
4. For an acknowledged ordinary table failure, first prove which insertion objects belong to this transaction. The smallest candidate rollback is deletion of the same batch's newly created table, plus only native-proven separator residue owned by that insertion, followed by **exact complete preimage readback**. A failure before table creation can authorize fallback only if an unchanged preimage is actually acknowledged. Do not infer it from an exception class or the absence of a returned table reference.
5. New middle recovery must preserve both sides, existing table/field/bookmark objects and boundary formatting. Deleting only the table may leave paragraph/style/bookmark changes; native evidence must decide whether this candidate suffices. If restoration requires unverified rich-range restore/Undo capabilities, leave recovery unimplemented and run a separate bounded feasibility probe; do not silently introduce them into production. No generic Undo, disk replacement of an attached document, or suffix deletion is authorized by this design.
6. Only an exact `restored-preimage` ACK permits the styled range fallback at the original position, with its own success ACK. Preserve the observed field-topology cache only if restoration proves the original topology exactly, following the existing acknowledged recovery pattern. Otherwise leave it invalidated.
7. Timeouts, invalid connection/cancellation (-1712/-609/-128), malformed/missing completion, a destructive ordinary recovery error, or failed postcondition retain evidence and quarantine. No further AppleEvents, fallback, save/publication, or “best effort” cleanup through the quarantined session. No dedup/cursor commit. Public method wraps ordinary exceptions using its frozen `DEGRADATION_INSERT_FAILED` message while retaining the internal error/evidence; propagate BaseException cancellation after retention.

This transaction is not atomic across an iterable of notices: already acknowledged earlier notices stay committed. It is local to a single notice. If the full middle recovery contract is not yet proved, the capability row remains incomplete even if empty/end success cases pass.

## Meaningful RED suite before implementation

Create a focused new test module; use real helper compiler/parsers and explicit transport doubles, not a mock replacing the whole operation.

1. Frozen signatures and None/DTO returns; selection End differs from document end and unrelated active selection. Empty/repeated reservation is invisible and keeps first title; changed title object must remain unconsumed. Native reservation failure leaves cursor/title uncommitted.
2. Dict/object field rules, default/invalid code differences, message versus fallbackText, private-text redaction, CJK/emoji UTF-16 lengths, duplicate with changed message, distinct placement/node identity, unhashable placement error timing, ignored node_id/object fallback attributes. Duplicate must not touch native transport or message getters.
3. Serial iterable fails on its second item: reservation and first insertion remain; second identity/cursor does not commit. Missing anchor wins over an issue object whose property access raises.
4. Persisted default bookmark Start differs from numeric upsert cursor; explicit bookmark first-paragraph End differs from bookmark End and document end. Repeated bookmark calls insert twice. Missing bookmark, case-variant native lookup, invalid page and numeric-name/ordinal cases retain coercion/typed-error semantics without a guessed target.
5. Middle ACK includes suffix and preexisting tables/fields/bookmarks; forged document delta, shifted wrong object, false prefix/suffix preservation, extra rows, boolean-as-integer, wrong actual ordinal or style must fail. Preserve the old append ACK tests unchanged.
6. Native table failure may submit fallback only after real recovery-helper compilation and a valid restoration ACK. A suffix mismatch, case-only text/field-code/path mismatch, crossing object, altered boundary format, or destructive rollback error prevents fallback and quarantines after mutation. Include a failure that inserts no table but alters text.
7. Stub subprocess at the actual submission boundary: timeout, ordinary nonzero exit, invalid JSON and malformed ACK invalidate prior field observations; local validation/deadline/script-write failure before launch does not. Dedup/cursor changes only on ACK. Do not use a fake `_execute` whose behavior bypasses this boundary as the sole proof.
8. Known synthetic document/window mismatch must be rejected before any destructive command. Generated scripts must use exact Foundation guards; source-string assertions supplement, not replace, behavior assertions.

## Required serialized native evidence

Begin with a throwaway owned-document feasibility fixture, separate from production enablement. The parent coordinates Word/UI; this worker ran none.

- Reserve at a nonterminal **bound selection End**, with another nonempty unsaved sentinel active. Prove exact target document/window, zero visible change, correct empty bookmark, and unchanged sentinel. Save/reopen and observe the actual persisted bookmark.
- Put content, a semantic bookmark, a field, a table and formatting on both sides of a middle quality anchor. Reserve empty, append later content, insert first/second document issues and duplicates; prove title once, distinct identities once, cursor advanced from each actual Range.End, public bookmark behavior and preserved suffix objects. Include CJK/emoji and a nonempty terminal paragraph.
- Patch after an explicit bookmark's first paragraph where its range spans multiple paragraphs; separately patch at the default bookmark Start. Repeated calls must create repeated notices. Include middle and terminal paragraph boundaries, existing neighboring tables, named bookmarks differing only by case, and native ordinal lookup cases supported by frozen behavior.
- Inject actual failure after creating/populating the middle native table, then prove exact local preimage restoration before one range fallback. Also inject failure before construction and during destructive cleanup. Retain raw failure scripts/logs and separate corrected reruns. A terminal-only rollback test cannot close this gate.
- Reopen native DOCX and inspect text order/count, table/cell/style/row flags, bookmarks, fields/results and unaffected native object identities; inspect PDF output for notice layout and unchanged surrounding formatting. Verify M5 title/outline/numbering/body indentation and final field convergence remain intact. Saved OOXML/PDF evidence is inspection, not a replacement generation backend.
- Parent performs actual notice edit → Undo → explicit Save/discard → close → exact-path reopen, with nonempty sentinel preservation and final owned-document inventory. Record source hashes before/after, artifact hashes, native version, submitted failure order and cleanup separately.

Only after meaningful RED/GREEN, source-bound middle success/failure/native evidence, independent review and affected full-suite checks may the four method rows be reconciled. This brief does not claim they are implemented, verified, installed, or complete.
