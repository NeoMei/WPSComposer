# Word pagination implementation report

## Brief (before production)

Scope is exactly `pagination_fragment_for_bookmark(node_id, bookmark_name)` and
`pagination_map_for_ranges(tracked_ranges)`. Frozen reference: writer.py 3020–3152.
Bookmark inspection reads the bookmark's first whole paragraph and returns its
native active-end page and optional 1 × 12 point bounds. It does not repaginate.
The map repaginates once, reads document page setup, clamps sampling points to
Content.End - 1, preserves original native offsets, emits all spanned pages and
visual-operation bounds only. Defaults, float/int coercion, raw-ID deduplication,
role conversion, redaction and invalid-coordinate geometry omission remain the
reference semantics. Errors are privacy-safe `PAGINATION_SNAPSHOT_FAILED`.

Implementation is a focused module that uses session `_execute` and its existing
exact document binding/deadline/quarantine behavior. There are no selection,
content, style, saved-state or topology writes. Read-only sessions are supported.
Range handles expose Start/End, including existing degradation semantic handles;
known tagged handles must belong to the current session. Windows executor's
internal COM range_factory and Windows transport handle decoding remain untouched.
No new range-string or dict grammar is introduced.

Tests will exercise actual decoding/geometry outputs, per-page boundaries,
Unicode/privacy, duplicate/default behavior, invalid native acknowledgements and
transport/lifecycle gates. Native fixture is guarded and requires root's Word
lease, exact source snapshots, unrelated unsaved sentinel, public save/PDF/reopen,
and read-only/source-preservation checks. No native execution by this worker.

## Session forwarding patch (parent-owned file)

```python
    def pagination_fragment_for_bookmark(self, node_id, bookmark_name):
        from .macos_word_pagination import pagination_fragment_for_bookmark
        return pagination_fragment_for_bookmark(self, node_id, bookmark_name)

    def pagination_map_for_ranges(self, tracked_ranges):
        from .macos_word_pagination import pagination_map_for_ranges
        return pagination_map_for_ranges(self, tracked_ranges)
```

## Validation

Pending RED/GREEN and independent parent review. Native/UI acceptance not run.

## Final implementation and unit evidence

- Production module and guarded native fixture are present. Session forwarding
  remains parent-owned and has not been edited by this worker.
- RED observed: initial 29 tests failed because pagination implementation was
  absent. GREEN: all 29 passed. Native bookmark integer-ordinal compatibility
  then had a separate observed RED (1 fail / 30 pass), followed by GREEN.
- Latest focused verification: **173 passed** across pagination (34), degradation
  and field tests, using Python 3.9 and
  `PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps` with
  `../wps-task-session-startup/.venv/bin/python -m pytest
  tests/msoffice/test_macos_word_pagination.py
  tests/msoffice/test_macos_word_degradation.py
  tests/msoffice/test_macos_word_fields.py -q`.
- Native fixture and module compile with Python; fixture `--help` succeeds and
  importing it launches no Office process. **No native/UI run, personal install,
  commit or push was performed by this worker.**
- Root ruling incorporated: ownership checks apply only to the project's defined
  `NativeDegradationRange` identity (`_degradation_session_id`). Arbitrary objects
  exposing Start/End keep their coercion behavior even if they also happen to
  expose a `session_id`. Their provenance cannot be proven by this method.
- Bookmark lookup supports literal names and native numeric ordinals, preserving
  the frozen collection lookup rather than introducing a name-only contract.
- `Word.sdef` inspected locally: active end page number (3), horizontal position
  relative to page (5), vertical position relative to page (6), repaginate and
  page setup properties exist. Dictionary inspection is not native verification.
- The output intentionally preserves frozen role strings; redaction applies to
  nodeId, while document body/handle Text are never included in commands/results.

## Parent native execution

After independent review and the two-method forwarding integration, under the
parent's sole Word lease:

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_pagination.py --execute --output-dir NEW_PATH
```

The new output directory records exact source copies/hashes, native runtime logs,
DOCX/PDF artifacts/hashes, unrelated-document inventory and synthetic unsaved
sentinel protection, 3-page spans, UTF-16 offsets, collapsed/past-end clamping,
original-ID dedup and redaction, first-paragraph bookmark geometry, missing native
bookmark privacy-safe failure, body/saved-state stability, public save/PDF, native
read-only reopen comparison and unchanged source bytes. Independent PDF text is
checked on each of the three pages. Coordinate availability remains a native gate;
negative native coordinates omit bounds exactly as the frozen implementation.
Independent review and native result status remain pending.

## Independent review correction and forwarding integration

Independent review reported SCOPED PASS for production, with one P2 native
fixture cleanup correction. Corrected the fixture to keep the synthetic sentinel
open throughout both owned and read-only reopen contexts. After each successful
context exit, the independent recovery-fixture inventory must prove the sentinel's
exact original name/path/saved/hash and that every remaining other document is
exactly the starting inventory. Only then does the independent `close_sentinel`
helper close the exact name/token/empty-path/unsaved document, followed by another
independent inventory equal to starting. A context close failure, a target that
remains despite `_closed=True`, or a changed sentinel retains evidence and does
not attempt sentinel cleanup. No production rebinding/retry was introduced.

Both inventory and close helpers are reused from
`fixtures/microsoft_parity/macos_word_recovery.py`; that source and its native
hash helper source are now part of retained source hashes. All inventories use
that same helper, avoiding the older sections-fixture SHA output's `  -` suffix
versus the recovery helper's canonical 64-character hash.

Four behavior-level fixture tests observed RED (34 passed / 4 failed), then
GREEN. They execute the actual fixture orchestration against a native-boundary
lifecycle double, covering the successful order, native target remaining after
`_closed=True`, sentinel hash changes, and context close failure. Existing recovery
fixture tests verify the reused independent exact name/token/path/saved guard and
finite timeout boundary. The differing legacy hash format was separately exposed
with a RED test transport observation before the inventory unification.

Root subsequently granted the two-forward edit lease. Added only those two lazy
forward methods to `MacWordSession`; no notices method or forward was modified.
Two tests call the public session methods in read-only mode and check semantic
results: observed RED (38 passed / 2 failed), then GREEN.

Latest combined verification: **230 passed**, covering pagination (**40**),
recovery, degradation and fields. No Office/native execution occurred during this
correction. The session file is frozen and its edit lease released to root.

Frozen hashes at handoff:

- session.py: `f8b2e89a4aefddb27158fefa958da967754e2e52c8e7d02491a05fa35faa4170`
- pagination fixture: `dc3de3b134d534a2cee2a37a56aa972ecc583b9cf085f523715a46b6f3443af9`
- pagination tests: `284604b743e1efb7cc635a469ba51af64fc3d18080446f478376945aad1fe43e`

P2 re-review and native acceptance remain parent gates. Earlier sections above
are the chronological pre-review/pre-integration evidence, not current pending
implementation work.

## Re-review native unsaved-name correction

The re-review accepted cleanup and forwarding, but identified a fixture-only P2:
Word's unsaved document `posix full name` can equal its document name. It is not
the distinct `path` property. Removed only the empty-string assertion on that
inventory column; saved=False, exact full inventory tuple/hash equality, and the
starting-document difference checks remain. The reused native `close_sentinel`
continues to verify actual `path` is empty before closing.

Changed the fixture lifecycle double to the observed `name/name/false/hash` shape.
Observed RED: **1 failed / 39 passed**. After the one-line fixture fix: **40 passed**.
No production module, session, helper or native state was changed in this repair.

Frozen correction hashes:

- pagination fixture: `ba831a654a7aec6ac3a2e45365f4251cf4a310be21a7fb80d2121b78ee25d237`
- pagination tests: `fe5a368361367d5e10432c8d6328a81330510bf3f443393d76a20f918719e413`
- session.py remains `f8b2e89a4aefddb27158fefa958da967754e2e52c8e7d02491a05fa35faa4170`

Ready for root re-review; native run remains unexecuted by this worker.
