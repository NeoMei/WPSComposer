# Word pagination independent review

Verdict: **SCOPED PASS for the two-method production module; one P2 fixture correction required before native acceptance.** No P1 finding. No production, session, tests, native Word state or personal installation changed during review.

## P2: close the owned context before the sentinel and independently verify final inventory

File: `fixtures/microsoft_parity/macos_word_pagination.py`, lines 153–165 (also 176–178).

The fixture closes its synthetic unsaved sentinel via the still-open target session inside the context's `finally`. The owned target context exits only afterward. This repeats the cleanup ordering observed in `macos-word-recovery/run-04`: sentinel close returned an ACK, then the target binding returned `missing value` although its exact private URL remained visibly open, quarantining the run and requiring root's explicit CUA discard. Pagination's fixture then uses `session._closed` as its only post-close proof, with no independent native document inventory after target/reopen close.

Apply the already reviewed/run05-verified fixture sequence: leave the sentinel open while the target context closes; independently inventory native documents to prove target disappearance and unchanged sentinel name/path/saved/hash; close the exact synthetic sentinel with an independent guarded finite-timeout script; verify the original inventory afterward and after reopen close. Preserve failure and quarantine evidence; do not add production rebinding/retry or infer that the prior incident was a production bug. Add pure checks for cleanup ordering and the independent exact-name/token/path/saved-state guard before authorizing this native fixture.

Evidence: recovery run04 raw close failure and `parent-ui-recovery.json`, contrasted with source-bound recovery run05's `owned_context_closed_with_unsaved_sentinel_preserved` and final empty inventory. This review did not repeat either native run.

## Production review results

- Compared directly to frozen `writer.py:3020–3152`. Bookmark lookup remains names/native numeric ordinals, reads the bookmark's first whole paragraph, uses native active-end page plus optional 1×12-point geometry, and does not repaginate.
- Map generation repaginates once, defaults page setup to the same values, clamps only sampled native offsets at Content.End−1 while retaining caller Start/End, handles collapsed ranges and exclusive-end sampling, emits every page in the span, and restricts visual bounds to the frozen operation set.
- Defaulting/coercion, raw-node-ID dedup before redaction, role string retention, nonvisual fragments, negative-margin clamping and unavailable-coordinate omission match the frozen methods. No Python visible text length is used as a native coordinate.
- Ownership validation is restricted to actual `NativeDegradationRange` values. Arbitrary Start/End objects with an incidental `session_id` retain their frozen duck-typed coercion; no stronger provenance claim is made.
- ACK envelopes/row arity and map ordinal/page-span validation exist. Numeric `int`/`float` coercion is deliberately retained. The bookmark method's lack of extra range/page validation is the same frozen contract, not an independently introduced tightening opportunity.
- Expected read-only data-shape/lookup failures become privacy-safe `PAGINATION_SNAPSHOT_FAILED` and do not require quarantining an otherwise usable session. `_execute` remains responsible for actual timeout, cancellation, malformed completion envelope, disconnect and stale-binding quarantine; existing native logs retain diagnostics. `KeyboardInterrupt` propagates. No request to blanket-quarantine snapshot decoding errors is made.
- Neither method modifies selection, content, styles, stored handles or field-topology caches. Native repagination saved-state/layout behavior still requires the pending real Word fixture.

## Consumer boundary

`WindowsLongformExecutor` consumes map output through `PaginationMap.from_dict` after field convergence and bookmark fragments after quality-notice patching. Returned version/node/fragments keys and geometry shapes match these consumers. Its `_dispatch_all` still records tracked ranges through `_native_position` and a COM `_doc.Range` factory. Merely forwarding these two direct methods on Mac does not establish whole-executor Mac range collection; that integration is explicitly outside this reviewed slice and must not be claimed from its passing direct tests.

The two `MacWordSession` forwards were not yet integrated at review capture. The guarded fixture calls the final public methods, so it should run only after parent integrates them and binds source hashes.

## Independent verification

- `PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_pagination.py -q`: **34 passed in 0.17s**.
- Actual frozen `WriterComposer.pagination_map_for_ranges` differential check with fake COM facts versus new module: **48 geometry cases matched**, covering 1/3-page spans, visual/nonvisual operations, negative/out-of-page/NaN coordinates and negative/excessive margins.
- `_map_commands` wrapped in the existing Foundation/Word script context compiled with `/usr/bin/osacompile`: **return code 0**. Compile-only; no native Word operation run.
- Implementer's 173 combined pass statement remains separately reported evidence, not duplicated as this review's own run.
- Guarded fixture records source copies and artifact hashes, checks Unicode/multipage/clamped offsets, repeated geometry and read-only reopen/source bytes, and checks independent PDF text on three pages. Geometry may legitimately omit unavailable native coordinates exactly as frozen behavior permits; no full visual geometry certification exists before native execution.

## Captured hashes

- `skills/WPSComposer/scripts/msoffice/macos_word_pagination.py`: `69f9a34e6a8b52d53f3367c6467b903a0260bba75055f716ccd85ee874c677d9`
- `tests/msoffice/test_macos_word_pagination.py`: `68e46b1f0419f3f335064cb24aeb294323805d88c0005e5b3f6ed8e40bd9ceb8`
- `fixtures/microsoft_parity/macos_word_pagination.py`: `bc96f7db09f43c3f67a28ba11e639446179662c75d5512f1218c8ec48eb0ca61`
- `.superpowers/sdd/2026-09-08-microsoft-wps-parity/word-pagination-report.md`: `6e1bde675e56ae1bac850b14c3c67062ae44d429ba80adaa7d633fba4cf0d023`
- `skills/WPSComposer/scripts/writer.py`: `b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205`
- `skills/WPSComposer/scripts/longform/windows_executor.py`: `9dc03d2b59271c434156b8ee096a652f5ec654ef96a8e2d7b731eb36af457148`

## P2 fix re-review: cleanup ordering ADDRESSED; one new fixture assertion remains

Captured session `f8b2e89a4aefddb27158fefa958da967754e2e52c8e7d02491a05fa35faa4170`, fixture `dc3de3b134d534a2cee2a37a56aa972ecc583b9cf085f523715a46b6f3443af9`, tests `284604b743e1efb7cc635a469ba51af64fc3d18080446f478376945aad1fe43e`.

The original P2 is **ADDRESSED**: target and read-only reopen contexts both exit before independent inventory verifies target disappearance and unchanged sentinel identity/hash. Independent guarded sentinel close follows those successful checks; final inventory must equal starting. Context failure, leftover target despite `_closed=True`, or changed sentinel prevents cleanup. Source dependencies for the reused helpers are retained. Two session forwards exactly preserve the public signatures and lazy routing; no new production P1/P2 finding. Independent pagination tests: **40 passed in 0.17s**.

**New P2 — do not require an empty `posix full name` for an unsaved native sentinel** (`fixtures/microsoft_parity/macos_word_pagination.py:128`). The new `assert sentinel_before[1] == ''` inspects column 1 from the reused `inventory` helper, which contains `posix full name`, not Word's `path` property. Existing real native recovery run03, run04 and final passing run05 returned respectively `["文档213","文档213",false,...]`, `["文档2","文档2",false,...]` and `["文档4","文档4",false,...]` for unsaved sentinels. Therefore the current fixture will fail immediately after sentinel creation on the already observed Word behavior and leave that synthetic sentinel retained, before exercising pagination. Its lifecycle test currently masks this by constructing `['Synthetic sentinel','',False,...]`.

Keep the `saved is False` check and exact baseline tuple/hash preservation; distinguish `posix full name` from `path` rather than requiring this inventory column to be empty. The reused independent `close_sentinel` already reads Word's actual `path` property and verifies it is empty before closing. Add a lifecycle case whose unsaved sentinel `posix full name` equals its display basename, as the retained native evidence shows. This is a fixture-only correction, not a request to loosen production identity checks or introduce native retries. No native run was made during re-review. Full native acceptance remains pending this correction.

## Final scoped re-review: SCOPED PASS

Both P2 fixture findings are **ADDRESSED**; no new P1/P2 findings in the requested correction/forwarding scope. The unsaved sentinel test now uses the observed `name/name/false/hash` shape. The fixture checks `saved is False`, preserves exact complete inventory tuples/hashes across target and reopen context close, and leaves the distinct native `path == ""` verification to the reused independent `close_sentinel` guard. The incorrect empty `posix full name` assertion is removed. Original target-before-sentinel ordering and independent final inventory protections remain intact. Both production forwards remain unchanged and signature-compatible.

Final frozen sources verified:

- Fixture: `ba831a654a7aec6ac3a2e45365f4251cf4a310be21a7fb80d2121b78ee25d237`
- Tests: `fe5a368361367d5e10432c8d6328a81330510bf3f443393d76a20f918719e413`
- Session: `f8b2e89a4aefddb27158fefa958da967754e2e52c8e7d02491a05fa35faa4170`

Independent focused command: `PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_pagination.py -q` — **40 passed in 0.25s**. This final re-review made no production/session/fixture/test changes and ran no native Word operation. Parent may schedule the source-bound native acceptance separately.
