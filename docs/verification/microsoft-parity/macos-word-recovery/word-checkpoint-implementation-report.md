# Word checkpoint implementation report

Status: DONE for the reviewed reference/paragraph/table append-recovery slice with safe refusal for changed/new shapes. Independent scoped review and final source-bound native run05 PASS. No commit or push. Worktree `.worktrees/office-description`, branch `codex/microsoft-parity`.

## Implemented scope

- Two explicit `MacWordSession` forwards, plain integer terminal native insertion coordinates; `int` coercion, retained preimages, same-coordinate ambiguity detection, repeated/no-op rollback, readonly checkpoint inventory and writable rollback.
- New focused recovery module inventories exact prefix hash and bounded 512 UTF-16-unit guard, terminal paragraph geometry/hash, tables, main-story fields, all bookmark identities/bounds/hash and exact tracked code-bookmark text. Snapshot tracking prefixes retain original frozen handles and exact code strings. Read-only observation-cache changes do not manufacture ambiguous same-coordinate tokens.
- SHA256 runs through Foundation NSTask/NSPipe, fixed `/usr/bin/shasum -a 256`, NSData UTF8 streamed after launch. Private full prefix never appears in shell, argv or persisted source. Terminal paragraph uses hash rather than full text to avoid persisting an entire one-paragraph document; this refinement was explicitly accepted by root.
- Transaction rechecks native preflight exactly using Foundation NSArray/NSString equality (Word dictionary shadows AppleScript `case`, so native `considering case` would not compile). Appended fields are deleted in descending native position before tables. Native field deletion deltas update live table bounds; surviving fields and adjusted table topology are re-enumerated/verified before table deletion. Residual range content is set to empty. Full postcondition ACK must match original immutable state before Python tracking commit.
- New/changed floating shapes or all inline shapes reject before any rollback deletion/fallback. Native identity/geometry inventory includes ordinal, native anchorID/editID, anchor/range, dimensions, shape name/type/z-order/visibility, text or alternate-text digest. This is a safe refusal guard; it is NOT figure/equation/object deletion rollback support.
- Required ACK summary counts shapes, all inline shapes, tables, fields and bookmarks; row count omissions and missing tracked identities quarantine. Failed or malformed rollback ACK leaves tracking/cache unchanged.
- Narrow session integration repair: exact stale-document error and explicit AppleEvent -609 connection-invalid error quarantine before cleanup/later AppleEvents. Existing field-identity ordinary-error distinction is preserved.

## Ownership

Only the new recovery module, focused tests/fixture/report/evidence and session two forwards/initializer/error mapping are this slice. Session topology invalidation changes and fields.py remain the other worker's changes; they were frozen before integration. No notice methods, Windows methods, Word global preferences, application lifecycle or user documents were changed.

## TDD and focused validation

Command: `PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_recovery.py -q`

- Initial RED: 17 failures, expected missing recovery module (`/tmp/word-recovery-red.txt`).
- Public forwards RED: 2 failures, expected missing methods (`/tmp/word-recovery-forwards-red.txt`).
- Readonly RED: expected readonly checkpoint gate rejection (`/tmp/word-recovery-readonly-red.txt`).
- Exact native comparison RED: expected missing exact comparison (`/tmp/word-recovery-case-red.txt`).
- Shape/disconnect RED: 3 expected failures, unimplemented inventory and nonquarantined -609 (`/tmp/word-recovery-shapes-red.txt`).
- Tracked-identity omission RED: expected acceptance of incomplete ACK (`/tmp/word-recovery-counts-red.txt`).
- Current GREEN: **46 passed in 2.03s**. This includes native Foundation synthetic SHA for empty/CJK/emoji/quotes/newline/700KB text (exceeding ARG_MAX), exact equality synthetic mismatches, compile-only Word scripts, real WindowsLongformExecutor recovery-controller calls on a real MacWordSession with recording transport, no-fallback failure, original handle identity, crossing objects, all-bookmark ordering and fixture AST signature binding checks.
- Full suite intentionally deferred to root's final reviewed candidate run.

## Native evidence

- `docs/verification/microsoft-parity/macos-word-recovery/run-01`: retained fixture signature failure (`add_heading(level=...)`), after successful empty/no-op native checkpoint; ordinary failure, exact owned/sentinel cleanup.
- `run-02`: retained fixture-only missing `add_heading_level_native` failure, ordinary failure and cleanup. Fixed to verified `add_heading_level('Chapter',1)` and added AST signature-binding preflight before further native work.
- `run-03/report.json`: **PASS** on prior source version. Actual WindowsLongformExecutor controller: partial tracked REF, untracked REF in new table, untracked REF outside table -> real checkpoint rollback -> static fallback once -> issue once. Preexisting REF/index handle instances preserved; snapshot drift guard after rollback/fallback valid. Native invalid bounds/noop/retry, DOCX save, PDF export/text, nonempty unsaved sentinel hash, native reopen/edit/rollback and source hash preserved; final inventory empty.
- Run03 predates the exact Foundation comparison, shape guard and expanded ACK-count changes. It is not final-source acceptance for the current hashes below. Final fixture includes a proven floating-textbox guard case (unchanged preexisting shape no-op, appended shape rejection, explicit discard of that synthetic document), pending root native lease after independent review.

## Remaining scope/limitations

- Appended shape/inline-object rollback fails typed rather than guessing deletion. Figure/equation rollback remains open and must not be marked complete from this REF/table slice.
- Other story mutation and rich-object binary content/format restoration are not claimed by this append-only main-story recovery contract.
- No WordArt test or application crash retry is performed by this fixture. Separate rules-probe -609 failure remains separate evidence.

## Current source hashes

- `skills/WPSComposer/scripts/msoffice/macos_word_recovery.py`: `8e84c69a775773250a4e75f98f209ebe30ce04596889ae1ddacc726599879dc4`
- `skills/WPSComposer/scripts/msoffice/macos_word_session.py`: `ee86d43761c7e503b999bddc09b25d583cd4ca494a04e83b2c6722f5c3c6fa7c`
- `tests/msoffice/test_macos_word_recovery.py`: `988622f10f5012b1934c4f34f08676614cf5e859151da4e62a77e08c6d328764`
- `fixtures/microsoft_parity/macos_word_recovery.py`: `e14f99fa783c49dba275fefafdce4c1e2056138139bfabfd8375564d0c443c8c`

## Independent-review fix: generic inline coverage

Reviewer reproduced that inline-picture-only enumeration excluded generic inline objects. The same refusal guard now enumerates the native `inline shapes` collection with exact ordinal, native IDs, range, kind, dimensions and alternate-text digest; mandatory summary counts the full collection. Generic inline identity change/new object prevents controller fallback; missing count/identity ACK quarantines. RED: 2 expected failures before implementation (`/tmp/word-recovery-inline-red.txt`). GREEN: 46 passed in 2.03s. The pending native guard fixture includes a preexisting real inline picture plus floating text box control. No generic-object deletion/restore support is claimed.

Latest freeze supersedes the earlier hashes above:
- `skills/WPSComposer/scripts/msoffice/macos_word_recovery.py`: `5ee6b2e65e4e13c6a963bd68d0d576fc24ab31cf6fe18fc586ff13356da9090f`
- `skills/WPSComposer/scripts/msoffice/macos_word_session.py`: `ee86d43761c7e503b999bddc09b25d583cd4ca494a04e83b2c6722f5c3c6fa7c`
- `tests/msoffice/test_macos_word_recovery.py`: `975336f85dc46c070540969866914916f343e96a0c917a44d481ac717b66b0e6`
- `fixtures/microsoft_parity/macos_word_recovery.py`: `5efad969d6acb04875c5bf86a2ffd9a6e36e0d1ba30bfe34d4f632a5d80eb7f7`

## Final-source run04 and fixture cleanup revision

Final-source run04 retained **FAIL/quarantine**, not acceptance. The main controller recovery, original handles, field snapshot after fallback, invalid bounds/repeat, native DOCX/PDF and nonempty unsaved sentinel hash all passed. Immediately after the fixture closed the sentinel through the still-open target session, the target close binding failed `active window of missing value (-1728)`. Native collection counts reported one document/window but exact and ordinal lookups returned missing value; CUA still displayed the exact private DOCX and FALLBACK-ONCE. Word PID/start time did not change during the run. These observations establish UI/native model divergence; they do not prove a production binding defect, alias close or application crash.

Root performed exact-URL CUA close with explicit Do Not Save. One subsequent native inventory returned `[]`; the quarantine marker was backed up to run04, existing `recover_quarantine` returned true, staging remained preserved. `run-04/parent-ui-recovery.json` explicitly records manual UI recovery and `native_acceptance_passed:false`. No production fix/rebind/retry was added for this incident.

Approved fixture-only adjustment: complete the owned context close while the nonempty unsaved sentinel remains open; read-only inventory must prove target disappearance and unchanged sentinel hash. Only then close the exact sentinel with an independent 30-second guarded script (exact name, unsaved path/state, token text and exact ACK), followed by read-only inventory. RED 2 expected cleanup-order/helper failures -> GREEN **48 passed in 2.38s**; independent close script compiles without launching Word. Production remains `5ee6b2e6...` / `ee86d437...`; fixture `7de1c9bb54f3d6fea6c6f90da049f7b74eb7a1ab8ceedb8846396f4e7de8900f`; tests `3c7710b146341b2c320d5ff7a1ca872a46ecac70511aa4b778341db120fc0051`. Run05 has not been executed at this report revision.


## Final source-bound acceptance: run05 PASS

Command: `PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_recovery.py --execute --output docs/verification/microsoft-parity/macos-word-recovery/run-05`

Result: `{"passed": true, "error": null}`. All **13 native checks** passed. Run05 source-copy hashes were compared to all nine current production/fixture sources and matched exactly; no production changes occurred during review or native run.

- Empty checkpoint, integer/coercion, invalid bounds, repeated/no-op rollback verified natively.
- Real WindowsLongformExecutor controller retained the original REF/index handle instances through partial tracked REF plus untracked REF inside/outside a new table, acknowledged rollback, static fallback once and issue once.
- Current Foundation exact state comparison and mandatory topology-count ACK executed against Word. Native snapshot after rollback/fallback remained valid.
- Existing real inline picture plus floating text box accepted exact no-op; adding another floating box caused typed pre-deletion rejection, and both boxes remained present until explicit synthetic document discard.
- Owned context closed while nonempty unsaved sentinel remained; its exact name/path/saved/hash stayed unchanged. Independent sentinel close acknowledged, inventory returned to baseline.
- Native DOCX/PDF generation, PDF content, native reopen, real edit and rollback, original saved source hash and close verified.
- Starting and final native document inventories were `[]`. Word lease released to root for its installed09 checks.

Artifacts: `run-05/recovery.docx`, `run-05/recovery.pdf`, `run-05/report.json`, `run-05/native-runtime`, `run-05/shape-guard-runtime`, `run-05/reopen-runtime`. Failure runs01/02/04 remain separate and unchanged; run04's root UI recovery is not relabeled as automatic success.

Final focused pure/compile/synthetic/controller validation remains **48 passed in 2.38s**. Root coordinates whole-suite/installer verification separately. No commit/push or personal plugin installation was performed by this worker.

Final frozen owned-source hashes:
- `skills/WPSComposer/scripts/msoffice/macos_word_recovery.py`: `5ee6b2e65e4e13c6a963bd68d0d576fc24ab31cf6fe18fc586ff13356da9090f`
- `skills/WPSComposer/scripts/msoffice/macos_word_session.py`: `ee86d43761c7e503b999bddc09b25d583cd4ca494a04e83b2c6722f5c3c6fa7c`
- `tests/msoffice/test_macos_word_recovery.py`: `3c7710b146341b2c320d5ff7a1ca872a46ecac70511aa4b778341db120fc0051`
- `fixtures/microsoft_parity/macos_word_recovery.py`: `7de1c9bb54f3d6fea6c6f90da049f7b74eb7a1ab8ceedb8846396f4e7de8900f`
