# Mac Excel worksheet deletion review

Date: 2026-09-09

Scope: read-only review of the current macOS Microsoft Excel whole-worksheet
delete path, its public preflight, the frozen WPS `SheetComposer` behavior and
retained native evidence. I changed no production, tests or fixtures; launched
no native application/UI; performed no installation; and made no commit.

## Verdict

**SCOPED PASS — no actionable P1/P2 found in the currently enabled deletion
subset.** General deletion of an existing worksheet remains an explicit,
accurately documented parity gap. The implementation does not claim that the
WPS baseline's unrestricted `sheet:N` remove is available through Mac Excel.

## Current enabled contract

`MacExcelSession.apply_structural_op()` validates the complete operation before
compiling a target. A whole-sheet remove is accepted only when its positional
index is present in `_fresh_empty_sheets`. That set is populated only from the
actual result of this same session's `insert(type='sheet')`; formatting,
move/clone, non-sheet insertion, business-sheet access and a completed delete
clear it. Thus an arbitrary existing sheet never reaches `_run()`.

The deletion AppleScript then:

1. rebinds `ownedBook` by its private filename and requires its exact native
   `full name` to equal the private owned path;
2. resolves the recorded worksheet index inside that bound workbook;
3. checks in the same script that more than one worksheet exists;
4. executes `delete obj`; and
5. returns only after the AppleEvent finishes.

Retained `full-01` and `full-02` evidence shows the supported sequence creating
`Temporary` as actual `sheet:3`, deleting `sheet:3`, and receiving the public
remove result. The current docs correctly call this “a worksheet created empty
and left untouched by this same session,” rather than general worksheet-delete
parity.

## Modal-confirmation evidence and baseline comparison

Frozen `SheetComposer._remove_element()` calls
`Worksheets(sheet_index).Delete()` directly and maps a COM exception to
`ValueError`; it has no separate last-visible-sheet or confirmation policy.
That is the frozen WPS behavior, but it does not prove Microsoft Excel for Mac
can execute the same call without UI.

The retained `final-primitives-01` native attempt deleted an existing nonempty
worksheet and timed out with `NATIVE_OFFICE_TIMEOUT` while Excel displayed a
confirmation. `business-probe-01` cleared the worksheet's used range first and
still timed out at deletion. Those owned-workbook incidents were retained and
manually recovered; the current source subsequently added the pre-native
existing-sheet rejection. This is evidence for retaining the narrow guard, not
for bypassing it with an application-global alert setting.

## Failure, ownership and deadline review

- **Exact document binding:** every enabled delete uses `_run(bind=True)`, whose
  script binds by the private workbook name and verifies the exact full path
  before selecting any worksheet. It does not rely on the active workbook.
- **Existing or missing target:** an existing worksheet not in the private
  fresh set is rejected locally with `NotImplementedError`, before an
  AppleEvent. A stale/nonexistent fresh positional target raises natively; it is
  never reported as a successful delete after a native error.
- **Last worksheet:** the supported flow inserts a fresh sheet first, so a
  normal target has at least one remaining sheet. The mutation script still
  checks `count <= 1` before `delete obj`. A failed check enters the normal
  native-error retention/isolation path.
- **Protected workbook/worksheet:** the public existing-sheet path is rejected
  before mutation. If Excel rejects an otherwise eligible fresh deletion due
  to protection, `_run()` treats the native error as failure, writes its private
  script/stdout/stderr plus `recovery.json`, quarantines the Office lock, and
  makes all later session operations fail locally. The current evidence does
  not certify a protection-specific error code or message, and this review does
  not invent one.
- **Timeout/cancellation:** `_run()` applies both the remaining session deadline
  and an AppleScript timeout. `TimeoutExpired`, ordinary native failure and
  other `BaseException` paths set `_failed`, retain recovery data and
  quarantine the lock. `close()` will not send another AppleEvent after such a
  failure. The focused expired-session test confirms an already expired session
  launches no event and quarantines before releasing its lock.

Worksheet identity within this enabled subset remains positional. The native
evidence proves the immediate insert-then-delete flow, while the public API does
not claim a persistent stable worksheet ID. A user/VBA/external reorder between
the two calls has not been natively accepted and should not be inferred safe.
This is an evidence limit shared with other positional sheet targets, not a new
P1/P2 demonstrated by the reviewed frozen flow. Any future general-delete work
should bind an immutable preimage rather than treating the current integer as a
stable identity.

## Independent pure verification

Using the required Python environment and
`PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps`:

- `tests/msoffice/test_macos_excel_session.py`: **65 passed**.
- Focused existing-delete preflight plus expired-session and edit-preflight
  cases: **2 passed, 131 deselected**.
- A direct state-machine check confirmed the actual insertion result `sheet:3`
  becomes the sole fresh entry, the next delete emits exact worksheet 3 plus
  the last-sheet guard in mutation order, and successful deletion clears the
  set.

The unit suite explicitly proves existing-sheet deletion cannot call native and
that deadline expiry quarantines without another event. The insert/delete
success is additionally backed by the retained source-bound native full-01 and
full-02 flows. These checks do not establish general existing/protected-sheet
deletion.

## Minimal native acceptance required before widening the contract

Use a newly created private owned workbook under the existing Excel lock, while
preserving a separately created unsaved synthetic sentinel by exact
name/saved-state/token hash. Do not operate on a user workbook, change global
security settings, globally quit Excel, or silently dismiss UI.

The bounded matrix should record immediate and reopen workbook topology and
source/artifact hashes for:

1. fresh empty insert followed immediately by delete, with exact count delta,
   generated name/index preimage and remaining-sheet identities;
2. a pre-existing nonempty sheet, under a bounded timeout, proving either a
   truly confirmation-free scoped primitive or the retained fail-closed modal
   result;
3. nonexistent and last-sheet targets, proving no count/name/content change;
4. workbook-structure protection and sheet protection separately, recording
   the native error and proving no deletion;
5. forced timeout/cancellation at the delete boundary, proving the session is
   quarantined, no later AppleEvent/automatic close is sent, the owned recovery
   target remains, and the sentinel is unchanged;
6. a deliberate reorder/rename between insertion and deletion, establishing
   whether a stronger name/topology preimage is required before positional
   deletion can be widened.

If Excel displays a confirmation, preserve the failed evidence and let the
parent perform an exact owned-document recovery; do not have the fixture click
the dialog or change application-wide alerts. Only after a source-bound run
proves the target preimage, count delta, persistence and cleanup should general
existing-sheet deletion be advertised.

## Reviewed identities

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_excel_session.py` | `a052c08efd2fb68b1f1812edc91609fc89424a9570321b2b89e3f9aafa6c2b3b` |
| `skills/WPSComposer/scripts/sheet.py` | `88313d34ba90ad7384c5bcd758eae59db76a32a1aab0bcf95181dc96cb96fdbe` |
| `skills/WPSComposer/scripts/msoffice/edit_preflight.py` | `1abe64ecc53e84e0deb0ab1817a9427f3d7686f675bb3c921ad98d150c5b7868` |
| `tests/msoffice/test_macos_excel_session.py` | `cb110f2bc0e5eab36c83b47a6285d2dfffae8679ab8d89258ede447697b0df16` |
| `tests/msoffice/test_edit_preflight.py` | `a8d5cb71ea51efb2dcb40718c1ca24678e6b7ddaee660f455f70b69fe63215eb` |
| `docs/verification/microsoft-parity/macos-excel-sessions/README.md` | `34f9a3fdd0dc86babc78057c44f5dad03022af01cbc1dd00dd6861e46a1bb4cc` |
