# Fresh Windows native acceptance, round 3

PASS on exact code candidate `d4884a55fe8a0db48d1f2f48543948959aa08dbf`:
new run11 completed all 14 native operations, current validator passed 12/12,
and sentinel04 passed preservation and process cleanup checks. All commands
exited 0. The generated DOCX/PDF and all nine page renders are retained.
This completes the previously pending fresh Windows native gate. Production
admission remains NO_GO; no production integration, master merge or release occurred.

## Scope clarification and preflight

The controller clarified that the existing authorization allows creating the
unique task-owned unsaved sentinel in registered empty Word. The restriction on
PID 9252 forbids Quit/kill or claiming ownership of that application's startup;
it does not prohibit the already approved sentinel operation. No new permission
was needed. The round2 interpretation was overly restrictive, not evidence of
read-only filesystem or missing user authorization. Its original registration
failure and logs remain unchanged as historical experimental evidence; no further
RegisterActiveObject or registration workaround was attempted.

Fresh read-only GetActiveObject preflight confirmed Microsoft Word, version 16.0,
build 16.0.17932, Office16 path, and Documents.Count=0. The unchanged d488 wrapper
independently repeated the empty Microsoft Word guard before adding any document.
Exact executed source snapshots, raw and normalized SHA256, commands, timestamps,
stdout/stderr, and postflight observations accompany this report.

## Native ownership, preservation and cleanup

- Existing registered sentinel application PID: 9252. Only the unique task-owned
  unsaved document `文档4` was created there. It was never saved to disk.
- Before/after: 62 characters, saved=false, identical content SHA256
  `8b29801468a9bed8e8d9a25a22b6163eb1ad85e5919688ce583a38e9e4222e4f`.
  Both native runner snapshots independently matched the wrapper snapshots.
- New task Word PID 3936 differed from the sentinel PID; its document HWND mapped
  to that verified process. All 14 operations succeeded without errors.
- Only the verified isolated task application was Quit after its document closed.
  Its PID disappeared after 2.797 seconds of read-only polling; no force kill.
- After ownership/content recheck, the sentinel was closed with SaveChanges=0.
  Registered Word remained responsive, PID 9252 remained present, and its document
  count returned to zero. No other pre-existing Word instance was quit or killed.

## Saved artifacts and visual review

The current validator passed all 12 checks, including effective native numbering,
four TOC page references and actual rendered heading/TOC page numbers, exact 80
records, repeated headers, FangSong 12-point body text, indentation, page bounds,
and the final end marker. Word and PDF both report nine pages.

All nine pages were inspected in the contact sheet. Hierarchy, wrapping and
repeated columns are consistent with prior output, without visible page-edge
clipping or overlapping columns. Known prototype limitations remain: TOC title
below its entries, borderless table, split rows 004/025/046/067, sparse last page.
The validator's raw `visual_review: pending` field is preserved; the separate
`visual-review.json` records the subsequent human-visible render inspection.

The already passing 56 Windows regressions and real mouse/keyboard UI audit were
not repeated, as requested. Timeout/failure regression coverage remains distinct
from this successful native path; this run does not claim native fault injection
or production readiness. Historical runs01-10, round2's failed registration
experiment and their raw outputs remain unchanged. The root manifest now points
to this fresh result as resolution of the earlier pending gate.
