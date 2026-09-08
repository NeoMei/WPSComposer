# Windows production API acceptance: first candidate

Native Windows public-API and real UI gates passed on `f48497d9e8d2c453a5075824229e287fab849b7e`
plus Windows identity fix `a745d2ac352f5f9d5c2037c9dd826ce54f797c91`. These results do not cover the newly requested
`89b39e8` candidate. Production admission remains NO_GO pending that verification.

## Root cause and fix

Initial smoke-01 failed before Word startup because Python 3.11 lacked pypdf.
Missing pypdf/reportlab were installed only under ignored build; pytest was reused
from the earlier ignored project dependency directory. No global settings changed.
Smoke-02 then reproduced `Native Word application HWND is unavailable`, before
Documents.Add. Its complete raw native worker diagnostics and request are retained.

Empty Word 16.0.17932.20910 raises AttributeError for Application.Hwnd, and the
tested IOleWindow/accessibility approaches did not supply an empty-host handle.
The successful identity probe used a random per-instance Caption challenge.
[Microsoft documents Application.Caption](https://learn.microsoft.com/en-us/office/vba/api/word.application.caption)
as the application title-bar property. After existing executable/new-PID,
IUnknown isolation and empty-document checks, the fix sets a unique temporary
caption and requires exactly one matching OpusApp HWND in that new PID. It restores
the original caption and rechecks IUnknown and zero documents before Add; the
document HWND/PID gate remains a second check. No registry, security preference,
Normal template or running-object registration is modified. Concurrent caption
changes are not overwritten, and proof/restore failures never grant Add or Quit.

TDD evidence: identity-red.log records 3 expected failures on the unfixed candidate;
the final msoffice group passes 70 tests, including missing HWND, zero/multiple
window matches, concurrent documents/caption changes and enumeration failure.
Only windows_host.py and its Windows tests were changed.

## Fresh official API runs

| Gate | Outcome |
| --- | --- |
| smoke-03, timeout 180 | PASS, 168.3 s: public generate(docx), shared M5 quality chain, public convert_to_pdf, source preservation and overwrite refusal |
| representative-01, timeout 600 | PASS, 188.8 s: official runner, five-page DOCX/PDF, six heading sizes, four numbered levels, native TOC, complete 37-row table and repeated header |
| public-pdf-01 | PASS, 136.3 s: public generate(format=pdf, engine=msoffice), requested PDF returned without an unrequested DOCX |
| Independent saved-artifact inspection | 11/11: native structure, actual numbered headings and TOC page references, all 36 records once, repeated headers, page bounds and PDF typography |

The two public PDF routes have byte-identical rendered pixels on all five pages.
Body text is FangSong 12pt with measured first-line offset 24.026pt. DOCX heading
sizes are exactly 16/15/15/14/14/12pt; exported PDF sizes differ by at most 0.12pt
due to Word export rounding. All five pages were visually reviewed. The cover,
three-level TOC and sparse final table page are retained as shared M5 layout.

## Sentinel, ownership and actual UI

Registered Word was Microsoft Word with zero documents before task sentinel
creation. Existing PID 9252 held only synthetic unsaved 文档5 (65 characters).
Before and after both representative/public-PDF calls, its name/path/saved state
and text SHA256 were identical: `3642637d7fa656c56ad2f444e51b181997d28dfe2cf52035d4503640088bc037`.
It was closed only with SaveChanges=0. PID 9252 remained responsive, present and
empty; it was never Quit or killed.

Eight successful native API instances and UI PID 21576 were verified absent after
cleanup. Pre-Add identity events and close results are in traces. The initial
failed acquisition and exploratory empty identity probes deliberately retained
their uncertain COM hosts; known probe PIDs 15152, 26492 and 6332 remained present.
Other old PIDs were also preserved. See remaining-processes.json: the successful
run cleanup PASS must not be read as a claim that all exploratory processes exited.

The UI used a task copy of representative generated.docx. Initial open used the
verified production host; editing, Undo, Save, Close and native file-dialog Reopen
used actual mouse/keyboard via the Computer Use API. COM was used only for setup,
identity/readback and final verified empty-instance cleanup. After input the
marker was present and Saved=false; Undo restored the original text hash. Save
and reopen retained identical DOCX bytes, one table, one TOC and four fields.
The final UI instance closed successfully. Raw evidence is in ui-01.

CUA focus text was repeatedly stale, some accessibility captures were null, one
element index failed, and one capture returned unrelated browser document_text.
Those fields are not treated as authoritative. Visible task-window screenshots,
retained COM identity and exact text/file hashes substantiate the UI result.
Published screenshots are document-only crops; account controls and unrelated
file-dialog listings are excluded. Unrelated UIA document_text is omitted from
the filtered event record; original captures stay in ignored build.

## Portable tests and outstanding limitations

Full pytest -v completed with 2681 passed / 45 failed / 39 skipped in 481.93 s.
The new worktree lacked the lockfile-pinned wpsjs templates. npm ci --ignore-scripts
installed the existing locked dependencies only into ignored node_modules;
Mac source files and lockfiles were not changed. Re-running all 45 failures yielded
44 passed / 1 failed. The remaining Mac-template test asserts POSIX mode 0600,
whereas Windows reports 0666. Its source equals the candidate; it remains a clearly
reported failure, not a silently skipped or modified Mac test. This is not a
clean full-suite pass. The separately requested Windows/routing/runner group is
70/70 passing. No successful native path claims native timeout/fault injection.

The audit-only sitecustomize observer records native host return events and
copies synthetic worker logs before routine cleanup. It does not replace APIs,
alter arguments/results or prevent normal cleanup. Source snapshots and hashes,
original errors, command exit codes, rendered artifacts and PID checks are kept.
No prototype generator was used to satisfy the production generation gates.
