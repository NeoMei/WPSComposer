# Microsoft parity: targeted Office.js supplement

Status: proposed after native probes; not enabled or installed.

## Why this is needed

The accepted native adapters now pass representative generation, conversion and file edits, but AppleScript exposes no verified non-rebinding save-copy for live Word/Excel/PowerPoint. Word window IDs are null on the installed Mac version, preventing trustworthy unsaved attachment. Excel nonempty worksheet deletion prompts even after clearing its range: two controlled probes timed out and the parent canceled their exact test dialogs. These remain baseline gaps, not waived requirements.

## Proposed bounded addition

Keep the native desktop COM/AppleScript route for supported operations. Add a local Office task-pane add-in only for explicitly selected native API gaps: document-bound attachment, native file snapshot/save-copy, and modal-free Excel worksheet deletion. It must keep engine=msoffice throughout; the application remains the layout and editing engine.

The user enables WPSComposer once from the Office add-ins menu in the target document. The add-in registers an ephemeral instance ID, host application, read-only state and runtime requirement sets with the local bridge. Public attachment selects one explicit connected instance; ambiguous/missing instances fail before mutation. No title-based guessing or changes to document custom properties merely to identify a session.

Transport uses a loopback endpoint with exact-origin checking, a short-lived pairing token, one outstanding command per instance, closed operation schemas and a monotonic job deadline. Never expose arbitrary script evaluation. File snapshots use Office's own getFileAsync compressed slices, size/hash validation and closeAsync in finally. Publish locally only after all slices validate. The snapshot itself must not rebind/save the open user document. A timed-out mutation is quarantined until its exact instance is reconciled.

Installation needs a per-app manifest in Office's supported wef location and a local web endpoint. Use a task-owned certificate only if required by the actual host; adding a trust-store certificate or altering Office security is not part of the existing authorization. Prepare files and tests first, then validate setup and removal in a controlled native run. No app restart while unrelated unsaved documents remain open.

## Gates before claiming support

1. Runtime API-set probes on Windows and macOS for each actual app, with unsupported versions explicit.
2. Saved and unsaved document binding, two-window ambiguity, host switch and stale-instance rejection.
3. getFileAsync snapshot preserves native shapes/formulas/fields, unsaved state and original path; reload saved copy in the actual desktop app.
4. Excel delete removes the exact requested sheet with no modal, preserves other sheets and never silently reports canceled work as removed.
5. Actual edit -> Undo -> explicit save/discard -> reopen, plus unrelated unsaved sentinel.
6. Authentication, origin, stale-token, oversized-slice, disconnect/timeout and atomic-publication tests; install/uninstall evidence.

Do not enable the bridge on the assumption that documentation equals native proof. This supplement does not waive other catalog gaps such as legacy formats, arbitrary paper sizes or all direct Composer operations.

## Official evidence checked 2026-09-08

- [Office.FileType](https://learn.microsoft.com/en-us/javascript/api/office/office.filetype?view=common-js-preview) documents compressed native DOCX/PPTX/XLSX bytes through getFileAsync. Actual host support still needs a requirement-set/runtime probe.
- [Excel.Worksheet](https://learn.microsoft.com/en-us/javascript/api/excel/excel.worksheet?view=excel-js-preview) exposes workbook-scoped stable worksheet IDs and worksheet operations.
- [Sideload add-ins on Mac](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac) describes per-app wef manifests and enabling the task pane from Office's add-ins menu.
