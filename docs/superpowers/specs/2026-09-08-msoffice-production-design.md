# Native Word production integration

User authorization: advance project to releasable state after native feasibility audit. First release scope assumed Word/DOCX/PDF on Windows and macOS pending optional user scope reply. Existing WPS Spreadsheet/Presentation and document editing remain available; no promise of MS Excel/PowerPoint or MS active-document editing in this release.

## Public contract

Add keyword-only engine="wps" | "msoffice" | "auto" to generate() and convert_to_pdf(). Default wps keeps existing routing; explicit wps must not intentionally select Word. auto probes installation/capability read-only, chooses WPS first then native MS Word for supported writer formats, pins that choice across generation, quality export, relayout and publication. Never fall back after starting a document or on content/runtime failures. Unsupported MS formats/legacy requests fail clearly before launching Office or publishing. Existing string artifact result and existing positional args preserved. Add read-only available-engine/capability diagnostic API if useful; no launches during detection. Open-result must use chosen engine for DOCX (system default can open wrong engine).

## Shared planning and lifecycle

Use existing build_longform_generation() -> GenerationPlan -> ExecutionOutcome -> M5 lifecycle for DOCX/PDF. Preserve guardrails including full body font/indent, six heading sizes, native numbering, one cover, named semantic bookmarks, TOC/index/ref fields, tables and images, resource preflight and visible approved degradation. Adapter supplies real M5 pagination from Word ranges; no fabricated page1 maps. Quality export and notice patch use same native engine; bounded relayout, package/PDF validation and atomic output publish remain mandatory. Explicit unsupported capabilities reject before native mutation if no documented visible degradation is possible.

## Windows

Reuse WindowsLongformExecutor through injected dedicated Word composer factory. Native factory verifies actual WINWORD.EXE and window PID plus isolation, closes only its own documents, never shared application. Native operations run in a bounded Python worker so parent deadlines apply even if COM blocks; preserve failure diagnostics and quarantine uncertain owned files instead of deleting open files. No global process kill/template/security/registration changes. Explicit WPS factory excludes MS ProgIDs. Same identity gate for PDF conversion. Keep existing WPS route behavior compatible.

## macOS

New native Word AppleScript compiler/executor consumes existing plan, uses exact task-owned document references and fresh Word-container staging, normalizes/stages resources locally, creates content/styles/fields/tables/images natively, repaginates and reads bookmark/range page information. Operations avoid active selection for document mutations. Snapshot existing names/paths/saved/full text in memory and verify afterwards. Serialize this plugin's Mac Word jobs across processes with an OS lock and one total deadline. One worker/AppleScript run must enforce remaining timeout and never quit shared Word. Safe cleanup only exact retained task-owned document; no permissions/security changes. Preserve user source bytes during convert by staging a copy. Native file format/complex object gaps must be proven or reported explicitly, never silent omission.

## Release acceptance

Meaningful routing/unsupported/timeout/ownership/publication regression tests and complete portable suite. Real public generate and convert on both OS, all supported semantics plus numbering edits/TOC pagination, overwrite refusal and atomicity, missing engine, no user doc interference, repeated runs, visible edit/Undo/save/reopen. Existing WPS real representative regression. Install a clean candidate in isolated destination and verify packaged modules/docs; version and release notes state exact capability matrix. No release/production readiness claim until these gates pass. Remote Windows uses existing task and evidence branch; do not interrupt ongoing jobs or infer stuck state from stale read_thread data. Preparing a releasable commit is authorized; publishing a release is not requested in this turn.
