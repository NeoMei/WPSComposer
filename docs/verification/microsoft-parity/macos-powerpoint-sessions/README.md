# macOS PowerPoint document sessions — native verification

Host: Microsoft PowerPoint 16.112.3, macOS. Native AppleScript dictionary from the installed application. No GUI layout, static PPTX generation, global Quit, security changes, or user-document cleanup. Explicit clone/move native clipboard operations are disclosed on results/exceptions; prior clipboard contents are never read or restored. All native open/save paths are isolated under PowerPoint `Data/Documents/wpscomposer/session-*`. The source is validated, copied with a total deadline, validated again, and bound by an exact unique native basename plus full path. Shared OfficeJobLock spans each session. Native results require a JSON success envelope; mutation errors/timeouts retain raw scripts/stdout/stderr and quarantine the app.

## Verified results

- `edit-04/report.json`: PASS. Source original SHA-256 `15782a04b3d7a3f528f5d099b61f16d49f944684cb94744731ecc86687d466ae` unchanged. All preexisting presentation names/paths/saved flags/slide counts/texts unchanged, including unsaved sentinels. Ten edits, PPTX save, PDF export, owned close, native reopen, semantic assertions and PDF text pass. Native slide count is 3.
- `edit-03`: independently caught and corrected ordering requirements: setting line weight after color resets color; switching master background after color resets background; automatic text size can override a previously set height. Final emitter applies toggles first, final colors/bounds last. Reopened values: title height 70, line `#445566`, background `#EEFFFF`.
- `runs-01`: a native character-font split produces two native `text flow` runs. Editing run 2 text/strikethrough/color survives save, PDF and native reopen. Run 1 remains `S`; run 2 becomes `ession native run`. Source and all preexisting documents unchanged.
- `attached-01`: one unique owned saved file was opened natively, attached, preflighted, edited, explicitly saved in place, detached without native close, then separately closed using exact identity and verified saved state. Native reopen confirms edited text. Original source SHA unchanged, preexisting docs unchanged. This does not claim attached save-copy support.
- `public-01`: actual `document_api.edit(engine='msoffice')` with text edit and native image insertion, PPTX + PDF publication and `inspect_after`; then actual public `inspect` of the output. Result `ok=true`; source and preexisting documents unchanged.
- `clipboard-02`: production session slide clone to start, cross-slide shape clone, safe copy/paste/confirm/delete move, final-index slide move, insertion at count+1 and removal. Save/PDF/reopen verifies native slide IDs `[256,257,258]`, shape counts `[2,3,3]`, retained notes and original image geometry/name. Clipboard side effect disclosed; source and all existing documents unchanged. `clipboard-01` retains the preceding primitive investigation.
- `public-02`: public slide clone and cross-slide image move with explicit `clipboard_changed`, native table-cell font edit, PPTX/PDF publication and public native reopen inspection. Source and preexisting documents unchanged.
- `selection-01`: native pasted shape-range selection inspected directly, with a full shape snapshot; no GUI selection. `selection-02` verifies one selected shape text/font/geometry edit, native save/PDF/reopen and unchanged existing documents.
- `semantic-03-public`: public `create_document('slide', engine='msoffice')`, native empty document, all 15 semantic method signatures, five slides, preset fonts/background, title/section/text/bullets/blank, textbox/autoshapes/image/table/notes and static layout elements. Native save, PDF and reopen pass; OOXML independently confirms editable table, embedded picture and native connector. Rendered page 5 checked. Existing documents unchanged. `semantic-01` retains the earlier direct-class run. `semantic-02-public` used an explicit owned native view-selection setup before layout; `semantic-03-public` removes that setup and runs only public semantic methods. The production new-slide method now establishes its owned current slide.
- `save-current-01`: explicit save_current updates a synthetic source copy; close(save_changes=True) saves a second edit to that original copy. Native reopen/PDF verifies the final text. A controlled external replacement is rejected by preflight and save_current before any native call; replacement bytes remain unchanged. Original input and existing documents unchanged. `saved-current.pptx` is an exact copy of the native reopened artifact, with provenance retained.
- `run-04`: first full native document snapshot success. `dictionary-read-01` verifies text-flow contents, no numeric shape ID (`missing value`), strike enum and ruler values.

## Current supported session semantics

Inspection: slides, unique shape-name targets, positional fallback for duplicate names, native slide ID, shape type/geometry/z-order, shape text and per-paragraph native runs, fonts, fill/line data, notes, table cells, and page width. Numeric public enums are mapped from dictionary names; unavailable attributes are null instead of fabricated values.

Formatting: text on shapes/paragraphs/runs/table cells; font name/size/bold/italic/underline/strikethrough/color; paragraph alignment/spacing; shape bounds/rotation; fill/background colors/visibility/transparency; line color/weight/dash/transparency; text-frame margins/wrap/auto-size/anchor; shape name; master background behavior; page width. Only fields actually observed above are native-accepted evidence; remaining enum variations need their own native checks.

Structure: slide insertion, native text-box insertion (`shape` uses the same textbox behavior as WPS structural API), image insertion, slide/shape removal, slide move, in-place shape duplicate, native complete slide clone, and cross-slide shape copy/move. Copy/paste destinations are bound before copying; source shapes are deleted only after destination native paste count verification. Explicit owned save, PDF and reopen. Attached saved-file `save_current` and detach; unsaved attachment save or attached save-copy is preflight-rejected before public mutations.

Semantic API: `new_document`, `slide_count`, title/section/text/bullets/blank slides, textbox, AutoShape, image (native aspect ratio), editable native table, notes, background, slide size with height 540, design preset, static layout template and `save_pptx`. Returned shape/slide references are public target strings rather than COM objects. Signature compatibility does not imply every enum variant is natively exercised.

Save preflight rejects unsupported fields and output formats or existing output files before mutation. Explicit source saves verify the original SHA before native save and again before publication. Closing with save uses this same original-file path; a preflight conflict leaves the session available for explicit discard.

## Open parity gaps and investigations

- Numeric `Shape.Id` returns `missing value`; no fabricated XML ID to native object mapping. Unique `@name` is provided; duplicate names use positional targets and need reinspection.
- Independent slide height absent from dictionary; arbitrary-height generation and editing remain unsupported.
- Direct native slide `duplicate` returned -50 in Task1; complete slide clone is now verified through native copy/paste. Clipboard side effects are disclosed for that route and cross-slide shapes. In-place native shape duplicate does not use the clipboard.
- Attached save-copy preserving native identity and saved-state is unsupported. Unsaved attached save is rejected. No reliable native Undo boundary; attached multi-primitive atomic edit remains rejected by public API.
- Native none and shape-range selections are verified. Single-shape selection editing is verified in `selection-02`. The installed dictionary exposes no native command to establish a text selection, so validating text-selection behavior requires an independently prepared selection. Multi-shape/child-group selection editing remains rejected because no per-selection atomic boundary is established. Text-selection reading is implemented from dictionary properties but not yet natively exercised.
- `line.visible`, paragraph indent/right indent/keep/widow/tab/line-rule dimensions, shape type mutation, unsupported shape/page attributes remain explicit rejected fields. Inspection now includes native run and cell fonts/paragraph/fill; it still does not expose every WPS textframe/paragraph dimension.
- `visible` is accepted for interface compatibility but no app-global visibility is changed; strictly hidden-native-open behavior is not proven.
- Strict macro-free input gate accepts PPTX only, excludes legacy presentations, embedded packages, macros, and external non-hyperlink relationships. This is not full legacy-format conversion parity.

## Raw failures retained

`run-01..03` JSON bridge failures and `raw-snapshot` narrowed native missing/enum values; `edit-01` native inventory variable `rows` collided with a PowerPoint property. `edit-02` passed initial coarse checks but independent deep artifact checking found reset fields; it is not evidence for those fields. All later verified results derive from their own original native artifacts, not edited ZIP substitutes or GUI make-up steps.

Reproduce core acceptance with a new output directory:

```sh
PYTHONPATH=/tmp/wps-spike-validator-audit-deps /Users/neomei/项目/codexprojects/WpsComposer/.worktrees/wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_powerpoint_sessions.py docs/verification/microsoft-parity/macos-powerpoint/run-19/native.pptx docs/verification/microsoft-parity/macos-powerpoint-sessions/NEW-RUN
```

Reproduce public semantic generation and explicit source-save acceptance using fresh directories:

```python
from fixtures.microsoft_parity.macos_powerpoint_sessions import run_semantic, run_save_current
run_semantic("docs/verification/microsoft-parity/macos-powerpoint-sessions/NEW-SEMANTIC")
run_save_current("docs/verification/microsoft-parity/macos-powerpoint-sessions/semantic-03-public/semantic.pptx", "docs/verification/microsoft-parity/macos-powerpoint-sessions/NEW-SAVE-CURRENT")
```

Final targeted verification: 103 passed across PowerPoint session (52), compiler (12), probe (11), and document engine routing (28). This is local macOS evidence, not Windows or full field-level WPS parity.
