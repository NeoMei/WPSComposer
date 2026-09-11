# macOS PowerPoint native feasibility and initial compiler evidence

Tested 2026-09-08, Microsoft PowerPoint 16.112.3. This is **partial parity evidence**, not a release or full UI edit/Undo acceptance.

## Accepted evidence

- `run-19/report.json`: a fresh, entirely automatic native AppleScript run created two slides, editable text, rectangle, embedded PNG, native 2x2 table and speaker notes; saved, closed, reopened, inspected native content, exported PDF, closed only the owned presentation, and verified the synthetic unsaved sentinel plus prior presentations. All 12 independent PPTX/PDF checks passed. `CORE_PASS_PARITY_GAPS` deliberately retains the structural/Undo gaps.
- `compiler-native-02/`: real execution of all nine currently closed presentation plan operations (reset, set_size, apply_preset, add_title, add_section, add_bullets, add_blank, add_image, add_table), producing four native slides and PDF. `stdout.txt` is the strict production success marker after owned close and prior-document checks. `artifact-inspection-v2.json` checks content across native font-run splits, editable image/table, page dimensions, PDF pages and text. The earlier inspector's artificial `|` separators caused a false negative on Chinese runs; both inspector results remain, and no native file was modified to repair it.
- `public-native-02/`: actual public `generate(..., format="pptx", preset="business", engine="msoffice")` followed by public `convert_to_pdf(..., engine="msoffice")`. Three pages, native table, text and bullets verified. The **original public PPTX** SHA-256 was `5ae678e65346db625589f30d9b0e1eeff3880a25871825968a7a901e0681dfde` both before and after conversion. PDF previews were inspected; Chinese text and bullets are visible.
- Portable tests: `tests/msoffice/test_parity_powerpoint_probe.py` (11) and `tests/msoffice/test_macos_powerpoint_script.py` (12), 23 passed. These are separate from the native evidence above.

The runner is `fixtures/microsoft_parity/macos_powerpoint.py --output-dir NEW_PATH --timeout 35`. Each output directory is new. Its scripts, stdout/stderr, runner hash, dictionary hash, source image and artifact checksums are retained. Core success still returns an unsuccessful parity status because unsupported rows remain.

## Native primitives and constraints

| Primitive | Native result |
|---|---|
| Create slide | `make new slide at end of ownedPresentation`; using `at end of slides of ...` failed `-2710` |
| Text/shape/image/table/notes | Passed in `run-19`; final 28 pt bold title, 200x90 pt rectangle at (60,140), `2060A0` fill, embedded picture, table cell values, notes verified independently in saved OOXML |
| Page width | Passed. Set dimensions before adding content: changing width afterward causes PowerPoint to rescale/reposition existing objects and fonts |
| Arbitrary page height | Dictionary does not expose an independent writable height. Compiler explicitly rejects height other than the verified 540 pt case; native on-screen preset then requested width gives 960x540 for current default plans |
| Slide movement | Move second before first, verify native slide IDs/order, restore, passed |
| Shape clone/removal/movement | Native duplicate/delete count assertions and coordinate edits passed |
| Slide clone | Direct `duplicate slide` returned `-50`; not enabled as a verified structural operation |
| Undo | `undo owned times 1` did not restore a script text assignment and invalidated the target shape; no reliable live atomic rollback claim |
| Selection | Native document-window selection inspection returned `selection type none`. Nonempty selection and UI edit/Undo remain unverified |
| Save identity | Save renames the presentation and invalidates the name-based reference returned from make. Rebind using a **unique output basename**, then assert its full path and saved state |
| Prior-document checks | Enumerate by index and materialize the comparison row into a variable before comparing it with the retained row. Direct named-reference snapshots fail with duplicate basenames; inline comparison also produced a false mismatch in compiler-native-01 |

## Staging investigation and failure evidence

`run-01` through `run-18`, `minimal`, `reopen-audit`, `relocated-open`, and `documents-location` retain the actual failures and recovery observations. They must not be erased or counted as successful automatic reruns.

Direct file access in worktree directories caused folder-access prompts. The parent coordinator authorized the exact low-sensitivity `run-10` directory using CUA; this was **manual recovery**, not an automatic probe pass. No child GUI automation, global application Quit, clipboard operation, macro execution or security-setting change was used.

PowerPoint could save PPTX/PDF in its own `Data/tmp` but native reopening returned `-9074`, including for a minimal one-slide file; LaunchServices reopening also did not produce the expected bound presentation. The parent observed an Open panel preview with a disabled Open button and no security warning. This is a location-related observation, **not a proven quarantine/security cause**.

Controlled relocation retained identical file bytes and restored the original stricter quarantine xattr exactly before opening. The same long basename opened successfully from `Data/Documents`, and a shorter worktree copy also opened. `relocated-open/restored-attribute-open.json` and `documents-location/report.json` record those checks. Python `shutil.copy2` on this host does not preserve the quarantine attribute; an initial `ditto` copy also did not retain its exact bytes, which is explicitly recorded rather than claimed as preservation. No attribute was removed or weakened. The relocated copy was finally closed after full-path/saved-state verification (`final-close.*`).

The accepted fresh `run-19` generated directly in **`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/`** and reopened there using ordinary native AppleScript `open` with an alias materialized outside the application tell block. No relocation, xattr change, folder prompt or LaunchServices workaround was used in this accepted run. The shared production runtime uses `Data/Documents/wpscomposer` for PowerPoint; Excel's staging path is independently owned.

## Remaining boundaries and recovery

The initial compiler implements the current closed plan, with an explicit 540 pt page-height limitation. It does not make arbitrary-height, new notes operations in GenerationPlan, slide cloning, nonempty selection, direct structural API, active-document rollback, legacy/macro-capable conversion, or complete Microsoft/WPS parity claims. Conversion is additionally gated by the shared native input validator; embedded workbook/OLE/chart-package scenarios remain outside the current closed plan.

Failed native sessions and their staging locations are retained for recovery. Synthetic unsaved sentinels from the investigation remain open intentionally. `run-19` specifically retained `演示文稿33` containing `SENTINEL-12cdf237e403491fa94f2aa1a5ad9432`, unchanged and unsaved; the successful generated presentation was closed. Initial baseline had zero user presentations. Do not bulk close/quit the application: any later user documents must be preserved, and earlier failed documents require exact ownership checks.

No production modules outside the two new compiler/test files were edited by this worker, and no commit or release was made. Shared runtime/routing changes are coordinated by the parent task.
