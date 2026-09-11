# macOS Excel native feasibility evidence

Candidate probe: `fixtures/microsoft_parity/macos_excel.py`.
Host: Microsoft Excel 16.112.3, `/Applications/Microsoft Excel.app`, bundle `com.microsoft.Excel`.
Full native feasibility passed in run-05. The public Microsoft generate/convert path also passed in public-run-01. These cover the current seven-operation spreadsheet plan grammar, not full SheetComposer parity.

## Reproduce

Use an interpreter with PyMuPDF available for independent PDF inspection:

```sh
PYTHONPATH=/tmp/wps-spike-validator-audit-deps \
/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/wps-task-session-startup/.venv/bin/python \
fixtures/microsoft_parity/macos_excel.py \
  --output-dir docs/verification/microsoft-parity/macos-excel/NEW_RUN \
  --timeout 60
```

The output directory must not exist. By default CLI native work happens in a unique child of Excel's own sandbox Documents directory, and Python copies the exact native bytes to the evidence directory. This avoids repeated grants for unrelated output directories without expanding file permissions. The runner writes the exact AppleScript,
runner source, raw logs, hashes and report. It never quits Excel, changes global
settings, uses the clipboard, executes VBA/document macros, or modifies unrelated
workbooks. A synthetic unsaved sentinel remains open on successful completion.
On uncertain native failures, owned workbooks remain available for recovery.
An AppleScript return code alone is not acceptance: independent XLSX XML and PDF
assertions must pass. Native artifacts are not rewritten by inspection.

## Raw runs and findings

| Run | Result | Evidence and implication |
| --- | --- | --- |
| `run-01` | Failed at worksheet creation, `-50` | `make new worksheet at end of worksheets of ownedBook` is rejected. Recovery experiment created the sheet inside `tell ownedBook` but returned a stale object reference; actual worksheet list showed Data and Summary. Explicit owned cleanup checked markers and sentinel unsaved state. |
| `run-02` | Failed at chart creation, `-50` | Formulas calculated to 60 and 120. Chart creation needs the worksheet context, like worksheet creation needs workbook context. Explicit owned cleanup recorded. |
| `run-03` | Timed out at native XLSX Save As after 60 seconds | Formula/cross-sheet calculation, chart creation/data binding, row insertion/deletion and selection passed in native runtime. Only an Excel lock file was emitted; a reported `saved=true`/full path is insufficient evidence of a saved artifact. A bounded HFS-path retry returned normally but emitted no output file and did not change identity; retained as a failed practical save experiment, not success. |

| `run-04` | Failed before creation, `-50` | Iterating `workbooks` directly failed when a prior unsaved workbook existed. Snapshot now obtains `name of every workbook` and rebinds each explicit workbook. No new workbook was created in this failed run. |
| `run-05` | Passed, 4.575 seconds | Container staging; two sheets, formulas and saved caches 60/120, formatting/merge, native chart XML plus cache [10,20,30], row insert/delete, sheet create/rename/clone/delete, selection, save/reopen and two PDF pages passed. Both synthetic unsaved sentinels preserved; owned native workbook closed. See report and supplemental-validation JSON. |

The unmodified run-03 logs prove these in-memory operations:

- Native values and `=SUM(B2:B4)` -> 60, `=Data!B6*2` -> 120.
- Merged A8:D8, font bold, fill, number formatting, column width and row height setters executed without errors. Saved/reopened persistence remains unverified in run-03.
- Embedded native chart, clustered column type, range source binding and title setters executed. Chart persistence/data cache require XLSX inspection.
- Entire row insert/delete preserved the owned marker after shifting down and back.
- Selection address read back as `$A$1:$B$4`.
- Sheet creation and rename executed; subsequent clone failed because the renamed worksheet reference was stale. The runner now rebinds by the known new name before copy. The correction passed in run-05.

## Native adapter details

Use a captured workbook object from creation, a unique content marker, and exact
full-path comparison after save/reopen. Do not infer ownership from `active workbook`.

Worksheet creation works with:

```applescript
tell ownedBook
 make new worksheet at end with properties {name:"Summary"}
end tell
set summarySheet to worksheet "Summary" of ownedBook
```

Returned creation references and references retained across rename can be stale.
Rebind the known worksheet after rename. Embedded chart creation works inside
`tell dataSheet`, then rebind `chart object 1 of dataSheet`. Set its `chart type`
to `column clustered`, and use `set source data` with a native range and
`plot by columns`. `chart title text of chart title` is writable.

Use `calculate dataSheet` / `calculate summarySheet` to avoid recalculating
unrelated workbooks. Structural row operations are `insert into range (range
"10:10" of dataSheet)` and `delete range (...)`. Native range selection can be
read using `get address selection` after selecting only the owned sheet/range.

## Permission recovery and current sentinel state

The parent inspected Excel UI and found a per-directory “Grant File Access”
dialog naming only run-03. After the parent coordinated this bounded grant, the
original native.xlsx appeared. `recovery-03` records subsequent exact-path/marker
binding, close, read-only reopen, native PDF export, and unchanged XLSX SHA256.
The original run-03 timeout report remains unchanged; recovery is separate evidence.
No global permissions were changed. A script return or workbook `saved=true`
does not prove that a native artifact exists.

Only these synthetic unsaved workbooks remain open after the public run:

- `工作簿5`, worksheet 1 A1 = `sentinel-58d9ceb7df8e`.
- `工作簿7`, worksheet 1 A1 = `sentinel-e9a7d3dce1ad`.

Both were read back unchanged and `saved=false` after public generation and
conversion; `public-run-01/sentinels-after.log` is the evidence. All owned native
output workbooks were closed. To discard fixture sentinels later, verify their
exact markers first; do not close by title alone.

## Task 2 compiler and public entry evidence

`skills/WPSComposer/scripts/msoffice/macos_excel_script.py` exports:

```python
compile_plan(plan, resources, native_path, pdf_path=None, timeout=60) -> str
compile_conversion(source_path, pdf_path, timeout=60) -> str
```

It validates the closed plan and Excel limits before launch, uses the captured
owned workbook and exact paths, preserves preexisting workbook count/path/saved
state, and emits `WPSCOMPOSER_MS_OFFICE_OK:spreadsheet` only after owned close.
Exceptions retain the document for shared-runtime recovery, without global Quit.
Caller responsibilities are app serialization, container staging, source macro
screening, deadline enforcement, independent artifact validation and publication.
The converter compiles XLSX/XLS syntax; the public runtime currently allows only
screened XLSX. Legacy XLS is not natively accepted by the public chain yet.

- `compiler-run-01`: native generation using all seven current plan operations;
  two sheets, formulas and caches 10/20, header font size and color/fill, native
  column widths/autofit and two-page PDF. Source snapshot and hashes retained.
- `compiler-conversion-01`: native read-only conversion, exact success marker,
  two-page PDF and identical XLSX source SHA256 before/after.
- `public-run-01`: actual `generate(..., format="xlsx", engine="msoffice")` then
  `convert_to_pdf(..., engine="msoffice")`; two sheets, formulas/caches 30/60,
  native header style, two rendered PDF pages with no business-content clipping,
  unchanged XLSX source SHA256, and both unsaved sentinels preserved. The report
  pins hashes of compiler, shared runtime, routes, orchestrator and conversion.

The 24 compiler tests and 12 probe-runner tests passed together. They do not
replace native evidence. The raw compiler run uses default white header font on
an explicitly unshaded second-sheet header, which matches the requested existing
WPS semantics; the public run uses the normal blue header and is visually checked.

## Shared runtime review

`runtime-review` contains a NONEXECUTABLE ZIP scanner fixture, never opened in
Office. The initial shared runtime scanner accepted an active vbaProject
relationship/content type when its payload part had an innocuous filename.
This was reported to the parent as a P1 before broad acceptance; the parent owns
the shared scanner fix and its regression tests. Do not treat the scanner fixture
as a native Office artifact or macro execution test.

## Scope still incomplete

The full SheetComposer union remains open beyond the current seven-operation
plan: sheet move, row/column move/clone, freeze panes, borders, conditional formats,
active attachment, UI Undo and transaction semantics. Native probe chart/merge
support is not yet exposed as new production plan operations. Keep these rows
open in the full Microsoft/WPS matrix.
