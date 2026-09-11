# macOS Excel sessions: native evidence

Host: macOS, Microsoft Excel **16.112.3**, `/Applications/Microsoft Excel.app`. All workbook creation, edits, chart construction, calculations, worksheet operations, saves and PDF exports use Excel AppleScript. OOXML/PDF readers only inspect native outputs; they never repair or manufacture them. Every run uses a fresh evidence directory and a private Excel container job. Original failed reports remain unchanged.

## Verified paths

| Evidence | Result |
|---|---|
| `full-02/report.json` | Owned XLSX inspection; cell value/font/fill/alignment/number format/size and native border edits; named chart/shape edits; sheet rename/page setup; row/column insertion/removal; new untouched empty sheet creation/removal; XLSX + 2-page PDF; reopen formula caches 65/130; source unchanged; staged abort produces no published file. |
| `shape-readback-01/report.json` | Native reopened shape fill `#F0F0F0`, transparency ~0.1 and line `#123456`, weight 2. |
| `public-01/report.json` | Public `inspect` → `edit` → native save/PDF → public reopen with engine `msoffice`; formula caches 75/150; source hash unchanged. |
| `active-01/report.json`, `active-03/report.json` | Saved active exact-path binding, active B2 selection, inspection restores selection, save-copy preflight rejects before mutation, native save-current, session close leaves workbook open; public attached edit saves B2=19 and reopened formula cache=78. |
| `structural-04/report.json` | Six native row/column clone/move/remove tests against multi-cell target B2:D4 operate on only its first row/column, matching SheetComposer. Extra rows/columns remain. Each copied workbook closes without a clipboard prompt. Sheet move/clone preserve selected-sheet business addressing, output is saved/PDF-exported/reopened, source hash unchanged. |
| `business-04/report.json` | All direct business methods exercised: native tables/cells/formulas, cell/range styles and BGR integer color, borders, merges/title, column/row sizing/autofit, chart, condition replacement, frozen panes, headers, multi-sheet creation/rename; independent OOXML checks and 3-page PDF. Native reopen and source hash checks pass. `page-1.png` visually inspected. |
| `new-02/report.json`, `new-03/report.json` | Public `from skills.WPSComposer import create_document`; empty native workbook creation, direct write-cell None clears native contents, data/formula/title/chart/freeze, native XLSX/PDF, reopen B5=24 with freeze state true. |

| `final-state-01/report.json` | Exact final source hashes; both original unsaved sentinel values/states preserved; no owned workbooks remain open; no Excel quarantine remains. |

Platform-independent verification: **82 passed** across `test_macos_excel_session.py` (45), `test_macos_excel_script.py` (25), and `test_parity_excel_probe.py` (12).

The immutable source snapshot/hash in each run identifies the exact implementation used. Later code changes are covered by later focused runs; earlier reports do not claim to validate later source.

## Reusable runner

Run from the repository root with Python 3.9+ and the existing native-validation dependencies:

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
  /Users/neomei/项目/codexprojects/WpsComposer/.worktrees/wps-task-session-startup/.venv/bin/python \
  -m fixtures.microsoft_parity.macos_excel_sessions \
  --mode business \
  --source docs/verification/microsoft-parity/macos-excel/public-run-01/public.xlsx \
  --output-dir /absolute/new/evidence/directory
```

Modes: `full`, `business`, `structural`, `active`, `new`. `full` uses `macos-excel/run-05/native.xlsx`; other modes use `macos-excel/public-run-01/public.xlsx`. Active mode opens only a synthetic copy in Excel's container and later closes that exact fixture; the adapter itself never closes an attached document. Do not run multiple Excel probes concurrently.

## Safety and lifecycle

- File sessions copy and validate the source before opening the private copy. A per-Excel lock spans the whole session, with a 600-second total deadline and 60-second maximum native step.
- Every native step checks the owned full path. Failed/uncertain native steps quarantine Excel and retain the workbook, script, stdout/stderr and recovery metadata. There is no automatic retry, global Quit, global security change, or hidden reopening of the source.
- XLSX/PDF publication uses deadline-bound atomic `publish_artifact`, not direct copying to the public destination. Existing destinations and original sources are rejected. Attached save-current preflights a saved writable path before mutation; unsupported save-copy is rejected before mutation.
- Only explicit row/column move or clone performs native copy/cut. Results report `clipboard_changed: true`; failure metadata additionally reports `clipboard_may_have_changed: true`. The operation then exits its own native cut/copy mode to avoid Excel's large-clipboard close dialog. Previous clipboard contents are never read or restored. Routine generation, formatting and inspection do not use copy/cut.
- Saved active inspection restores the originally bound worksheet and cell selection on success and on native errors. Timeout errors do not launch restoration retries. Attached close only releases the session lock.
- Two original unsaved sentinels remain: `工作簿5`, A1=`sentinel-58d9ceb7df8e`; `工作簿7`, A1=`sentinel-e9a7d3dce1ad`; both saved=false. Recovery logs contain native token/state checks. Generic recovery correctly rejects unidentified unsaved workbooks; explicitly coordinated fixture cleanup verified both known sentinels before clearing only the matching quarantine. Original quarantine/refusal files remain in the evidence directories.

## Failed runs and remaining gaps

- `run-01`/`run-02`: Excel returned another active worksheet's cell properties even with explicit worksheet object references. Owned worksheet activation before inspection fixed the semantic failure; `run-03`/`full-02` verify correct cross-sheet values. Active inspection restores the prior selection afterward.
- `probe-01` records initial Foundation JSON/properties experiments and the observed clipboard change. The original stricter clipboard ban was an implementation preference, not the approved user scope. Explicit move/clone was subsequently authorized and its side effect is disclosed.
- `structural-01`: native worksheet copy succeeded, but comparing a mutable before-name list failed to identify the result. Count plus destination-index rebinding fixed it. The raw failure and owned recovery remain.
- `final-primitives-01`: unsaved active attachment correctly rejected without mutation. Deleting an existing nonempty worksheet displayed a confirmation and timed out. `business-probe-01` proved clearing its used range still does not avoid that confirmation. Both dialogs were inspected and cancelled only for the owned task workbook. **Existing worksheet deletion now rejects before native mutation.** Only a worksheet created empty and left untouched by this same session has the verified deletion path. Full worksheet-deletion parity remains open.
- `structural-03`: copying a whole row then immediately closing displayed Excel's large-clipboard prompt. The task-owned prompt was dismissed; explicit cut/copy-mode cleanup fixes the path in `structural-04`.
- `business-probe-01`/`02`, `business-01`/`02`: raw errors establish the required native references for charts, range merges and conditional formats. Collection deletion is not supported; condition replacement deletes individual native rules. `business-03` completed native output/reopen but its initial validator accepted only XML `state="frozen"`; Excel's valid `frozenSplit` representation is independently accepted in `business-04`.
- `new-01`: native Save As succeeded, but the original new-workbook reference became stale. Rebinding the unique saved name before checking its full path fixes the constructor (`new-02`).
- `active-02`: a read-only identity query returned native osascript `-36/-1700` with an identical script to passing `active-01`. The runtime quarantined rather than retrying. Raw diagnosis and exact fixture cleanup remain; a later explicitly initiated read-only diagnostic and the fresh complete `active-03` run succeeded. This intermittent native failure is not represented as fixed.
- Unsaved active workbook binding, non-rebinding attached save-copy/PDF, stable numeric shape IDs, raw COM worksheet/range return objects, general worksheet deletion, advanced conditional-format types beyond cell-value/expression and the remaining shape/page enum variants are not claimed as complete parity. Known unique shape names provide stable targets for the verified shape edits.
- Native accurate inspection activates owned worksheets. `visible=False` does not promise that Excel's application/window stays hidden. No global visibility settings are changed.
- A real user-facing UI edit→Undo→save/discard acceptance and Windows native parity are separate tasks. This evidence does not claim either.
