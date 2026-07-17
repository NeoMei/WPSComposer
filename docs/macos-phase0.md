# macOS WPS JSAPI Phase 0 evidence

## Decision

**NO-GO for Phase 1 on the tested WPS version.** Do not route the public
`generate()` API to the macOS backend.

The installed Mac WPS can load authenticated JS add-ins and can serialize the
Presentation and Spreadsheet smoke files. Writer exposes document creation,
formatting, image, `SaveAs2`, and `ExportAsFixedFormat` APIs, but it did not
complete DOCX or PDF serialization. A minimal Writer experiment containing
only one text paragraph failed identically, so the result is not caused by the
table or image fixture.

Phase 1 may be reconsidered only after a WPS update or a documented Writer API
sequence produces and validates all four independent outputs. The production
Windows COM backend remains unchanged.

## Tested environment

| Item | Value |
|---|---|
| Test time | 2026-07-17 21:46 CST (+0800) |
| macOS | 26.5.2 (25F84) |
| WPS Office | 12.1.26035 |
| Node.js | 24.14.0 (Codex bundled runtime) |
| `wpsjs` | 2.2.3 |
| `wps-jsapi-declare` | 2.2.0 |
| npm audit | 0 known vulnerabilities across 152 dependencies |

The real probe command was:

```bash
.venv/bin/python -m skills.WPSComposer.scripts.macos_probe \
  --node /Users/neomei/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node \
  --output-dir build/macos-phase0-save-events \
  --timeout 45
```

Primary failure report:
`build/macos-phase0-save-events/phase0-report.json`.

Two additional runs isolate and corroborate the result:

- `build/macos-phase0-minimal-writer/phase0-report.json` removes Writer tables
  and images and still times out waiting for `smoke.docx` to save.
- `build/macos-phase0-http-image-clean/phase0-report.json` records the complete
  capability map and independently validated PPTX/XLSX artifacts.

All paths above are runtime evidence under the gitignored `build/` directory.

## Artifact evidence

| Format | Actual test path | Size | WPS API | Validation |
|---|---|---:|---|---|
| DOCX | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-save-events/smoke.docx` | missing | `Document.SaveAs2(path, 12)` | **Failed:** no file and no `FileAfterSave` event within 15 seconds |
| PPTX | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-http-image-clean/smoke.pptx` | 73,006 | `Presentation.SaveAs(path, 24)` | **Passed:** ZIP signature, size, and `ppt/presentation.xml` verified |
| XLSX | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-http-image-clean/smoke.xlsx` | 13,181 | `Workbook.SaveAs(path, 51)` | **Passed:** ZIP signature, size, and `xl/workbook.xml` verified |
| PDF | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-save-events/smoke.pdf` | missing | `Document.ExportAsFixedFormat(path, 17, ...)` | **Failed:** Writer could not serialize its independent temporary source document; earlier direct export also returned without a PDF file |

The all-four artifact structure check therefore fails by design. Partial files
are not reported as successful outputs.

## Capability matrix

Classification reflects observed behavior, not merely the presence of a JSAPI
member.

| Component | Capability | Class | Evidence/detail |
|---|---|---|---|
| Writer | `writer.documents` | native | `Application.Documents.Count` succeeded |
| Writer | `writer.font_enumeration` | native | `Application.FontNames.Count` succeeded; the probe did not resolve a concrete requested-font list |
| Writer | `writer.template_enumeration` | native | `Application.Templates.Count` succeeded |
| Writer | `writer.image_source` | mapped | Allowed image served from the fixed Writer loopback origin |
| Writer | `writer.font_ascii` | mapped | `Font.NameAscii` accepted as the Latin-font mapping |
| Writer | `writer.tables` | native | `Document.Tables.Add` succeeded |
| Writer | `writer.image_return` | native | HTTP image insertion returned an `InlineShape` |
| Writer | `writer.image_api` | native | `Document.InlineShapes.AddPicture` accepted the loopback URL |
| Writer | `writer.images` | native | Inline image creation and sizing succeeded before save |
| Writer | `writer.save_docx` | unsupported | `SaveAs2` is exposed but produced no file and no completion event, including for a one-paragraph document |
| Writer | `writer.export_pdf` | unsupported | No standalone PDF was produced; the independent Writer source save also failed |
| Presentation | `presentation.presentations` | native | Presentations collection accessible |
| Presentation | `presentation.slides` | native | `Slides.Add` succeeded |
| Presentation | `presentation.shapes` | native | Shapes collection accessible |
| Presentation | `presentation.template_resolution` | mapped | `Presentations.Add` used the WPS-native blank presentation |
| Presentation | `presentation.create` | native | `Presentations.Add` succeeded |
| Presentation | `presentation.text` | native | `Shapes.AddTextbox` succeeded |
| Presentation | `presentation.image` | native | `Shapes.AddPicture` succeeded |
| Presentation | `presentation.table` | native | `Shapes.AddTable` succeeded |
| Presentation | `presentation.save_pptx` | native | Format 24 produced a valid 73,006-byte PPTX |
| Spreadsheet | `spreadsheet.workbooks` | native | Workbooks collection accessible |
| Spreadsheet | `spreadsheet.range` | native | `Worksheet.Range` succeeded |
| Spreadsheet | `spreadsheet.charts` | native | `Worksheet.ChartObjects` accessible |
| Spreadsheet | `spreadsheet.create` | native | `Workbooks.Add` succeeded |
| Spreadsheet | `spreadsheet.values` | native | Cell `Value2` assignment succeeded |
| Spreadsheet | `spreadsheet.formulas` | native | `Range.Formula` and `FillDown` succeeded |
| Spreadsheet | `spreadsheet.formatting` | native | Font, fill, borders, and `AutoFit` succeeded |
| Spreadsheet | `spreadsheet.chart` | native | `ChartObjects.Add` and `SetSourceData` succeeded |
| Spreadsheet | `spreadsheet.save_xlsx` | native | Format 51 produced a valid 13,181-byte XLSX |
| Spreadsheet | `spreadsheet.template_resolution` | mapped | `Workbooks.Add` used the WPS-native default workbook |

## Safety and rollback evidence

The registration path was absent before every real run:

```text
~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml
ABSENT
```

It was also absent after success, command failure, timeout, and an interrupted
run. Probe-created WPS processes were identified as the exact PID difference
from the pre-run WPS process snapshot and terminated without touching
pre-existing instances. Temporary recovery directories were removed only after
registration verification succeeded.

If a crash bypasses normal cleanup:

1. Stop only the Node processes whose command points into
   `macos/wps-jsapi-probe/node_modules/wpsjs`; do not quit an unrelated WPS
   session.
2. Open the `WPS registration recovery: ...` directory printed by the probe and
   inspect `registration.json`.
3. When `existed` is `true`, copy `publish.xml.original` to a temporary file in
   the registration directory and atomically rename it to `publish.xml`. When
   `existed` is `false`, remove only the probe-created `publish.xml`.
4. Reopen only the WPS instances used for the probe and verify the registration
   file matches its pre-run state.

## Re-running the gate

Install the pinned dependencies with `npm ci` in
`macos/wps-jsapi-probe`, then run the command above into a new, empty output
directory. A future **GO** requires valid DOCX, PPTX, XLSX, and PDF artifacts,
plus a supported Writer font/template resolution path. API availability alone
is insufficient.
