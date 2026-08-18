# macOS WPS JSAPI Phase 0 evidence

## Decision

**GO for document generation on macOS (supersedes the 2026-07-17 NO-GO).** On
2026-07-25, `generate()` was routed through the gated macOS JSAPI backend on
the same installed Mac (WPS 12.1.26035) and all four formats serialized: docx,
xlsx, pptx, and pdf each produced a valid package via `generate_macos`. The
earlier Writer `SaveAs2` failure no longer reproduces, so
`MACOS_GENERATION_ENABLED` is open for all four formats. The original NO-GO
analysis and artifact table are retained below for provenance.

**GO for the Phase 0 smoke probe (2026-07-26 fix).** The smoke `smoke_docx` /
`smoke_pdf` commands previously timed out at `waitForFileAfterSave`
(`writer.js`) because the `FileAfterSave` JSAPI event is never emitted for
`SaveAs2` on this WPS build (it is emitted for `ExportAsFixedFormat`, which is
why `convert_to_pdf` always worked). The two `SaveAs2` call sites now call
`SaveAs2` directly — mirroring the production `generateWriterDocument` path
(`Save()` then `Close()`, host validates via filesystem poll) — and a full
probe run produces all four smoke artifacts: docx, pptx, xlsx, pdf. See the
"2026-07-26 smoke re-run" table below.


**GO for existing Office-to-PDF conversion.** On 2026-07-18, two consecutive
installed-Mac runs converted `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, and
`.pptx` through the container-staged typed backend. The two-visible-sheet XLSX
produced a two-page PDF, every output passed `pdfinfo`, registration was
restored, and no staging session remained. Public `convert_to_pdf()` is enabled
on macOS.

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
| Test time | 2026-07-17 21:52 CST (+0800) |
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
  --output-dir build/macos-phase0-after-manual-authorization \
  --timeout 60
```

Primary failure report:
`build/macos-phase0-after-manual-authorization/phase0-report.json`.

This control run was performed after the user had completed the visible macOS
authorization prompts. It still produced and validated PPTX and XLSX while
Writer failed both new-document saves with the same completion-event timeout. The result
therefore is not attributed to a prompt being approved too late. Recent TCC
logs showed an Apple Events denial for the optional Codex Computer Use helper,
not a WPS Writer file-access denial; that denial affected automated window
inspection only and was not used by the probe.

Two additional runs isolate and corroborate the result:

- `build/macos-phase0-minimal-writer/phase0-report.json` removes Writer tables
  and images and still times out waiting for `smoke.docx` to save.
- `build/macos-phase0-http-image-clean/phase0-report.json` records the complete
  capability map and independently validated PPTX/XLSX artifacts.

All paths above are runtime evidence under the gitignored `build/` directory.

## 2026-07-26 smoke re-run (post SaveAs2-event fix)

After dropping the `FileAfterSave` event wait from the two `SaveAs2` call sites
in `macos/wps-jsapi-probe/addin/writer.js` (`saveDocx`, `savePdf`'s source
save), a full `macos_probe` run on the same installed Mac produced all four
smoke artifacts with `status: passed` and zero failures:

| Format | Path | Size | Validation |
|---|---|---:|---|
| DOCX | `build/mac-recheck2/smoke.docx` | 11,885 | ZIP, 18 parts, `word/document.xml` present |
| PPTX | `build/mac-recheck2/smoke.pptx` | 73,474 | ZIP, valid |
| XLSX | `build/mac-recheck2/smoke.xlsx` | 13,179 | ZIP, valid |
| PDF | `build/mac-recheck2/smoke.pdf` | 745,649 | `%PDF-` header, valid |

Re-run command:

```bash
.venv/bin/python -m skills.WPSComposer.scripts.macos_probe \
  --node <node-path> --output-dir build/mac-recheck2 --timeout 120
```

## Artifact evidence


| Format | Actual test path | Size | WPS API | Validation |
|---|---|---:|---|---|
| DOCX | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-after-manual-authorization/smoke.docx` | missing | `Document.SaveAs2(path, 12)` | **Failed:** no file and no `FileAfterSave` event within 15 seconds |
| PPTX | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-after-manual-authorization/smoke.pptx` | 73,006 | `Presentation.SaveAs(path, 24)` | **Passed:** ZIP signature, size, and `ppt/presentation.xml` verified |
| XLSX | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-after-manual-authorization/smoke.xlsx` | 13,179 | `Workbook.SaveAs(path, 51)` | **Passed:** ZIP signature, size, and `xl/workbook.xml` verified |
| PDF | `/Users/neomei/项目/WpsComposer/.worktrees/macos-jsapi-phase0/build/macos-phase0-after-manual-authorization/smoke.pdf` | missing | `Document.ExportAsFixedFormat(path, 17, ...)` | **Failed for generation:** Writer could not first serialize the independent temporary source document. A later direct export of an existing staged DOC/DOCX succeeded and is the basis of `convert_to_pdf()` support. |

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
| Spreadsheet | `spreadsheet.save_xlsx` | native | Format 51 produced a valid 13,179-byte XLSX in the authorization-control run |
| Spreadsheet | `spreadsheet.template_resolution` | mapped | `Workbooks.Add` used the WPS-native default workbook |

## Safety and rollback evidence

The registration path was absent before every real run:

```text
~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons/publish.xml
ABSENT
```

The runtime restores the exact pre-run registration bytes after partial artifact
generation, command failure, timeout, and interruption. WPS is activated through
LaunchServices and is not force-terminated; the add-in closes only its activation
fixture, while the runtime reaps its own `wpsjs` server processes. Temporary
recovery directories are removed only after registration verification succeeds.

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
