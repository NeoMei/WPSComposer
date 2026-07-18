# WPS Sandbox Staging and Office-to-PDF Conversion Design

## Status

Approved approach: stage WPS inputs and generated artifacts inside the WPS
application container, then publish validated bytes to the requested
destination from the Python host. Add one public cross-platform, single-file
Office-to-PDF conversion API.

This design refines the Phase 0 probe and adds public conversion support for
`doc`, `docx`, `xls`, `xlsx`, `ppt`, and `pptx`. It does not enable the public
macOS `generate()` backend.

### Installed-Mac acceptance result (2026-07-18)

The two gates have different outcomes on macOS 26.5.2 with WPS 12.1.26035:

- **Office-to-PDF conversion: GO.** Two consecutive runs converted all six
  suffixes through the typed public backend without a save dialog or new
  permission interaction. The output directories were
  `/tmp/wpscomposer-conversion-gate-run-1-20260718` and
  `/tmp/wpscomposer-conversion-gate-run-2-20260718`.
- **Public document generation: NO-GO.** The staged Phase 0 run at
  `/tmp/wpscomposer-phase0-staged-run-1-20260718/phase0-report.json` published
  valid PPTX and XLSX, but Writer `SaveAs2` still opened the native save path
  and timed out for DOCX and the temporary DOCX used by standalone PDF
  generation. No broader permission was requested and public `generate()`
  remains disabled on macOS.

The conversion outputs were identical in size across both runs:

| Source | Component | PDF bytes | Pages |
|---|---|---:|---:|
| `.doc` | Writer | 1,052 | 1 |
| `.docx` | Writer | 1,051 | 1 |
| `.xls` | Spreadsheet | 321,446 | 1 |
| `.xlsx` (two visible worksheets) | Spreadsheet | 199,724 | 2 |
| `.ppt` | Presentation | 1,212 | 1 |
| `.pptx` | Presentation | 1,212 | 1 |

The legacy fixtures are valid Compound File Binary documents bundled with WPS
itself. The modern two-sheet XLSX fixture declares both sheets visible, and
the resulting two-page PDF proves workbook-level export. Every run restored
registration, removed its WPS-container session, and left the pre-existing WPS
process untouched. `MACOS_CONVERSION_ENABLED` is therefore enabled while the
unrelated macOS generation gate stays closed.

## Problem

Mac WPS is sandboxed and declares user-selected file access rather than broad
arbitrary-path write access. When Writer receives an external `SaveAs2` path,
it opens the macOS save panel instead of completing the headless save. The save
panel causes repeated TCC requests, including input monitoring, screen capture,
Contacts, Calendar, and Reminders checks. The probe then times out because no
Writer save-completion event is delivered.

Presentation and Spreadsheet have saved directly in the tested version, but
allowing any WPS component to receive an arbitrary external output path leaves
the same prompt behavior possible. The artifact transport should therefore be
uniform across all four formats.

## Goals

- Prevent WPS from opening a native save panel during the probe.
- Avoid Computer Use, AppleScript, Accessibility, keyboard, and screen-capture
  automation.
- Keep WPS responsible for DOCX, PPTX, XLSX, and PDF serialization.
- Publish only validated artifacts to the requested output directory.
- Preserve registration rollback, process ownership, typed commands, and the
  behavior of existing `generate()`, `edit()`, and PDF-editing APIs.
- Make consecutive probe runs non-interactive after any unavoidable first-use
  macOS permission has already been granted.
- Provide one public `convert_to_pdf()` entrypoint on Windows and macOS.
- Export every visible Excel worksheet into one PDF.

## Non-goals

- Granting Full Disk Access or broad Automation permission to WPS or Codex.
- Persisting user-selected security-scoped bookmarks.
- Adding a non-WPS serializer or format fallback.
- Enabling the production macOS backend before all four smoke outputs pass.
- Batch, list, directory, or recursive conversion.
- Password entry, dialog automation, or conversion of macro-enabled formats.

## Architecture

### Staging location

Each run creates a random private directory below:

```text
~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer/{random-session-id}/
```

The directory is mode `0700`. Probe artifact filenames remain fixed and typed:
`smoke.docx`, `smoke.pptx`, `smoke.xlsx`, `smoke.pdf`, and the private
`smoke-pdf-source.docx` used only by the standalone PDF command. Conversion
uses a random collision-resistant filename while preserving the validated
source extension and a paired `.pdf` output name.

The Python runtime already reads and restores WPS registration in this
container. macOS may require a one-time Codex App Data decision for that
container, but the design must not trigger a new WPS save-panel decision on
every run.

### Command flow

1. The runner validates the requested final output directory and refuses to
   overwrite any existing target.
2. The runtime allocates the private WPS-container staging directory.
3. The runner sends only staged `outputPath` values to the closed set of WPS
   smoke commands. Add-ins never receive the final external destination.
4. WPS serializes each artifact in its own container. Writer waits for the
   documented `FileAfterSave` event; the other components retain their current
   save APIs.
5. The Python host validates the staged signature, minimum size, expected
   package member, and the command-reported staged path.
6. The host publishes the staged artifact to the requested final path using a
   temporary file in the destination directory, flushes and `fsync`s it, then
   atomically replaces the final target.
7. The final file is validated again. Reports expose only the requested final
   path, never the WPS staging path or the PDF source document.

The destination-side temporary file makes publication work even when the WPS
container and requested destination are on different filesystems. `os.replace`
is used only for the final same-directory rename.

### Public conversion API

The package exports:

```python
def convert_to_pdf(
    source: str,
    output: Optional[str] = None,
    *,
    overwrite: bool = False,
) -> str:
    ...
```

The source extension is matched case-insensitively:

- Word: `doc`, `docx`
- Excel: `xls`, `xlsx`
- PowerPoint: `ppt`, `pptx`

When `output` is omitted, the destination is the source path with its suffix
replaced by `.pdf`. The destination must have a `.pdf` suffix. Existing output
raises `FileExistsError` unless `overwrite=True`. Success returns the absolute
final PDF path.

Missing source files raise `FileNotFoundError`; unsupported extensions and an
invalid output suffix raise `ValueError`. WPS runtime and conversion failures
raise `ConversionError`, which carries a stable code, source path, component,
backend, and readable message.

### Conversion data flow

The public facade validates paths and routes by extension and platform.

On Windows, the adapter opens the existing Office file through the matching
composer and asks WPS or Microsoft Office to export a host-private staged PDF.
Writer uses `Document.ExportAsFixedFormat`, Presentation uses
`Presentation.SaveAs` with format 32, and Excel uses workbook-level
`ExportAsFixedFormat` so every visible worksheet is included. The common host
publisher validates and atomically installs the PDF at the final destination.

On macOS:

1. The Python host copies the validated source into the private WPS-container
   staging directory. WPS never receives the original external source path.
2. The bridge routes one of three typed methods to its fixed component:
   `convert_writer_pdf`, `convert_workbook_pdf`, or
   `convert_presentation_pdf`.
3. The add-in opens the staged source, exports a staged PDF with the native WPS
   API, waits for the relevant completion event, and closes only that source
   document without saving changes.
4. The host validates and atomically publishes the staged PDF using the same
   destination-side publisher as the probe.

Password prompts, damaged-file dialogs, compatibility dialogs, and any other
interactive state fail with `INTERACTIVE_INPUT_REQUIRED` or the nearest
component-specific error. The implementation never clicks or dismisses them.

### Path policy

The existing path policy gains the resolved per-run staging directory as an
allowed root. The requested output directory remains a separate allowed root.
The runner constructs all staged filenames; no path supplied by the add-in or
input document can select an arbitrary location.

The bridge method allow-list remains closed and expands only with the three
conversion operations:

- `probe_capabilities`
- `smoke_docx`
- `smoke_pdf`
- `smoke_pptx`
- `smoke_xlsx`
- `convert_writer_pdf`
- `convert_workbook_pdf`
- `convert_presentation_pdf`

### Lifecycle and cleanup

The runtime owns the staging directory. On success or ordinary failure it
removes staged artifacts and the session directory after final publication or
reporting. Registration recovery remains higher priority: if registration
restore verification fails, the recovery material is retained and reported,
while artifact staging can still be removed independently.

Only WPS processes absent from the pre-run exact-process snapshot are stopped.
No pre-existing WPS process or document is closed.

The conversion facade opens one source document per call and closes it in
`finally`. It never modifies or saves the source. The Windows adapter follows
the same ownership rule.

## Error model

The report distinguishes these cases:

- `STAGING_SAVE_FAILED`: WPS did not create a valid staged artifact.
- `STAGED_ARTIFACT_INVALID`: WPS created bytes with an invalid signature,
  package structure, or size.
- `ARTIFACT_PUBLISH_FAILED`: validated staged bytes could not be copied or
  atomically installed at the requested destination.
- `FINAL_ARTIFACT_INVALID`: destination bytes failed post-publication
  validation.
- `INTERACTIVE_INPUT_REQUIRED`: WPS requires a password or modal user action.
- `NO_VISIBLE_WORKSHEETS`: an Excel workbook has no visible sheet to export.
- `CONVERSION_COMMAND_FAILED`: the native component export returned an error or
  timed out.

A failed publication never reports an artifact as successful. Destination
temporary files are removed in `finally` cleanup.

## Reporting

Successful artifact entries continue to contain the requested format, final
absolute path, and final size. The report adds:

```json
{
  "artifactTransport": "wps-container-staging"
}
```

Staging paths, bearer tokens, and the temporary PDF source path are omitted.
Capabilities continue to describe WPS rendering and serialization behavior;
the transport is classified separately rather than misreported as a native
WPS formatting feature.

The public conversion call returns only the final absolute path. Internal
staging paths are available to diagnostic logs only and are never included in
the successful return value or a user-facing error message.

## Testing strategy

Implementation follows test-driven development.

Platform-independent tests prove:

- staging directories resolve below the exact WPS container root and use a
  random per-run child;
- all four WPS commands receive staged paths, not final external paths;
- a staged artifact is validated before publication and again afterward;
- publication uses a destination-local temporary file and atomic rename;
- publication refuses existing targets and cleans partial temporary files;
- reports contain final paths and `wps-container-staging`, with no staging or
  PDF-source path leakage;
- staging cleanup occurs on success, command failure, publication failure, and
  registration rollback failure;
- the typed method allow-list remains closed, and all existing Windows/public
  API tests remain passing;
- public API export and extension routing for all six accepted input suffixes;
- default output naming, explicit output, `overwrite=False`, and
  `overwrite=True` behavior;
- macOS source copying and PDF output remain inside the per-run WPS staging
  directory until publication;
- Windows routes Writer, workbook, and Presentation exports to their correct
  native APIs;
- Excel conversion invokes a workbook-level export rather than the existing
  first-worksheet helper;
- source documents are closed without modification on success and failure;
- password, invalid source, no-visible-sheet, and native conversion failures do
  not publish partial PDFs;
- the existing `generate()`, `edit()`, and PDF-editing tests remain unchanged.

The installed-Mac acceptance run must then prove:

1. DOCX, PPTX, XLSX, and standalone PDF all pass independent structural
   validation.
2. Registration is identical before and after the run.
3. Two consecutive runs complete without a WPS save panel or new permission
   prompt.
4. No Computer Use, AppleScript, Accessibility, or screen-capture helper is
   invoked.
5. `doc`, `docx`, `xls`, `xlsx`, `ppt`, and `pptx` fixtures each convert to a
   structurally valid PDF through the public API.
6. The Excel fixture contains at least two visible worksheets and both are
   present in the resulting PDF.

If Writer still opens a save panel for a path inside its own container, the
gate remains NO-GO and the implementation records that result without asking
for broader permissions.

Public macOS conversion is enabled only when the installed-Mac conversion gate
passes. Windows conversion does not depend on the macOS gate.

## Alternatives considered

### Persistent user-selected folder bookmark

This could convert one user-approved directory into reusable access, but it
requires native security-scoped bookmark management and couples a headless
probe to a user-selected location. It is unnecessary while the WPS container
provides a private staging boundary.

### Broader macOS permissions

Full Disk Access, Accessibility, Automation, or screen-capture permission could
mask the save-panel problem but grants capabilities unrelated to document
serialization. This is rejected.

### Keep direct external output paths

This preserves the current code but continues to invoke the native save panel
and repeated authorization flow. It does not satisfy the non-interactive probe
requirement.
