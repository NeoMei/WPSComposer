# macOS WPS Capability Parity and Generation Design

## Status

Approved by the user's standing instruction to auto-confirm designs and choose
the recommended option. This document defines the macOS parity roadmap and the
first implementation slice: public document generation.

The target is behavioral parity with the public Windows APIs, not identical
implementation internals or byte-for-byte output. macOS continues to use WPS
JSAPI; Windows continues to use COM.

## Current state

The installed macOS WPS 12.1.26035 has already passed repeated real gates for
converting existing `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, and `.pptx`
files to PDF. The runtime provides private WPS-container staging, authenticated
loopback commands, registration rollback, owned-process cleanup, typed errors,
atomic publication, and an inter-process lock.

The remaining public APIs are Windows-only because the current renderers create
new COM documents and call `SaveAs`. On macOS, Writer `Document.SaveAs2` opens
the native save panel even when authorization prompts have already been
answered. Presentation and Spreadsheet creation passed the original probe, but
public generation stayed disabled because the all-format gate was incomplete.

## Scope decomposition

Capability parity is delivered as four independently gated slices:

1. **Generation parity:** `generate()` produces validated DOCX, PPTX, XLSX,
   and PDF through WPS on macOS.
2. **Existing-file parity:** `inspect(path)` and `edit(path, ...)` support the
   same addressable elements and patch contract as Windows.
3. **Active-document parity:** `attach_active()`, `inspect(path=None)`, and
   `edit(path=None, ...)` operate on the user's current WPS document without
   closing or replacing it.
4. **Capability-depth parity:** close remaining per-component gaps in styles,
   sections, headers/footers, charts, notes, geometry, and selection handling.

Each slice gets its own implementation plan and real WPS acceptance gate. A
later slice cannot weaken rollback, path isolation, or non-interactive behavior
established by an earlier slice.

This specification fully designs slice 1 and establishes interfaces that
slices 2-4 will reuse.

## Alternatives considered

### A. WPS-native template clone plus typed operation plan — selected

Bundle minimal WPS-produced DOCX, PPTX, and XLSX templates. For every request,
copy the matching template into the private WPS container, open the existing
copy as writable, clear its document-specific content, apply a closed typed
operation plan, and call native `Save()` in place. The Python host validates
and atomically publishes the resulting file.

This avoids `SaveAs2`, keeps WPS responsible for layout and serialization, and
uses the already-proven staging boundary. Sending one operation plan per
artifact avoids a fragile high-volume RPC conversation.

### B. One remote RPC for every COM composer method

Expose a long-lived document handle and mirror every `WriterComposer`,
`SheetComposer`, and `SlideComposer` call over the bridge. This closely mirrors
the Windows object model but creates hundreds of round trips, complicated
session recovery, and partially mutated documents when a late call fails.

### C. Generate OOXML outside WPS and use WPS only for PDF/render checks

Use `python-docx`, `openpyxl`, and `python-pptx` on macOS. This would be easier
to implement but would not preserve the project's defining guarantee that the
real WPS engine owns layout and final Office serialization.

## Parity contract

For the same public call, Windows and macOS must agree on:

- accepted inputs, defaults, overwrite behavior, and return shape;
- document structure: sections, headings, paragraphs, tables, lists, images,
  slides, sheets, formulas, charts, notes, headers, footers, and page fields;
- supported design presets and layout templates;
- stable error categories and absence of partial final output;
- inspection target identifiers and patch semantics in later slices.

Parity does not mean identical ZIP bytes or pixel-identical rendering. WPS can
select different fonts or line breaks across operating systems. Real-gate
comparison therefore checks package structure, extracted text, element counts,
formula/chart presence, PDF page count, and absence of clipped or empty pages.

## Architecture

### Renderer-neutral operation plans

Add immutable JSON-compatible plans for Writer, Spreadsheet, and Presentation.
An operation has a fixed name, typed arguments, and a monotonically increasing
index. Arbitrary property paths, arbitrary JavaScript, `eval`, and dynamic
method names are forbidden.

Examples:

```json
{"op": "writer.add_paragraph", "text": "Summary", "style": "Heading 1"}
{"op": "sheet.write_table", "sheet": 1, "startRow": 1, "startCol": 1,
 "values": [["Item", "Amount"], ["A", 10]]}
{"op": "slide.add_textbox", "slide": 2, "text": "Result",
 "left": 48, "top": 80, "width": 600, "height": 60}
```

The existing Markdown parser, document model, presets, numbering, and layout
decisions remain in Python. Existing renderers gain an injected composer
factory. Windows uses the existing COM composers. macOS uses recording
composers with the same renderer-facing methods; they validate arguments and
record operations instead of opening COM.

Only methods actually used by public renderers are included initially. New
operations are added with validation and executable add-in tests, never via a
generic escape hatch.

### Native templates

The repository contains one minimal, WPS-produced template for each OOXML
format. Every template must pass `validate_office_package`, contain no personal
metadata beyond a fixed WPSComposer marker, and be documented with its WPS
version and SHA-256 digest.

At runtime the host copies the template to a private session filename:

```text
.../WPSComposer/session-*/generated.docx
.../WPSComposer/session-*/generated.pptx
.../WPSComposer/session-*/generated.xlsx
```

The add-in opens that existing file as writable and calls the component's
native `Save()` after applying all operations. It never calls `SaveAs`, never
receives the user's external output path, and never handles an arbitrary input
path.

Writer first deletes the template body while retaining the native document and
style container. Spreadsheet removes extra sheets and clears the retained first
sheet. Presentation removes all template slides. The plan then builds the
requested artifact from this known baseline.

### Command protocol

Add three typed bridge methods:

- `generate_writer_document`
- `generate_spreadsheet_workbook`
- `generate_presentation_deck`

Each command contains only the staged template path, a fixed format identifier,
and the validated operation plan. The bridge enforces component routing and a
maximum serialized plan size. Each add-in validates every operation again
before mutating WPS state.

The result contains the staged artifact path, applied operation count, and a
component capability report. A returned path outside the active staging
session is a protocol error.

### PDF generation

macOS `generate(format="pdf")` reuses the Writer operation plan:

1. clone and populate the staged DOCX template;
2. save the DOCX in place and validate it;
3. reopen or retain that staged document;
4. export a staged PDF with `ExportAsFixedFormat`;
5. validate and publish only the requested PDF.

The temporary DOCX remains private and is removed with the session.

### Public routing

`generate()` keeps its current signature. After parsing and preset resolution,
it selects the backend by platform:

- Windows: existing COM renderer path;
- macOS: build an operation plan and call the staged JSAPI generation backend;
- other platforms: stable backend-unavailable error.

No public macOS format is enabled merely because unit tests pass. A per-format
gate constant is enabled only after two consecutive installed-WPS runs produce
valid artifacts without a save panel or new macOS permission prompt.

## Failure handling

All failures occur before final publication whenever possible. Stable public
codes include:

- `MACOS_GENERATION_GATE_NOT_PASSED`
- `OPERATION_PLAN_INVALID`
- `MACOS_CAPABILITY_UNAVAILABLE`
- `INTERACTIVE_INPUT_REQUIRED`
- `GENERATION_COMMAND_FAILED`
- `STAGED_ARTIFACT_INVALID`
- `ARTIFACT_PUBLISH_FAILED`
- `FINAL_ARTIFACT_INVALID`
- `REGISTRATION_RESTORE_FAILED`

Unknown WPS or vendor codes normalize to `GENERATION_COMMAND_FAILED`. Messages
redact private staging paths. `DisplayAlerts` is restored and opened documents
are closed without saving in nested `finally` blocks. A failed in-place save
never publishes the template or a partially mutated package.

Registration recovery failure retains its durable recovery directory and
blocks cleanup of that evidence. Concurrent Phase 0, conversion, generation,
inspection, and editing sessions share the same inter-process WPS runtime lock.

## Security and permission model

- All WPS file access stays within its private application container.
- The Python host alone reads source Markdown and publishes final artifacts.
- The add-in accepts authenticated loopback commands from the current random
  token and fixed origins only.
- No AppleScript, Computer Use, Accessibility, Input Monitoring, screen
  capture, Full Disk Access, keyboard simulation, or save-panel automation is
  part of the backend.
- Images are copied to staging or served from a fixed authenticated loopback
  resource; arbitrary remote URLs are rejected.
- Plans have operation-count, string-length, table-size, and image-size limits.

## Testing strategy

### Deterministic tests

- plan schema validation, serialization limits, and rejection of unknown ops;
- renderer recording tests for representative Markdown and every preset;
- executable Node tests for every add-in operation and all cleanup paths;
- template digest and OOXML structural validation;
- output publication races, overwrite behavior, path redaction, and typed
  error normalization;
- platform routing tests proving Windows behavior is unchanged;
- clean plugin installation proving templates and production JS dependencies
  are present.

### Installed-WPS gates

The feasibility gate runs before the full renderer port:

1. copy each native template into a fresh private session;
2. open it writable, make one visible mutation, call `Save()` in place, close;
3. validate the same path from Python;
4. repeat twice with fresh sessions;
5. confirm no save panel, permission prompt, registration residue, staging
   session, or unrelated-process termination.

If Writer in-place `Save()` fails, generation remains disabled and no GUI
automation fallback is introduced. The failure report must distinguish an
interactive prompt from an unsupported API.

After the feasibility gate, the full acceptance matrix generates twice:

- a styled multi-section DOCX with TOC, table, image, header/footer, and page
  number;
- a multi-sheet XLSX with formulas, merged cells, borders, frozen panes,
  conditional formatting, and a chart;
- a multi-slide PPTX with text, bullets, image, table, notes, and a preset;
- a Writer-derived PDF from the same representative Markdown.

Every Office artifact passes ZIP validation and semantic inspection. Every PDF
passes signature, page-count, and extracted-text checks. No output is enabled
until its two consecutive runs pass.

## Later slices

### Existing-file inspect and edit

Reuse staging and the operation protocol. Inspection returns the current public
snapshot tree from read-only staged copies. Editing applies ordered typed
patches to a writable staged copy, calls `Save()` in place, validates it, and
atomically publishes either to `output` or back over the original when the
caller explicitly requests in-place behavior. Optional PDF export occurs before
the Office document is closed.

Target identifiers stay compatible with Windows (`paragraph:N`, `table:N`,
`sheet:N/range:A1`, `slide:N/shape:N`, and their existing nested forms).

### Active-document attachment

Temporary add-in registration cannot reliably attach to an already-running WPS
process. This slice therefore uses an explicitly installed dormant
WPSComposer add-in shell. The shell performs no work when no authenticated
loopback bridge is present. `attach_active(kind)` starts a short-lived bridge,
the already-loaded add-in registers the active component, and commands operate
on `ActiveDocument`, `ActiveWorkbook`, or `ActivePresentation` without opening,
closing, or saving unless requested.

Persistent registration must merge only the WPSComposer entry, preserve other
add-ins byte-for-byte where possible, support transactional uninstall, and pass
a separate security review before activation.

### Capability-depth parity

Maintain a machine-readable matrix mapping every public composer method to
Windows COM evidence, macOS JSAPI evidence, tests, and real-gate status. Missing
JSAPI features return `MACOS_CAPABILITY_UNAVAILABLE`; they are never silently
ignored. Equivalent native sequences are allowed when WPS exposes different
object models across platforms.

## Completion criteria

Slice 1 is complete when all four public `generate()` formats pass deterministic
tests and two consecutive real WPS gates, the Windows suite remains green, and
the macOS documentation no longer labels generation as globally NO-GO.

Overall macOS parity is complete when generation, existing-file inspection and
editing, active-document attachment, and every capability marked supported in
the public matrix pass their respective real gates. Platform-specific WPS
rendering differences and explicitly unsupported vendor APIs remain documented
rather than hidden.
