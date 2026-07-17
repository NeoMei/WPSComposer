# macOS WPS Sandbox Staging Design

## Status

Approved approach: stage WPS-generated artifacts inside the WPS application
container, then publish validated bytes to the requested destination from the
Python host.

This design refines the Phase 0 probe only. It does not enable the public
macOS `generate()` backend.

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
  existing public API boundary.
- Make consecutive probe runs non-interactive after any unavoidable first-use
  macOS permission has already been granted.

## Non-goals

- Granting Full Disk Access or broad Automation permission to WPS or Codex.
- Persisting user-selected security-scoped bookmarks.
- Adding a non-WPS serializer or format fallback.
- Enabling the production macOS backend before all four smoke outputs pass.

## Architecture

### Staging location

Each run creates a random private directory below:

```text
~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer/{random-session-id}/
```

The directory is mode `0700`. Artifact filenames remain fixed and typed:
`smoke.docx`, `smoke.pptx`, `smoke.xlsx`, `smoke.pdf`, and the private
`smoke-pdf-source.docx` used only by the standalone PDF command.

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

### Path policy

The existing path policy gains the resolved per-run staging directory as an
allowed root. The requested output directory remains a separate allowed root.
The runner constructs all staged filenames; no path supplied by the add-in or
input document can select an arbitrary location.

The bridge method allow-list remains unchanged:

- `probe_capabilities`
- `smoke_docx`
- `smoke_pdf`
- `smoke_pptx`
- `smoke_xlsx`

### Lifecycle and cleanup

The runtime owns the staging directory. On success or ordinary failure it
removes staged artifacts and the session directory after final publication or
reporting. Registration recovery remains higher priority: if registration
restore verification fails, the recovery material is retained and reported,
while artifact staging can still be removed independently.

Only WPS processes absent from the pre-run exact-process snapshot are stopped.
No pre-existing WPS process or document is closed.

## Error model

The report distinguishes these cases:

- `STAGING_SAVE_FAILED`: WPS did not create a valid staged artifact.
- `STAGED_ARTIFACT_INVALID`: WPS created bytes with an invalid signature,
  package structure, or size.
- `ARTIFACT_PUBLISH_FAILED`: validated staged bytes could not be copied or
  atomically installed at the requested destination.
- `FINAL_ARTIFACT_INVALID`: destination bytes failed post-publication
  validation.

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
- the typed method allow-list and all Windows/public API tests remain
  unchanged.

The installed-Mac acceptance run must then prove:

1. DOCX, PPTX, XLSX, and standalone PDF all pass independent structural
   validation.
2. Registration is identical before and after the run.
3. Two consecutive runs complete without a WPS save panel or new permission
   prompt.
4. No Computer Use, AppleScript, Accessibility, or screen-capture helper is
   invoked.

If Writer still opens a save panel for a path inside its own container, the
gate remains NO-GO and the implementation records that result without asking
for broader permissions.

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
