# macOS Word document sessions

Native host: Microsoft Word 16.112.3 on macOS. This is a bounded session adapter
gate, not a claim of complete Microsoft/WPS parity.

`native-09/report.json` passed all eight checks. Its `source.docx`,
`edited.docx`, `edited.pdf`, and `moved.docx` are native Word artifacts. Raw
AppleScripts and responses are retained under `runtime/` and `setup-runtime/`.

Verified through the new `MacWordSession` class:

- Open a validated DOCX private copy; inspect paragraph text/font/paragraph
  formatting, tables/cells, shapes and section page setup.
- Patch paragraph font name/size/bold/color and indentation/spacing, then read
  the properties back from Word.
- Insert paragraphs, clone formatted paragraphs, create a native table,
  fill/alignment in a native cell, native text box and inline image.
- Move a formatted paragraph, remove a paragraph, replace paragraph text while
  preserving its terminating boundary.
- Save a DOCX copy without rebinding the open task document, export PDF, close,
  reopen, and verify saved formatting/object counts. The source hash remains
  unchanged. PDF text contains the document and text-box text.
- Preserve a synthetic unsaved document throughout; close it only after exact
  name/path/content/saved-state assertions. No global Quit or clipboard calls.

The adapter uses the same `WordJobLock` as generation/conversion and one
600-second session deadline, with each AppleEvent capped at 60 seconds. A
strict completion envelope is required even when a command returns no data.
Private document basenames include a UUID. `save` and `save_copy` require new
output paths; explicit `save_current` verifies the original source hash before
saving and again before replacement. PDF output requires a new `.pdf` path.
Uncertain
AppleEvent timeouts or close failures retain staging and quarantine later jobs.
Ordinary native operation errors retain diagnostics even after successful task
document closure. Portable tests cover these paths; this session gate did not
inject an actual hung AppleEvent.

Native-09 reruns all eight checks after the deadline, completion-envelope,
unique-basename, and publication-safety changes. Source-conflict and publication
failure paths are additionally covered by portable transport-injection tests.

## Explicit limitations

Word returned `null` for `id of active window` in the native probes. Owned file
sessions use their unique private path, with native document-name resolution and
an exact path check before every operation. `attach_active()` fails before any
mutation on this host because a stable native window identity is unavailable;
the synthetic sentinel remains unmodified. No title-only attachment fallback is
enabled. Attached non-rebinding save-copy/PDF is also unsupported.

The following baseline facets remain incomplete: floating-object/table/inline
image move and clone, non-DOCX input/output, exotic underline and paragraph enum
values, line dash styles, section paper-size enum, and arbitrary insertion
properties. These are rejected explicitly, not silently omitted. Native testing
of every supported formatting key, merged-cell edits, range selection, and the
public installed-module orchestration is still required for full parity.

Stable `w14:paraId` targets are exposed only while the private saved document and
live paragraph counts align, before structural changes. Structural changes
return `reinspect_required`; positional inspection remains available afterward.

## Failure ledger

- `native-01`: result variable `rows` collided with Word's table-row property;
  failure occurred before document creation. Renamed the script variable.
- `native-02`: new document reference became stale after SaveAs. Rebound the
  fixture by exact saved path before closing it; exact-path recovery retained.
- `native-03`: Word returned null window ID; collection iteration also needed
  indexed native access. `recovery/` records closing only the private path and
  the exact synthetic sentinel, then resolving the task's quarantine marker.
- `native-04`/`native-05`: COM-style expansion of `font.name` to every script
  family rejected `east asian name = Arial`. The Word adapter now sets native
  `name` directly. Native-04 diagnostics were deleted after successful close;
  a regression fix retains operation-error evidence, visible in native-05.
- `native-06`: table-cell iteration produced an unusable range. Indexed native
  cell access fixed the snapshot, with a contract regression test.
- `native-07`: operations/reopen/source checks passed, but the fixture text box
  at its default location overlapped body text; PDF extraction interleaved the
  glyphs. This failure is retained. Native-08 places the box below body content
  and passes the PDF check; no backend success criterion was weakened.

Run explicitly, with exclusive native Word access:

```sh
python fixtures/microsoft_parity/macos_word_sessions.py --output-dir NEW_DIRECTORY
```

The runner refuses an existing output directory. This evidence does not include
GUI edit/Undo acceptance or native Windows sessions.
