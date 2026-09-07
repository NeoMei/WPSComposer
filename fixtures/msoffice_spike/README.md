# Native Microsoft Word feasibility probes

These standalone investigation runners exercise Microsoft Word itself. They do
not change WPSComposer production code, install add-ins, enable macros, or use WPS.
Python 3.9+ is required. Run from the repository root with a **new output directory**.
Existing directories are rejected, including empty ones, to prevent stale artifacts.

macOS requires an installed Microsoft Word and permission for the calling terminal
or Codex to automate Word. The runner uses its installed AppleScript dictionary:

```sh
python3 fixtures/msoffice_spike/mac_word.py --output-dir build/msoffice-spike/mac-run-01
```

For compilation without native execution, use a different new directory:

```sh
python3 fixtures/msoffice_spike/mac_word.py --output-dir build/msoffice-spike/mac-compile-01 --compile-only
```

Windows requires an interactive desktop, Microsoft Word and `pywin32`. Copy both
Python files together: the Windows runner imports the shared fixture content from
`mac_word.py`, which does not execute platform operations on import.

```powershell
python -m pip install pywin32
python fixtures/msoffice_spike/windows_word.py --output-dir build/msoffice-spike/windows-run-01
```

Each native run attempts Chinese body font 仿宋 at 12 points, first-line indent 24
points/two characters, built-in Heading 1/2/3 with sizes 16/15/15 points, a
document-owned multilevel numbering template, a native TOC and field refresh, an
82-row table with a repeating header, repagination, DOCX save/reopen and native
PDF export. These are capability probes: setters returning without exceptions
are only operation evidence. Inspect persisted DOCX XML and PDF layout/text before
claiming formatting, numbering, page counts or visual acceptance.

`result.json` contains the actual engine/version, individual operation states,
raw errors, snapshots and artifact presence. The Mac run additionally retains the
exact AppleScript, compiled script, fixture input, compile error stream, stdout
and stderr. Native Mac save/reopen/export uses a fresh task-owned directory in
Word's sandbox `Data/tmp`; Python then copies the artifacts byte for byte into
the requested output directory. Native paths and hashes are retained in the
result, and the staging directory is retained for evidence. This avoids asking
Word to save directly into every fresh caller output directory; any remaining
automation or native UI permission requirement must still be recorded. Exit status is 0 when all attempted native operations succeeded (or
when `--compile-only` succeeded), 2 when core save/reopen succeeded but capability
gaps remain, and 1 for core/ownership failure. A compile-only run explicitly says
native execution was unrun; it is not native acceptance. Windows preparation on a
Mac is likewise not Windows acceptance.

The Mac probe snapshots preexisting document names, paths, saved flags and full
text in memory, comparing them after closing only the task document. Raw snapshots
retain only metadata and text length, not user text. It never quits Word. The
Windows probe inspects a registered existing Word instance via `GetActiveObject`,
then uses explicit `DispatchEx("Word.Application")` and refuses mutations if that
instance is shared or already contains documents. It quits only its isolated
instance after the owned document is closed and zero documents remain. Snapshots
compare user-text hashes; other unregistered Word instances are outside this
snapshot scope. Neither runner changes global security or templates. These checks
cover text and saved state, not all application UI/selection or rich document
state.

The Windows runner verifies one new `WINWORD.EXE` process whose image directory
matches the COM application's path and checks `Microsoft Word` plus its version.
It compares COM `IUnknown` identities before creating a document and rejects an
instance with existing documents. Word exposes `Hwnd` on `Window`, not
`Application`: after creating its own document, the runner maps that document's
window to the verified process before writing content. Ambiguous process identity
fails closed; a concurrently launched Word can cause a conservative refusal.
Document-level `RemovePersonalInformation` prevents author-profile metadata in
published synthetic evidence. No global privacy or security setting is changed.

Run portable identity-guard tests with
`python fixtures/msoffice_spike/test_windows_identity.py -v`. After native runs,
`python fixtures/msoffice_spike/validate_windows_artifacts.py <output-directory>`
checks saved DOCX XML, native PDF text/fonts/page bounds and hashes, and renders
pages for separate visual review. It requires existing PyMuPDF and Pillow;
passing its structural checks does not imply polished layout or production admission.

Do not edit other Word documents during a probe: any observed state change is
reported as a preservation failure. A Mac timeout retains partial error output and
marks cleanup unverified. Inspect and close only the task-owned document if such
a timeout leaves it open; never broadly quit or kill Word. The probe does not
automatically bypass automation permission dialogs or native Word errors.
