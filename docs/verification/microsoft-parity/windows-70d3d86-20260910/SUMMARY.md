# Windows execution checkpoint for 70d3d86 — PARTIAL / BLOCKED

This is fresh Windows command and startup-preflight evidence from September 10,
2026. It is not native acceptance, a release approval, or an accepted-dispatch
status. Normal Office and UI gates remain blocked.

## Fresh entry and environment

Entry command exit code **0**. Actual initial output:

```text
2026-09-10T22:33:19.1647711+08:00
D:\MyDocuments\codexprojects\WpsComposer
f7e3ed72289c52b2aa8fd6c36129bf6f11f8f501 Record fresh Windows sentinel and native acceptance
```

Commands: current date, cwd, `git status --short`, `git log -1`, process inventory,
installed executable version inspection, and GitHub commit resolution.
`git status --short` emitted no lines. The original workspace is preserved.

- Active original Word PID **13308** and retained Excel PID **19360**.
- Word, Excel and PowerPoint executables: `C:/Program Files/Microsoft Office/Root/Office16/`, version **16.0.17932.20910**.
- WPS Writer/ET/WPP executables: `C:/Users/1/AppData/Local/12.1.0.28505/office6/`, version **12.1.0.28505**; none running at entry.
- Python: existing `D:/wpsc168-env/Scripts/python.exe`, Python **3.14.3**.

## Exact candidate source

Candidate: **70d3d86ff6b03945dacccffd1688922830ad6879**.
Isolated checkout: `D:/wpsc70`, branch `codex/windows-acceptance-70d3d86`.
Actual imported API: `D:/wpsc70/skills/WPSComposer/__init__.py`.
The current remote candidate branch was also read and matched this SHA.

All **490 materialized tracked files** match the candidate Git blob identities.
Native source digest:
`0c2ae9e66b36daa9f34729bd5e45f6bb208e015dbd799b73ee24cf8fb2cd0a87`.
Final source comparison is included in `source-verification.json`.
Historical evidence files (18,692) are excluded from materialization; full
historical Git object closure is not claimed. The acquisition log retains a
transient gitattributes-missing diagnostic during object import; final
materialized source hashes are independently checked.

Read current `AGENTS.md`, `status.md`, `windows-candidate.md` and native runner.
The private Excel fix is implemented in Mac session/process/transport code;
the Windows native runner and Windows host paths are unchanged from 6653e14.
Shared preflight/document API changes are included in this candidate's source
checks. Mac full42 and hosted CI are not Windows native acceptance.

## Actual current startup checks

```powershell
$env:PYTHONPATH='D:/wpsc237-evidence/guard'
$env:WPSC_NO_TERMINATION_LOG='D:/wpsc70-evidence/process-preservation.jsonl'
D:/wpsc168-env/Scripts/python.exe D:/wpsc70-evidence/preflight.py
```

Execution exit code **2**. All three actual current-source public
`create_document(kind, engine='msoffice')` calls failed before creating a worker:

```text
NativeOfficeError: Native Office cleanup is unverified; explicit recovery is required.
code: NATIVE_OFFICE_QUARANTINED
```

| Component | Current result | Normal create/generate/edit/save/PDF/reopen | UI edit/Undo |
|---|---|---|---|
| Word | BLOCKED by retained writer quarantine | Not run | Not run |
| Excel | BLOCKED by retained spreadsheet quarantine and live uncertain Excel instance | Not run | Not run |
| PowerPoint | BLOCKED by retained presentation quarantine | Not run | Not run |

No new native session directory appeared, and all marker bytes remain unchanged.
This is an application lifecycle guard, **not a filesystem permission refusal**.
The public API's exact error stacks and console output are published here.
No quarantine recovery, marker removal, new fault injection or safety-setting
change was attempted in this checkpoint.

The standard native runner creates sentinels before reaching the guarded public
session. It was not launched over known unresolved quarantine. In particular,
Excel's recorded uncertain instance is still alive with an unsaved workbook;
the dead-process recovery used in a previous turn does not apply to it.

## Preserved state and actionable existing failures

The two original Word documents have identical before/after path, window, saved
state and content hashes. The original unsaved Word document remains unsaved.
Excel PID 19360 has the retained unsaved workbook with one `Details` worksheet;
its current before/after content hash and saved state match. Neither application
was closed, quit, edited or saved. Detailed user-document inventories stay local;
only preservation counts/results are published in `preservation.json`.

`historical-6653e14/` contains explicitly **older** raw failure records, supplied
to make the unresolved Windows defects reviewable, not as current native runs:

- Word `add_table()` fails native handle registration; main worker closes with
  an ACK, while sentinel cleanup reports disconnected COM object.
- Excel's second worksheet move changes the expected workbook collection;
  close returns `Unrelated Office documents changed; application left open`.
  Diagnostic copying additionally raises WinError 32 on live Excel logs.
- PowerPoint cannot bind its native window to one application process before
  adding a presentation.

The current-source normal/UI matrix requires resolving these Windows failures
and performing exact recovery of the retained live Excel state. No generic
`passed` result, successful source check or portable suite closes those gates.

## Publication scope

This checkpoint is published as an evidence-only child commit of 70d3d86 on
`codex/windows-evidence-70d3d86-20260910`. Production code and candidate branch
are untouched. No merge, tag, release or personal installation is performed.
The publication method is GitHub Git Data API; remote parent, tree, branch and
every published file are read back and checked after creation.

Files include startup report/console/exit/stacks, all materialized source hashes,
preservation summary, full Windows pytest log and exit code, acquisition log,
and clearly separated historical failures. Local evidence root:
`D:/wpsc70-evidence`. Local publication verification receipt is
`publication-verified.json`; SHA-256 list is `manifest.json` in this directory.

## Windows pytest result

Ran `D:/wpsc168-env/Scripts/python.exe -m pytest -v` in the exact isolated
checkout. Pytest collected **5,182 items**, but the run did **not complete**.
At approximately 41% it stopped progressing in
`test_profile_server_close_releases_request_handler`.

A read-only py-spy stack captured main thread blocked in
`ProfileServer.close -> server_close -> Thread.join` and a request thread
blocked in `socket.readinto -> handle_one_request`. The original raw log and
stack are retained. A Ctrl+C was sent to the exact owned execution session
90031; it returned exit **1**. Both recorded pytest PIDs 3472 and 20452 were
then absent. No independent pytest final exit status or final count summary
was produced, and no full-suite pass is claimed. `pytest-interruption.json`
records observed log counts separately from a completed pytest result.

The available log records 2,105 passed, 18 failed and 18 skipped items before
interruption; these are observed lines, not final suite totals. A separate
`pytest tests/macos_probe/test_generation.py -x -q` reproduction exited 1 with
one failure in 0.35s: the pinned WPS generation template could not be staged.
Its complete exception chain is in `pytest-first-failure.log`; the JavaScript
template resources were not installed in this isolated checkout. This setup
failure and the profile-server shutdown hang are recorded separately.

This new Windows test hang is additional to the native quarantine blockers.
It is not an Office UI/native acceptance result. Only the owned pytest session
was interrupted; both original Word documents and the retained unsaved Excel
workbook were rechecked unchanged afterwards.
