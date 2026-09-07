# Nonempty registered unsaved-document protection gate

**PASS for the explicitly tested synthetic sentinel, 2026-09-08.** This supplements
the original successful runs 05/06 without rerunning or replacing their evidence.
The native runner was unchanged: working-tree SHA256
`46fb262161af8f4a24f6cd1127a5542f0ba228cbe2ca86423c5722327d30505d`.
Production admission remains NO_GO for the limitations in the main report.

## Setup and scope

The registered Microsoft Word application was read first and had zero documents.
The wrapper created exactly one task-owned document in that application and wrote
a random unique `WPSCOMPOSER-UNSAVED-SENTINEL-...` marker. It retained that exact COM
document reference. `GetActiveObject` returned the same COM IUnknown, and its only
document matched the sentinel's name/path/hash and `Saved=False`. The sentinel's
window mapped to real WINWORD.EXE PID 9252. No application settings were changed.

The native runner ran as a separate Python process and started an isolated Word
instance. Only the synthetic sentinel appeared in its nonempty existing snapshots.
The wrapper never called Quit on the registered application, never killed any
process, and never saved the sentinel to disk. After matching the unique content
and original name/path, it explicitly called `sentinel.Close(SaveChanges=0)`.

## Final passing observation: sentinel-02 / native run-08

- Wrapper exit 0; native runner exit 0; all 14 native operations succeeded.
- Sentinel name and unsaved full name: `文档2` before and after; 62 characters,
  `saved=false` before and after.
- Full-text SHA256 before and after:
  `3e217f181393f52c708159a3cb7007fdc167bdc1bf32578e4cd8ef35068bd2e4`.
- Sentinel PID 9252; separate task WINWORD.EXE PID 21212. The task document HWND
  resolved to its task PID. Runner `existing_before` and `existing_after` each
  contain exactly the sentinel snapshot and match the wrapper's observations.
- Owned task document cleanup and isolated-instance Quit completed. Timestamped
  read-only PID observations show task process exit 2.422 seconds after the
  wrapper began observing. No forced termination was attempted.
- Sentinel closed with `SaveChanges=0`; registered document count returned to 0;
  the original application remained responsive. PID 9252 was not quit.
- Real native DOCX/PDF: all 11 artifact checks passed, 82 rows/3 columns, 80
  records, native numbering and TOC, Chinese font/indent, repeated headers, final
  marker and 9 PDF pages. Every rendered page is byte-identical to previously
  visually reviewed run 06. The same recorded layout limitations apply.

## Preserved initial observation: sentinel-01 / native run-07

The first sentinel attempt also preserved its nonempty unsaved document exactly
and closed it without saving. Name/full name `文档1`, 62 characters, `saved=false`,
SHA256 `3f132b52aa4c011c399c95a7738af7b9658ceff2fd202132f26ace93ee6a407c`
matched before/after. The native runner exited 0 and its 11 artifact checks passed.

However, the wrapper exited 1: it tested for PID 15164 disappearance immediately
after Quit, while Word was still exiting. Its original `overall=failed` and
`runner_process_absent_after_cleanup=false` remain untouched. A subsequent
read-only process query found only PID 9252; PID 15164 was absent. That observation
is retained separately in `windows-sentinel-01/process-exit-followup.json`.

The wrapper was corrected to observe the verified task PID for up to 10 seconds,
at 250 ms intervals, recording elapsed time and presence. It does not force
termination. A fresh sentinel and fresh native output directory were then used for
run 08, producing the passing observation above. No changes were made to the
native runner or the original successful artifacts.

## Reproduction and limits

Use an existing 64-bit Python with pywin32. The registered Word application must
already exist and be empty; otherwise the wrapper refuses to create the sentinel.
Do not edit other Word documents during this test.

```powershell
python fixtures/msoffice_spike/windows_unsaved_sentinel.py --evidence-dir build/msoffice-spike/windows-sentinel-new --output-dir build/msoffice-spike/windows-run-new
python fixtures/msoffice_spike/validate_windows_artifacts.py build/msoffice-spike/windows-run-new
```

Both output directories must be new. Public evidence contains only synthetic
sentinel metadata. Original wrapper results/stdout/stderr, native results,
DOCX/PDF, validation JSON, contact sheets and SHA256 are retained. This proves
preservation of this synthetic unsaved document's plain text, name, path and saved
state during this probe; it does not prove arbitrary rich content, selection/UI
state, all unregistered instances, or general concurrency safety. The old PID
9252 remains intentionally untouched at the application level. Full pytest is
still unavailable locally; no global environment or production plugin was installed.
