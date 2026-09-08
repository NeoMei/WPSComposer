# Final cleanup patch: Windows native verification

**PASS: two fresh native runs and the registered unsaved-sentinel gate on the
exact reviewed candidate `f0dc91f279d3b6d8d61f43a02046964dbc30a1d7`.**
Verification date: 2026-09-08, Asia/Shanghai. This is native prototype validation;
production integration/admission remains NO_GO and no release or master merge
was performed.

## Candidate and source identity

The evidence branch was clean at `233b24057ab20f1a3f7fdff43009d4c940b34952`.
The review branch was fetched, its SHA and parent checked, and the evidence branch
fast-forwarded to `f0dc91f279d3b6d8d61f43a02046964dbc30a1d7`. No local runner patch
was made. Its only runtime change defers `exclusive_instance=True` and
`Visible=True` until the task document's HWND maps to the verified candidate PID.

The raw executed runner is preserved byte for byte as
`executed-windows-word.py.txt` (14386 bytes, 273 CRLF, zero bare LF):

- Raw SHA256: `0770f34d18e4bb2ecccf15572efb017183c8e99cb1588e4328b6d73e52d028ed`.
- CRLF-to-LF SHA256: `5dc4f102fb1f0243181b4eaea2362b75b1b6b5eebb634f0692fd7ca2671cb65a`.
- Git blob: `8c2fc817a5ee54a1ef5f9fcc4f132642456823ed`.

The normalized bytes equal that blob. The snapshot was taken before execution;
the source was rechecked after both runs. The sentinel wrapper independently
recorded the same raw SHA. Source provenance is separate from the historical
05-08 raw SHA `46fb262...`; neither the older raw results nor their executed
source snapshot was overwritten.

## Native runs

Existing Python 3.11.9 x64 and pywin32 312 were used. The process image was
`C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE`;
COM reported Microsoft Word 16.0 / build 16.0.17932. The previously checked
installed file version was 16.0.17932.20910. No packages or plugins were installed.

| Run | Command exit | Task PID | Outcome |
| --- | --- | --- | --- |
| 09 | 0 | 3192 | All 14 native operations passed; empty registered snapshots preserved. 57.812 s. |
| 10 | Wrapper 0, native 0 | 15836 | All 14 operations passed; nonempty unsaved sentinel preserved and closed without saving. Wrapper 62.641 s. |

Actual commands, UTC start/end times, stdout and stderr are retained in
`command-logs/`; native run 10 stdout/stderr are also retained with
`windows-sentinel-03/`. There were no native errors. No previous output directory
was reused. `windows_word.py` was identical for both runs.

**Deferred visibility works on this installed Word:** run 09 obtained task HWND
33490176 and mapped it to PID 3192; run 10 obtained HWND 1444194 and mapped it to
PID 15836. In the verified source, these operations occur before granting Quit
permission or setting Visible. Both then completed content generation and export.
No safety condition was moved back or bypassed.

## Nonempty unsaved sentinel and cleanup

The sentinel was created only in the existing registered application after
confirming it had zero documents. Its owned document window mapped to PID 9252.
`GetActiveObject` returned the same COM identity and exposed exactly that sentinel.

- Before and after: name/full name `文档3`, 62 characters, `saved=false`.
- Both full-text hashes:
  `7b0e17f865fa0118027644c2ed8a4138c1b6f842d007ef5d5bdd96ae6e023e66`.
- Runner `existing_before` and `existing_after` each contain exactly this
  synthetic snapshot and equal the wrapper's independently read snapshots.
- Task PID 15836 differed from sentinel PID 9252; its document PID matched.
- Task document cleanup and verified isolated-instance Quit succeeded. PID 15836
  disappeared after 2.453 seconds of bounded read-only exit observation.
- The wrapper rechecked the exact unique sentinel content and name/path, then
  explicitly called `Close(SaveChanges=0)`. The sentinel was never saved to disk.
  The registered application remained responsive with document count 0.
- A final read-only WINWORD process query returned PID 9252 only; both task PIDs
  3192 and 15836 were absent. No attempt was made to quit or kill PID 9252.

## Independent saved-artifact and visual checks

Each run passed all 11 checks using existing PyMuPDF/Pillow and saved DOCX XML:
exact 82-row/3-column fixture with all 80 records and end marker; repeating header;
12 pt 仿宋 and 24 pt/two-character first-line indent; Heading 1/2/3 style sizes and
native numbering links; TOC with four PAGEREF fields; all records once in the PDF;
headers on every page; final marker; native/PDF page counts agreeing at 9;
FangSong 12 pt in the PDF; text bounds within page boxes.

The run 09 contact sheet was visually reviewed. Every full-page PNG of run 09
equals previously reviewed run 06, and every run 10 page equals run 09 byte for
byte. Thus the first-page hierarchy/TOC, wrapping/repeated headers and page-9 ending
remain unchanged. The same prototype limits remain: TOC title below its entries,
borderless table, configured row splitting across pages, and a sparse last page.
This is not approval of a polished publication template. Author metadata is empty.

The Windows host reran five existing portable identity tests: 5 passed. It still
has no pytest installation. Separately, the checked-in macOS controller review
records six main-flow cleanup regressions passing and the full suite result
**2630 passed, 12 skipped, 175.21 s** for this candidate; that is controller-side
evidence, not a Windows pytest execution. See the immutable review report in
candidate f0dc91f. No install was performed to change this distinction.

`manifest.json`, `source-provenance.json` and `SHA256SUMS.txt` describe only this
final candidate's new evidence. The parent manifest links here while retaining
its original source hashes. Local raw originals are under
`build/msoffice-spike/windows-run-09`, `windows-run-10`, `windows-sentinel-03` and
`windows-final-review-logs`. Published copies preserve the original bytes.
