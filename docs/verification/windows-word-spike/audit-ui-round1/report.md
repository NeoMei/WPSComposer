# Independent Windows UI audit, round 1

Baseline: `760704fc4b976dcc49f88d0e34835b1f5b7abbea`, 2026-09-08.
**PASS for the synthetic-copy UI edit/Undo/save/close/reopen workflow and five
portable output-path rejection cases.** No runner or production source was edited.
This audit precedes subsequent controller fixes and does not certify them.

## Actual interaction and ownership

PID 9252 was queried before work: WINWORD.EXE at the Microsoft Office path, created
01:18:13 local. Its launch ownership was not inferred from the PID. The registered
Word instance had zero documents. It was not quit or killed.

A byte-identical copy of the existing synthetic final-run-10 `probe.docx` was
created at ignored `build/msoffice-spike/audit-ui-round1/audit-ui-probe.docx`.
COM created a separate empty Word application, rejected shared COM identity,
opened only this copy, and mapped the document HWND 1968272 to new WINWORD.EXE
PID 19732 before setting Visible. The original evidence remained read-only.

Native `@oai/sky` CUA then selected the exact returned `audit-ui-probe - Word`
window. The screenshot showed native numbered headings 1 / 1.1 / 1.1.1 / 2 and
the TOC. The native accessibility tree exposed TOC/PAGEREF fields, numbered list
items, tables, nine pages and the end marker. COM setup and later readbacks are
explicitly distinguished from the actual mouse/keyboard actions below.

| Step | Actual transport | Observed outcome |
| --- | --- | --- |
| Initial open and isolation | COM | Exact copy/path and window/PID established; text hash recorded. |
| Focus body and go to end | CUA mouse click, then Ctrl+End | Visible caret below the final table, end marker intact. |
| Insert unique marker | CUA type_text | `WPS-AUDIT-UI-760704F-ROUND1` appeared in native document screenshot and accessibility text. COM readback: marker present, Saved=False. |
| Undo | CUA Ctrl+Z | Marker disappeared; COM full-text hash restored exactly; Saved=True. |
| Explicit save | CUA Ctrl+S | Chosen action was save the reverted copy; no discard/save prompt. COM Saved=True. |
| Close | CUA Ctrl+W | Owned window became empty Word, no document body. No application Quit. |
| Reopen | CUA Ctrl+F12, type exact copy path, Return in native Open dialog | Document reopened in the same task PID. COM readback verified original hash, structure and Saved=True. |
| Final close | CUA Ctrl+W | Empty Word window remained. Registered Word document count was still zero. |

Baseline/Undo/reopen full-text SHA256:
`e74f28843df6174118e3341ed08ad7fb5639932883383ec63610326d52bda751` (3741 characters).
After marker insertion:
`07054323edce2468fdb2d9cf1fa38b259087cd07c29e5a6437fc3180602a75a1` (3768 characters).
At every COM readback the table remained 82 rows / 3 columns with repeating
header, one TOC and five fields. End marker persisted. The saved copy is byte
identical to the original DOCX (SHA256
`90776e2747736b37834c7858cd1383c0cfa52567824324f115f7d31087f6028c`);
all ZIP parts, including document/styles/numbering XML, are identical.

## Real UI/tool limits

The first requested window screenshot incorrectly showed the foreground Codex
window while the accessibility tree described Word. No action used that image;
explicit Word activation and recapture produced the correct native window.
The erroneous capture was not published.

CUA's focused_element continued to report the navigation search field after body
clicks. Visible document caret, Ctrl+End navigation, insertion screenshot,
accessibility document text and independent COM readbacks confirmed actual body
input. Thus the stale focus field is retained as a tool limitation, not silently
treated as accurate. A Sharp import failed before any input while preparing
cropped evidence; Python/Pillow later cropped saved screenshots for publication.

Published screenshots are exact crops of the synthetic document area
`(260,185,1260,738)` from 1268x739 native captures, excluding account/ribbon and
navigation areas. No document pixels were altered. Original native captures
remain in ignored build. No Open-dialog screenshot or unrelated directory listing
was published. A Word paste-help tooltip appeared; no option/security setting was
changed. No claim is made about all UI flows, arbitrary user documents or complete
UI state preservation. Both the old PID 9252 and empty task application PID 19732
remained running; neither was quit or killed.

## Portable rejection tests

Five real CLI invocations of unchanged `windows_word.py` passed the expected
rejection contract: existing empty output directory, existing nonempty evidence,
same output repeated, output already a file, and child path under a file. Each
returned nonzero with FileExistsError at `out.mkdir`, before COM initialization.
All prior sentinel-file hashes were unchanged. Raw stdout/stderr and command
details are in `portable-gates/`; these are filesystem tests, not native Word runs.

The numbered headings, TOC-title placement, borderless table and configured row
splitting remain the prior prototype layout. No full native generation was repeated
solely for this UI audit. The included DOCX is the tested synthetic copy; original
evidence hashes were checked before and after. Production admission remains NO_GO.
