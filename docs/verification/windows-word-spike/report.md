# Windows native Microsoft Word/PDF evidence

**Final cleanup patch verification:** the exact reviewed candidate
`f0dc91f279d3b6d8d61f43a02046964dbc30a1d7` passed fresh native runs 09/10,
including the nonempty unsaved-sentinel gate. See
[final-cleanup-review/report.md](final-cleanup-review/report.md) and its independent
manifest/source snapshot. The 01-08 records and source hashes below are historical
evidence for the earlier runner, not evidence for the final cleanup patch.

Date: 2026-09-08 (Asia/Shanghai). Candidate baseline:
`88a0ecc4b5bc41f12ccc44e1410fa50ce1ece630`, from
`codex/msoffice-native-spike`. Evidence and minimal probe fixes are on the
independent `codex/windows-word-spike-evidence` branch.

## Result and scope

**Native capabilities and persisted artifact checks passed in two final fresh
runs (05 and 06), followed by a successful registered unsaved-sentinel gate
(run 08). Production admission remains NO_GO.** Both original final runs saved/reopened
a real DOCX and exported a real 9-page PDF through Microsoft Word. Each completed
all 14 probe operations with exit 0 and no native errors; each separately passed
11 saved-artifact checks. The 9 rendered page images are byte-identical between
these runs. This is a Windows feasibility result, not a public engine release,
full product regression pass, or proof of arbitrary user-document preservation.
The additional synthetic unsaved-document preservation evidence and its initial
asynchronous-exit observation are detailed in [sentinel-report.md](sentinel-report.md).

No registry, macro-security, Normal-template, installation, release, merge, or
global process termination action was taken. No preexisting document was edited.
The initial rejected connectivity-file test was not retried. Ordinary clone,
source edits, native generation and evidence writes subsequently succeeded.

## Environment and actual engine

- Windows build reported by Python: `Windows-10-10.0.26100-SP0`.
- Python 3.11.9, 64-bit; pywin32 312. Existing PyMuPDF 1.27.2.2 and Pillow 12.2.0
  were used for independent artifact analysis and rendering.
- Word COM reports `Microsoft Word`, version `16.0`, build `16.0.17932`.
- Actual executable: `C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE`;
  file/product version `16.0.17932.20910`.
- Run 05: newly observed Word PID 26788; owned document window HWND 54330552 maps
  to the same PID. Run 06: PID 25840; HWND 5965298 maps to the same PID.
- The 64-bit Word registration points to WINWORD.EXE. The 32-bit registration
  points to WPS; ProgID alone is therefore insufficient evidence on this machine.

## Command history and fixes

The directory was empty before executing exactly:

```powershell
git clone --branch codex/msoffice-native-spike --single-branch https://github.com/NeoMei/WPSComposer.git .
git log -1 --format=%H
git switch -c codex/windows-word-spike-evidence
```

Clone exited 0 and HEAD matched the candidate baseline. Each native invocation
used the existing 64-bit Python 3.11 executable and this command pattern:

```powershell
python fixtures/msoffice_spike/windows_word.py --output-dir build/msoffice-spike/windows-run-NN
```

| Run | Exit | Observation |
| --- | --- | --- |
| 01 | 1 | `AttributeError('Word.Application.Hwnd')` before document creation; DOCX/PDF absent. |
| 02 | 0 | After replacing the Application.Hwnd assumption, all native operations succeeded; 9-page PDF. Intermediate artifacts retained locally. |
| 03 | 0 | Repeat of 02 in a fresh directory; all native operations succeeded. Intermediate artifacts retained locally. |
| 04 | 1 | New document's HWND/PID check succeeded; experimental author-property setter failed with `com_error(-2147352573, '找不到成员。', None, None)`. Owned document and isolated instance cleaned up. |
| 05 | 0 | Final runner, document-level RemovePersonalInformation, complete native and artifact checks passed. |
| 06 | 0 | Same final runner, fresh directory, complete native and artifact checks passed. |
| 07 | 0 | Nonempty unsaved sentinel preserved; native/artifact checks passed. Wrapper exited 1 because its immediate post-Quit PID check was premature; raw result retained. |
| 08 | 0 | Same native runner; unsaved sentinel preserved, task PID exit observed after 2.422 seconds, sentinel closed without saving. Wrapper exited 0. |

All eight original `result.json` files are retained in the corresponding evidence
subdirectories, including full tracebacks for failures. No failure was overwritten.
Final native binaries are copied byte for byte from ignored build outputs.
Run 02/03 binaries remain local because they predate the document privacy setting.

The first issue is a probe API bug: Microsoft documents
[Hwnd on Word Window](https://learn.microsoft.com/en-us/office/vba/api/word.window.hwnd),
not Word Application. The fix checks new native process images and COM application
path/name, compares COM IUnknown identities, rejects nonempty/shared instances,
and then confirms the owned document Window.Hwnd maps to the verified new process
before writing content. Ambiguous concurrent startup fails closed. Run 04's
author setter was removed; final evidence uses only the document-level
`RemovePersonalInformation` setting. Published DOCX creator/lastModifiedBy fields
are empty and published PDF Author is empty.

## Saved-artifact verification

`validate_windows_artifacts.py` read the saved ZIP/XML and PDF, independently of
the native runner's success states. It did not reopen Office or alter DOCX/PDF.

- DOCX: one exact 82-row, 3-column table matching the synthetic fixture: header,
  records 001-080 and final `MSOFFICE-SPIKE-END` row. Repeating-header XML exists.
- Body sample: East Asian font 仿宋, size 24 half-points (12 pt), firstLine 480
  twips (24 pt), firstLineChars 200 (2 characters). PDF body font is FangSong,
  12 pt. This checks the designated body sample, not every paragraph's typography.
- Native Heading 1/2/3 and second Heading 1 retain sizes 16/15/15/16 pt;
  styles link through numId/abstractNum to matching levels and `%1`, `%1.%2`,
  `%1.%2.%3` numbering. PDF displays 1, 1.1, 1.1.1 and 2.
- Native TOC field (`TOC \\o "1-3"`) and four PAGEREF fields persist. All four
  headings are on page 1 and displayed TOC page numbers are 1.
- Native Word page count and independent PDF count are both 9. No Mac page-count
  equality requirement was used. PDF contains all 80 records once, header labels
  on every page, and final marker/结束标记/完成 on page 9. Text bounds stay within
  page boxes. Chinese wrapping and repeated headers are visible in rendered pages.
- Visual review: all-page contact sheet plus full first/second/last page views;
  both final runs have identical page raster bytes. No missing ending, page-edge
  clipping or overlapping columns was observed. Full rendered pages remain in build;
  contact sheets are published alongside the final DOCX/PDF.

### Layout limits

The TOC is inserted before the literal `目录` heading, leaving that heading below
the entries. Table borders are not applied. Rows are allowed to split across
pages: records 004/025/046/067 continue below the next page's repeated header.
This is the configured behavior and text is retained, but it is not polished
publication layout. The final page contains records 078-080 and the end marker
with substantial white space. These limits were not silently edited away.

## Ownership, cleanup and validation limits

Final runs compared empty registered-instance document snapshots before/after;
each owned instance started empty. Only the task document was closed without
saving, followed by Quit on that verified isolated instance after zero documents
remained. Additional runs 07/08 used one uniquely marked task-owned unsaved
sentinel in the registered instance: nonempty before/after snapshots matched,
and the sentinel was explicitly closed without saving. Run 08 also observed the
owned process exit. This covers synthetic text/name/path/saved state, not arbitrary
rich user-document state or unregistered Word instances.

The failed first attempt left Word PID 9252 (created at 01:18:13 local). It was
still present after the sentinel gate. The sentinel's owned document window now
maps the registered application to PID 9252, but ownership of that application's
initial launch remains unverified. This process was not quit or killed.
Therefore whole-session cleanup is **not fully verified**, even though successful
owned-instance cleanup passed. Published snapshots are either empty (01-06) or
contain only the task-created synthetic sentinel (07/08). No user document names,
paths, text or hashes are included.

Five portable identity-guard tests passed (unittest). They test fail-closed
decisions and do not substitute for native evidence. `git diff --check` passed.
The requested full pytest suite could not start: the selected Python returned
`No module named pytest`, exit 1; the clone has no `.venv`. No dependency was
installed. This is another reason not to claim full product/production acceptance.

## Reproduction and evidence

```powershell
python fixtures/msoffice_spike/test_windows_identity.py -v
python fixtures/msoffice_spike/windows_word.py --output-dir build/msoffice-spike/windows-next-01
python fixtures/msoffice_spike/validate_windows_artifacts.py build/msoffice-spike/windows-next-01
```

Always use an absent output directory and a 64-bit Python with pywin32. Do not
edit other Word documents during the probe. Review the images separately; exit 0
alone is not acceptance. `manifest.json` records all eight run command exit codes,
baseline/source hashes, artifact SHA256, and the final visual comparison.
`SHA256SUMS.txt` covers the published evidence files. The local originals remain
under `build/msoffice-spike/windows-run-01` through `windows-run-08`; wrapper evidence
is under `windows-sentinel-01` and `windows-sentinel-02` in the same build directory.

These results support continuing a native Windows adapter feasibility track for
Word/PDF fidelity. They do not compare Office.js experimentally or authorize
production integration until layout, broader rich-document preservation, cleanup
failure handling and full regression gates are completed.
