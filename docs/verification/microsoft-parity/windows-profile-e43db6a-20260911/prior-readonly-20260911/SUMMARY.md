# Windows bounded read-only recovery diagnosis — 2026-09-11

**Recovery not performed; release remains NO_GO.** This run did not run pytest, native acceptance, installation, sentinel creation, UI input, document save/close, Office Quit/termination, merge, push or publication. Only new local diagnostic files were written. No document body is included in this report or the minimal correlation bundle.

## Fresh checkout and commands

Entry at 2026-09-11T08:25:34.9942555+08:00: Get-Date -Format o; Get-Location; git status --short; git rev-parse HEAD — exit 0.

- Actual cwd: D:/MyDocuments/codexprojects/WpsComposer
- Actual HEAD: f7e3ed72289c52b2aa8fd6c36129bf6f11f8f501; status empty.
- Existing isolated candidate D:/wpsce322: ce322b29cef43df68551e387dcc4e6e94bab5495; status empty.
- Failure source D:/wpsc665: 6653e14e9b45ff8c753c457adecbd3d77857f707; status empty.
- Read-only diagnose.py executed with D:/wpsc168-env/Scripts/python.exe (3.14.3), exit 0, 08:28:42–08:28:46 +08:00. Nested CIM process query also exit 0. Full command/exit/timing: current-state.command.json.
- Four relevant source files have both failure-commit and ce322 hashes in minimal-correlations.json. Existing failures are not relabeled as fresh candidate native runs.

## Current quarantine and process identity

All three active markers exist, match the retained 665 failure bytes, and were unchanged during diagnosis.

| Component | Marker SHA-256 | Recorded/current process findings |
| --- | --- | --- |
| Writer | 502fe12d0206d8ffc964c233bc3180b96841aea406c9179beecc27db8e5e04bc | Marker has evidence_path but no PID fields. Linked logs: worker 26192 absent, worker WINWORD 12416 absent. Old sentinel PID 8004 now belongs to C:/Program Files/nodejs/node.exe (created 08:11:15), so it is PID reuse, not a retained Word instance. |
| Spreadsheet | 21fab62c82ed0301bea1cbb58823de7c020f1a16bf695686172384a924bec3cd | Marker worker 7324 absent. Office 19360 alive at C:/Program Files/Microsoft Office/Root/Office16/EXCEL.EXE, created 2026-09-10 21:11:06. Old sentinel PID 7436 now belongs to the Codex cua_node runtime node_repl.exe (created 08:12:23), so it is PID reuse. |
| Presentation | 381a23a92c0445ff60027dda126bf5a9901866a95c21a8430ca1a9f8b3cf9076 | Marker has evidence_path but no worker/Office PID fields. Startup failed before unique native PID binding; no historical PID can truthfully be invented. No POWERPNT.EXE is currently observed. |

Full current executable paths, creation times, marker paths and evidence links are in current-state.json. No reused PID was acted on.

## Retained document state

Excel is independently bound by HWND 788508 -> PID 19360 and Microsoft application path. Its single workbook remains unsaved with one worksheet; name, saved state and value/formula content hashes match both the original retained_inventory.json and the later ce322 baseline. Worksheet content SHA-256: d628d553f844db93b69241977dd5a9f48b294f6288ca16648f4ff9fdee529661. Raw cell values/formulas were not written to output.

Original Word PID 13308 is absent, and there is no WINWORD.EXE process. Its two historical live document objects are no longer available. Their current saved/content state cannot be checked and is **not reported as unchanged**. This matches the preceding ce322 run's absence of WINWORD; this diagnosis did not close them. Word.Application currently resolves to WPS 12.0 in the WPS office6 path with zero registered Documents, which cannot substitute for Microsoft Word identity or prove all WPS UI windows empty.

## Minimal failure correlation

1. **Word add_table:** request 3 reached native result encoding and failed in NativeHandleRegistry._location, before a public UUID handle was allocated. The code compares the returned object's IUnknown with the session document's Tables items after session verification. Request 4 close was acknowledged; worker exit was 0. The returned table token, its owning-document token, session document token, Tables counts and individual comparisons were never logged. The evidence cannot choose between COM identity mismatch, wrong/empty collection or a truly foreign/stale object. A dedicated later probe is required to collect those fields without logging body text.

2. **Excel Move:** request 6 (move before) succeeded from sheet:2 to sheet:1; request 7 (move after) failed. Source uses named arguments source.Move(Before=anchor) then source.Move(After=anchor), through **{key:anchor}, without positional None. The failing assertion compares **Worksheets.Count**, despite the error saying workbook collection; it does not compare Application.Workbooks.Count. Exact workbook/sheet counts before and after Move, source/anchor Parent identity, and effective COM optional-argument binding are missing. Current 1-workbook/1-sheet state is a later retained snapshot, not the missing before/after data. Request 8 close failed and did not ACK.

3. **PowerPoint identity:** startup enumerates paneClassDC child HWNDs under POWERPNT.EXE, extracts native DocumentWindow objects, compares Application IUnknown and requires exactly one matching distinct PID. It failed before presentation Add. Neither raw candidate-window count nor matching PID set was recorded. The only justified condition is matching PID count != 1; zero versus multiple remains unknown. Dedicated later native instrumentation must log HWND, PID, executable, accessibility success and identity-equality booleans.

Exact source ranges, sanitized protocol rows and original file hashes are preserved in source-snippets.json and minimal-correlations.json. Missing fields remain missing; no repeated native experiment was run to manufacture them.

## Three recovery prerequisites

| Component | Concrete prerequisites before a future native retry |
| --- | --- |
| Word | Reconcile both worker and sentinel lifecycle separately; verify original Microsoft processes are gone by executable/creation identity (exclude reused node PID 8004), re-baseline current user documents, verify immutable marker/evidence and obtain component lock before any future archive decision. Add the missing table/document/collection identity diagnostics. Main worker close ACK alone does not clear sentinel quarantine. |
| Excel | Preserve live PID 19360 and its unsaved workbook. Re-establish exact workbook HWND/application binding, name/content fingerprint and provenance of the surviving sheet after Move before designing any owned-object recovery. Inventory all workbooks and compare parent identities/counts; a process PID or generic workbook name cannot justify discard/Quit. Obtain verified native close ACK and unchanged unrelated-document inventory before future marker recovery. Exclude reused node_repl PID 7436. |
| PowerPoint | Reconcile evidence with a fresh executable/process baseline; absent current POWERPNT is useful but no historical unique PID was recorded. Add candidate HWND/PID/native-application correlation at the startup boundary and require a unique verified application before creating a presentation. Verify marker/evidence under the component lock before any later recovery; do not recreate sentinels merely to diagnose this old failure. |

## Can existing helpers strictly identify their test objects?

- NativeSentinels uses live held IUnknown references plus application/window/executable bindings, saved state and content hashes. Within the original live process this is a strict guard that refuses uncertain cleanup. Its numeric token indexes are process-local; JSON token 4/5 cannot reattach an old object in a new worker. These old references are now unavailable.
- NativeHandleRegistry checks collection membership by IUnknown; it rejected the fresh Word table. It currently lacks the diagnostics needed to prove why that supposedly owned return value failed its guard.
- _OwnedOffice verifies executable/app IUnknown/document IUnknown and a document HWND/PID (or PowerPoint native window), and requires the unrelated-document baseline before Quit. Those guards are useful but do not reconstruct ownership after worker death or a sheet migrating between workbooks. Excel's close refusal correctly remains unverified.
- Prior recover.py is a one-off script pinned to older marker/session IDs and original Word baselines. It checks recorded numeric PID absence; today's PID reuse and changed original Word state make blind reuse invalid. It is not a generic recovery tool and was not executed.
- The WPS probe's post-Close cached collection failed with <unknown>.Count; the preceding UI run also observed zero registered COM Documents while a UI document was visibly open. Neither that collection count nor a WPS process list strictly establishes ownership of every visible test object.

No automatic recovery helper can currently be certified to rediscover and safely dispose of all these retained objects across processes. Current read-only checks establish state and gaps; they do not grant a cleanup pass.

## Delivery

This report is kept locally per this turn's explicit no-publication constraint. No new evidence push was made. The previous independent GitHub evidence branch was not changed. Portable profile-server fixes and full pytest were not repeated.
