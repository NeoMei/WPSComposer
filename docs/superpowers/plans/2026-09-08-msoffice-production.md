# Native Word production release implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use subagent-driven-development task-by-task, independent task and final review.

**Goal:** Releasable explicit Word/PDF support on Windows and macOS via public APIs.
**Architecture:** Shared GenerationPlan and M5 lifecycle, pinned engine, native platform workers.
**Tech Stack:** Python3.9+, pywin32 on Windows, AppleScript on Mac, existing PDF quality dependencies.
**Spec:** docs/superpowers/specs/2026-09-08-msoffice-production-design.md

## Global constraints

Defaults wps; native Word only when explicitly selected/resolved. No document-library replacement for native layout. No shared Word Quit, no global template/security/registry changes. Preserve existing public positional args and string results. Atomic publish only after quality check. Never certify skipped native gates.

### Task1: Native platform executors

Windows files: new scripts/msoffice/windows_host.py, windows_runtime.py, tests/msoffice/test_windows_host.py and runtime tests. Consume LongformBuild and existing WindowsLongformExecutor(composer_factory=...), produce WindowsWordAdapter(build) with execute/export_pdf/patch_quality_notices/close compatible with _BaseAdapter. Native convert(request, timeout=...) -> Path. Verify process identity before touching documents; one worker deadline guards COM and isolates failures. Test known fake shared/empty/wrongPID paths before implementation; execute targeted tests to green; retain native acceptance for exact candidate.

Mac files: scripts/msoffice/macos_script.py, macos_runtime.py and tests/msoffice. Consume GenerationPlan and PreparedLongformResource -> ExecutionOutcome. Produce MacWordAdapter(build) lifecycle interface; convert(request, timeout=...) -> Path. Compile actual Word dictionary commands, TDD serializers/unsupported operations/strict parsing. Add concrete native smoke fixture to prove section/header/image/bookmark page mapping before wiring public API. Test source escaping and ensure no user content becomes AppleScript code.

- [x] Implement Windows native adapter and regression tests.
- [x] Implement Mac plan adapter and regression/native capability checks.
- [x] Independent reviews of each adapter; fix findings.

### Task2: Public engine routing and lifecycle

Files: new scripts/office_engines.py, scripts/orchestrator.py, conversion.py, longform/platform_runtime.py, presentation.py, __init__.py. Tests tests/msoffice/test_engine_routing.py.

Tests first: invalidengine fails without launches; explicit msoffice xlsx fails; auto chooses WPS first and pins fallback beforeexecution only; default positional compatibility preserved; requested Word open-result does not call systemdefault app; total deadline isfinitepositive. Implement engine resolver and lazy adapter imports. Pass explicit selections into native quality lifecycle; preserve old injected testfactory signatures for default route if necessary. Conversion timeout new keyword only.

- [x] Routing tests RED then GREEN.
- [x] Public generate/convert integration and presentation behavior.
- [x] Independent routing and cross-platform lifecycle review.

### Task3: Native acceptance and install/release preparation

Files docs/verification/msoffice-production/, fixtures/msoffice_release/, SKILL.md, references/api.md, README.md, package manifests/install metadata as required. Public fixtures generate Chinese body/sixlevels/cover/numbered chapters/longtable/image/refs/formula and separatePDF; verify actualOOXML/PDF with no helper make-up. Test unsupportedcapabilities rather than silentlydegradeunconfigured content. Windows remote run exactsourceand reportdiagnostics/UI. Mac localsame. Compare WPS representative output to protectedguardrails. Test cleaninstallmodule availability inisolateddestination. Versioncandidate0.9.0 onlywhenfeaturegatespass; existing0.8.1 untoucheduntilthen.

- [x] Mac public native acceptance and UI save/reopen.
- [x] Windows public native acceptance and UI save/reopen.
- [x] Existing WPS native representative regression.
- [x] Native heading insertion/deletion renumbering and updated TOC pagination after save/reopen on both operating systems.
- [x] Update exact API/capability/install/release docs.
- [x] Full pytest, repeated code/task audit and exact candidate clean install.
- [x] Commit/push reviewed candidate; report release readiness with concrete evidence and no automatic release publication.

## Execution ledger

Ruling: reuse current linked review worktree, create production branch from c4762c5; audit evidence remains ancestor. The user approved implementation and release readiness, so do not add repeated approval gates for these reversible changes.

2026-09-08 integration checkpoint: f48497d pushed on codex/msoffice-production for Windows native acceptance. Windows adapter independent review found and fixed pre-document application PID binding, A4 Normal-template drift, and resource quarantine after failed cleanup; re-review 78 tests passed. Portable integration snapshot 2757 passed, 12 native-gated skips. Isolated installer at /tmp/wpscomposer-production-install-0908 succeeded and both native adapters imported from the installed package; repeat on final candidate. Mac public smoke succeeded, but complex section/header and semantic coverage still under active fixes. WPS public representative generation/conversion succeeded but artifact gate caught real Heading 1–3 style alias drift (22/16/16 pt vs 16/15/15), now being fixed with native regression required. Remote Windows task dispatched exact f48497d plus public runner; do not confuse historical spike results with production acceptance. Release is not yet ready.

2026-09-08 integrated checkpoint: both production adapters and repeated independent reviews completed. Full suite: 2795 passed, 12 native-gated skips (build/msoffice-production/pytest-native-integrated.log). Mac Word final public representative generated DOCX, converted PDF and directly generated PDF, all 5 pages; WPS final representative all 4 pages with native style alias fix. Real keyboard edit, Undo, explicit save, close and reopen passed in both applications. Mac Word unsaved sentinel, native image/REF/SEQ, transparent headings, legacy DOC conversion and timeout/quarantine/recovery passed. Committed evidence is under docs/verification/msoffice-production. Windows production acceptance remains pending; historical spike evidence does not satisfy this gate. Six missing WPS bootstrap recovery tabs observed during UI verification are under investigation. Keep version 0.8.1 until remaining gates pass.

2026-09-08 final-review checkpoint: independent review and native contact sheets found and fixed safe native-error propagation and WPS builtin Title entering the TOC. Shared Title explicitly uses body outline level; stronger runner rejects cached TOC duplicates. Mac public representative-final-2 and WPS regression-05 pass, including actual unsupported-equation failure and no newly open WPS document tabs. Windows first production public API/UI/sentinel evidence and verified Caption identity fix were imported from 80b2c46 (local cherry-picks eef84e3, 6be339e, c2e24f7); repeat acceptance on final merged code remains required. WPS cleanup review caught and is correcting a non-isolated inspect/edit compatibility regression. Snapshot inspection confirms only Title outlineLevel=10 changed in four M3 snapshots. Windows-only POSIX 0600 assertion is now conditional while clone integrity still runs.

Windows native review correction: the initial report did not enforce section page-number policy. Parent/independent review found a cover PAGE field, missing section numbering restart values and physical page numbers 1–5 instead of none/i/1. Remote Windows owns the native section/footer correction and revalidation. Windows Title outlineLevel camel/snake aliases were separately fixed locally with two RED cases and 80 focused passing tests. This is a release blocker until the actual final PDF/TOC and section XML pass.

2026-09-08 section correction integrated (ead655a): exact production/test source matches remote c0604a8. Native setter evidence proves VT_BOOL True works where integer -1 is silently ignored. Footer unlink precedes PAGE creation; strict readback prevents publishing an incorrect section policy. Independent review passes and complete portable suite is 2846 passed, 12 skipped. Exact candidate isolated installation passes. Final Windows representative, semantic edge cases, process postflight and available WPS COM regression remain required before version promotion.

Additional final task review required native structural edits and exposed two real Mac defects: stale TOC entries despite successful generic field update, and independent heading list instances. Both corrected through native collection refresh and list-level LinkedStyle binding using the localized style name. Real four-level chapter insertion/deletion, changed-page TOC and save/reopen now pass. New shared-list runner gate rejects the previously accepted split-list samples. Latest local suite: 2855 passed, 12 skipped. Windows round3 WPS and Word strict pagination/media pass, but the same Word multilevel issue remains under remote correction.

Latest superseding checkpoint (d12b458): Windows round4 three-scheme tests pass in both Word and WPS (30 native checks across insertion, deletion, save/reopen and PDF), and real auto selection correctly discovers and uses WPS across COM registry views. Mac WPS Chinese/hybrid numbering is now aligned and independently reviewed, with 18 native checks and final complete representative regression passing; CUA shows home and zero document tabs. Current complete portable suite: 2874 passed, 12 skipped. Prior ledger entries are historical checkpoints. Remaining Windows work is only final representative structural pagination/TOC, final UI/postflight and complete Windows suite. Version promotion, final installation and PR readiness follow those gates; draft PR #8 is open and mergeable.

Final Windows checkpoint (1114df5): final native admission passes after importing 12debd1 and independently verifying 181 artifact checksums, production source identity, representative/structural PDFs, actual synchronous UI reopen, 28 task-owned PID exits and complete Windows test coverage (2838 + 1 serial passed, 45 skipped). Original helper failures and old/uncertain probes remain recorded. 0.9.0 metadata promotion passes its RED/GREEN check; final local full-suite and promoted-source installation are in progress.

Final 0.9.0 verification: 2874 passed / 12 native-gated skipped, exit 0 in 331.87 s. Isolated installation matches 105 files; real installed Word generation and conversion preserve source bytes and produce the expected PDF content. Source archive and wheel build successfully. Final incremental metadata/docs review and independent Windows evidence review find no blocker. No further production code changed after native acceptance; final candidate push and PR review-state transition remain the last administrative step.

Release readiness complete: promoted candidate and native installation artifacts pushed (fd2d308, 5dfbb20); PR #8 marked ready for review. No merge, tag, upload or installation into the user's live plugin directory was performed. The complete diff whitespace check passes outside one preserved historical raw TSV stderr file whose trailing empty event fields are intentional. All task checkboxes are complete.
