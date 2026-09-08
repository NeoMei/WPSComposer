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

- [ ] Implement Windows native adapter and regression tests.
- [ ] Implement Mac plan adapter and regression/native capability checks.
- [ ] Independent reviews of each adapter; fix findings.

### Task2: Public engine routing and lifecycle

Files: new scripts/office_engines.py, scripts/orchestrator.py, conversion.py, longform/platform_runtime.py, presentation.py, __init__.py. Tests tests/msoffice/test_engine_routing.py.

Tests first: invalidengine fails without launches; explicit msoffice xlsx fails; auto chooses WPS first and pins fallback beforeexecution only; default positional compatibility preserved; requested Word open-result does not call systemdefault app; total deadline isfinitepositive. Implement engine resolver and lazy adapter imports. Pass explicit selections into native quality lifecycle; preserve old injected testfactory signatures for default route if necessary. Conversion timeout new keyword only.

- [ ] Routing tests RED then GREEN.
- [ ] Public generate/convert integration and presentation behavior.
- [ ] Independent routing and cross-platform lifecycle review.

### Task3: Native acceptance and install/release preparation

Files docs/verification/msoffice-production/, fixtures/msoffice_release/, SKILL.md, references/api.md, README.md, package manifests/install metadata as required. Public fixtures generate Chinese body/sixlevels/cover/numbered chapters/longtable/image/refs/formula and separatePDF; verify actualOOXML/PDF with no helper make-up. Test unsupportedcapabilities rather than silentlydegradeunconfigured content. Windows remote run exactsourceand reportdiagnostics/UI. Mac localsame. Compare WPS representative output to protectedguardrails. Test cleaninstallmodule availability inisolateddestination. Versioncandidate0.9.0 onlywhenfeaturegatespass; existing0.8.1 untoucheduntilthen.

- [ ] Mac public native acceptance and UI save/reopen.
- [ ] Windows public native acceptance and UI save/reopen.
- [ ] Existing WPS native representative regression.
- [ ] Update exact API/capability/install/release docs.
- [ ] Full pytest, repeated code/task audit and exact candidate clean install.
- [ ] Commit/push reviewed candidate; report release readiness with concrete evidence and no automatic release publication.

## Execution ledger

Ruling: reuse current linked review worktree, create production branch from c4762c5; audit evidence remains ancestor. The user approved implementation and release readiness, so do not add repeated approval gates for these reversible changes.
