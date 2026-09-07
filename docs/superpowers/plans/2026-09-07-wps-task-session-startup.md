# WPS task session startup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove demo-window churn from macOS WPS startup and show the final artifact only when explicitly requested.

**Architecture:** Keep existing authenticated per-task bridges, ownership and atomic publication. Replace wpsjs debug with a loopback-only profile service; activate a native blank task document and render into it; expose opt-in final presentation after cleanup.

**Tech Stack:** Python >=3.9, macOS WPS JSAPI, JavaScript, pytest/Node test harness.

**Spec:** docs/superpowers/specs/2026-09-07-wps-task-session-startup-design.md

## Global Constraints

- Python >=3.9; package remains importable without WPS/pywin32 on non-Windows hosts.
- Only requested artifact formats are publicly returned. The separate both-formats feature is deferred.
- Keep the existing 600-second generate deadline, capability authentication, registration snapshot/restore, fixed component ports/origins and PID identity checks. Do not reuse an arbitrary user WPS process for generation/conversion.
- No broad process termination, no modifications to unrelated user documents, no deletion of pre-existing staging sessions or history entries.
- Keep public generate()/convert_to_pdf() return values and overwrite/atomic-publication rules. Add an optional keyword-only open_result=False for explicit desktop presentation; interactive skill guidance uses True when the user expects to see the result. Existing unattended callers stay unchanged.
- The known numbered-heading style mismatch is outside this change. Preserve source-derived fonts/numbering/style behavior and record it during native acceptance.

### Task 1: Scoped loopback profile hosting

**Files:**
- Create: skills/WPSComposer/scripts/macos_probe/profile_server.py
- Modify: skills/WPSComposer/scripts/macos_probe/runtime.py
- Modify: skills/WPSComposer/scripts/longform/platform_runtime.py
- Modify: skills/WPSComposer/scripts/macos_probe/generation.py and conversion.py for scoped caller selection where present
- Test: tests/macos_probe/test_profile_server.py, test_runtime.py, test_generation.py, test_conversion.py; tests/longform_m5/test_platform_runtime.py

**Interfaces:**
- Consumes: existing generated profile files and COMPONENT_CONFIG.
- Produces: ProfileServer(profile_root: Path, port: int), start(), close(); ProbeRuntime(..., components: Optional[set[str]]=None), default all; callers select requested component.

- [ ] Add RED tests: real ephemeral-port HTTP serves index/JS/JSON, rejects traversal (including encoded and symlink), directory listing and writes, binds loopback; runtime only prepares/registers/serves selected components and never starts wpsjs debug.
- [ ] Run focused new tests and record actual failing output.
- [ ] Implement profile_server using Python standard library, validate names/subsets before filesystem or registration effects, restrict runtime loops to selected components. Preserve restore and owned-host termination semantics. Keep find_node/find_wpsjs_cli compatibility helpers if externally tested, but new runtime hosting cannot call them. No npm/dependency changes.
- [ ] Adapt existing mocks to the real hosting abstraction; keep failure/cleanup tests meaningful. Only selected component preflight ports are required free.
- [ ] Run focused suites, then .venv/bin/python -m pytest -v before commit. Report failures from unrelated unchanged code separately, do not change contracts merely to satisfy mocks.
- [ ] Commit explicit changed paths and write report with RED/GREEN evidence.

### Task 2: Native blank task activation and document ownership

**Files:**
- Create: macos/wps-jsapi-probe/resources/writer-blank.docx (controller provides native WPS seed)
- Modify: skills/WPSComposer/scripts/macos_probe/templates.py, runtime.py, models.py as needed
- Modify: skills/WPSComposer/scripts/longform/platform_runtime.py, macos_executor.py
- Modify: macos/wps-jsapi-probe/addin/bridge-client.js, writer.js, writer-longform-v2.js and asset-manifest.json
- Test: tests/macos_probe/test_runtime.py, test_templates.py, test_addin_assets.py; relevant tests/longform_m3 and tests/longform_m5 tests

**Interfaces:**
- Consumes: Task 1 selected-component runtime and pinned native blank seed.
- Produces: optional activation_document and retain flag in runtime/bootstrap; authenticated M5 activation-document claim carried only for first generation. Existing callers retain defaults. Add corresponding closed request schema field only if needed, validate both Python and JS sides.

- [ ] RED tests: correct private blank document is opened once; external/symlink/incorrect component paths rejected; claim resolves the owned document despite focus change; a different/missing path aborts; first M5 run reuses claimed document and relayout creates a fresh owned document; fixture close enumerates exact matching path only.
- [ ] Verify controller seed pinned hash, empty visible content, valid DOCX, no fields/macros/external relationships; runtime clones into its private staging path using existing validation/publication utilities. Do not generate or alter seed XML by hand.
- [ ] Implement runtime/bootstrap/command ownership end-to-end. Initial document copy has task-related safe basename. All user documents remain untouched; never use ActiveDocument as fallback. Maintain protocol/schema determinism and asset hash updates. A bootstrap claim failure must be observable before rendering.
- [ ] Run targeted tests including JS behavioral harness; test old request compatibility and relayout. Cover failure cleanup and generation against non-active owned document.
- [ ] Run full pytest before commit; do not run WPS acceptance concurrently with controller.
- [ ] Commit explicit changed paths and report RED/GREEN, seed hash, interface details.

### Task 3: Explicit final-artifact presentation

**Files:**
- Create: skills/WPSComposer/scripts/presentation.py (small platform launcher helper; choose another unambiguous focused name if conflict)
- Modify: skills/WPSComposer/scripts/orchestrator.py, conversion.py
- Modify: skills/WPSComposer/SKILL.md and references/api.md, README.md where helpful
- Test: tests/test_orchestrator.py, tests/test_conversion.py or equivalent existing public API test files, and presentation helper tests

**Interfaces:**
- Consumes: successful public artifact path after underlying native call and cleanup.
- Produces: generate(..., *, open_result: bool=False) and convert_to_pdf(..., open_result: bool=False), same return paths; platform helper to open only an existing finalized artifact.

- [ ] RED tests: default performs no opening; True opens exactly once after success/cleanup; generation/conversion failure never opens; opener failure emits a warning but returns published artifact; all formats and platform dispatch use shell-free argv, no shell interpolation.
- [ ] Implement bounded explicit display option without changing default behavior. Validate option before effects. Keep legacy and non-macOS public paths compatible.
- [ ] Update skill/API docs with guarded Python script example using if __name__ == '__main__', interactive open_result=True, unattended default False, no dual-format output promise and best-effort background behavior.
- [ ] Run focused/full tests, commit and report.

## Integration and acceptance (controller)

- [ ] Review each task diff independently for spec and code quality before next dependent task.
- [ ] Run fresh complete suite after final integration if later changes justify it; verify git diff --check.
- [ ] Run native fixtures sequentially and preserve screenshot/OOXML/PDF evidence. Check no demo files are opened; selected-component startup count; repeat generation; standalone conversion; explicit final output; unrelated disposable unsaved document preserved.
- [ ] Whole-branch review, fix evidence-backed findings, record precise local/remote/install status. Keep branch/worktree available; no implicit merge or push.

Writer activation uses the new native blank seed in every route, including standalone conversion; only M5 generation retains it. Other components keep their existing seeds in this Word-focused stage.
