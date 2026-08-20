# WPSComposer Long-form M0 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build and run an isolated dual-platform M0 probe that proves or rejects every native WPS capability required by the long-form 0.8.0 design before any M1 product implementation begins.

**Architecture:** Add a platform-neutral evidence contract and matrix verifier, then drive one closed Writer probe command through the existing authenticated macOS bridge and one supervised worker through Windows COM. Both probes create WPS-native DOCX/PDF evidence, close and reopen the document, refresh fields, collect redacted structured snapshots, and feed the same verifier; the probe code remains outside the public generation API.

**Tech Stack:** Python 3.9+, pytest, WPS Office Writer, macOS WPS JSAPI through the existing loopback bridge, Windows `pywin32` COM with `DispatchEx`, JSON evidence, `pypdf`, `pdfplumber`, Pillow, Poppler command-line inspection when available.

## Global Constraints

- M0 is a go/no-go gate. Do not change the public `generate()` routing or begin M1 while any required capability 1-14 lacks passing Windows and macOS evidence.
- Use only WPS-native APIs and WPS-native editable objects. Do not create or repair the probe DOCX with hand-written OOXML and do not use a non-WPS rendering fallback.
- Capability 15 (SVG) is optional; an unsupported result excludes SVG from 0.8.0 without failing capabilities 1-14.
- Content-bearing capabilities must be verified after save, close, reopen, field refresh, save, and PDF export.
- A missing WPS engine, protocol mismatch, resource-manifest mismatch, save failure, export failure, or unprovable Windows process ownership is fatal for the platform probe.
- A local object failure may degrade only inside the checkpoint assigned to that object; the probe must prove cleanup, visible fallback insertion, and continuation.
- The macOS probe may write WPS runtime state only under `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/`; registration must be restored byte-for-byte on success, error, and timeout.
- The Windows probe must use `DispatchEx` or a separately proven equivalent, must identify its worker-owned WPS process tree, and must never terminate a pre-existing WPS process.
- Evidence must not contain source paths, staging paths, document text, bookmark mappings, field hashes, resource hashes, environment variables, or process command lines.
- Keep `uv.lock` untouched.

## File and Interface Map

- Create `skills/WPSComposer/scripts/longform_m0/__init__.py` — internal M0 exports only.
- Create `skills/WPSComposer/scripts/longform_m0/contracts.py` — capability IDs, statuses, evidence records, redaction, canonical JSON, and matrix validation.
- Create `skills/WPSComposer/scripts/longform_m0/host_checks.py` — dependency, PDF font/page-box, artifact, and evidence-directory checks shared by both platforms.
- Create `skills/WPSComposer/scripts/longform_m0/macos.py` — macOS orchestration using `ProbeRuntime` and `LoopbackBridge`.
- Create `skills/WPSComposer/scripts/longform_m0/windows.py` — Windows supervisor and worker entry point with lazy `pywin32` imports.
- Create `skills/WPSComposer/scripts/longform_m0/__main__.py` — explicit `--platform macos|windows|verify` CLI.
- Create `macos/wps-jsapi-probe/addin/writer-longform-m0.js` — isolated Writer JSAPI probe implementation.
- Modify `macos/wps-jsapi-probe/addin/index.html` — load the isolated probe module before `writer.js`.
- Modify `macos/wps-jsapi-probe/addin/bridge-client.js` — preserve typed protocol and resource-manifest mismatch errors.
- Modify `macos/wps-jsapi-probe/addin/writer.js` — register exactly one typed `probe_longform_m0` handler.
- Modify `skills/WPSComposer/scripts/macos_probe/models.py` — route the new method to Writer.
- Modify `skills/WPSComposer/scripts/macos_probe/runtime.py` — copy the isolated probe asset into every private add-in profile.
- Create `tests/longform_m0/test_contracts.py` — evidence and matrix contract tests.
- Create `tests/longform_m0/test_host_checks.py` — PDF/artifact/dependency tests.
- Create `tests/longform_m0/test_macos.py` — bridge orchestration, cleanup, and error tests with fakes.
- Create `tests/longform_m0/test_windows.py` — worker ownership and timeout tests with fake process/COM adapters.
- Create `tests/longform_m0/test_addin_assets.py` — static closed-protocol and JS safety assertions.
- Create `docs/longform-m0.md` — operator commands, evidence interpretation, and recovery instructions.
- Create runtime evidence only under `build/longform-m0/<platform>-<run-id>/`; keep it ignored and do not commit native artifacts.

---

### Task 1: Define the closed M0 evidence contract

**Files:**
- Create: `skills/WPSComposer/scripts/longform_m0/__init__.py`
- Create: `skills/WPSComposer/scripts/longform_m0/contracts.py`
- Create: `tests/longform_m0/test_contracts.py`

**Interfaces:**
- Consumes: JSON-compatible mappings returned by either native executor.
- Produces: `CAPABILITIES`, `CapabilityEvidence`, `PlatformEvidence`, `validate_platform_evidence(raw)`, `merge_platform_evidence(windows, macos)`, and `write_canonical_json(path, value)`.

- [ ] **Step 1: Write failing contract tests**

```python
from skills.WPSComposer.scripts.longform_m0.contracts import (
    REQUIRED_IDS,
    merge_platform_evidence,
    validate_platform_evidence,
)


def passing(platform: str) -> dict:
    return {
        "schemaVersion": 1,
        "probeVersion": "0.8.0-m0.1",
        "platform": platform,
        "wpsVersion": "12.1.test",
        "protocolVersion": 2,
        "resourceManifestVersion": 1,
        "capabilities": [
            {
                "id": capability_id,
                "status": "passed",
                "checks": ["native", "reopened", "refreshed"],
                "artifacts": ["probe.docx", "probe.pdf"],
                "metrics": {},
            }
            for capability_id in range(1, 16)
        ],
        "artifacts": {
            "docx": {"name": "probe.docx", "sha256": "a" * 64},
            "pdf": {"name": "probe.pdf", "sha256": "b" * 64},
        },
        "failures": [],
    }


def test_required_ids_are_exactly_one_through_fourteen():
    assert REQUIRED_IDS == frozenset(range(1, 15))


def test_rejects_duplicate_or_missing_capability_ids():
    raw = passing("macos")
    raw["capabilities"].pop()
    raw["capabilities"].append(dict(raw["capabilities"][0]))
    try:
        validate_platform_evidence(raw)
    except ValueError as exc:
        assert "capability ids" in str(exc)
    else:
        raise AssertionError("invalid evidence was accepted")


def test_optional_svg_does_not_block_go_decision():
    windows = passing("windows")
    macos = passing("macos")
    macos["capabilities"][14]["status"] = "unsupported"
    matrix = merge_platform_evidence(
        validate_platform_evidence(windows),
        validate_platform_evidence(macos),
    )
    assert matrix["decision"] == "go"
    assert matrix["svg"] == "excluded"


def test_required_failure_forces_no_go():
    windows = passing("windows")
    macos = passing("macos")
    macos["capabilities"][9]["status"] = "failed"
    matrix = merge_platform_evidence(
        validate_platform_evidence(windows),
        validate_platform_evidence(macos),
    )
    assert matrix["decision"] == "no-go"
    assert matrix["blockingCapabilities"] == [10]
```

- [ ] **Step 2: Run the focused tests and confirm the import failure**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_contracts.py -q`

Expected: collection fails with `ModuleNotFoundError` for `longform_m0`.

- [ ] **Step 3: Implement the immutable contract**

Use these exact public constants and value domains:

```python
PROBE_VERSION = "0.8.0-m0.1"
SCHEMA_VERSION = 1
PROTOCOL_VERSION = 2
RESOURCE_MANIFEST_VERSION = 1
REQUIRED_IDS = frozenset(range(1, 15))
OPTIONAL_IDS = frozenset({15})
ALL_IDS = REQUIRED_IDS | OPTIONAL_IDS
STATUSES = frozenset({"passed", "failed", "unsupported", "not-run"})
```

`validate_platform_evidence()` must reject unknown root keys, unknown capability keys, duplicate/missing IDs, absolute artifact names, non-hex digests, required `unsupported`, and a nominally passed capability without `native`. Capability IDs 3-15 additionally require `reopened` and `refreshed`; capability 2 is Windows-only and macOS records it as `passed` with check `not-applicable-macos`.

- [ ] **Step 4: Implement deterministic merge semantics**

`merge_platform_evidence()` accepts exactly one `windows` and one `macos` record, sorts entries by numeric ID, sets `decision=no-go` when any ID 1-14 is not passed on either platform, and sets `svg=included` only when ID 15 passes on both platforms. It must omit platform-private metrics from the merged summary.

- [ ] **Step 5: Run tests**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_contracts.py -q`

Expected: all tests pass.

- [ ] **Step 6: Commit the contract**

```bash
git add skills/WPSComposer/scripts/longform_m0 tests/longform_m0/test_contracts.py
git commit -m "Add longform M0 evidence contract"
```

### Task 2: Add host-side artifact and PDF evidence checks

**Files:**
- Create: `skills/WPSComposer/scripts/longform_m0/host_checks.py`
- Create: `tests/longform_m0/test_host_checks.py`

**Interfaces:**
- Consumes: native DOCX/PDF paths and the executor's redacted snapshots.
- Produces: `require_probe_dependencies()`, `validate_native_artifacts()`, `snapshot_pdf(path)`, `sha256_file(path)`, and `prepare_evidence_directory(path)`.

- [ ] **Step 1: Write failing tests for dependency and artifact boundaries**

Tests must prove that a missing `pypdf`, `pdfplumber`, or Pillow dependency fails before an executor callback is invoked; malformed DOCX/PDF artifacts fail; output directories containing prior artifacts are rejected; PDF snapshots expose only page count, normalized MediaBox/CropBox/rotation, and embedded font names.

```python
def test_dependency_failure_happens_before_executor(monkeypatch, tmp_path):
    called = False
    def executor():
        nonlocal called
        called = True
    monkeypatch.setattr(host_checks, "_find_missing_dependencies", lambda: ["pypdf"])
    with pytest.raises(RuntimeError, match="pypdf"):
        host_checks.run_after_dependency_gate(executor)
    assert called is False
```

- [ ] **Step 2: Run and observe the missing-module failure**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_host_checks.py -q`

Expected: collection fails because `host_checks.py` does not exist.

- [ ] **Step 3: Implement the checks using existing validators**

Reuse `artifact_transport.validate_office_package(path, "docx")` and `artifact_transport.validate_pdf(path)`. Read PDF page boxes and fonts through `pypdf`; use `pdfplumber` only for per-page character bounds needed by coordinate evidence. Normalize boxes to four finite point values and rotation to `0|90|180|270`.

- [ ] **Step 4: Add privacy assertions**

`snapshot_pdf()` must never return extracted text. Add a recursive validator that rejects values containing the evidence directory's absolute parent, the WPS staging root, keys matching `text|bookmarkMap|fieldHash|sourcePath|stagingPath|commandLine`, or strings beginning with `/Users/` or a Windows drive prefix.

- [ ] **Step 5: Run focused tests**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_host_checks.py tests/test_artifact_transport.py -q`

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add skills/WPSComposer/scripts/longform_m0/host_checks.py tests/longform_m0/test_host_checks.py
git commit -m "Validate longform M0 native evidence"
```

### Task 3: Extend the closed macOS bridge for one M0 command

**Files:**
- Modify: `skills/WPSComposer/scripts/macos_probe/models.py`
- Modify: `skills/WPSComposer/scripts/macos_probe/runtime.py`
- Modify: `tests/macos_probe/test_models.py`
- Modify: `tests/macos_probe/test_runtime.py`
- Create: `macos/wps-jsapi-probe/addin/writer-longform-m0.js`
- Modify: `macos/wps-jsapi-probe/addin/index.html`
- Modify: `macos/wps-jsapi-probe/addin/bridge-client.js`
- Modify: `macos/wps-jsapi-probe/addin/writer.js`
- Create: `tests/longform_m0/test_addin_assets.py`

**Interfaces:**
- Consumes: `{stagedDocxPath, stagedPdfPath, manifest, protocolVersion, resourceManifestVersion, probeVersion}`.
- Produces: `window.WPSComposerLongformM0.run(params)` returning the raw platform evidence payload.

- [ ] **Step 1: Add failing protocol tests**

Assert `METHOD_COMPONENT["probe_longform_m0"] == "writer"`, cross-component routing fails, the add-in contains no `eval`, `Function(`, `fetch(`, filesystem path, or dynamic method name, and `writer.js` exposes exactly `window.WPSComposerLongformM0.run(params)` through its handler table.

- [ ] **Step 2: Run and confirm failures**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_models.py tests/longform_m0/test_addin_assets.py -q`

Expected: failures identify the missing method and asset.

- [ ] **Step 3: Add the typed route and strict envelope checks**

The JS module must accept exactly these root keys:

```javascript
const M0_KEYS = [
  "manifest", "probeVersion", "protocolVersion",
  "resourceManifestVersion", "stagedDocxPath", "stagedPdfPath"
];
```

Reject protocol versions other than `2`, manifest versions other than `1`, non-empty resource entries, a manifest digest that is not SHA-256 hex, unknown keys, paths not already validated by the host bridge, and a probe version other than `0.8.0-m0.1`. Run the negative handshake before opening or resetting a document.

- [ ] **Step 4: Implement the shared JS envelope helpers**

Provide closed helpers for exact-key comparison, typed protocol failure, path-shape validation, empty-manifest digest binding, and a non-native scaffold result. The scaffold result must state that the native probe has not run and cannot satisfy the platform evidence contract.

- [ ] **Step 5: Run static tests**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_models.py tests/longform_m0/test_addin_assets.py -q`

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/models.py tests/macos_probe/test_models.py macos/wps-jsapi-probe/addin tests/longform_m0/test_addin_assets.py
git commit -m "Add closed Mac longform M0 command"
```

### Task 4: Implement macOS native capabilities 1 and 3-15

**Files:**
- Modify: `macos/wps-jsapi-probe/addin/writer-longform-m0.js`
- Create: `skills/WPSComposer/scripts/longform_m0/macos.py`
- Create: `tests/longform_m0/test_macos.py`

**Interfaces:**
- Consumes: existing `ProbeRuntime`, `LoopbackBridge`, the M0 contract, and two private staged output paths.
- Produces: `run_macos_probe(output_dir: Path, timeout: float = 600) -> Path` returning `platform-evidence.json`.

- [ ] **Step 1: Write failing orchestration tests with fake bridge/runtime objects**

Cover dependency preflight, no-clobber output, canonical empty resource manifest binding, method routing, typed timeout, rejection before document open, artifact validation, evidence validation, registration restoration, staging deletion, and private-path redaction.

- [ ] **Step 2: Run and confirm the missing implementation**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_macos.py -q`

Expected: collection fails for the missing `longform_m0.macos` module.

- [ ] **Step 3: Implement native document construction and lifecycle**

The JS probe must add closed helpers for collection access, safe property reads, bounded attempts, checkpoint cleanup, field result snapshots, UTF-16 offsets, section/page snapshots, and redacted errors. It must then create all test content through Writer JSAPI, save the staged DOCX, close it, reopen it, refresh fields, save again, export the staged PDF, and close only its own document. Record `reopened` and `refreshed` only after the final object-count and field-result snapshots match the expected probe-local invariants.

- [ ] **Step 4: Implement the capability checks exactly**

Use this mapping; each row is one independently reported capability:

| ID | Native action and pass invariant |
|---:|---|
| 1 | Accept valid v2/manifest envelope; reject wrong protocol, manifest version, digest, and unknown key before opening a document. |
| 2 | Report `passed` with `not-applicable-macos`; do not claim Windows ownership. |
| 3 | Enumerate fonts; apply logical CJK/Latin/mono fonts; snapshot effective range fonts; PDF exposes non-empty font resources. |
| 4 | Insert NFC/decomposed Chinese, emoji ZWJ, variation selector, and extension-Han paragraphs; UTF-16 ranges locate every marker after reopen. |
| 5 | Insert a positioned shape and paragraph; WPS point coordinates and normalized PDF character bounds agree within 1 point or report low-confidence evidence without passing the coordinate subcheck. |
| 6 | Create three independent H1-H4 native list definitions; insert, move, and delete headings; numbering results update after refresh and reopen. |
| 7 | Create cover/front/body page sections, Roman-to-Arabic restart, explicit page break, and first/middle/last landscape objects without an empty tail section. |
| 8 | Modify TOC 1/2/3 through stable IDs/enums; create a non-outline TOC title; refresh TOC and prove it does not list itself. |
| 9 | Insert independent figure/table `SEQ` fields and native figure/table indexes; refresh and prove chapter reset/global behavior used by the fixture. |
| 10 | Create deterministic ASCII bookmarks and cross-reference fields; exercise collision retry with forced first-candidate collision; move targets before refresh. |
| 11 | Insert editable Office Math/WPS formula content inside a borderless layout table; prove the formula object remains native after reopen. |
| 12 | Snapshot node IDs to sections, physical page spans, UTF-16 ranges, and per-page fragments; a cross-page table must have at least two page-local fragments. |
| 13 | Fail the second child of a two-child composite, remove only that child's partial object, insert a visible fallback, and successfully execute a later paragraph operation. |
| 14 | Run bounded field convergence in normal generation and one notice-only patch; save/export counts stay within 2 full passes, 1 patch, and 3 exports. |
| 15 | Insert static local SVG with no scripts/external references; save/reopen/export. Record `unsupported` without failing the command when native insertion is unavailable. |

- [ ] **Step 5: Implement bounded execution and cleanup**

Use one bridge deadline covering the complete run. The JS command must restore `Application.DisplayAlerts` and `Application.ScreenUpdating` in nested `finally` blocks. Python must restore registration, remove the WPS staging session, preserve native artifacts in the requested evidence directory, and write a failed evidence record even when the native command fails.

- [ ] **Step 6: Run platform-independent macOS tests**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_macos.py tests/longform_m0/test_addin_assets.py tests/macos_probe -q`

Expected: all tests pass without launching WPS.

- [ ] **Step 7: Commit**

```bash
git add macos/wps-jsapi-probe/addin/writer-longform-m0.js skills/WPSComposer/scripts/longform_m0/macos.py tests/longform_m0/test_macos.py
git commit -m "Probe Mac WPS longform semantics"
```

### Task 5: Implement the supervised Windows native probe

**Files:**
- Create: `skills/WPSComposer/scripts/longform_m0/windows.py`
- Create: `tests/longform_m0/test_windows.py`

**Interfaces:**
- Consumes: output directory, total deadline, and lazy-loaded `win32com.client`/process adapters.
- Produces: `run_windows_probe(output_dir: Path, timeout: float = 600) -> Path` and private `_worker_main(request_path, result_path)`.

- [ ] **Step 1: Write failing supervisor tests**

Use fake process identities and fake COM adapters to prove: `DispatchEx` is required; an ordinary `Dispatch` path is never called; a pre-existing WPS PID is never owned; the worker must report a newly created PID plus immutable identity; timeout first sends cooperative cancellation, waits at most five seconds, and terminates only the still-matching owned process tree; missing ownership is a fatal capability-2 failure.

- [ ] **Step 2: Run and confirm the missing implementation**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_windows.py -q`

Expected: collection fails for the missing Windows module while the rest of the suite remains importable on macOS/Linux.

- [ ] **Step 3: Implement the supervisor/worker boundary**

The supervisor writes a canonical private request, launches `.venv` Python with `-m skills.WPSComposer.scripts.longform_m0.windows --worker`, receives structured progress over a private file/pipe, and validates PID, executable identity, creation time, and parent relationship before treating any WPS process as owned. Import `pythoncom`, `win32api`, `win32com.client`, and Windows process helpers only inside Windows runtime functions.

- [ ] **Step 4: Implement capabilities 1-15 through COM**

Use the same fixture markers, numbering schemes, section lifecycle, fields, bookmarks, formulas, pagination fragments, checkpoint failure, convergence bounds, and SVG decision as Task 4. Create objects through COM only. Capability 2 additionally runs the real cooperative-timeout fixture while a separate user WPS document is open and proves that document/PID survives.

- [ ] **Step 5: Add Windows artifact and reopen verification**

The worker saves DOCX, closes it, reopens it with the dedicated instance, refreshes fields, saves, exports PDF, closes its documents, quits its WPS instance, and emits the shared evidence payload. The supervisor validates artifacts and evidence after the worker exits.

- [ ] **Step 6: Run cross-platform unit tests**

Run: `.venv/bin/python -m pytest tests/longform_m0/test_windows.py tests/test_com_lifecycle.py tests/test_windows_conversion.py -q`

Expected on macOS: fake-adapter tests pass and existing real-COM tests remain skipped as before.

- [ ] **Step 7: Commit**

```bash
git add skills/WPSComposer/scripts/longform_m0/windows.py tests/longform_m0/test_windows.py
git commit -m "Add supervised Windows longform M0 probe"
```

### Task 6: Add CLI, evidence verifier, and operator documentation

**Files:**
- Create: `skills/WPSComposer/scripts/longform_m0/__main__.py`
- Create: `docs/longform-m0.md`
- Modify: `tests/test_documentation.py`

**Interfaces:**
- Consumes: platform output directories or two completed `platform-evidence.json` files.
- Produces: CLI commands `--platform macos`, `--platform windows`, and `--platform verify --windows-evidence ... --macos-evidence ...`.

- [ ] **Step 1: Write failing CLI and documentation tests**

Assert the parser requires explicit platform, positive timeout, new output directory, and both evidence paths for verification. Documentation tests require the exact 1-15 matrix, current go/no-go rule, unrestricted local-run requirement for macOS WPS container access, Windows prerequisites, recovery paths, and evidence privacy rules.

- [ ] **Step 2: Run and confirm failures**

Run: `.venv/bin/python -m pytest tests/test_documentation.py tests/longform_m0 -q`

Expected: failures name the absent CLI/doc sections.

- [ ] **Step 3: Implement the explicit CLI**

```text
.venv/bin/python -m skills.WPSComposer.scripts.longform_m0 \
  --platform macos --output-dir build/longform-m0/macos-<run-id> --timeout 600

.venv/Scripts/python.exe -m skills.WPSComposer.scripts.longform_m0 \
  --platform windows --output-dir build/longform-m0/windows-<run-id> --timeout 600

.venv/bin/python -m skills.WPSComposer.scripts.longform_m0 \
  --platform verify \
  --windows-evidence build/longform-m0/windows-<run-id>/platform-evidence.json \
  --macos-evidence build/longform-m0/macos-<run-id>/platform-evidence.json \
  --output-dir build/longform-m0/matrix-<run-id>
```

The CLI exits `0` only for a passed platform probe or a merged `go` decision; it exits `1` for native failure/no-go and `2` for invalid arguments/evidence.

- [ ] **Step 4: Document recovery and visual evidence capture**

Describe how to inspect registration recovery before rerunning, how to verify no unrelated WPS document closed, and how to capture representative cover/TOC, numbering, landscape, indexes, formula, cross-page table, and degradation pages as screenshots. Record screenshots by relative filename only in evidence JSON.

- [ ] **Step 5: Run tests**

Run: `.venv/bin/python -m pytest tests/test_documentation.py tests/longform_m0 -q`

Expected: all tests pass.

- [ ] **Step 6: Commit**

```bash
git add skills/WPSComposer/scripts/longform_m0/__main__.py docs/longform-m0.md tests/test_documentation.py
git commit -m "Document longform M0 verification"
```

### Task 7: Run both native gates and make the go/no-go decision

**Files:**
- Runtime only: `build/longform-m0/macos-<run-id>/`
- Runtime only: `build/longform-m0/windows-<run-id>/`
- Runtime only: `build/longform-m0/matrix-<run-id>/`
- Modify only after evidence exists: `docs/superpowers/specs/2026-08-20-wpscomposer-longform-engine-design.md`

**Interfaces:**
- Consumes: completed Tasks 1-6 and real WPS installations on both platforms.
- Produces: two native evidence bundles, merged `capability-matrix.json`, screenshots, and an evidence-based design review.

- [ ] **Step 1: Run the complete platform-independent suite**

Run: `.venv/bin/python -m pytest -v`

Expected: all platform-independent tests pass; platform-specific skips are explained.

- [ ] **Step 2: Run macOS twice from an unrestricted local shell**

Use two new output directories and verify both runs have identical capability statuses, protocol/manifest versions, and structural counts. Artifact hashes may differ because WPS may serialize timestamps; semantic evidence must match.

- [ ] **Step 3: Inspect macOS native artifacts and screenshots**

Open the DOCX in WPS, refresh all fields once more, and inspect the exported PDF in Preview. Check columns/wrapping, heading renumbering, compact TOC spacing, non-self-listing title, page labels, landscape transitions, fields, formula editability, cross-page fragments, fallback notices, and fonts.

- [ ] **Step 4: Run Windows twice on the recorded WPS baseline**

Use fresh evidence directories, keep an unrelated user WPS document open during the timeout fixture, and confirm the dedicated worker owns and closes only its own process/document.

- [ ] **Step 5: Inspect Windows native artifacts and screenshots**

Perform the same WPS reopen/refresh and PDF visual checks as macOS. Record WPS version, Windows version, hardware, page count, operation count, and stage timings; do not record usernames or paths.

- [ ] **Step 6: Merge evidence and enforce the gate**

Run the `--platform verify` command. If the decision is `no-go`, stop 0.8.0 implementation and document only the failed capability plus native evidence. If the decision is `go`, run one targeted evidence-based review of the design and adjust only contracts contradicted by real native behavior before M1.

- [ ] **Step 7: Update the design status and commit only the status change**

The status must name both WPS versions, evidence run IDs, capability result, SVG decision, and whether M1 is permitted. Do not commit `build/` artifacts.

```bash
git add docs/superpowers/specs/2026-08-20-wpscomposer-longform-engine-design.md
git commit -m "Record longform M0 gate result"
```

## Plan Self-Review

- Every M0 requirement 1-15 maps to Tasks 4 and 5 and to one shared evidence ID.
- Save/close/reopen/refresh/export is mandatory in both native implementations and cannot be inferred from unit tests.
- Windows-only ownership is explicit; macOS records it as not applicable without making a false capability claim.
- SVG is optional and has an explicit exclusion result.
- No task changes the public API or begins M1.
- Protocol, staging, privacy, timeout, field convergence, pagination fragments, and visual inspection all have executable acceptance steps.
- The plan contains no unresolved implementation placeholders.
