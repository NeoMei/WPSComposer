# WPS Sandbox Staging and Office-to-PDF Conversion Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a stable `convert_to_pdf()` API for Word, Excel, and PowerPoint files on Windows and macOS while moving every macOS WPS-produced artifact through a private WPS-container staging directory before atomic publication.

**Architecture:** A platform-neutral facade validates source and destination paths, then delegates to a Windows COM converter or a macOS WPS JSAPI converter. Both backends write a staged PDF, validate it, and use one destination-local atomic publisher; the macOS probe also adopts the same WPS-container transport so the real feasibility gate proves that no save panel or broader permission is required.

**Tech Stack:** Python 3, pytest, Windows `pywin32` COM automation, macOS WPS JSAPI add-ins, authenticated loopback bridge, Node.js 20+, `wpsjs` 2.2.3.

## Global Constraints

- Support source suffixes case-insensitively: `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx` only.
- The public signature is `convert_to_pdf(source: str, output: Optional[str] = None, *, overwrite: bool = False) -> str`.
- Default output is the source sibling with the same stem and `.pdf`; a supplied output must end in `.pdf`.
- Existing output raises `FileExistsError` unless `overwrite=True`; missing source raises `FileNotFoundError`; invalid suffixes raise `ValueError`.
- Runtime failures raise `ConversionError` with stable `code`, `source`, `component`, `backend`, and `message` fields.
- WPS or Office remains the sole document serializer; Python may only copy, validate, flush, and atomically publish bytes.
- macOS staging root is `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer/{random-session-id}/` with mode `0700`.
- macOS source and output paths exposed to WPS must both be inside that session directory.
- Publication must validate the staged artifact, copy to a destination-local temporary file, flush and `fsync`, call `os.replace`, then validate the final artifact.
- No Full Disk Access, Computer Use, AppleScript, Accessibility, or automatic modal-dialog interaction.
- The macOS public backend stays disabled unless the real conversion gate passes twice consecutively without a save panel or new permission prompt.
- Windows Writer uses `Document.ExportAsFixedFormat(output_path, 17)`; Presentation uses `Presentation.SaveAs(output_path, 32)`; Spreadsheet uses workbook-level `Workbook.ExportAsFixedFormat(0, output_path)` so every visible worksheet is included.
- Opened sources are always closed without saving changes, including failure paths.
- Existing generation, editing, and PDF-editing APIs retain their current behavior.

---

## File Structure

- Create `skills/WPSComposer/scripts/artifact_transport.py`: PDF/package validation and destination-local atomic publication only.
- Create `skills/WPSComposer/scripts/conversion.py`: public request validation, stable errors, platform routing, and backend enablement policy.
- Create `skills/WPSComposer/scripts/windows_conversion.py`: Windows COM source opening and format-specific PDF export.
- Create `skills/WPSComposer/scripts/macos_probe/conversion.py`: one-component macOS WPS runtime orchestration and typed conversion command execution.
- Modify `skills/WPSComposer/scripts/macos_probe/runtime.py`: create, expose, and remove the private WPS-container session.
- Modify `skills/WPSComposer/scripts/macos_probe/runner.py`: stage every smoke artifact, atomically publish it, and report only final paths.
- Modify `skills/WPSComposer/scripts/macos_probe/models.py`: closed allowlist for three conversion methods.
- Modify `macos/wps-jsapi-probe/addin/{writer,spreadsheet,presentation}.js`: open staged sources, export staged PDFs, close without saving.
- Modify `skills/WPSComposer/scripts/wps_engine.py` and `skills/WPSComposer/__init__.py`: export the stable conversion API.
- Create `tests/test_artifact_transport.py`, `tests/test_conversion.py`, `tests/test_windows_conversion.py`, and `tests/macos_probe/test_conversion.py`; extend existing macOS probe tests.
- Modify `README.md`: document supported formats, errors, overwrite behavior, and platform gates.

---

### Task 1: Validated Atomic Artifact Publication

**Files:**
- Create: `skills/WPSComposer/scripts/artifact_transport.py`
- Create: `tests/test_artifact_transport.py`

**Interfaces:**
- Consumes: filesystem `Path` objects and a validator callback.
- Produces: `ArtifactValidationError`, `validate_pdf(path: Path) -> None`, `validate_office_package(path: Path, format_name: str) -> None`, and `publish_artifact(staged: Path, destination: Path, *, overwrite: bool, validator: Callable[[Path], None]) -> Path`.

- [ ] **Step 1: Write failing validator and publisher tests**

```python
def test_publish_validates_both_copies_and_replaces_atomically(tmp_path, monkeypatch):
    staged = tmp_path / "stage" / "result.pdf"
    staged.parent.mkdir()
    staged.write_bytes(b"%PDF-1.7\n" + b"x" * 2048)
    destination = tmp_path / "out" / "result.pdf"
    destination.parent.mkdir()
    seen = []
    monkeypatch.setattr(os, "replace", lambda source, target: (
        seen.append((Path(source), Path(target))),
        Path(target).write_bytes(Path(source).read_bytes()),
        Path(source).unlink(),
    ))
    result = publish_artifact(
        staged, destination, overwrite=False,
        validator=lambda path: (seen.append(path), validate_pdf(path)),
    )
    assert result == destination.resolve()
    assert seen[0] == staged.resolve()
    assert seen[-1] == destination.resolve()
    assert seen[1][1] == destination.resolve()
    assert seen[1][0].parent == destination.parent.resolve()

def test_publish_failure_leaves_no_partial_destination(tmp_path):
    staged = tmp_path / "bad.pdf"
    staged.write_bytes(b"broken")
    destination = tmp_path / "result.pdf"
    with pytest.raises(ArtifactValidationError):
        publish_artifact(staged, destination, overwrite=False, validator=validate_pdf)
    assert not destination.exists()
    assert not list(tmp_path.glob(".wpscomposer-*.tmp"))
```

- [ ] **Step 2: Run the tests and verify the missing module failure**

Run: `python -m pytest tests/test_artifact_transport.py -q`

Expected: FAIL during collection with `ModuleNotFoundError: No module named 'skills.WPSComposer.scripts.artifact_transport'`.

- [ ] **Step 3: Implement strict validators and the destination-local publisher**

```python
class ArtifactValidationError(RuntimeError):
    pass

def validate_pdf(path: Path) -> None:
    target = path.expanduser().resolve()
    if not target.is_file() or target.stat().st_size < 1024:
        raise ArtifactValidationError(f"PDF artifact is missing or too small: {target}")
    with target.open("rb") as stream:
        if not stream.read(5).startswith(b"%PDF-"):
            raise ArtifactValidationError(f"Invalid PDF signature: {target}")

def publish_artifact(staged, destination, *, overwrite, validator):
    source = Path(staged).expanduser().resolve()
    target = Path(destination).expanduser().resolve()
    validator(source)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and not overwrite:
        raise FileExistsError(f"Output already exists: {target}")
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb", dir=target.parent, prefix=".wpscomposer-", suffix=".tmp", delete=False
        ) as stream:
            temporary = Path(stream.name)
            with source.open("rb") as incoming:
                shutil.copyfileobj(incoming, stream)
            stream.flush()
            os.fsync(stream.fileno())
        validator(temporary)
        os.replace(temporary, target)
        temporary = None
        validator(target)
        return target
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
```

- [ ] **Step 4: Run publication tests**

Run: `python -m pytest tests/test_artifact_transport.py -q`

Expected: all tests PASS, including invalid PDF, corrupt ZIP package, overwrite refusal, overwrite success, and cleanup after validator failure.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/artifact_transport.py tests/test_artifact_transport.py
git commit -m "Add validated atomic artifact publication"
```

### Task 2: Private WPS-Container Staging Sessions

**Files:**
- Modify: `skills/WPSComposer/scripts/macos_probe/runtime.py`
- Modify: `tests/macos_probe/test_runtime.py`

**Interfaces:**
- Consumes: `ProbeRuntime` lifecycle and optional test-only `staging_root: Path`.
- Produces: `WPS_STAGING_ROOT`, `create_staging_session(root: Path = WPS_STAGING_ROOT) -> Path`, and `ProbeRuntime.staging_dir: Path`.

- [ ] **Step 1: Add failing session security and cleanup tests**

```python
def test_staging_session_is_private_and_removed_on_close(tmp_path):
    root = tmp_path / "WPSComposer"
    probe = runtime.ProbeRuntime(
        tmp_path, tmp_path / "runtime", "http://127.0.0.1:45678", "token",
        wps_app=tmp_path / "wpsoffice.app", staging_root=root,
    )
    probe.__enter__()
    session = probe.staging_dir
    assert session.parent == root.resolve()
    assert stat.S_IMODE(session.stat().st_mode) == 0o700
    (session / "artifact.pdf").write_bytes(b"data")
    probe.close()
    assert not session.exists()

def test_default_staging_root_is_inside_wps_container():
    assert str(runtime.WPS_STAGING_ROOT).endswith(
        "Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer"
    )
```

- [ ] **Step 2: Run the focused runtime tests**

Run: `python -m pytest tests/macos_probe/test_runtime.py -q`

Expected: FAIL because `staging_root` and `staging_dir` do not exist.

- [ ] **Step 3: Implement session allocation and unconditional cleanup**

```python
WPS_STAGING_ROOT = (
    Path.home() / "Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer"
)

def create_staging_session(root: Path = WPS_STAGING_ROOT) -> Path:
    parent = root.expanduser().resolve()
    parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    os.chmod(parent, 0o700)
    session = Path(tempfile.mkdtemp(prefix="session-", dir=parent))
    os.chmod(session, 0o700)
    return session

# ProbeRuntime.__enter__ allocates self.staging_dir.
# ProbeRuntime.close() calls shutil.rmtree(self.staging_dir, ignore_errors=True)
# after WPS processes are terminated and before registration recovery is removed.
```

- [ ] **Step 4: Run runtime tests**

Run: `python -m pytest tests/macos_probe/test_runtime.py -q`

Expected: all runtime tests PASS and existing process/registration ownership behavior remains unchanged.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/runtime.py tests/macos_probe/test_runtime.py
git commit -m "Add private WPS container staging sessions"
```

### Task 3: Move Phase 0 Artifacts Through Staging

**Files:**
- Modify: `skills/WPSComposer/scripts/macos_probe/runner.py`
- Modify: `tests/macos_probe/test_runner.py`

**Interfaces:**
- Consumes: `ProbeRuntime.staging_dir`, `publish_artifact()`, package/PDF validators.
- Produces: Phase 0 reports with `artifactTransport: "wps-container-staging"`, final output paths only, and no staging leaks.

- [ ] **Step 1: Rewrite the command test around staged and final paths**

```python
def test_execute_commands_publishes_staged_outputs(tmp_path):
    staging = tmp_path / "container" / "session"
    output = tmp_path / "output"
    staging.mkdir(parents=True)
    output.mkdir()
    report = run_fake_commands(staging=staging, output=output)
    assert report["artifactTransport"] == "wps-container-staging"
    assert {Path(item["path"]).parent for item in report["artifacts"]} == {output}
    assert not any(staging.iterdir())
    serialized = json.dumps(report)
    assert str(staging) not in serialized
```

- [ ] **Step 2: Run the focused runner tests**

Run: `python -m pytest tests/macos_probe/test_runner.py -q`

Expected: FAIL because commands currently write directly to `output`.

- [ ] **Step 3: Stage, validate, publish, and sanitize each artifact**

```python
for component, method, filename, format_name in OUTPUT_COMMANDS:
    staged_path = policy.require_allowed(staging_dir / filename)
    final_path = output_dir / filename
    command = bridge.issue(component, method, {
        "outputPath": str(staged_path),
        "imagePath": str(allowed_image),
        "imageUrl": WRITER_IMAGE_URL,
        "sourcePath": str(staging_dir / "smoke-pdf-source.docx"),
    })
    # After the result path equals staged_path:
    wait_for_artifact(staged_path, format_name, timeout=min(timeout, 10))
    validator = validate_pdf if format_name == "pdf" else (
        lambda path, name=format_name: validate_office_package(path, name)
    )
    published = publish_artifact(
        staged_path, final_path, overwrite=False, validator=validator,
    )
    artifacts.append({"format": format_name, "path": str(published), "size": published.stat().st_size})
```

Add `artifactTransport` at the execution/report level and ensure failure messages replace the private session prefix with `<wps-staging>` before JSON serialization.

- [ ] **Step 4: Run all macOS probe unit tests**

Run: `python -m pytest tests/macos_probe -q`

Expected: all tests PASS; reports contain only final artifact paths and cleanup removes every staged file.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/runner.py tests/macos_probe/test_runner.py
git commit -m "Stage and atomically publish Mac probe artifacts"
```

### Task 4: Public Conversion Request and Error Contract

**Files:**
- Create: `skills/WPSComposer/scripts/conversion.py`
- Create: `tests/test_conversion.py`
- Modify: `skills/WPSComposer/scripts/wps_engine.py`
- Modify: `skills/WPSComposer/__init__.py`

**Interfaces:**
- Consumes: `sys.platform`, Windows and macOS backend callables.
- Produces: `ConversionError` and `convert_to_pdf(source: str, output: Optional[str] = None, *, overwrite: bool = False) -> str`.

- [ ] **Step 1: Add failing public-contract and routing tests**

```python
@pytest.mark.parametrize("suffix", ["doc", "DOCX", "xls", "XLSX", "ppt", "PPTX"])
def test_convert_routes_all_supported_suffixes(tmp_path, monkeypatch, suffix):
    source = tmp_path / f"input.{suffix}"
    source.write_bytes(b"source")
    observed = []
    monkeypatch.setattr(conversion, "_select_backend", lambda: (
        "test-backend", lambda request: observed.append(request) or write_pdf(request.output)
    ))
    result = convert_to_pdf(str(source))
    assert result == str((tmp_path / "input.pdf").resolve())
    assert observed[0].source == source.resolve()

def test_conversion_error_exposes_stable_fields(tmp_path, monkeypatch):
    source = tmp_path / "input.docx"
    source.write_bytes(b"source")
    monkeypatch.setattr(conversion, "_select_backend", lambda: (
        "mac-wps-jsapi", lambda request: (_ for _ in ()).throw(RuntimeError("boom"))
    ))
    with pytest.raises(ConversionError) as caught:
        convert_to_pdf(str(source))
    assert caught.value.code == "CONVERSION_FAILED"
    assert caught.value.source == str(source.resolve())
    assert caught.value.component == "writer"
    assert caught.value.backend == "mac-wps-jsapi"
    assert caught.value.message == "boom"
```

- [ ] **Step 2: Run contract tests**

Run: `python -m pytest tests/test_conversion.py -q`

Expected: FAIL because the public conversion module does not exist.

- [ ] **Step 3: Implement validation, dispatch, and stable exceptions**

```python
@dataclass(frozen=True)
class ConversionRequest:
    source: Path
    output: Path
    component: str
    overwrite: bool

class ConversionError(RuntimeError):
    def __init__(self, *, code, source, component, backend, message):
        super().__init__(message)
        self.code, self.source = code, source
        self.component, self.backend, self.message = component, backend, message

def convert_to_pdf(source, output=None, *, overwrite=False):
    request = _build_request(source, output, overwrite=overwrite)
    backend_name, backend = _select_backend()
    try:
        result = backend(request)
        validate_pdf(result)
        return str(Path(result).resolve())
    except (FileNotFoundError, FileExistsError, ValueError, ConversionError):
        raise
    except Exception as exc:
        raise ConversionError(
            code="CONVERSION_FAILED", source=str(request.source),
            component=request.component, backend=backend_name, message=str(exc),
        ) from exc
```

`_build_request` maps Word suffixes to `writer`, Excel to `spreadsheet`, and PowerPoint to `presentation`; it performs all path and overwrite checks before selecting a backend. `_select_backend` supports `win32` and `darwin`, and raises `ConversionError(code="BACKEND_UNAVAILABLE", source=str(request.source), component=request.component, backend=sys.platform, message=f"Unsupported platform: {sys.platform}")` elsewhere.

- [ ] **Step 4: Export the API and run tests**

Run: `python -m pytest tests/test_conversion.py tests/test_public_api.py -q`

Expected: all tests PASS and both `from skills.WPSComposer import convert_to_pdf` and the compatibility `wps_engine` facade expose the same function.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/conversion.py skills/WPSComposer/scripts/wps_engine.py skills/WPSComposer/__init__.py tests/test_conversion.py tests/test_public_api.py
git commit -m "Add stable Office to PDF conversion API"
```

### Task 5: Windows COM Conversion Backend

**Files:**
- Create: `skills/WPSComposer/scripts/windows_conversion.py`
- Create: `tests/test_windows_conversion.py`
- Modify: `skills/WPSComposer/scripts/conversion.py`

**Interfaces:**
- Consumes: `ConversionRequest`, composers `WriterComposer`, `SheetComposer`, `SlideComposer`, and `publish_artifact()`.
- Produces: `convert_windows(request: ConversionRequest) -> Path`.

- [ ] **Step 1: Add failing fake-COM API-selection tests**

```python
def test_writer_uses_export_as_fixed_format(fake_composer, request):
    convert_windows(request("input.docx", "writer"), composers={"writer": fake_composer})
    fake_composer.doc.ExportAsFixedFormat.assert_called_once_with(ANY_STAGED_PATH, 17)
    fake_composer.doc.Close.assert_called_once_with(False)

def test_spreadsheet_exports_the_whole_workbook(fake_composer, request):
    convert_windows(request("input.xlsx", "spreadsheet"), composers={"spreadsheet": fake_composer})
    fake_composer.doc.ExportAsFixedFormat.assert_called_once_with(0, ANY_STAGED_PATH)
    assert not fake_composer.doc.Worksheets.called

def test_presentation_uses_save_as_pdf(fake_composer, request):
    convert_windows(request("input.pptx", "presentation"), composers={"presentation": fake_composer})
    fake_composer.doc.SaveAs.assert_called_once_with(ANY_STAGED_PATH, 32)
```

- [ ] **Step 2: Run Windows backend tests**

Run: `python -m pytest tests/test_windows_conversion.py -q`

Expected: FAIL because `windows_conversion` does not exist.

- [ ] **Step 3: Implement format-specific COM export with guaranteed close**

```python
def convert_windows(request, composers=None):
    classes = composers or {
        "writer": WriterComposer,
        "spreadsheet": SheetComposer,
        "presentation": SlideComposer,
    }
    with tempfile.TemporaryDirectory(prefix="wpscomposer-convert-") as directory:
        staged = Path(directory) / "converted.pdf"
        with classes[request.component].open_document(
            str(request.source), read_only=True, visible=False
        ) as composer:
            if request.component == "writer":
                composer.doc.ExportAsFixedFormat(str(staged), 17)
            elif request.component == "spreadsheet":
                composer.doc.ExportAsFixedFormat(0, str(staged))
            else:
                composer.doc.SaveAs(str(staged), 32)
        return publish_artifact(
            staged, request.output, overwrite=request.overwrite, validator=validate_pdf,
        )
```

Use the context manager as the single owner so `BaseComposer.close(save_changes=False)` closes the document and application in every path.

- [ ] **Step 4: Run Windows and facade tests**

Run: `python -m pytest tests/test_windows_conversion.py tests/test_conversion.py -q`

Expected: all tests PASS, including export failure cleanup and workbook-level multi-sheet API selection.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/windows_conversion.py skills/WPSComposer/scripts/conversion.py tests/test_windows_conversion.py
git commit -m "Add Windows Office to PDF conversion backend"
```

### Task 6: Closed macOS Conversion Protocol and Add-in Handlers

**Files:**
- Modify: `skills/WPSComposer/scripts/macos_probe/models.py`
- Modify: `macos/wps-jsapi-probe/addin/writer.js`
- Modify: `macos/wps-jsapi-probe/addin/spreadsheet.js`
- Modify: `macos/wps-jsapi-probe/addin/presentation.js`
- Modify: `tests/macos_probe/test_models.py`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: bridge commands with `sourcePath` and `outputPath` inside one allowed session.
- Produces: `convert_writer_pdf`, `convert_workbook_pdf`, and `convert_presentation_pdf` result objects `{path: outputPath}`.

- [ ] **Step 1: Add failing allowlist and static handler tests**

```python
@pytest.mark.parametrize("component,method", [
    ("writer", "convert_writer_pdf"),
    ("spreadsheet", "convert_workbook_pdf"),
    ("presentation", "convert_presentation_pdf"),
])
def test_conversion_methods_are_component_scoped(component, method):
    assert ProbeCommand.create(component, method, {}).method == method
    wrong = next(name for name in COMPONENTS if name != component)
    with pytest.raises(ProtocolError):
        ProbeCommand.create(wrong, method, {})

def test_addins_open_export_and_close_sources_without_saving():
    assert_writer_handler("Documents.Open", "ExportAsFixedFormat", "Close(0)")
    assert_spreadsheet_handler("Workbooks.Open", "ExportAsFixedFormat", "Close(false)")
    assert_presentation_handler("Presentations.Open", "SaveAs(outputPath, 32)", "Close()")
```

- [ ] **Step 2: Run protocol/add-in tests**

Run: `python -m pytest tests/macos_probe/test_models.py tests/macos_probe/test_addin_assets.py -q`

Expected: FAIL because conversion methods are rejected and handlers are absent.

- [ ] **Step 3: Add exactly three allowlisted methods and handlers**

```javascript
function convertWriterPdf(params) {
  const sourcePath = requirePath(params, "sourcePath");
  const outputPath = requirePath(params, "outputPath");
  const document = Application.Documents.Open(sourcePath, false, true);
  try {
    document.ExportAsFixedFormat(outputPath, 17, false, 0, 0);
    return {path: outputPath};
  } finally {
    document.Close(0);
  }
}

function convertWorkbookPdf(params) {
  const workbook = Application.Workbooks.Open(requireOutputPath({outputPath: params.sourcePath}), 0, true);
  try {
    workbook.ExportAsFixedFormat(0, requireOutputPath(params));
    return {path: params.outputPath};
  } finally {
    workbook.Close(false);
  }
}

function convertPresentationPdf(params) {
  const sourcePath = requirePath(params, "sourcePath");
  const outputPath = requirePath(params, "outputPath");
  const presentation = Application.Presentations.Open(sourcePath, true, false, false);
  try {
    presentation.SaveAs(outputPath, 32);
    return {path: outputPath};
  } finally {
    presentation.Close();
  }
}
```

Register only these names in the existing `handlers` maps; arbitrary method names remain rejected.

- [ ] **Step 4: Run protocol/add-in tests**

Run: `python -m pytest tests/macos_probe/test_models.py tests/macos_probe/test_addin_assets.py -q`

Expected: all tests PASS, including unknown-method rejection and close-without-save checks.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/models.py macos/wps-jsapi-probe/addin/writer.js macos/wps-jsapi-probe/addin/spreadsheet.js macos/wps-jsapi-probe/addin/presentation.js tests/macos_probe/test_models.py tests/macos_probe/test_addin_assets.py
git commit -m "Add typed Mac WPS PDF conversion commands"
```

### Task 7: macOS WPS Conversion Backend

**Files:**
- Create: `skills/WPSComposer/scripts/macos_probe/conversion.py`
- Create: `tests/macos_probe/test_conversion.py`
- Modify: `skills/WPSComposer/scripts/conversion.py`

**Interfaces:**
- Consumes: `ConversionRequest`, `ProbeRuntime`, `LoopbackBridge`, typed conversion commands, and atomic publisher.
- Produces: `convert_macos(request: ConversionRequest) -> Path`.

- [ ] **Step 1: Add failing orchestration tests with fake runtime/bridge**

```python
def test_macos_copies_source_and_issues_only_staged_paths(tmp_path):
    request = make_request(tmp_path, "input.docx", "writer")
    result, command = run_with_fakes(request, staging_root=tmp_path / "container")
    assert command.method == "convert_writer_pdf"
    assert Path(command.params["sourcePath"]).parent == Path(command.params["outputPath"]).parent
    assert "container" in command.params["sourcePath"]
    assert result == request.output

def test_macos_failure_publishes_no_partial_pdf(tmp_path):
    request = make_request(tmp_path, "input.xlsx", "spreadsheet")
    with pytest.raises(RuntimeError, match="password or modal"):
        run_with_fakes(request, command_error="password or modal")
    assert not request.output.exists()
```

- [ ] **Step 2: Run macOS backend tests**

Run: `python -m pytest tests/macos_probe/test_conversion.py -q`

Expected: FAIL because `macos_probe.conversion` does not exist.

- [ ] **Step 3: Implement one-component staged conversion orchestration**

```python
METHODS = {
    "writer": "convert_writer_pdf",
    "spreadsheet": "convert_workbook_pdf",
    "presentation": "convert_presentation_pdf",
}

def convert_macos(request, *, runtime_factory=ProbeRuntime, bridge_factory=LoopbackBridge):
    with bridge_factory(ORIGINS) as bridge:
        repository_root = Path(__file__).resolve().parents[4]
        probe_root = repository_root / "macos/wps-jsapi-probe"
        with tempfile.TemporaryDirectory(prefix="wpscomposer-macos-convert-") as directory:
            runtime = runtime_factory(
                probe_root,
                Path(directory) / "runtime",
                bridge.url,
                bridge.token,
            )
        with runtime:
            runtime.prepare_profiles()
            runtime.start_servers()
            staged_source = runtime.staging_dir / request.source.name
            staged_pdf = runtime.staging_dir / "converted.pdf"
            shutil.copy2(request.source, staged_source)
            runtime.activate_component(request.component, fixture=staged_source)
            bridge.wait_registered({request.component}, 30)
            command = bridge.issue(request.component, METHODS[request.component], {
                "sourcePath": str(staged_source), "outputPath": str(staged_pdf),
            })
            result = bridge.wait_result(command.id, 90)
            if not result.ok:
                raise RuntimeError(str((result.error or {}).get("message", "WPS conversion failed")))
            if Path(str(result.value.get("path", ""))).resolve() != staged_pdf.resolve():
                raise ProtocolError("WPS returned a path outside the conversion session")
            return publish_artifact(
                staged_pdf, request.output, overwrite=request.overwrite, validator=validate_pdf,
            )
```

Extend `ProbeRuntime.activate_component` with an optional `fixture` parameter so conversion opens the copied source while Phase 0 retains official fixtures. Start all three add-in servers because current WPS registration is shared, but activate only the selected component.

- [ ] **Step 4: Run macOS backend and facade tests**

Run: `python -m pytest tests/macos_probe/test_conversion.py tests/test_conversion.py -q`

Expected: all tests PASS; every WPS-visible path is inside `runtime.staging_dir`, failures leave no final or temporary PDF, and registration is restored.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/conversion.py skills/WPSComposer/scripts/macos_probe/runtime.py skills/WPSComposer/scripts/conversion.py tests/macos_probe/test_conversion.py tests/macos_probe/test_runtime.py
git commit -m "Add Mac WPS Office to PDF conversion backend"
```

### Task 8: Documentation, Full Regression, and Real macOS Gate

**Files:**
- Modify: `README.md`
- Create: `macos/wps-jsapi-probe/fixtures/README.md`
- Create: `tests/fixtures/office_conversion/.gitkeep`
- Modify: `docs/superpowers/specs/2026-07-17-macos-wps-sandbox-staging-design.md`

**Interfaces:**
- Consumes: completed public API and both backends.
- Produces: user-facing usage docs, reproducible fixture instructions, and recorded GO/NO-GO evidence.

- [ ] **Step 1: Document the exact public behavior**

```markdown
from skills.WPSComposer import convert_to_pdf

pdf = convert_to_pdf("quarterly.xlsx")
pdf = convert_to_pdf("deck.pptx", "exports/deck.pdf", overwrite=True)
```

State the six supported suffixes, default output naming, `FileExistsError`, `ConversionError` fields, Windows COM prerequisite, macOS WPS prerequisite, and the fact that macOS requires no Full Disk Access or GUI-control permission.

- [ ] **Step 2: Run the complete deterministic suite**

Run: `python -m pytest -q`

Expected: all tests PASS with no skips newly introduced by this feature.

- [ ] **Step 3: Run Phase 0 twice into fresh output directories**

Run:

```bash
python -m skills.WPSComposer.scripts.macos_probe /tmp/wpscomposer-phase0-staged-run-1
python -m skills.WPSComposer.scripts.macos_probe /tmp/wpscomposer-phase0-staged-run-2
```

Expected: both commands exit 0; each report is `passed`, contains all four valid final artifacts and `"artifactTransport": "wps-container-staging"`; no WPS save panel or new macOS permission prompt appears; the WPS staging root contains no session left by either run.

- [ ] **Step 4: Exercise all six real conversions**

Run:

```bash
python - <<'PY'
from pathlib import Path
from skills.WPSComposer import convert_to_pdf

root = Path("tests/fixtures/office_conversion").resolve()
out = Path("/tmp/wpscomposer-conversion-gate").resolve()
out.mkdir(parents=True, exist_ok=True)
for name in ("sample.doc", "sample.docx", "sample.xls", "sample.xlsx", "sample.ppt", "sample.pptx"):
    print(convert_to_pdf(str(root / name), str(out / f"{Path(name).stem}-{Path(name).suffix[1:]}.pdf")))
PY
```

Expected: six valid PDFs; the XLS/XLSX PDFs include both visible worksheets; the command is repeated a second time into a fresh directory without save panels or new permission prompts. Legacy fixtures are generated by WPS itself from the modern fixtures using its matching native SaveAs format and committed only after package validation.

- [ ] **Step 5: Record the gate result and enforce backend enablement**

If both Phase 0 runs and all twelve conversion calls pass without prompts, record WPS/macOS versions, report paths, artifact sizes, and `GO` in the design document and set `MACOS_CONVERSION_ENABLED = True`. If Writer still presents a save panel for an in-container path, record `NO-GO`, keep `MACOS_CONVERSION_ENABLED = False`, retain the Windows backend, and do not request broader permissions.

- [ ] **Step 6: Re-run regression and inspect repository state**

Run:

```bash
python -m pytest -q
git diff --check
git status --short
```

Expected: all tests PASS, `git diff --check` emits no output, and status lists only the intended documentation/gate changes.

- [ ] **Step 7: Commit**

```bash
git add README.md macos/wps-jsapi-probe/fixtures/README.md tests/fixtures/office_conversion docs/superpowers/specs/2026-07-17-macos-wps-sandbox-staging-design.md skills/WPSComposer/scripts/conversion.py
git commit -m "Verify and document Office to PDF conversion"
```
