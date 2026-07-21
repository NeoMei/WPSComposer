# macOS WPS Generation Parity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Enable the existing public `generate()` API on macOS for DOCX, PPTX, XLSX, and Writer-derived PDF without save panels, broader macOS permissions, or changes to Windows behavior.

**Architecture:** Python keeps parsing Markdown and choosing layout, but injected recording composers convert renderer calls into a validated, closed operation plan. The macOS backend clones pinned WPS-produced fixtures into the WPS container, sends one typed plan to the matching JSAPI add-in, saves the already-existing staged file in place with `Save()`, validates it, and atomically publishes it.

**Tech Stack:** Python 3.9+, pytest, WPS JSAPI, Node.js 20+, `wpsjs==2.2.3`, `wps-jsapi-declare==2.2.0`, JSON operation plans, OOXML ZIP validation, existing authenticated loopback bridge.

## Global Constraints

- Keep Windows COM generation behavior and public signatures backward compatible.
- Do not use `SaveAs`, `SaveAs2`, AppleScript, Computer Use, Accessibility, Input Monitoring, screen capture, Full Disk Access, keyboard simulation, or save-panel automation.
- WPS receives only paths inside `~/Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/WPSComposer/`.
- Only fixed operation names and typed arguments are accepted; no arbitrary property paths, JavaScript, `eval`, or dynamic method invocation.
- Use the pinned `wpsjs==2.2.3` fixtures and verify these SHA-256 digests: Writer `95c2da9c75b65f7da18da345847a65ecab21512978c15e65c5b601512d52fd8e`, Spreadsheet `999138c4d4d22c2eb7c80e114d623da0ac310406ab71317256758754aae02dc9`, Presentation `5bafe9e14e99c7b1f7f81e0d1e32c59eb2d52c93ccfdd4ca786b0c8616174600`.
- A format stays disabled until two consecutive installed-WPS runs pass with a fresh session, valid final artifact, no interactive panel, restored registration, zero staging sessions, and no unrelated-process termination.
- Preserve stable typed errors, private-path redaction, no-clobber publication, durable registration recovery, and the shared WPS runtime lock.

---

## File and interface map

- Create `skills/WPSComposer/scripts/generation_plan.py`: immutable plan model, schemas, size limits, and validation.
- Create `skills/WPSComposer/scripts/recording_composers.py`: renderer-facing Writer/Sheet/Slide recorders.
- Create `skills/WPSComposer/scripts/macos_probe/templates.py`: pinned template lookup, digest validation, and private cloning.
- Create `skills/WPSComposer/scripts/macos_probe/generation.py`: staged generation lifecycle and artifact publication.
- Modify `skills/WPSComposer/scripts/renderers/*.py`: inject composer factories and replace direct COM access with explicit composer methods.
- Modify `skills/WPSComposer/scripts/{writer,sheet,slide}.py`: add the small high-level methods needed to remove renderer access to private COM fields.
- Modify `macos/wps-jsapi-probe/addin/{writer,spreadsheet,presentation}.js`: validate and execute component operation plans, then save in place.
- Modify `skills/WPSComposer/scripts/macos_probe/models.py`: allow only the three typed generation methods.
- Modify `skills/WPSComposer/scripts/orchestrator.py`: platform routing and stable generation errors.
- Modify public exports and documentation only after the real gate passes.

---

### Task 1: Pinned WPS template catalog and private clone

**Files:**
- Create: `skills/WPSComposer/scripts/macos_probe/templates.py`
- Create: `tests/macos_probe/test_templates.py`
- Modify: `macos/wps-jsapi-probe/fixtures/README.md`

**Interfaces:**
- Consumes: installed probe root and existing `validate_office_package(path, format_name)`.
- Produces: `TemplateSpec`, `template_for_component(component)`, and `clone_template(probe_root, staging_dir, component) -> Path`.

- [ ] **Step 1: Write failing catalog and clone tests**

```python
def test_clone_template_verifies_digest_and_creates_private_copy(tmp_path):
    source = tmp_path / "probe/node_modules/wpsjs/src/lib/res/wpsDemo.docx"
    source.parent.mkdir(parents=True)
    shutil.copy2(REAL_WRITER_FIXTURE, source)
    staging = tmp_path / "session"
    staging.mkdir(mode=0o700)

    cloned = clone_template(tmp_path / "probe", staging, "writer")

    assert cloned == staging / "generated.docx"
    assert cloned.read_bytes() == source.read_bytes()
    assert stat.S_IMODE(cloned.stat().st_mode) == 0o600
    validate_office_package(cloned, "docx")


def test_clone_template_rejects_changed_fixture(tmp_path):
    # Arrange the expected path with altered bytes.
    with pytest.raises(TemplateError, match="digest mismatch"):
        clone_template(tmp_path / "probe", tmp_path / "session", "writer")
```

- [ ] **Step 2: Run tests and verify the expected import failure**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_templates.py -q`

Expected: collection fails because `macos_probe.templates` does not exist.

- [ ] **Step 3: Implement the fixed catalog and clone**

```python
@dataclass(frozen=True)
class TemplateSpec:
    component: str
    filename: str
    output_name: str
    format_name: str
    sha256: str


TEMPLATES = {
    "writer": TemplateSpec("writer", "wpsDemo.docx", "generated.docx", "docx", "95c2da9c75b65f7da18da345847a65ecab21512978c15e65c5b601512d52fd8e"),
    "spreadsheet": TemplateSpec("spreadsheet", "etDemo.xlsx", "generated.xlsx", "xlsx", "999138c4d4d22c2eb7c80e114d623da0ac310406ab71317256758754aae02dc9"),
    "presentation": TemplateSpec("presentation", "wppDemo.pptx", "generated.pptx", "pptx", "5bafe9e14e99c7b1f7f81e0d1e32c59eb2d52c93ccfdd4ca786b0c8616174600"),
}


def clone_template(probe_root: Path, staging_dir: Path, component: str) -> Path:
    spec = template_for_component(component)
    source = probe_root / "node_modules/wpsjs/src/lib/res" / spec.filename
    if _sha256(source) != spec.sha256:
        raise TemplateError(f"Pinned {component} template digest mismatch")
    validate_office_package(source, spec.format_name)
    target = staging_dir / spec.output_name
    shutil.copyfile(source, target)
    os.chmod(target, 0o600)
    return target
```

- [ ] **Step 4: Run focused tests**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_templates.py tests/test_artifact_transport.py -q`

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/templates.py tests/macos_probe/test_templates.py macos/wps-jsapi-probe/fixtures/README.md
git commit -m "Add pinned Mac WPS generation templates"
```

---

### Task 2: Closed generation operation-plan protocol

**Files:**
- Create: `skills/WPSComposer/scripts/generation_plan.py`
- Create: `tests/test_generation_plan.py`
- Modify: `skills/WPSComposer/scripts/macos_probe/models.py`
- Modify: `tests/macos_probe/test_models.py`

**Interfaces:**
- Consumes: JSON-compatible renderer arguments.
- Produces: `GenerationOperation`, `GenerationPlan`, `GenerationResource`, `RecordedGeneration`, `validate_generation_plan(raw, component)`, `MAX_PLAN_BYTES = 2_000_000`, and three bridge methods.

- [ ] **Step 1: Write failing model tests**

```python
def test_generation_plan_round_trips_valid_writer_operations():
    plan = GenerationPlan(
        "writer",
        (
            GenerationOperation("writer.reset", {}),
            GenerationOperation("writer.add_paragraph", {
                "text": "macOS parity", "style": "Heading 1"
            }),
        ),
    )
    assert validate_generation_plan(plan.to_dict(), "writer") == plan


def test_generation_plan_rejects_unknown_or_cross_component_operation():
    with pytest.raises(OperationPlanError, match="unsupported operation"):
        validate_generation_plan({
            "component": "writer",
            "operations": [{"op": "sheet.write_table", "args": {}}],
        }, "writer")


def test_recorded_resources_are_not_serialized_into_bridge_plan(tmp_path):
    plan = GenerationPlan(
        "writer",
        (GenerationOperation("writer.reset", {}),),
    )
    resource = GenerationResource("image-1", tmp_path / "figure.png", "image/png")
    recorded = RecordedGeneration(plan=plan, resources=(resource,))
    assert "source_path" not in json.dumps(recorded.plan.to_dict())
    assert recorded.resources == (resource,)


def test_generation_command_is_component_typed():
    with pytest.raises(ProtocolError, match="requires writer"):
        ProbeCommand.create("spreadsheet", "generate_writer_document", {})
```

- [ ] **Step 2: Run and verify failure**

Run: `.venv/bin/python -m pytest tests/test_generation_plan.py tests/macos_probe/test_models.py -q`

Expected: import/allow-list assertions fail.

- [ ] **Step 3: Implement immutable models and exact schemas**

The initial allowed operations are:

```python
ALLOWED_OPERATIONS = {
    "writer": {
        "writer.reset", "writer.configure_page", "writer.ensure_styles",
        "writer.add_paragraph", "writer.add_heading", "writer.add_list",
        "writer.add_table", "writer.add_image", "writer.add_page_break",
        "writer.add_section", "writer.add_horizontal_line",
        "writer.insert_toc", "writer.set_page_number", "writer.update_fields",
    },
    "spreadsheet": {
        "sheet.reset", "sheet.rename", "sheet.add", "sheet.select",
        "sheet.write_table", "sheet.set_column_width", "sheet.autofit",
    },
    "presentation": {
        "slide.reset", "slide.set_size", "slide.apply_preset",
        "slide.add_title", "slide.add_section", "slide.add_bullets",
        "slide.add_blank", "slide.add_image", "slide.add_table",
    },
}
```

Validate: component equality, non-empty list, maximum 10,000 operations, exact `{op,args}` keys, operation membership, JSON size no greater than 2,000,000 bytes, strings no longer than 100,000 characters, tables no larger than 10,000 cells, and image paths represented only as logical `imageId` values.

`GenerationResource` contains `id`, resolved host `source_path`, and an
allow-listed media type. `RecordedGeneration` pairs a serializable
`GenerationPlan` with a tuple of host-only resources. Resource source paths are
never included in `GenerationPlan.to_dict()` or sent to WPS.

Extend `METHOD_COMPONENT` with:

```python
"generate_writer_document": "writer",
"generate_spreadsheet_workbook": "spreadsheet",
"generate_presentation_deck": "presentation",
```

- [ ] **Step 4: Run focused protocol tests**

Run: `.venv/bin/python -m pytest tests/test_generation_plan.py tests/macos_probe/test_models.py tests/macos_probe/test_bridge.py -q`

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/generation_plan.py skills/WPSComposer/scripts/macos_probe/models.py tests/test_generation_plan.py tests/macos_probe/test_models.py
git commit -m "Add closed Mac generation plan protocol"
```

---

### Task 3: In-place WPS save feasibility backend

**Files:**
- Create: `skills/WPSComposer/scripts/macos_probe/generation.py`
- Create: `tests/macos_probe/test_generation.py`
- Modify: `macos/wps-jsapi-probe/addin/writer.js`
- Modify: `macos/wps-jsapi-probe/addin/spreadsheet.js`
- Modify: `macos/wps-jsapi-probe/addin/presentation.js`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: a cloned template and a validated minimal generation plan.
- Produces: `GenerationError`, `GenerationRequest`, `execute_generation_plan(request: GenerationRequest, recorded: RecordedGeneration, *, enabled: Optional[Mapping[str, bool]] = None, bridge_factory=LoopbackBridge, runtime_factory=ProbeRuntime, timeout: float = 90) -> Path`, and three add-in handlers that use native `Save()` in place.

- [ ] **Step 1: Add executable Node failure/success tests**

For each component, execute the add-in in Node with a stub document. Assert:

```javascript
const result = await window.WPSComposerProbe.handleCommand({
  method: "generate_writer_document",
  params: {
    stagedPath: "/staged/generated.docx",
    plan: {component: "writer", operations: [
      {op: "writer.reset", args: {}},
      {op: "writer.add_paragraph", args: {text: "IN_PLACE_MARKER", style: "Body Text"}}
    ]}
  }
});
assert.equal(result.path, "/staged/generated.docx");
assert.equal(result.appliedOperations, 2);
assert.equal(document.saveCalls, 1);
assert.equal(document.saveAsCalls, 0);
assert.equal(Application.DisplayAlerts, 7);
assert.equal(document.closeArgument, 0);
```

Also assert unknown operations fail as `OPERATION_PLAN_INVALID`, save failure becomes `GENERATION_COMMAND_FAILED`, and `DisplayAlerts` is restored after open, mutation, save, or close failure.

- [ ] **Step 2: Add Python lifecycle tests with fake bridge/runtime**

```python
def test_generation_uses_only_staged_path_and_publishes_valid_package(tmp_path):
    request = GenerationRequest(output=tmp_path / "final.docx", component="writer", overwrite=False)
    result = execute_generation_plan(
        request,
        RecordedGeneration(WRITER_MARKER_PLAN, ()),
        enabled={"docx": True},
        bridge_factory=lambda origins: bridge,
        runtime_factory=lambda *args, **kwargs: runtime,
        timeout=2,
    )
    command = bridge.commands[0]
    assert Path(command.params["stagedPath"]).parent == runtime.staging_dir
    assert "output" not in command.params
    assert result == request.output.resolve()
```

- [ ] **Step 3: Run tests and verify failure**

Run: `.venv/bin/python -m pytest tests/macos_probe/test_generation.py tests/macos_probe/test_addin_assets.py -q`

Expected: missing handlers/backend failures.

- [ ] **Step 4: Implement minimal in-place handlers**

Define the backend's stable request and error types before implementing the
lifecycle:

```python
@dataclass(frozen=True)
class GenerationRequest:
    output: Path
    component: str
    format_name: str
    overwrite: bool = False


class GenerationError(RuntimeError):
    def __init__(self, *, code, output, component, backend, message):
        super().__init__(message)
        self.code = code
        self.output = output
        self.component = component
        self.backend = backend
        self.message = message
```

Each handler follows this structure:

```javascript
function generateWriterDocument(params) {
  const previousAlerts = Application.DisplayAlerts;
  let document = null;
  let failure = null;
  try {
    Application.DisplayAlerts = 0;
    document = Application.Documents.Open(params.stagedPath, false, false);
    const applied = executeWriterOperations(document, params.plan.operations);
    document.Save();
    return {path: params.stagedPath, appliedOperations: applied};
  } catch (error) {
    failure = generationError(error);
    throw failure;
  } finally {
    try {
      if (document !== null) document.Close(0);
    } catch (closeError) {
      if (failure === null) throw generationError(closeError);
    } finally {
      Application.DisplayAlerts = previousAlerts;
    }
  }
}
```

Spreadsheet uses `Workbooks.Open(params.stagedPath, 0, false)`, `workbook.Save()`, and `Close(false)`. Presentation uses `Presentations.Open(params.stagedPath, true, false, false)`, `presentation.Save()`, and `Close()`.

The minimal executors implement only reset plus one visible text/value operation required by the feasibility marker. The host validates the command-reported path, polls until the OOXML package contains `IN_PLACE_MARKER`, publishes atomically, and always uses `ProbeRuntime` so locking and rollback remain shared.

- [ ] **Step 5: Run deterministic tests and JS syntax checks**

Run:

```bash
.venv/bin/python -m pytest tests/macos_probe/test_generation.py tests/macos_probe/test_addin_assets.py -q
for f in macos/wps-jsapi-probe/addin/{writer,spreadsheet,presentation}.js; do node --check "$f"; done
```

Expected: all pass.

- [ ] **Step 6: Run the installed-WPS feasibility gate twice**

Run a temporary Python gate that clones each pinned template, submits reset plus `IN_PLACE_MARKER`, validates the marker inside OOXML XML members, and writes final artifacts under:

```text
/tmp/wpscomposer-in-place-save-gate-1-20260719
/tmp/wpscomposer-in-place-save-gate-2-20260719
```

Expected: six package saves total pass; no save panel or permission prompt; `publish.xml` restored; WPS staging session count zero. If Writer fails, stop the plan here, retain the failure evidence, and keep public generation disabled.

- [ ] **Step 7: Commit only after the real feasibility gate passes**

```bash
git add skills/WPSComposer/scripts/macos_probe/generation.py tests/macos_probe/test_generation.py macos/wps-jsapi-probe/addin tests/macos_probe/test_addin_assets.py
git commit -m "Prove Mac WPS in-place document saves"
```

---

### Task 4: Writer renderer seam and complete public Writer plan

**Files:**
- Create: `skills/WPSComposer/scripts/recording_composers.py`
- Modify: `skills/WPSComposer/scripts/writer.py`
- Modify: `skills/WPSComposer/scripts/renderers/writer_renderer.py`
- Modify: `macos/wps-jsapi-probe/addin/writer.js`
- Create: `tests/test_recording_composers.py`
- Create: `tests/test_writer_renderer.py`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: existing `StructuredDocument`, styles, presets, and `GenerationPlan`.
- Produces: `RecordingWriterComposer`, injected `writer_renderer.render(doc, output_path, preset=None, composer_factory=WriterComposer)`, and all Writer operations from Task 2.

- [ ] **Step 1: Test renderer injection and semantic plan output**

```python
def test_writer_renderer_records_heading_table_image_and_footer(sample_document):
    recorder = RecordingWriterComposer()
    recorded = writer_renderer.render(
        sample_document,
        "ignored.docx",
        preset=PRESETS["business"],
        composer_factory=lambda: recorder,
    )
    names = [operation.op for operation in recorded.plan.operations]
    assert names[0] == "writer.reset"
    assert "writer.ensure_styles" in names
    assert "writer.add_heading" in names
    assert "writer.add_table" in names
    assert "writer.add_image" in names
    assert names[-2:] == ["writer.set_page_number", "writer.update_fields"]
```

- [ ] **Step 2: Refactor direct COM access behind explicit methods**

Add these methods to `WriterComposer`: `apply_heading_text_color(color)`, `add_rich_paragraph(spans, style_name)`, and `add_code_lines(lines)`. Move the current selection/font implementation into those methods without changing output. Change renderer calls from `w.doc`, `w.selection`, `_set_font_family`, `_set_line_spacing`, and `_reset_selection_to_normal` to these public renderer-facing methods.

Change the render signature to:

```python
def render(doc, output_path, preset=None, composer_factory=WriterComposer):
    with composer_factory() as w:
        _render_into(w, doc, preset)
        return w.save_docx(output_path)


def _render_into(w, doc, preset):
    _configure_document(w, preset)
    _render_title_page(w, doc, preset)
    _render_body(w, doc, preset)
    w.set_page_number_in_footer()
    w.update_fields()
```

- [ ] **Step 3: Implement `RecordingWriterComposer`**

It mirrors only renderer-used methods, appends normalized operations, maps
image paths to `GenerationResource` objects with logical IDs, and returns
`RecordedGeneration(GenerationPlan("writer", tuple(operations)),
tuple(resources))` from `save_docx`. Context exit performs no external I/O.

- [ ] **Step 4: Implement every Writer operation in the add-in**

Use a literal dispatch object:

```javascript
const writerOperations = {
  "writer.reset": resetWriter,
  "writer.configure_page": configureWriterPage,
  "writer.ensure_styles": ensureWriterStyles,
  "writer.add_paragraph": addWriterParagraph,
  "writer.add_heading": addWriterHeading,
  "writer.add_list": addWriterList,
  "writer.add_table": addWriterTable,
  "writer.add_image": addWriterImage,
  "writer.add_page_break": addWriterPageBreak,
  "writer.add_section": addWriterSection,
  "writer.add_horizontal_line": addWriterHorizontalLine,
  "writer.insert_toc": insertWriterToc,
  "writer.set_page_number": setWriterPageNumber,
  "writer.update_fields": updateWriterFields
};
```

Reject missing handlers before opening the document. Use the document object and ranges belonging to it; do not assume another user document's selection.

- [ ] **Step 5: Run Writer regression and executable JS tests**

Run:

```bash
.venv/bin/python -m pytest tests/test_recording_composers.py tests/test_writer_renderer.py tests/macos_probe/test_addin_assets.py -q
.venv/bin/python -m pytest tests/test_markdown_parser.py tests/test_public_api.py -q
node --check macos/wps-jsapi-probe/addin/writer.js
```

Expected: all pass and Windows renderer snapshots/call assertions remain unchanged.

- [ ] **Step 6: Commit**

```bash
git add skills/WPSComposer/scripts/recording_composers.py skills/WPSComposer/scripts/writer.py skills/WPSComposer/scripts/renderers/writer_renderer.py macos/wps-jsapi-probe/addin/writer.js tests/test_recording_composers.py tests/test_writer_renderer.py tests/macos_probe/test_addin_assets.py
git commit -m "Record and execute Mac Writer generation plans"
```

---

### Task 5: Spreadsheet renderer seam and public Sheet plan

**Files:**
- Modify: `skills/WPSComposer/scripts/recording_composers.py`
- Modify: `skills/WPSComposer/scripts/sheet.py`
- Modify: `skills/WPSComposer/scripts/renderers/sheet_renderer.py`
- Modify: `macos/wps-jsapi-probe/addin/spreadsheet.js`
- Modify: `tests/test_recording_composers.py`
- Create: `tests/test_sheet_renderer.py`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: TableBlock-based public Sheet renderer.
- Produces: `RecordingSheetComposer` and exact execution for reset, rename, add/select sheet, table values/styles, widths, and autofit.

- [ ] **Step 1: Write a multi-table recording test**

Assert two Markdown table sections produce `sheet.reset`, rename the first sheet, add the second, write both tables, set widths, select sheet 1, and autofit.

- [ ] **Step 2: Remove renderer access to `s._doc`**

Add `SheetComposer.rename_sheet(index, name)` and change the renderer to call it. Inject `composer_factory=SheetComposer` and return `s.save_xlsx(output_path)` as before.

- [ ] **Step 3: Implement recorder and literal JS dispatch**

The add-in clears all worksheets to a known one-sheet baseline, sanitizes names to 31 characters, writes rectangular `Value2` arrays, applies the current header font/fill/border choices, sets widths, selects the requested sheet, and calls `UsedRange.Columns.AutoFit()`.

- [ ] **Step 4: Run focused tests**

Run:

```bash
.venv/bin/python -m pytest tests/test_recording_composers.py tests/test_sheet_renderer.py tests/macos_probe/test_addin_assets.py -q
node --check macos/wps-jsapi-probe/addin/spreadsheet.js
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/recording_composers.py skills/WPSComposer/scripts/sheet.py skills/WPSComposer/scripts/renderers/sheet_renderer.py macos/wps-jsapi-probe/addin/spreadsheet.js tests/test_recording_composers.py tests/test_sheet_renderer.py tests/macos_probe/test_addin_assets.py
git commit -m "Record and execute Mac Sheet generation plans"
```

---

### Task 6: Presentation renderer seam and public Slide plan

**Files:**
- Modify: `skills/WPSComposer/scripts/recording_composers.py`
- Modify: `skills/WPSComposer/scripts/slide.py`
- Modify: `skills/WPSComposer/scripts/renderers/slide_renderer.py`
- Modify: `macos/wps-jsapi-probe/addin/presentation.js`
- Modify: `tests/test_recording_composers.py`
- Create: `tests/test_slide_renderer.py`
- Modify: `tests/macos_probe/test_addin_assets.py`

**Interfaces:**
- Consumes: current title/section/content/table/image renderer behavior.
- Produces: `RecordingSlideComposer` and exact execution for the Presentation operations in Task 2.

- [ ] **Step 1: Write representative recording tests**

Assert a title, H1 section, seven bullets, table, image, and preset produce the same slide pacing as Windows: title slide, section slide, two bullet slides, table slide, and image slide.

- [ ] **Step 2: Remove renderer access to `p._doc`**

Add `SlideComposer.slide_count` and use it after `add_blank_slide()`. Inject `composer_factory=SlideComposer` into the render signature.

- [ ] **Step 3: Implement recorder and add-in dispatch**

Reset deletes all template slides. Implement native size, preset colors/fonts, title, section, bullets, blank, image, and table operations. Every geometry argument must be finite, non-negative, and bounded to four times the slide dimensions before calling JSAPI.

- [ ] **Step 4: Run focused tests**

Run:

```bash
.venv/bin/python -m pytest tests/test_recording_composers.py tests/test_slide_renderer.py tests/macos_probe/test_addin_assets.py -q
node --check macos/wps-jsapi-probe/addin/presentation.js
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/recording_composers.py skills/WPSComposer/scripts/slide.py skills/WPSComposer/scripts/renderers/slide_renderer.py macos/wps-jsapi-probe/addin/presentation.js tests/test_recording_composers.py tests/test_slide_renderer.py tests/macos_probe/test_addin_assets.py
git commit -m "Record and execute Mac Slide generation plans"
```

---

### Task 7: macOS generation facade and public platform routing

**Files:**
- Modify: `skills/WPSComposer/scripts/macos_probe/generation.py`
- Modify: `skills/WPSComposer/scripts/orchestrator.py`
- Modify: `skills/WPSComposer/scripts/__init__.py` if present
- Modify: `skills/WPSComposer/__init__.py`
- Create: `tests/test_generation.py`
- Modify: `tests/test_public_api.py`

**Interfaces:**
- Consumes: `RecordedGeneration`, injected renderers, template cloning, bridge runtime, artifact publisher.
- Produces: `generate_macos(doc, format_name, output, preset) -> Path`, public exports for `GenerationError`, and platform-routed `generate()`.

- [ ] **Step 1: Write platform routing and stable error tests**

```python
def test_generate_routes_macos_without_importing_pywin32(monkeypatch, tmp_path):
    monkeypatch.setattr(orchestrator.sys, "platform", "darwin")
    monkeypatch.setattr(orchestrator, "generate_macos", fake_generate)
    result = generate("# Report", format="docx", output=tmp_path / "r.docx", source_is_text=True)
    assert result == str((tmp_path / "r.docx").resolve())


def test_unknown_vendor_generation_code_is_normalized():
    assert normalize_generation_error_code("WPS_E_9001") == "GENERATION_COMMAND_FAILED"
```

Also test default paths, overwrite refusal, unsupported platform, missing gate, path redaction, and no partial final file.

- [ ] **Step 2: Implement request validation and backend lifecycle**

`GenerationRequest` contains resolved output, component, format name, overwrite. `generate_macos` creates one durable runtime root, one bridge, one `ProbeRuntime`, clones the component template, builds/validates a plan, issues the typed method, validates the returned path and applied count, waits for the saved semantic content, then calls `publish_artifact` with the matching validator.

Before issuing the command, resolve host-only resources as follows:

```python
def stage_generation_resources(recorded, runtime):
    resolved = {}
    for resource in recorded.resources:
        validate_generation_resource(resource, max_bytes=50 * 1024 * 1024)
        safe_name = f"resource-{resource.id}{resource.source_path.suffix.lower()}"
        if recorded.plan.component == "writer":
            target = runtime.profiles["writer"] / safe_name
            shutil.copyfile(resource.source_path, target)
            resolved[resource.id] = f"http://127.0.0.1:3889/{safe_name}"
        else:
            target = runtime.staging_dir / "resources" / safe_name
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(resource.source_path, target)
            resolved[resource.id] = str(target)
    return resolved
```

Send this `resourceId -> internal URL/path` manifest next to the serialized
plan. Writer accepts only its fixed `http://127.0.0.1:3889/` origin;
Presentation accepts only paths under the staged document's session directory.
Unknown, unused, duplicate, oversized, missing, or unsupported resources fail
as `OPERATION_PLAN_INVALID` before WPS mutation.

Use per-format gates:

```python
MACOS_GENERATION_ENABLED = {
    "docx": False,
    "xlsx": False,
    "pptx": False,
    "pdf": False,
}
```

Tests inject `enabled=True`; production values change only in Task 9.

- [ ] **Step 3: Route orchestrator by platform**

Keep parsing, preset resolution, and output naming shared. On `darwin`, call `generate_macos`; on `win32`, retain the existing renderer imports and behavior. Preserve the existing `generate()` signature and return absolute strings.

- [ ] **Step 4: Run focused and complete deterministic suites**

Run:

```bash
.venv/bin/python -m pytest tests/test_generation.py tests/test_public_api.py tests/macos_probe/test_generation.py -q
.venv/bin/python -m pytest -q
```

Expected: all pass with production macOS gates still false.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/generation.py skills/WPSComposer/scripts/orchestrator.py skills/WPSComposer/__init__.py tests/test_generation.py tests/test_public_api.py tests/macos_probe/test_generation.py
git commit -m "Add gated Mac WPS generation backend"
```

---

### Task 8: Writer-derived PDF generation

**Files:**
- Modify: `skills/WPSComposer/scripts/macos_probe/generation.py`
- Modify: `macos/wps-jsapi-probe/addin/writer.js`
- Modify: `tests/macos_probe/test_generation.py`
- Modify: `tests/macos_probe/test_addin_assets.py`
- Modify: `tests/test_generation.py`

**Interfaces:**
- Consumes: a successful Writer plan and private generated DOCX.
- Produces: staged `generated.pdf` and public `generate(source, format="pdf", preset=None, output=None, source_is_text=False)` behavior.

- [ ] **Step 1: Write failing PDF lifecycle tests**

Assert the host never publishes the private DOCX, the add-in exports only after `document.Save()`, returned PDF path is inside staging, invalid PDFs map to `STAGED_ARTIFACT_INVALID`, and cleanup removes both private files.

- [ ] **Step 2: Extend Writer generation command with an explicit output mode**

Accept only `outputFormat: "docx" | "pdf"`. For PDF, save the opened template in place, call `ExportAsFixedFormat(stagedPdfPath, 17, false, 0, 0)`, and return both `path` (PDF) and `sourcePath` (private DOCX). The host constructs both paths; the add-in may not choose them.

- [ ] **Step 3: Validate and publish PDF only**

Use `validate_office_package` for the private DOCX before accepting the PDF, `validate_pdf` before and after publication, and the stable generation error allow-list.

- [ ] **Step 4: Run focused tests**

Run:

```bash
.venv/bin/python -m pytest tests/test_generation.py tests/macos_probe/test_generation.py tests/macos_probe/test_addin_assets.py -q
node --check macos/wps-jsapi-probe/addin/writer.js
```

Expected: all pass.

- [ ] **Step 5: Commit**

```bash
git add skills/WPSComposer/scripts/macos_probe/generation.py macos/wps-jsapi-probe/addin/writer.js tests/test_generation.py tests/macos_probe/test_generation.py tests/macos_probe/test_addin_assets.py
git commit -m "Generate PDF through staged Mac Writer"
```

---

### Task 9: Real four-format gate, enablement, documentation, and review

**Files:**
- Modify: `skills/WPSComposer/scripts/macos_probe/generation.py`
- Modify: `README.md`
- Modify: `skills/WPSComposer/references/api.md`
- Modify: `docs/macos-phase0.md`
- Modify: `docs/superpowers/specs/2026-07-19-macos-wps-capability-parity-design.md`
- Create: `macos/wps-jsapi-probe/fixtures/generation-gate.md`
- Modify: `tests/test_documentation.py`

**Interfaces:**
- Consumes: completed generation backend and representative Markdown/image fixtures.
- Produces: evidence-backed per-format gate values and user-facing macOS generation documentation.

- [ ] **Step 1: Create deterministic representative inputs and semantic inspectors**

The gate input includes title metadata, H1/H2 sections, mixed bold/italic/code spans, bullet and numbered lists, task list, table, block quote, horizontal rule, and a local PNG. Add inspectors that unzip Office packages and assert required text, worksheet count/table values, slide count/text/table/image relations, and Writer relationships/styles/numbering parts.

- [ ] **Step 2: Run the full deterministic verification before WPS**

Run:

```bash
.venv/bin/python -m pytest -q
.venv/bin/python -m compileall -q skills install.py
for f in macos/wps-jsapi-probe/addin/{bridge-client,writer,spreadsheet,presentation}.js; do node --check "$f"; done
npm --prefix macos/wps-jsapi-probe audit
git diff --check
```

Expected: all tests pass, compilation/checks emit no errors, npm reports zero vulnerabilities.

- [ ] **Step 3: Run two fresh installed-WPS generation gates**

Use output directories:

```text
/tmp/wpscomposer-generation-gate-run-1-20260719
/tmp/wpscomposer-generation-gate-run-2-20260719
```

For each run call the public API with gate injection enabled for DOCX, XLSX, PPTX, and PDF. Validate semantic structure and use `pdfinfo` plus extracted text for PDF. Record WPS/macOS/Node versions, output sizes, element/page counts, registration state, staging count, and process ownership.

- [ ] **Step 4: Enable only passing formats**

Set a format to `True` only if both real runs passed without a native panel or new permission prompt. If one format fails, keep only that format false and document its exact typed failure; do not disable formats that independently pass.

- [ ] **Step 5: Update public documentation and tests**

Remove the global macOS generation NO-GO wording only when all four formats pass. Document that macOS uses native-template in-place saves, that output may differ in line wrapping/fonts from Windows, and that existing-file editing and active-document attachment remain subsequent parity slices.

- [ ] **Step 6: Request code review and address all Critical/Important findings**

Run the requesting-code-review workflow over the complete generation diff. Apply technically valid findings through failing tests, then request a focused re-review until no Critical/Important issue remains.

- [ ] **Step 7: Run fresh completion verification**

Repeat the commands from Step 2 after the final review fixes and confirm `git status --short` lists only intended documentation/gate changes.

- [ ] **Step 8: Commit**

```bash
git add README.md skills/WPSComposer/references/api.md docs/macos-phase0.md docs/superpowers/specs/2026-07-19-macos-wps-capability-parity-design.md macos/wps-jsapi-probe/fixtures/generation-gate.md skills/WPSComposer/scripts/macos_probe/generation.py tests/test_documentation.py
git commit -m "Enable verified Mac WPS document generation"
```

---

## Plan completion boundary

This plan ends after public macOS generation is independently gated for DOCX,
XLSX, PPTX, and PDF. Existing-file `inspect/edit`, persistent active-document
attachment, and the remaining low-level composer matrix deliberately begin in
separate specs and plans so each can be reviewed, rolled back, and accepted
without destabilizing generation.
