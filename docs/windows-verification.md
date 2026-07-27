# Windows COM verification handoff

> **Start here on Windows.** Read `AGENTS.md` first, then this file.
> Verification **completed on Windows** — see "Windows run results" below.

## Status

**Verified on Windows (2026-07-27).** All checklist items A–G pass against a
live WPS Office host. The four pre-known COM bugs were fixed as directed, and
seven further COM-behaviour bugs found during the live run were fixed (listed
in "Bugs fixed in the Windows run" below). The full platform-independent
suite is green on Windows (600 passed, 11 skipped — the skips are POSIX-only
macOS-probe tests plus the pypdf skip), and the COM verification scripts in
`fixtures/verify_*.py` pass end-to-end.

Historical context: a sequence of macOS sessions added agent-friendly
conversational-edit features to
`skills/WPSComposer/scripts/document_api.py` and the three COM composers
(`writer.py` / `sheet.py` / `slide.py`):

- **Orchestration (pure-Python, tested on macOS):** structured patch results +
  `PatchError`, atomic `edit()`, `validate_target` / `patch_grammar`,
  `snapshot_to_patches` (dump→replay), and structural verbs via `apply_ops` /
  `validate_op` (`insert` / `remove` / `move` / `clone`).
- **COM bodies (written blind, verified and fixed on Windows):** stable-ID
  readback (`@paraId` / `@id` / `@name`) and resolution, and the full
  `apply_structural_op` implementations on all three composers.

The macOS suite is green (**609 passed, 1 skipped**; 77 tests in
`tests/test_document_api.py`) and a 7000-iteration fuzz of the orchestration
layer produced 0 crashes. The COM-coupled behaviour can only be verified against
a live WPS/Office host on Windows.

## Why macOS could not finish

The conversational API (`inspect` / `edit` / `apply_format_patch` /
`apply_structural_op`) drives the WPS COM object model. macOS uses the WPS
JSAPI bridge for *generation* only (`generate()` / `convert_to_pdf()`); the
inspect/edit path has no macOS backend. The COM bodies therefore need a Windows
host with `pywin32` + WPS (or MS) Office.

## What was implemented and already verified (macOS, pure-Python)

All in `skills/WPSComposer/scripts/document_api.py`, exported via
`skills/WPSComposer/scripts/wps_engine.py`:

| Addition | File reference | Tests |
|---|---|---|
| `PATCH_GRAMMAR` data table + `patch_grammar(kind=None)` | `document_api.py` | `test_patch_grammar_*` |
| `validate_target(target, kind)` with closest-form suggestion | `document_api.py` | `test_validate_target_*` |
| `PatchError` carrying `.reports` / `.errors` | `document_api.py` | `test_apply_patches_*` |
| `apply_patches(..., atomic=True)` structured reports + atomic raise | `document_api.py` | `test_apply_patches_*` |
| `edit(..., atomic=True, raise_on_error=False)` structured result, no-save on failure | `document_api.py` | `test_edit_*` |

Run on macOS:

```bash
.venv/bin/python -m pytest tests/test_document_api.py -v   # 77 passed
```

## Windows verification checklist

Environment is recorded at the bottom of this file.

> **The verify scripts below are illustrative.** Runnable equivalents live in
> `fixtures/verify_*.py` (see "Windows run results"); `fixtures/make_fixtures.py`
> recreates the sample documents.

The COM-coupled changes were tagged `# WINDOWS-VERIFY:` in source; all markers
were verified and cleared in the Windows run. The high-value sites are called
out below with symbol references.

### A. Error-classification heuristic parity (orchestration layer)

`document_api.py::_error_report` tags a `ValueError` as `invalid_target`
when its message contains `Unsupported`, `target`, or `Unknown kind`;
otherwise `invalid_value`. This relies on the real composers' static
`ValueError(f"Unsupported {Kind} target: {target}")` messages:

- `writer.py` — `apply_format_patch` final `raise ValueError(...)`
- `sheet.py` — idem
- `slide.py` — idem

**Verify:** trigger each with a bad target on a real doc; confirm the message
still contains `Unsupported` (no localized override in installed WPS). If WPS
localizes it, loosen the heuristic in `_error_report`.

### B. Atomic no-save guarantee (core correctness claim)

`edit(atomic=True)` must NOT write a file when any patch fails.

```python
from skills.WPSComposer import edit
r = edit("fixtures/sample.docx", output="out.docx",
         patches=[{"target": "paragraph:1", "font": {"size": 20}},
                  {"target": "paragraph:99999"}])   # out-of-range -> apply_failed
assert r["ok"] is False and r["saved"] is False and r["saved_path"] is None
import os; assert not os.path.exists("out.docx")     # the guarantee
# and the happy path still saves:
r = edit("fixtures/sample.docx", output="out.docx",
         patches=[{"target": "paragraph:1", "font": {"size": 20}}])
assert r["ok"] is True and r["saved_path"].endswith("out.docx")
```

Also reopen the original source after a failing atomic `edit()` and confirm
it is unchanged.

### C. Attach-active caveat (documented behaviour)

For `edit(path=None)` atomic mode skips `save_current()` on failure (no disk
write), but the live WPS window may still show partially-applied formatting
(COM calls are not rolled back). Verify both halves and confirm against the
"Atomicity and the attach-active caveat" note in `skills/WPSComposer/references/api.md`.

### D. Stable-ID readback + resolution (NEW — borrowable #1, the core)

This is the main COM work added this session. Three hosts:

| Host | What to verify | Source (symbol; grep `WINDOWS-VERIFY` for the exact site) |
|---|---|---|
| Writer `@paraId` | `inspect()` emits `paragraph:@paraId=HEX` for saved docx; `apply_format_patch("paragraph:@paraId=...")` hits the right paragraph after a structural change | `writer.py`: `_read_paraid_map`, `_paragraph_index_for_paraid`, `_extract_paraids` (pure), `inspect_document` paragraph loop, `apply_format_patch` `@paraId=` branch |
| Slide `@id` / `@name` | `inspect()` emits `slide:N/shape:@id=K`; `@id=` and `@name=` resolve to the right shape | `slide.py`: `_shape_snapshot`, `apply_format_patch` `@id=`/`@name=` branches, `_find_shape_in_slide` |
| Sheet `@id` / `@name` | same for `sheet:N/shape:@id=K` / `@name=` | `sheet.py`: shape snapshot in `inspect_document`, `apply_format_patch` `@id=`/`@name=` branches, `_find_shape_in_sheet` |

**Writer assumptions to confirm** (the riskiest piece):

1. The document-order walk of `<w:p>` in `word/document.xml` matches
   `doc.Paragraphs(1..N)`. `_read_paraid_map` enforces this with a count
   check and falls back to positional on mismatch — confirm the count matches
   on a few real docs (with tables) so stable ids actually emit.
2. `w14:paraId` is present on modern .docx (Word 2010+). Older docs have none
   → positional fallback (expected, not a bug).
3. Reading the docx zip while WPS holds the doc open works read-only (it
   should; WPS does not read-lock the package).

**Resolution parity smoke:**

```python
from skills.WPSComposer import inspect, edit
snap = inspect("fixtures/sample.docx")
# grab the first stable paragraph id, mutate structure, re-apply by id
pid = next(p["id"] for p in snap["paragraphs"] if p.get("para_id"))
# (insert a paragraph here via the GUI or another edit, then:)
edit("fixtures/sample.docx",
     patches=[{"target": pid, "font": {"bold": True}}])   # must still hit the ORIGINAL paragraph
```

### E. `validate_target` grammar parity (positional + stable forms)

Confirm each accepted form resolves on a real document, especially edge forms:
sheet `sheet:1/cell:$A$1` (COM `Range()` accepts `$`), slide run path
`slide:1/shape:2/paragraph:3/run:1`, slide table cell
`slide:1/shape:2/table/cell:1,1`, writer `range:0-10`.

### F. `snapshot_to_patches` round-trip fidelity (NEW — borrowable #4)

`snapshot_to_patches()` converts an inspect snapshot to a patch list. Fidelity
depends on snapshot/apply key symmetry for each dimension. Round-trip check:

```python
from skills.WPSComposer import inspect, snapshot_to_patches, edit
snap = inspect("fixtures/sample.docx")
patches = snapshot_to_patches(snap, dimensions=("font",))
r = edit("fixtures/blank.docx", output="replay.docx", patches=patches)
assert r["ok"]
# then visually/diff confirm replay.docx fonts match sample.docx
```

If a snapshot key (e.g. a `font_snapshot` field) is not accepted by
`apply_font`, it lands in the patch's `rejected` list — inspect
`r["ops"][*]["rejected"]` (or the `patches` alias) to find any asymmetry.

### G. Structural verbs (NEW — insert / remove / move / clone, full coverage)

`apply_structural_op(op)` was added to all three composers and covers all
addressable element types (gaps from the first pass are filled). The
orchestration (`apply_ops` / `edit(ops=...)` / `validate_op`) is tested on macOS
with a fake composer (77 tests); the COM bodies are tagged `# WINDOWS-VERIFY`.

| Host | insert | remove | move | clone |
|---|---|---|---|---|
| Writer | paragraph / heading / page_break / table / image / textbox | paragraph / shape / table | paragraph / shape / table (clipboard) | paragraph / shape / table (clipboard) |
| Slide | slide / textbox / image | slide / shape | slide (reorder) / shape (`{"slide": N}`) | slide / shape |
| Sheet | row / column / sheet | row / column / shape / chart / sheet | row / column / sheet | row / sheet |

**Verify each verb end-to-end** on a docx with ≥3 paragraphs + a table + a
shape, a pptx with ≥3 slides + shapes, an xlsx with ≥3 rows/cols + a chart:

```python
from skills.WPSComposer import edit
# Writer: full set
edit("f.docx", output="o.docx", ops=[
    {"op": "insert", "type": "paragraph", "props": {"text": "appended"}},
    {"op": "insert", "type": "table", "props": {"rows": 2, "cols": 2, "data": [["a","b"],["c","d"]]}},
    {"op": "insert", "type": "image", "props": {"path": "logo.png"}},
    {"op": "insert", "type": "textbox", "props": {"text": "note", "left": 100, "top": 100, "width": 200, "height": 50}},
    {"op": "remove", "target": "table:1"},
    {"op": "remove", "target": "shape:1"},
    {"op": "clone", "target": "paragraph:1", "to": "end"},
    {"op": "move", "target": "paragraph:2", "to": {"after": "paragraph:4"}},
])
# Slide: shape move/clone to another slide
edit("f.pptx", output="o.pptx", ops=[
    {"op": "insert", "type": "slide", "props": {"layout": 12}},
    {"op": "clone", "target": "slide:1/shape:@id=7", "to": {"slide": 2}},
    {"op": "move", "target": "slide:1/shape:2", "to": {"slide": 3}},
    {"op": "remove", "target": "slide:2"},
])
# Sheet: column/sheet insert+remove, row/sheet clone
edit("f.xlsx", output="o.xlsx", ops=[
    {"op": "insert", "parent": "sheet:1", "type": "column",
     "props": {"values": ["x", "y"]}, "position": {"index": 2}},
    {"op": "insert", "type": "sheet", "props": {"name": "Summary"}},
    {"op": "remove", "target": "sheet:1/cell:C1"},          # axis=row default -> removes row 1
    {"op": "remove", "target": "sheet:1/cell:C1", "axis": "column"},  # removes column C (3)
    {"op": "remove", "target": "sheet:2/chart:1"},
    {"op": "clone", "target": "sheet:1/cell:A2", "to": {"index": 5}},
    {"op": "clone", "target": "sheet:1", "to": {"after": 1}},
    {"op": "move", "target": "sheet:1", "to": {"after": 2}},
])
```

**Assumptions to confirm** (the risky blind-COM parts):

- Writer `_structural_target` resolves paragraph (positional/`@paraId`) / shape /
  table to a COM object; `move`/`clone` use `obj.Cut()`/`obj.Copy()` +
  `Range.Paste()`. Confirm shape paste via `Range.Paste()` works (shapes are
  floating); if not, paste via a `Selection` or `Shapes.Paste`.
- Writer `image` insert uses `InlineShapes.AddPicture`; the returned path is
  `shape:N` (InlineShapes share the Shapes index on some hosts — confirm the
  returned index actually resolves).
- Slide shape move = `shape.Cut()` + `Slides(N).Shapes.Paste()`; clone =
  `shape.Duplicate()` (same slide) or `Copy`+`Paste` (cross-slide). Confirm
  `Shapes.Paste()` returns/positions the new shape.
- Sheet `Rows(n).Insert(0)` / `Columns(n).Insert(1)` shift direction constants
  (`xlShiftDown=0`, `xlShiftToRight=1`); `Worksheets.Add()` / `.Copy(After=)` /
  `.Move(After=)` for sheets; `ChartObjects(n).Delete()` for charts. Cell/range
  `remove`/`move` take `"axis": "row"|"column"` (default row) to pick the axis.
- Sheet whole-`sheet` remove is **dangerous** (WPS requires ≥1 visible sheet) —
  the call raises if it would remove the last sheet; confirm the error surfaces
  as `invalid_target`/`apply_failed` and not a crash.
- Inserted/cloned element `path` is **best-effort positional**; re-`inspect()`
  for a stable id. Structural ops shift sibling positional indices — address
  later ops in the same batch by stable id or re-inspect between batches.

**Known COM bugs found by code review (fixed in the Windows run):**

- **Writer `_resolve_insert_range` end/start** (`writer.py`): used bare
  `doc.Range` (a *method*, not a property) → would crash on `.Collapse()`.
  Fixed: `doc.Content` collapsed to end/start.
- **Writer `_insert_element` heading at end** (`writer.py`): the positional
  `return` fired *before* the Heading style was applied. Fixed: style applied
  before the early return, using the locale-independent built-in style id
  (`Styles(-1 - level)`, wdStyleHeading1..9 = -2..-10) — the English name
  "Heading N" does not exist in localized WPS builds.
- **Sheet `chart:N` remove** (`sheet.py`): non-numeric refs no longer hit the
  bare `int()`; the pattern is restricted to digits and falls through to the
  standard `Unsupported Sheet target` ValueError.
- **Writer image insert path** (`writer.py`): returned `shape:N` from
  `InlineShapes.Count`, but `shape:N` resolved via `doc.Shapes(N)` (a different
  collection). Fixed: returns `inline_shape:N`, and `_structural_target`
  resolves that form via `doc.InlineShapes(N).Range`.

## Bugs fixed in the Windows run (found live, not by review)

1. **`_extract_paraids` did not match `doc.Paragraphs`** (`writer.py`). The
   naive `root.iter(w:p)` walk (a) counted text-box-story paragraphs that COM
   excludes, and (b) missed that COM counts one paragraph per table
   *end-of-row mark* (no `<w:p>` in the XML). Fixed by pruning
   `w:txbxContent` subtrees and appending one `None` per `<w:tr>`. Verified
   aligned on docs with tables + textboxes (30 == 30).
2. **WPS ET ignores keyword `Before=`/`After=` on `Worksheet.Move`/`Copy`**
   (`sheet.py`). With kwargs, `Copy` silently falls back to "new workbook" —
   clones leaked into a second workbook and the save lost them. Fixed by
   calling positionally: `ws.Copy(None, target)` / `ws.Move(None, target)`.
3. **Sheet `insert type=sheet` crashed without a parent** (`sheet.py`):
   `_insert_element` parsed `parent` unconditionally before dispatch. The
   whole-workbook sheet insert is now handled before parent parsing.
4. **Slide image insert crashed without explicit width/height** (`slide.py`):
   `None` was passed to `Shapes.AddPicture`. Now defaults to `-1` (native
   size).
5. **`snapshot_to_patches` emitted document-specific stable ids**
   (`document_api.py`): replaying a dump on *another* document could never
   resolve `@paraId`/`@id`. Targets are now rewritten to positional form via
   the element's `index` (stable ids still work for same-document addressing
   through `edit()` directly). `None`-valued snapshot keys are dropped (they
   mean "host reported no value" and landed in `rejected`).
6. **Positional out-of-range targets raised raw COM errors** (`writer.py`):
   `paragraph:99999` etc. surfaced as `apply_failed` with an opaque HRESULT.
   Bounds checks now raise `ValueError` → classified `invalid_target` with
   self-heal `valid_forms`, consistent with the `@paraId` path.
7. **Flaky `AttributeError: KWpp.Application.Presentations`** (`_base.py`):
   WPS's single-process model occasionally hands back an app object whose
   collection property is not ready. `__enter__` now re-dispatches and
   retries up to 3 times (0.6 s backoff); 4 consecutive clean verification
   runs after the fix (was ~1-in-3 failure).

## Test-suite fixes for Windows (platform-independent suite)

- `macos/wps-jsapi-probe` requires `npm ci` once — several `macos_probe` and
  `test_generation.py` tests need the real `wpsjs` template fixtures
  (`node_modules/wpsjs/.../res/wpsDemo.docx`). Without it: 43 failures.
- `tests/macos_probe/test_addin_assets.py` ran 100 KB scripts through
  `node -e` — over the Windows 8191-char command-line limit. Scripts now go
  through a temp file.
- 10 macOS-probe tests assert POSIX-only behaviour (`fcntl`, `ps`, `0o700`
  file modes, `/` paths, executable shell stubs) and are now
  `skipif(os.name != "posix")`.

## WPS quirks worth knowing (discovered this run)

- WPS preserves `w14:paraId` across saves and assigns ids to new paragraphs.
- zh-CN WPS localizes built-in style names (`标题 1`); use style ids -2..-10.
- `Range.WordOpenXML` exists but returns the whole Flat-OPC package per call —
  too heavy for per-paragraph id readback; the zip walk is the right source.
- `edit()` batches run patches first, then ops. `@paraId` resolution after an
  unsaved structural change in the same batch cannot work (the id map comes
  from the on-disk package) and fails cleanly as `invalid_target`; address
  post-insert elements in a later batch (re-inspect).
- Headless WPS automation leaks processes over long sessions (Quit does not
  always reap them) — kill stray `wps`/`et`/`wpp` processes after big runs.
- WPS `Slides.MoveTo(p)` inserts BEFORE the slide at pre-move index `p`
  (source still counted); `MoveTo(count+1)` appends to the end without
  raising. `_move_slide`/`_clone_slide` convert the desired final position
  `F` to `F+1 when F > source else F`. Sample decks generated from markdown
  can repeat a title across slides ("Deck Title" on slides 1-2) — use unique
  stamps, not titles, to assert landing positions.

## Implementation notes (all borrowable points now implemented)

All six borrowable points are implemented **and COM-verified on Windows**. The
pure-Python layers are tested cross-platform; the COM-coupled layers passed the
live checklists A–G and the `# WINDOWS-VERIFY` tags are cleared.

| # | Borrowable point | Status | Where |
|---|---|---|---|
| 1 | Stable ID addressing (`@paraId` / `@id` / `@name`) | implemented + COM-verified | checklist D |
| 2 | Structured errors + self-heal suggestions | implemented + tested | `document_api.py::_error_report`, `PatchError` |
| 3 | Atomic batch | implemented + tested | `edit()` / `apply_ops()` |
| 4 | dump → replay | implemented + COM-verified | `snapshot_to_patches()` (checklist F) |
| 5 | Built-in help | implemented + tested | `patch_grammar()` / `validate_target()` |
| 6 | Structural editing (insert/remove/move/clone) | implemented + COM-verified | checklist G |

### Known limitations to keep in mind on Windows

- **Word `@paraId` requires a saved .docx.** Unsaved / never-saved documents
  have no on-disk package to read paraIds from, so `inspect()` falls back to
  positional ids for them. Saving once enables stable ids. This is by design,
  not a bug.
- **Word paraId count must match `doc.Paragraphs.Count`.** `_read_paraid_map`
  enforces this and falls back to positional on mismatch. The alignment
  (text-box story pruned, one `None` per table row-end mark) was validated on
  real documents with tables and textboxes; exotic stories (nested tables,
  SDT-wrapped rows) may still mismatch — positional fallback covers them.
- **`snapshot_to_patches` fidelity is bounded by snapshot/apply key symmetry.**
  A snapshot field that `apply_font`/`apply_fill` does not accept lands in the
  patch's `rejected` list. `None`-valued fields are dropped at emission time
  (they mean "host reported no value").
- **`shape.Id` / `shape.Name` resolution compares Python int / str equality.**
  If a COM host returns `Id` as a non-int variant, `_find_shape_in_*` won't
  match; coerce in the helper if that surfaces.

## Re-running

Platform-independent suite (must stay green on Windows too):

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -e ".[dev]"
.venv\Scripts\python -m pytest -v
```

The full suite was 533 passed + 1 skipped (pypdf) on macOS before this change;
after this change it is 609 passed + 1 skipped (77 tests in
`tests/test_document_api.py`: orchestration, stable-id grammar, paraId
extraction, snapshot_to_patches, and structural verbs across all hosts).
On Windows it is 600 passed + 11 skipped (10 POSIX-only macOS-probe tests +
pypdf). One-time setup: `npm ci` in `macos/wps-jsapi-probe` (test fixtures).

COM smoke (Windows only, ad hoc):

```powershell
.venv\Scripts\python -c "from skills.WPSComposer import inspect; print(inspect('fixtures/sample.docx')['counts'])"
```

## Tested environment

| Item | Value |
|---|---|
| Test date | 2026-07-27 |
| Windows build | 10.0.26200 |
| WPS Office / MS Office version | WPS Office 12.1.0.26899 (zh-CN) |
| `pywin32` version | 312 |
| Python | 3.14.3 |
| `tests/test_document_api.py` | pass (77) |
| Full suite | 600 passed, 11 skipped (10 POSIX-only + pypdf) |
| Items A–G above | all pass |

## Windows run results

Live COM verification scripts (runnable, kept in `fixtures/`):

| Script | Covers | Result |
|---|---|---|
| `fixtures/verify_ab.py` | A (error classification ×3 hosts), B (atomic no-save, happy path, source unchanged) | 6/6 |
| `fixtures/verify_c.py` | C (attach-active caveat, both halves) | 4/4 |
| `fixtures/verify_d.py` | D (`@paraId` readback 27/30 ids, stability across insert+save; slide/sheet `@id`/`@name`) | 4/4 |
| `fixtures/verify_e.py` | E (edge target forms: `$A$1`, run path, table cell, `range:0-10`) | 5/5 |
| `fixtures/verify_f.py` | F (font dump→replay round-trip, no rejected keys, title-font fidelity) | 1/1 |
| `fixtures/verify_g.py` | G (31 checks: full insert/remove/move/clone matrix on all hosts + last-sheet guard + exact slide move/clone landing positions) | 31/31 |
| `fixtures/verify_cr1.py` | Code-review R1: attach + `output=` does not rebind the live document (save_copy guard) | pass |
| `fixtures/verify_r1.py` | Completeness R1: writer move table/shape + clone shape, slide table-cell target, paraId alignment on more document shapes | pass |

Fixtures: `fixtures/make_fixtures.py` + `fixtures/add_extras.py` build
`sample.docx` (≥3 paragraphs + table + textbox), `sample.pptx` (4 slides),
`sample.xlsx` (4×3 data + chart) via the real engine.

## Files touched this session

- `skills/WPSComposer/scripts/document_api.py` — grammar data (`PATCH_GRAMMAR`
  with `@paraId`/`@id`/`@name` forms), `validate_target`, `patch_grammar`,
  `snapshot_to_patches`, `PatchError`, `apply_ops`/`validate_op` (structural
  verbs), rewritten `apply_patches`/`edit`
- `skills/WPSComposer/scripts/writer.py` — pure `_extract_paraids` /
  `read_paraids_from_docx`; COM wiring `_read_paraid_map` /
  `_paragraph_index_for_paraid`; `@paraId=` branch in `apply_format_patch`;
  paraId readback in `inspect_document`; `apply_structural_op` (insert/remove/
  move/clone)
- `skills/WPSComposer/scripts/slide.py` — `shape.Id` readback in
  `_shape_snapshot`; `@id=` / `@name=` branches + `_find_shape_in_slide` /
  `_apply_shape_patch`; `apply_structural_op` (insert slide/textbox/image,
  remove, move, clone)
- `skills/WPSComposer/scripts/sheet.py` — `shape.Id` readback; `@id=` /
  `@name=` branches + `_find_shape_in_sheet` / `_apply_shape_patch`;
  `apply_structural_op` (insert/remove/move row, remove shape)
- `skills/WPSComposer/scripts/wps_engine.py` — new exports
- `skills/WPSComposer/references/api.md` — structured-result / atomic /
  stable-id / replay / structural-verbs docs
- `skills/WPSComposer/SKILL.md` — pointer to new functions + this doc
- `tests/test_document_api.py` — fake-composer + pure-helper suite (77 tests)
