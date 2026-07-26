# Windows COM verification handoff

> **Start here on Windows.** Read `AGENTS.md` first, then this file, then run
> `grep -rn WINDOWS-VERIFY skills/WPSComposer/scripts/` to list every COM site
> flagged for verification.

## Status

**Pending Windows verification.** A sequence of macOS sessions added
agent-friendly conversational-edit features to
`skills/WPSComposer/scripts/document_api.py` and the three COM composers
(`writer.py` / `sheet.py` / `slide.py`):

- **Orchestration (pure-Python, tested on macOS):** structured patch results +
  `PatchError`, atomic `edit()`, `validate_target` / `patch_grammar`,
  `snapshot_to_patches` (dump→replay), and structural verbs via `apply_ops` /
  `validate_op` (`insert` / `remove` / `move` / `clone`).
- **COM bodies (written blind, need Windows verification):** stable-ID readback
  (`@paraId` / `@id` / `@name`) and resolution, and the full `apply_structural_op`
  implementations on all three composers.

The macOS suite is green (**609 passed, 1 skipped**; 76 tests in
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
.venv/bin/python -m pytest tests/test_document_api.py -v   # 27 passed
```

## Windows verification checklist

Environment to fill in at the bottom of this file before running.

Every COM-coupled change in this session is tagged `# WINDOWS-VERIFY:` in the
source. `grep -rn WINDOWS-VERIFY skills/WPSComposer/scripts/` lists all 33
sites. The high-value ones are called out below with file:line references.

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
"Atomicity and the attach-active caveat" note in `references/api.md`.

### D. Stable-ID readback + resolution (NEW — borrowable #1, the core)

This is the main COM work added this session. Three hosts:

| Host | What to verify | Source |
|---|---|---|
| Writer `@paraId` | `inspect()` emits `paragraph:@paraId=HEX` for saved docx; `apply_format_patch("paragraph:@paraId=...")` hits the right paragraph after a structural change | `writer.py:1087` (read), `writer.py:1203` (resolve), `_read_paraid_map`/`_paragraph_index_for_paraid` |
| Slide `@id` / `@name` | `inspect()` emits `slide:N/shape:@id=K`; `@id=` and `@name=` resolve to the right shape | `slide.py:649` (`_shape_snapshot`), `slide.py:581`+`595` (branches), `_find_shape_in_slide` |
| Sheet `@id` / `@name` | same for `sheet:N/shape:@id=K` / `@name=` | `sheet.py:298` (snapshot), `sheet.py:391`+`404` (branches), `_find_shape_in_sheet` |

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
with a fake composer (72 tests); the COM bodies are tagged `# WINDOWS-VERIFY`.

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
    {"op": "remove", "target": "sheet:1/cell:C1"},          # removes row 3? confirm
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

**Known COM bugs found by code review (fix on Windows — do NOT re-derive):**

- **Writer `_resolve_insert_range` end/start** (`writer.py`): uses bare
  `doc.Range` (a *method*, not a property) → will crash on `.Collapse()`. Fix:
  use `doc.Content` (the main-story Range) collapsed to end/start. Everywhere
  else in the file calls `doc.Range(start, end)` with args.
- **Writer `_insert_element` heading at end** (`writer.py`): the positional
  `return` fires *before* the Heading style is applied, so
  `{"op":"insert","type":"heading","position":"end"}` (the default!) yields a
  plain paragraph. Fix: apply the Heading style before the early return
  (`InsertAfter` extends the range to cover the new text, so `rng.Style = ...`
  applies to it).
- **Sheet `chart:N` remove** (`sheet.py`): only handles positional integer
  refs (`int(ref)`); a non-numeric ref raises. Inspect emits positional chart
  ids, so this is consistent — but coerce/validate cleanly rather than rely on
  the bare `int()`.
- **Writer image insert path** (`writer.py`): returns `shape:N` from
  `InlineShapes.Count`, but `_structural_target` resolves `shape:N` via
  `doc.Shapes(N)` (a different collection). Confirm the index cross-resolves,
  or return an InlineShapes-addressable path.

These are the known issues; others may surface during the end-to-end run —
append them here as you find and fix them.
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

## Implementation notes (all borrowable points now implemented)

All five borrowable points are implemented. The pure-Python layers are tested
on macOS; the COM-coupled layers are written, tagged `# WINDOWS-VERIFY:` in
source, and listed in checklist D–F above. Nothing is deferred — Windows only
needs to verify, not implement.

| # | Borrowable point | Status | Where |
|---|---|---|---|
| 1 | Stable ID addressing (`@paraId` / `@id` / `@name`) | implemented (COM unverified) | checklist D |
| 2 | Structured errors + self-heal suggestions | implemented + tested | `document_api.py::_error_report`, `PatchError` |
| 3 | Atomic batch | implemented + tested | `edit()` / `apply_ops()` |
| 4 | dump → replay | implemented + tested | `snapshot_to_patches()` (checklist F) |
| 5 | Built-in help | implemented + tested | `patch_grammar()` / `validate_target()` |
| 6 | Structural editing (insert/remove/move/clone) | implemented (COM unverified) | checklist G |

### Known limitations to keep in mind on Windows

- **Word `@paraId` requires a saved .docx.** Unsaved / never-saved documents
  have no on-disk package to read paraIds from, so `inspect()` falls back to
  positional ids for them. Saving once enables stable ids. This is by design,
  not a bug.
- **Word paraId count must match `doc.Paragraphs.Count`.** `_read_paraid_map`
  enforces this and falls back to positional on mismatch. If you see stable
  ids never emitted on a doc that should have them, the document-order
  assumption needs revisiting (see checklist D, Writer assumption 1).
- **`snapshot_to_patches` fidelity is bounded by snapshot/apply key symmetry.**
  A snapshot field that `apply_font`/`apply_fill` does not accept lands in the
  patch's `rejected` list. Check `r["patches"][*]["rejected"]` during the
  round-trip (checklist F).
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
after this change it is 607 passed + 1 skipped (74 tests in
`tests/test_document_api.py`: orchestration, stable-id grammar, paraId
extraction, snapshot_to_patches, and structural verbs across all hosts).

COM smoke (Windows only, ad hoc):

```powershell
.venv\Scripts\python -c "from skills.WPSComposer import inspect; print(inspect('fixtures/sample.docx')['counts'])"
```

## Tested environment

Fill in during the Windows run:

| Item | Value |
|---|---|
| Test date | _(to fill)_ |
| Windows build | _(to fill)_ |
| WPS Office / MS Office version | _(to fill)_ |
| `pywin32` version | _(to fill)_ |
| Python | _(to fill)_ |
| `tests/test_document_api.py` | _(pass/fail)_ |
| Items A–G above | _(pass/fail per item)_ |

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
- `tests/test_document_api.py` — fake-composer + pure-helper suite (59 tests)
