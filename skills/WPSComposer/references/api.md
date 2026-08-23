# WPS Composer API Reference

## Long-form offline plan and optional native execution

This Python API is for WPSComposer/SuperWriter orchestration code, plugin
maintainers, and advanced integrations that need a deterministic plan before
starting WPS. Normal users should keep using `generate()`; M4 does not reroute
that public entry point.

```python
from skills.WPSComposer.scripts.longform import (
    LongformBuild,
    build_longform_generation,
    execute_longform_plan,
)

build = build_longform_generation(markdown, base_dir="assets")
outcome = execute_longform_plan(build, executor, deadline=deadline)
```

```python
def build_longform_generation(
    markdown: str,
    base_dir: str = "",
) -> LongformBuild:
    ...

def execute_longform_plan(
    build: LongformBuild,
    executor: LongformExecutor,
    deadline: Optional[float] = None,
) -> ExecutionOutcome:
    ...
```

`build_longform_generation()` parses and normalizes Markdown, preflights local
media once, and returns a closed protocol-v2 plan without launching WPS or
writing a file. `execute_longform_plan()` validates that plan and passes an
immutable tuple of private `PreparedLongformResource` objects to the caller's
dedicated executor. Source paths are not re-read at the executor boundary;
the pipeline's transient normalized-payload tuple is released on every exit
path. Executor-staged copy cleanup runs on success, error, timeout, and save
failure; cleanup failure is itself fatal and cannot silently publish a result.

### M4 Markdown contract

| Object | Attributes / syntax | Native result |
|---|---|---|
| Figure | `#id`, `caption`, `width=auto\|column\|full\|Npt`, `orientation=portrait\|landscape`, `kind`, `layout=stack\|columns`, `columns=2` | native image container; caption below; one or two children |
| Table | `#id`, `caption`, `style=three-line\|grid`, `orientation`, `merges=A2:A3;B2:C2`, `repeat_header` | native table; caption above; repeated header and direct alignment |
| Formula | `:::equation {#eq:id fallback_image="relative.png"}`; legacy `:::formula` | validated editable Office Math attempt plus the M3 equation-number shell, or marked image/source fallback |
| Reference | `{{ref:target-id}}` outside code/math spans | `REF <safe-bookmark> \\h` in the same paragraph |
| Citation | `{{cite:id}}` in visible inline content | deterministic numeric `[n]` in the same paragraph or table cell |
| Bibliography | `:::bibliography` with `[id] text` lines; legacy `:::references` | cited-first native hanging paragraphs with fixed spacing |

References in abstracts and mixed ordered/unordered lists use the same native
field contract. List items retain controlled markers, tabs, `List Paragraph`
style, and a 24 pt hanging indent on both Windows COM and macOS JSAPI.

Front matter accepts `caption_numbering: auto|global|chapter`, `figure_index`,
and `table_index`. `auto` is object-local: before the first numbered H1 it is
global, after that H1 it is chapter-based, and an unnumbered H1 is transparent
to sequence reset. Explicit `chapter` falls back to global before a numbered
H1. Controlled sequence identifiers are `WPSC_FIG`, `WPSC_TAB`, and
`WPSC_EQ`; user text never enters field codes or generated bookmark names.

Formula source uses one bounded restricted-LaTeX grammar (10,000 Unicode code
points and 64 brace levels). It accepts Unicode variables, scripts, fractions,
roots, sums/products/integrals with limits, scalable delimiters, common
Greek/operators/relations, matrices, cases, and bounded nesting. It rejects
custom/unknown commands, packages, file/URL/shell access, malformed or oversized
structures, and external resources before WPS starts. Optional `fallback_image`
is resolved relative to `base_dir`, normalized privately under the M3 media
limits, bound to exactly one formula, and never serialized as a path or payload.

Citation numbers follow first visible semantic occurrence across abstract/body
paragraphs, lists, block quotes, page-break paragraphs, and table cells.
Repeated citations reuse the number. Bibliography output is cited-first;
remaining declarations retain source order unless front matter sets
`bibliography_include_uncited: false`. Missing citations remain visibly marked
in their owning paragraph/cell and never disappear silently.

Figure and table index operations are native and populated from those internal
sequence labels. Field finalization runs numbering, bookmark/reference,
index, and page-field phases in order. Two adjacent equal snapshots must occur
within three mutation rounds; otherwise a read-only fourth snapshot freezes
the document and emits one `FIELD_REFRESH_UNSTABLE` issue.

### Media, sizing, and privacy

Accepted decoded formats are PNG, JPEG, TIFF, BMP, GIF, and restricted static
SVG. WebP, remote resources, active SVG content, corrupt media, payloads over
50 MiB, images over 80,000,000 pixels, and either side over 32,768 pixels are
rejected. EXIF transforms, the first GIF frame, and the first TIFF page use a
private lossless-PNG normalized payload. Width is resolved in points against
the current portrait/explicit-landscape content slot; paired columns use a
12 pt gap and never trigger orientation changes heuristically.

Plan and diagnostic JSON omit payload bytes/base64, source paths, source and
payload hashes, private locators, bookmark lookup maps, and visible field
values/hashes. The executable plan retains only its required manifest digest,
controlled bookmark descriptors, and opaque resource IDs.

### Table and degradation contract

Academic tables default to `three-line`: 1.5 pt top/bottom borders, 0.75 pt
header-bottom border, no vertical/interior body borders, zero cell indent,
direct alignment, repeated header, and no row splitting. Other presets default
to `grid`. Merge ranges are one-based A1 rectangles and are applied only when
the complete declaration validates. Any invalid range discards every requested
merge, preserves the full grid, and places one `TABLE_MERGE_INVALID` notice
after the caption. An over-page vertical group recovers to the one specified
unmerged/splittable grid form.

M4 uses one closed recovery matrix on Windows and macOS. Formula code
`EQUATION_INSERT_FAILED` permits one `explicit-image-then-source-notice` ladder:
attempt native math, then at most one validated image; if unavailable or failed,
keep readable source in the terminal marked notice. `CROSS_REFERENCE_FAILED`
stays inline, and `BIBLIOGRAPHY_INSERT_FAILED` stays at the bibliography
paragraph. Block notices use a restrained local box; document notices
deduplicate at the single empty `生成质量提示` anchor. Empty optional content leaves
that anchor invisible and creates no issue, paragraph, spacing, or notice.

On verified macOS WPS 12.1.26055, professional `BuildUp` remains a structural
no-op for the supported formula families. The executor verifies the native
structure and honestly uses image/source degradation instead of claiming
Type-20 linear text as native success. This is a runtime capability result, not
a parser relaxation. Windows receives the same descriptor and postcondition;
real Windows M4 verification is deferred to the M5/final gate.

Only named object-local failures may recover. `ENGINE_LOST`, acquisition or
configuration failure, protocol/capability mismatch, unknown native errors,
rollback failure, required field/index/repagination failure, staging/hash or
cleanup failure, terminal source/notice fallback failure, and
save/export/validation/publication failure are fatal. A recoverable image-rung
failure may still roll back into the declared source notice. Diagnostics and
notices never expose paths, payload/base64,
resource hashes or locators, bookmark maps, field values/hashes, or exception
representations.

M5/final work owns PDF geometry/bbox quality checks, deterministic full
re-layout and notice-only patches, performance gates, public `generate()`
default migration, the real Windows cross-platform gate, and release/version
publication. M4 still returns only the artifact format requested by the caller.

## Office-to-PDF conversion

```python
from skills.WPSComposer import ConversionError, convert_to_pdf

result = convert_to_pdf("report.docx")
result = convert_to_pdf("book.xlsx", "exports/book.pdf", overwrite=True)
```

```python
def convert_to_pdf(
    source: str,
    output: Optional[str] = None,
    *,
    overwrite: bool = False,
) -> str:
    ...
```

Accepted source suffixes are `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, and
`.pptx`, matched case-insensitively. The default destination is the source
sibling with a `.pdf` suffix. Success returns its absolute path. Excel export
is workbook-level and includes every visible worksheet.

| Condition | Exception |
|---|---|
| Source does not exist | `FileNotFoundError` |
| Unsupported source or non-PDF destination suffix | `ValueError` |
| Destination exists and `overwrite=False` | `FileExistsError` |
| WPS/Office, protocol, modal, or validation failure | `ConversionError` |

`ConversionError` exposes `code`, `source`, `component`, `backend`, and
`message`, plus `to_dict()`. Windows uses WPS/MS Office COM. macOS uses typed
WPS JSAPI commands with source and output staged in WPS's private application
container, followed by validated destination-local atomic publication.
Vendor-specific error codes are normalized to the stable public set. Common
codes include `CONVERSION_COMMAND_FAILED`, `INTERACTIVE_INPUT_REQUIRED`,
`NO_VISIBLE_WORKSHEETS`, `STAGED_ARTIFACT_INVALID`,
`ARTIFACT_PUBLISH_FAILED`, `FINAL_ARTIFACT_INVALID`, and
`REGISTRATION_RESTORE_FAILED`. Private staging paths are redacted from public
messages.

## Color format

All color args accept **either** a `#RRGGBB` hex string (e.g. `"#4472C4"`) **or**
a BGR integer (`0xBBGGRR`). The engine auto-converts hex strings to COM's BGR
longs internally.

## Unified existing/current document API

Use this API for conversational formatting work. It supports three workflows:
opening an existing file, attaching to the document currently visible in WPS,
and applying ordered element-level patches.

```python
from wps_engine import open_document, attach_active, inspect, edit

# Inspect an existing file. The snapshot contains addressable element ids.
snapshot = inspect("report.docx")

# Work directly on the document currently open in WPS Writer.
w = attach_active("writer")
selection = w.inspect_selection()
w.apply_format_patch(
    "paragraph:3",
    font={"name": "等线", "size": 11, "color": "#333333"},
    paragraph={"alignment": 3, "line_spacing": 20, "space_after": 6},
)
w.save_current()

# Batch-edit a copy while preserving the original. Atomic by default: if any
# patch fails, nothing is saved and the result carries ok=False + errors.
result = edit(
    "report.docx",
    output="report-revised.docx",
    overwrite=False,
    export_pdf="report-revised.pdf",
    patches=[
        {"target": "paragraph:1", "font": {"size": 20, "bold": True}},
        {"target": "table:1/cell:2,1", "fill": {"color": "#FFF2CC"}},
    ],
)
# result == {"ok": bool, "saved": bool, "saved_path": str|None,
#            "pdf_path": str|None, "patches": [...], "errors": [...],
#            "before": ...|None, "after": ...|None}

# Validate a target before applying -- lets an agent self-correct instead of
# guess-fail-retry. Returns {"valid": bool, "element": ..., "error": {...}}.
validate_target("paragraph:3", "writer")
# -> {"valid": True, "kind": "writer", "element": "paragraph", "form": "paragraph:N"}

validate_target("paragraf:3", "writer")
# -> {"valid": False, "error": {"code": "invalid_target",
#       "valid_forms": [...], "closest": "paragraph:N"}}
```

Common functions:

| Function | Purpose |
|---|---|
| `open_document(path, kind=None, read_only=False, visible=False)` | Open a supported existing file; returns a context-manageable composer |
| `attach_active(kind=None)` | Attach to the user's active Writer/Sheet/Slide without closing it later; auto-detects when omitted |
| `inspect(path=None, kind=None, selection=False, **options)` | Return a JSON-compatible document or selection snapshot |
| `edit(path=None, kind=None, patches=None, ops=None, output=None, export_pdf=None, atomic=True, raise_on_error=False, overwrite=False)` | Apply patches and/or ops and save in place or to a copy. `patches` is sugar for `{"op":"set",...}` and runs before `ops`; one atomic transaction. Atomic by default: on any failure the document is **not** saved and a structured `{"ok": False, "errors": [...]}` result is returned. Existing outputs require `overwrite=True`. With `export_pdf`, both artifacts are staged and validated before group publication with rollback; attached/macOS edit export is unsupported. |
| `apply_ops(composer, ops, atomic=True)` | Unified op executor (`set`/`insert`/`remove`/`move`/`clone`); raises `PatchError` in atomic mode |
| `apply_patches(composer, patches, atomic=True)` | Back-compat wrapper: `set`-only patches, normalised to `apply_ops` |
| `validate_op(op, kind=None)` | Validate one op dict against the schema; returns `{valid, error:{code,...}}` |
| `validate_target(target, kind)` | Validate a patch target against the grammar; returns a suggestion on miss |
| `patch_grammar(kind=None)` | Return the address grammar as plain data (agent help / discovery) |
| `snapshot_to_patches(snapshot, dimensions=("font","paragraph","fill"))` | Convert an inspect snapshot to a replayable patch list (dump→replay) |
| `PatchError` | Exception raised in atomic mode; `.reports` and `.errors` carry structured per-op results |
| `supported_formats()` | Return recognized Writer/Sheet/Slide extensions |
| `snapshot_json(snapshot)` | Serialize a snapshot without losing Chinese text |

Every composer also exposes `inspect_document()`, `inspect_selection()`,
`apply_format_patch(target, **patch)`, `save_current()`, and `save(path)`.
Patches are partial: properties omitted by the agent are not reset.

Each patch report is structured for agent consumption:

```python
{"index": 0, "target": "paragraph:1", "ok": True,
 "accepted": ["font"], "rejected": []}
# on failure:
{"index": 1, "target": "paragraf:2", "ok": False,
 "error": {"code": "invalid_target", "message": "...",
           "valid_forms": ["paragraph:N", ...], "closest": "paragraph:N"}}
```

Error codes: `missing_target`, `invalid_target`, `invalid_value`,
`apply_failed`, `unsupported_kind`.

### Atomicity and the attach-active caveat

`edit()` is **atomic by default** (`atomic=True`): if any patch fails, the
document is not saved (no file is written for `output=`, and `save_current()`
is not called for the active document). Set `atomic=False` for the legacy
best-effort behaviour that saves whatever succeeded.

For `path=None` (attach-active), only a single `set` operation containing at
most one leaf property is accepted in atomic mode. Composite patches and all
structural operations are rejected **before the first mutation** because
WPS/Office exposes no reliable rollback boundary for the live window. Use
`atomic=False` explicitly for best-effort live mutation, or edit a file-backed
copy. An existing `output=` is rejected before attaching unless
`overwrite=True`; after attachment, its document family is validated before
mutation. The live document keeps its binding and the output is published from
a validated destination-local staging file.

File-backed edits preserve the source document family. A distinct destination
must not exist unless `overwrite=True`; successful saves are validated in a
same-directory staging file and then atomically published. On macOS, editing is
currently restricted to `.pptx` input and `.pptx` output so legacy, slideshow,
and macro-enabled presentation semantics cannot be silently lost.

### Address grammar

| Host | Targets |
|---|---|
| Writer | `selection`, `paragraph:N`, `paragraph:@paraId=HEX` *(stable)*, `range:START-END`, `table:N/cell:R,C`, `shape:N`, `section:N` |
| Sheet | `selection`, `sheet:N`, `sheet:N/cell:A1`, `sheet:N/range:A1:C20`, `sheet:N/shape:N`, `sheet:N/shape:@id=N` *(stable)*, `sheet:N/shape:@name=NAME` *(stable)*, `sheet:N/chart:N` |
| Slide | `selection`, `presentation`, `slide:N`, `slide:N/shape:N`, `slide:N/shape:@id=N` *(stable)*, `slide:N/shape:@name=NAME` *(stable)*, `.../paragraph:N`, `.../paragraph:N/run:N`, `.../table/cell:R,C`; stable `@id` supports all three nested suffixes |

All indices are 1-based. **Stable forms** (`@paraId` / `@id` / `@name`) survive
structural edits between `inspect()` and `edit()` and are emitted by
`inspect()` whenever the host exposes a native id — prefer them over
positional indices. Inspection snapshots return these exact ids so agents
do not need to reconstruct selectors manually. `validate_target(target, kind)`
checks any candidate against this grammar and returns the closest valid form
on a miss; `patch_grammar(kind)` returns the table above as plain data.

### Replay a snapshot's formatting (dump → replay)

`snapshot_to_patches(snapshot, dimensions=("font", "paragraph", "fill"))` walks
an `inspect()` snapshot and emits a patch list that reproduces the captured
formatting on another document — the WpsComposer analogue of officecli's
`dump` → `batch`, scoped to formatting (full structural cloning is
`generate()`'s job). Only non-empty requested dimensions are emitted.
Stable ids (`@paraId` / `@id` / `@name`) are document-specific, so emitted
targets are rewritten to their positional form (`paragraph:N`, `.../shape:N`)
via the element's index; fields the host reported as `None` are dropped.

```python
from skills.WPSComposer import inspect, snapshot_to_patches, edit
snap = inspect("styled-sample.docx")
patches = snapshot_to_patches(snap, dimensions=("font", "fill"))   # copy fonts + fills
edit("target.docx", output="target-styled.docx", patches=patches)
```

### Structural operations (insert / remove / move / clone)

`edit()` and `apply_ops()` accept an `ops=[...]` list of typed operations
alongside (or instead of) formatting `patches`. Each op carries an `op` verb:

| Verb | Purpose | Required keys |
|---|---|---|
| `set` | apply a formatting patch (== `patches` entries) | `target`, + format keys |
| `insert` | add a new element | `type`, optional `parent`, `props`, `position` |
| `remove` | delete an element | `target` |
| `move` | relocate an element | `target`, `to` (position) |
| `clone` | duplicate an element | `target`, `to` (position) |

`patches=[...]` is sugar for `ops=[{"op":"set", ...}, ...]` and runs first;
pass `ops=` for structural verbs or to interleave order. All ops share one
atomic transaction.

Insert-able `type` per host:

| Host | Insert types |
|---|---|
| Writer | `paragraph`, `heading` (`text`, `level`), `page_break`, `table` (`rows`, `cols`, `data`), `image` (`path`), `textbox` (`text`, `left/top/width/height`) |
| Slide | `slide` (`layout`), `textbox`/`shape` (`text`, geometry), `image` (`path`, geometry) |
| Sheet | `row`, `column` (`values`, `position.index`), `sheet` (`name`) |

`remove` / `move` / `clone` accept any addressable element the host resolves:

| Host | remove | move | clone |
|---|---|---|---|
| Writer | paragraph / shape / table | paragraph / shape / table (clipboard) | paragraph / shape / table (clipboard) |
| Slide | slide / shape | slide (reorder) / shape (to `{"slide": N}`) | slide / shape (`Duplicate`, or to `{"slide": N}`) |
| Sheet | row / column / shape / chart / sheet | row / column / sheet | row / sheet |

`position` / `to` (for insert/move/clone): `"end"` (default), `"start"`,
`{"after": target}`, `{"before": target}`, or `{"index": N}`; Slide shape
move/clone accept `{"slide": N}` to target another slide; Sheet sheet
move/clone accept `{"before": N}` / `{"after": N}`. Sheet `remove`/`move` on a
cell/range target take `"axis": "row"` (default) or `"axis": "column"` to pick
which axis the address resolves to.

```python
from skills.WPSComposer import edit
edit("deck.pptx", output="deck2.pptx", ops=[
    {"op": "insert", "type": "slide", "props": {"layout": 12}},
    {"op": "insert", "parent": "slide:3", "type": "textbox",
     "props": {"text": "New callout", "left": 100, "top": 100, "width": 300}},
    {"op": "clone", "target": "slide:1", "to": "end"},
    {"op": "remove", "target": "slide:2"},
    {"op": "set", "target": "slide:1/shape:@id=7", "font": {"size": 18}},
])
```

`validate_op(op, kind)` checks any op against the schema and returns an
`error.code` (`unknown_verb` / `missing_type` / `unsupported_type` /
`missing_target` / `invalid_anchor` / `invalid_position`) on a miss.

**Positional-id drift caveat:** structural verbs shift sibling indices. Address
subsequent ops by stable id (`@paraId` / `@id`) or re-`inspect()` between
batches. Returned `path` values for inserted/cloned elements are best-effort
positional; re-inspect for a stable id after a save. Writer `image` inserts
return an `inline_shape:N` path (InlineShapes are a separate collection from
floating `shape:N`); it resolves for `remove` / `move` / `clone`.

### Recognized input formats

- Writer: DOC/DOCX/DOCM/DOT/DOTX/DOTM, RTF, TXT, HTML/MHTML, XML, ODT,
  and WPS/WPT.
- Sheet: XLS/XLSX/XLSM/XLSB/XLTX/XLTM, CSV/TSV, ODS, XML/HTML, and ET/ETT.
- Slide: PPT/PPTX/PPTM, PPS/PPSX/PPSM, POT/POTX/POTM, ODP, and DPS/DPT.

Host support still depends on the installed WPS/Office version. PDF is handled
by `PdfComposer`; PDF does not expose editable WPS layout elements.

## WriterComposer (docx) — KWps.Application

### Page setup
| Method | Args | Notes |
|---|---|---|
| `set_columns(count)` | int | multi-column layout |
| `set_orientation(landscape)` | bool | True=landscape |
| `set_margins(top,bottom,left,right)` | float pts | |
| `set_page_size(width, height)` | float, float | custom page size |
| `add_section()` | | insert section break (independent page layout) |

### Header / Footer
| Method | Args | Notes |
|---|---|---|
| `set_header(text)` | str | primary header |
| `set_footer(text)` | str | primary footer |
| `set_page_number_in_footer()` | | "Page N" with field |

### Text
| Method | Args | Notes |
|---|---|---|
| `add_heading(text, size=None, bold=True, color=None)` | str,int,bool,color | Heading 1 style; default size comes from the centralized heading map |
| `add_heading2(text, size=None, color=None)` | str,int,color | Heading 2 style; default size comes from the centralized heading map |
| `add_paragraph(text, size,bold,italic,color,align,indent_first,line_spacing,line_spacing_rule,space_before,space_after,font_name,font_name_ascii)` | str,... | align: 0L,1C,2R,3justify; semantic rules include `single`, `one_and_half`, `exact`, `multiple` |
| `add_centered(text, size=24, bold=True, color=None)` | str,... | centered paragraph |
| `add_bullet_list(items)` | list[str] | bullet style |
| `add_numbered_list(items)` | list[str] | numbered style |

### Tables
| Method | Args | Notes |
|---|---|---|
| `add_table(rows,cols,data,shade_header,header_color,font_size,col_widths,alignments,auto_fit)` | ... | styled table; `auto_fit=True` uses content-aware column widths, while explicit `col_widths` remain point values |
| `add_merged_table(data, merges, shade_header)` | list, list[(r1,c1,r2,c2)] | merged cells |

### Shapes & images
| Method | Args | Notes |
|---|---|---|
| `add_floating_textbox(text,left,top,width,height,wrap,fill_color,font_size,bold)` | ... | floating box w/ fill |
| `add_wordart(text,left=200,top=400,preset=0)` | str,... | WordArt |
| `add_image(path, width=None, height=None, wrap=0, *, max_width=None, max_height=None, inline=True, preserve_aspect=True, alt=None)` | str,... | inline by default; proportional fit within optional bounds |
| `add_page_break()` | | page break |
| `add_horizontal_line()` | | horizontal rule |
| `insert_toc(title="Table of Contents")` | str | auto TOC (levels 1-3) |
| `update_fields()` | | refresh TOC/fields |

### Output
| Method | Args | Notes |
|---|---|---|
| `save_docx(path)` | str | returns abs path |
| `export_pdf(path)` | str | returns abs path |

---

## SheetComposer (xlsx) — Ket.Application

### Sheets
| Method | Args | Notes |
|---|---|---|
| `select_sheet(index)` | int | switch active sheet |
| `add_sheet(name=None)` | str | new sheet, becomes active |

### Cells
| Method | Args | Notes |
|---|---|---|
| `write_cell(row, col, value)` | int,int,any | single cell |
| `set_formula(row, col, formula)` | int,int,str | e.g. "=SUM(B2:B4)" |
| `write_table(start_row,start_col,data,header_shade,header_font_color,font_size)` | ... | styled table w/ colored header |

### Styling
| Method | Args | Notes |
|---|---|---|
| `set_cell_style(row,col,bold,italic,font_size,font_color,fill_color,align,number_format)` | ... | per-cell |
| `set_range_style(range_str,bold,italic,font_size,font_color,fill_color,align,number_format)` | str,... | range e.g. "A1:C3"; align: -4108 C,-4131 L,-4152 R |
| `set_borders(range_str,style=1,weight=2,color)` | str,... | all 6 borders |
| `merge_cells(range_str)` | str | merge range |
| `set_column_width(col, width)` | str/int, float | |
| `set_row_height(row, height)` | int, float | |
| `freeze_panes(cell)` | str | e.g. "A3" |

### Advanced
| Method | Args | Notes |
|---|---|---|
| `add_chart(chart_type,left,top,width,height,source_range,title)` | int,...,str | xlColumnClustered=51, xlLine=4, xlPie=5 |
| `add_title_row(range_str, text, fill_color, font_color, size)` | ... | merged title bar |
| `conditional_format(range_str, rule_type, operator, formula, fill_color)` | ... | highlight rule |
| `set_header_footer(left,center,right)` | str... | print headers |

### Output
| Method | Args | Notes |
|---|---|---|
| `autofit()` | | auto-width columns |
| `save_xlsx(path)` | str | returns abs path |
| `export_pdf(path)` | str | returns abs path |

---

## SlideComposer (pptx) — KWpp.Application

### Setup
| Method | Args | Notes |
|---|---|---|
| `set_slide_size(width_pt=960, height_pt=540)` | float,float | 16:9 default |

### Slides
| Method | Args | Notes |
|---|---|---|
| `add_title_slide(title,subtitle,title_size,sub_size,title_color)` | str... | title layout |
| `add_section_slide(title)` | str | section divider |
| `add_text_slide(title,body,title_size,body_size,title_color,body_color,bullets)` | str... | body=list for bullets |
| `add_bullets_slide(title, items, title_size, body_size)` | str,list... | convenience |
| `add_blank_slide()` | | returns (slide, idx) |

### Content
| Method | Args | Notes |
|---|---|---|
| `set_background_color(slide_index, color)` | int,color | per-slide bg |
| `add_textbox(slide_idx,text,left,top,width,height,size,bold,color,align,fill_color,shape_type)` | ... | align:1L,2C,3R |
| `add_shape(slide_idx,shape_type,left,top,width,height,fill_color,line_color,text,size)` | ... | MSO shape constants |
| `add_image(slide_index, path, left, top, width=None, height=None)` | int,str... | picture |
| `add_table(slide_idx,rows,cols,left,top,width,height,data,header_shade,header_font,font_size)` | ... | slide table |
| `set_notes(slide_index, text)` | int,str | speaker notes |


### Design presets & layout templates (NEW)
| Method | Args | Notes |
|---|---|---|
| `apply_design_preset(preset)` | DesignPreset | set slide master bg color |
| `apply_layout_template(layout)` | LayoutTemplate | best-effort render text elements onto current slide |

### Shape type constants (class attributes)
`MSO_TEXTBOX=17`, `MSO_RECTANGLE=1`, `MSO_ROUNDED_RECTANGLE=5`, `MSO_OVAL=9`,
`MSO_RIGHT_ARROW=13`, `MSO_LEFT_RIGHT_ARROW=37`

### Output
| Method | Args | Notes |
|---|---|---|
| `save_pptx(path)` | str | returns abs path |
| `export_pdf(path)` | str | via SaveAs(fmt=32) |

---

## PdfComposer (edit existing PDFs) — pypdf + pdfplumber

Cross-platform pure-Python, no COM host. All methods static.

| Method | Args | Notes |
|---|---|---|
| `merge(input_paths, output_path, *, overwrite=False)` | list[str], str | concatenate PDFs; explicit overwrite |
| `split(input_path, output_dir, stem=None, *, overwrite=False)` | str,str | atomic one-page PDF set; explicit overwrite |
| `extract_pages(input_path, page_indices, output_path, *, overwrite=False)` | str, list[int], str | 1-based indices; explicit overwrite |
| `rotate(input_path, angle, output_path, *, overwrite=False)` | str,{90,180,270},str | rotate all; explicit overwrite |
| `extract_text(input_path, pages=None)` | str, list[int]\|None | pdfplumber |
| `page_count(input_path)` | str | returns int |
| `add_text_watermark(..., *, overwrite=False)` | ... | diagonal text watermark; explicit overwrite |

---


---

## Export format constants (NEW)

All composers expose these constants for direct `SaveAs` / `ExportAsFixedFormat` calls:

| Constant | Value | Format |
|---|---|---|
| `FMT_DOCX` | 12 | Word Document (.docx) |
| `FMT_DOC` | 0 | Word 97-2003 (.doc) |
| `FMT_PDF_FROM_DOC` | 17 | PDF from Writer |
| `FMT_TXT` | 2 | Plain text |
| `FMT_HTML` | 8 | HTML |
| `FMT_RTF` | 6 | Rich Text Format |
| `FMT_ODT` | 23 | OpenDocument Text |
| `FMT_XLSX` | 51 | Excel Workbook (.xlsx) |
| `FMT_XLS` | -4143 | Excel 97-2003 (.xls) |
| `FMT_CSV` | 62 | CSV (UTF-8) |
| `FMT_PPTX` | 24 | PowerPoint (.pptx) |
| `FMT_PPT` | 2 | PowerPoint 97-2003 (.ppt) |
| `FMT_PPSX` | 36 | PowerPoint Show (.ppsx) |
| `FMT_PDF_FROM_PPT` | 32 | PDF from Presentation |

---

## ProgID fallback order
- Writer: `KWps.Application` → `Wps.Application` → `Word.Application`
- Sheet: `Ket.Application` → `Excel.Application`
- Slide: `KWpp.Application` → `Wpp.Application` → `PowerPoint.Application`

## Notes / gotchas
- Do NOT set `Visible` on WPP — some builds raise on assignment. Engine skips it.
- PPTX→PDF uses `SaveAs(path,32)` not `ExportAsFixedFormat`.
- `Quit()` may raise on some WPP builds; engine swallows it.
- Non-Windows / no COM host → raises `WPSUnavailable`.
