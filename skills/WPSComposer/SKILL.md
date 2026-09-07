---
name: WPSComposer
description: 'Generate and edit rich-layout DOCX, PPTX, XLSX, and PDF documents by driving the real WPS Office layout engine — via COM on Windows (full generation + conversational editing) or the WPS JSAPI bridge on macOS (generation and Office-to-PDF conversion). Use when the user wants to create documents that need real layout control (multi-column, floating text boxes with text wrapping, WordArt, shaded/merged tables, charts, auto-updated TOC and fields) that python-docx or openpyxl cannot produce. Triggers on "WPS", "rich layout document", "排版文档", "用 WPS 生成", or when output quality requires a real layout engine rather than static OOXML. Covers all three WPS apps: Writer (docx), Spreadsheets (xlsx), Presentation (pptx).'
---

# WPS Composer

Generate DOCX / PPTX / XLSX with full layout control by driving the WPS Office
layout engine — COM on Windows, the JSAPI loopback bridge on macOS. WPS computes
the layout (columns, wrapping, field results, chart rendering), so output
matches what you'd see in the WPS GUI — no hand-rolled OOXML, no guessing about
wrapping or page breaks. Conversational inspect/edit works on both platforms:
Windows uses COM directly; macOS uses the JSAPI loopback bridge to read and
edit PPT/DOCX/XLSX through the real WPS engine (no PDF extraction fallback).

## Quick start -- Markdown to document

The ``generate()`` function is the single entry point for all document
generation. For an interactive script, explicitly request presentation of the
final artifact after WPS cleanup:

```python
from skills.WPSComposer import generate


def main() -> None:
    generate(
        "report.md",
        format="docx",
        preset="academic",
        open_result=True,
    )


if __name__ == "__main__":
    main()
```

Library and unattended calls leave `open_result` at its default `False`, so
they publish without opening a desktop application. Each call returns only the
one requested format. Opening is an asynchronous, best-effort request to the
platform default application; a launcher warning does not invalidate an
artifact that was already published.

Under the hood:
1. ``md_parser.py`` parses Markdown into a ``StructuredDocument``
2. A format-specific renderer drives the WPS Composer
3. Design presets control colours, fonts, and spacing

## Existing and currently open documents

For conversational formatting, inspect first and then patch the returned
element ids. All patches are partial, so omitted properties stay unchanged.
```python
from skills.WPSComposer import inspect, edit, attach_active

# Existing file -> addressable structure + effective formatting
tree = inspect("report.docx")

# Current visible WPS document -> selection-aware editing
w = attach_active("writer")
print(w.inspect_selection())
w.apply_format_patch(
    "paragraph:3",
    font={"size": 11, "color": "#333333"},
    paragraph={"alignment": 3, "line_spacing": 20},
)
w.save_current()

# Ordered patches, save a copy, and export PDF for visual verification
edit(
    "report.docx",
    output="report-revised.docx",
    export_pdf="report-revised.pdf",
    patches=[
        {"target": "paragraph:1", "font": {"size": 20, "bold": True}},
        {"target": "table:1/cell:2,1", "fill": {"color": "#FFF2CC"}},
    ],
)
```

Addressable elements include Writer paragraphs/ranges/tables/cells/shapes/
sections, Sheet cells/ranges/shapes/charts/page setup, and Slide shapes/text
paragraphs/runs/table cells/backgrounds. See `references/api.md` for the target
grammar and patch fields.

`edit()` is atomic by default: a failed patch blocks the save and returns
`{"ok": False, "errors": [...]}` instead of raising. `validate_target(target,
kind)` checks a target against the grammar and suggests the closest valid form
on a miss, and `patch_grammar(kind)` returns the grammar as data for agent
discovery. A distinct existing output is rejected unless `overwrite=True`, and
the output must remain in the source document family. File-backed results are
validated and atomically published. In attach-active atomic mode, only one
`set` operation containing at most one leaf property is accepted; composite
patches and structural operations are rejected before mutation because the
live host has no reliable rollback boundary. macOS editing currently supports
verified `.pptx` input/output only.

For **structural editing** (insert/remove/move/clone), pass `ops=` instead of
(or alongside) `patches`. Each op carries a verb: `set` (formatting, == a
patch), `insert` (paragraph/table/shape/slide/row/…), `remove`, `move`,
`clone`. Patches run first, then ops, in one atomic transaction. See
`references/api.md` → "Structural operations" for the full verb/type matrix.

```python
edit("deck.pptx", output="deck2.pptx", ops=[
    {"op": "insert", "type": "slide", "props": {"layout": 12}},
    {"op": "clone", "target": "slide:1", "to": "end"},
    {"op": "remove", "target": "slide:2"},
])
```

The COM-coupled parts of this work (real `inspect`/`edit`/structural ops
against a live WPS host, native stable IDs) are written but need Windows
verification — tracked in `docs/windows-verification.md`.

## When to use

- Documents needing **multi-column** layout, **floating shapes with text wrap**,
  **WordArt**, **shaded/merged tables**, or **auto-populated TOC / page numbers**.
- Spreadsheets with **formulas, styled cells, autofit columns**.
- Slides with **title/text layouts** and 16:9 sizing.
- When python-docx / openpyxl / python-pptx output "looks wrong" because they
  cannot compute layout.

## Requirements

- Windows + WPS Office installed (Writer/Spreadsheets/Presentation), **or** MS
  Office (Word/Excel/PowerPoint). ProgIDs fall back automatically.
- `pywin32` available in the runtime.

## Core engine

The engine is now split into focused modules with `wps_engine.py` as a
backward-compatible re-export facade:

| Module | Class | Purpose |
|---|---|---|
| `_dispatch.py` | — | COM dispatch, ProgID chains, format constants, WPS path search |
| `_colors.py` | — | Unified colour model: hex→BGR, semantic name resolution |
| `_base.py` | BaseComposer | Shared COM lifecycle + save/export |
| `writer.py` | WriterComposer | WPS Writer → docx (+ paraId helpers) |
| `sheet.py` | SheetComposer | WPS Spreadsheets → xlsx |
| `slide.py` | SlideComposer | WPS Presentation → pptx |
| `pdf.py` | PdfComposer | PDF edit (pypdf/pdfplumber, cross-platform) |
| `document_api.py` | — | Conversational API: `inspect`/`edit`/`apply_ops`/`validate_op`/`validate_target`/`snapshot_to_patches` + atomic orchestration |
| `orchestrator.py` | — | `generate()` — markdown → document |

See `AGENTS.md` for the full module layout (including the macOS JSAPI backend
under `macos_probe/`).

Use the stable public package for new code::

    from skills.WPSComposer import (
        WriterComposer, SheetComposer, SlideComposer, PdfComposer,
    )

### Usage
```python
from skills.WPSComposer import WriterComposer, SheetComposer, SlideComposer

with WriterComposer() as w:      # KWps.Application
    w.set_columns(2)
    w.add_heading("Title")
    w.add_table(3, 3, [[...]])
    w.add_floating_textbox("note", 80, 280, 150, 50, wrap=0)
    w.save_docx("out.docx")
    w.export_pdf("out.pdf")

with SheetComposer() as s:       # Ket.Application
    s.write_table(1, 1, [[...]])
    s.set_formula(5, 2, "=SUM(B2:B4)")
    s.save_xlsx("out.xlsx")

with SlideComposer() as p:       # KWpp.Application
    p.set_slide_size(960, 540)   # 16:9
    p.add_title_slide("Title", "subtitle")
    p.save_pptx("out.pptx")
```

Each composer is a context manager: `__enter__` dispatches the COM app and
creates a fresh document; `__exit__` always closes + quits, even on error.


## Companion modules

Three new modules provide design presets, layout templates, and quality checks
(integrated from the harness-anything project):

`scripts/design_presets.py` — 5 colour+font presets for quick styling:
```python
from skills.WPSComposer.scripts.design_presets import get_preset, list_presets
preset = get_preset("academic")  # consultant, business, tech, or proposal
preset.get_color("primary")           # "#1A3C8B"
preset.get_font("title")              # ("Arial", 40, "#1A3C8B")
```

`scripts/layout_templates.py` — 14 PPT slide layouts + 4 talk-type sequences:
```python
from skills.WPSComposer.scripts.layout_templates import get_layout, get_talk_preset
layout = get_layout("cover")          # 14 layouts total
talk = get_talk_preset("defense")     # conference/business/defense/school
```

`scripts/quality_checks.py` — slide-level validation against design rules:
```python
from skills.WPSComposer.scripts.quality_checks import validate_slide, review_deck, REVIEW_DIMENSIONS
result = validate_slide(elements, preset.rules)   # {"pass": bool, "score": 0-100, ...}
```

SlideComposer gained two convenience methods:
```python
p.apply_design_preset(preset)         # set slide master background
p.apply_layout_template(layout)       # best-effort render text elements
```


## API reference

See `references/api.md` for the full method list per composer.

## Layout capabilities (COM-verified, end-to-end tested)

| Capability | Writer | Sheet | Slide |
|---|---|---|---|
| Multi-column / page setup | yes | page setup | yes |
| Tables (merge/shade/borders) | yes | yes | yes |
| Floating shapes + wrap | yes | - | yes |
| WordArt | yes | - | - |
| Charts (bar/line/pie) | - | yes | - |
| Formulas + number format | - | yes | - |
| Freeze panes / conditional format | - | yes | - |
| Images | yes | - | yes |
| Shapes (rect/arrow/oval) | - | - | yes |
| Bullet / numbered lists | yes | - | - |
| Headers / footers | yes | yes (print) | - |
| Speaker notes | - | - | yes |
| Field refresh (TOC/PAGE) | yes | - | - |
| Background color | - | cell fill | yes (per-slide) |
| Inspect existing/current document | yes | yes | yes |
| Element-level partial formatting patch | yes | yes | yes |
| Hex color support | `#RRGGBB` | `#RRGGBB` | `#RRGGBB` |

## PDF editing (existing PDFs)

PDF *generation* comes from the three WPS composers above (`export_pdf`).
PDF *editing* (merge / split / extract / rotate / watermark) uses `PdfComposer`,
backed by pypdf + pdfplumber — NOT the WPS PDF COM host (KPDF.Application
Dispatch blocks headless, so it's unsuitable for automation).

```python
from skills.WPSComposer import PdfComposer
PdfComposer.merge(["a.pdf","b.pdf"], "out.pdf")  # overwrite=True to replace
PdfComposer.split("in.pdf", "out_dir")            # one-page PDFs
PdfComposer.extract_pages("in.pdf", [1,3], "out.pdf")  # keep pages 1 & 3
PdfComposer.rotate("in.pdf", 90, "out.pdf")
PdfComposer.add_text_watermark("in.pdf", "CONFIDENTIAL", "out.pdf")
print(PdfComposer.extract_text("in.pdf"))           # via pdfplumber
```

PDF outputs are staged and validated before atomic publication. Existing
outputs are refused by default; pass `overwrite=True` explicitly to replace.

## Output discipline

Generate and deliver only the requested artifact format. During development,
create PDF evidence separately when a native WPS layout change needs visual
verification; do not make that PDF an automatic public companion output.

## Long-form generation and PDF quality lifecycle (M5)

Public `generate()` now routes DOCX/PDF through the M5 long-form engine by
default. PPTX/XLSX retain their existing route. An explicit deprecated
frontmatter escape hatch, `layout_engine: legacy`, selects the old Writer route;
engine loss, protocol mismatch, save/export failure, or validation failure must
never fall back to it automatically. The lifecycle uses the public 600-second
timeout as one absolute deadline and publishes only the requested artifact.

The lower-level API remains available to WPSComposer/SuperWriter agents,
plugin maintainers, and applications that need to inspect a deterministic
offline plan before choosing a native executor:

```python
from skills.WPSComposer.scripts.longform import (
    build_longform_generation,
    execute_longform_plan,
)

build = build_longform_generation(markdown, base_dir="assets")
# Supply a dedicated Windows COM or macOS JSAPI executor only when native
# generation is explicitly required.
outcome = execute_longform_plan(build, executor, deadline=deadline)
```

Long-form Markdown preserves the M3 figure, table, numbering, reference, and
resource contracts and adds these closed M4 forms:

- Figures: `#id`, `caption`, `width="auto|column|full|Npt"`,
  `orientation="portrait|landscape"`, `kind`, `layout="stack|columns"`, and
  `columns=2`. Figure captions are native fields below the image container.
- Tables: `#id`, `caption`, `style="three-line|grid"`, explicit
  `orientation`, `merges="A2:A3;B2:C2"`, and `repeat_header`. Table captions
  are native fields above the first row.
- Formula blocks use `:::equation {#eq:id}` (legacy `:::formula` is accepted).
  Optional `fallback_image="relative/path.png"` names one local, validated
  fallback image. Formula text is a bounded restricted-LaTeX subset: Unicode
  variables, scripts, fractions, roots, sums/products/integrals, scalable
  delimiters, common Greek/operators/relations, matrices, cases, and bounded
  nesting. Custom/unknown commands, packages, file/URL/shell access, malformed
  groups/environments, and external renderers are rejected before execution.
- `{{cite:id}}` emits numeric citations. Numbers follow first visible semantic
  occurrence across paragraphs, lists, block quotes, page-break paragraphs,
  and table cells; repeats reuse the number. `:::bibliography` (legacy
  `:::references`) declares one `[id] text` entry per line. Cited entries come
  first, followed by uncited declarations unless front matter sets
  `bibliography_include_uncited: false`.
- `{{ref:target-id}}` remains the native hyperlinking object reference. Formula
  references reuse the M3 equation number/bookmark shell. References and
  citations in abstracts and lists retain their paragraph and list geometry.
- Front matter controls `caption_numbering: auto|global|chapter` plus
  `figure_index` and `table_index`. `auto` is resolved per object: content
  before the first numbered H1 is global, later content is chapter-numbered,
  and an unnumbered H1 does not reset a sequence.

Controlled native sequences are `WPSC_FIG`, `WPSC_TAB`, and `WPSC_EQ`.
Bookmarks wrap only the visible number, and references use `REF ... \\h`.
Figure/table indexes are populated native fields in the front matter. Field
finalization is bounded: numbering, references, indexes, and page fields must
produce two adjacent equal snapshots within three rounds; otherwise one frozen
fourth snapshot records `FIELD_REFRESH_UNSTABLE`.

Accepted media are decoded PNG, JPEG, TIFF, BMP, GIF, and restricted static
SVG. WebP and network media are rejected. EXIF-oriented images, the first GIF
frame, and the first TIFF page are normalized privately to lossless PNG.
Limits are 50 MiB per resource, 80,000,000 pixels, and 32,768 pixels on either
side. The normalized bytes and source paths never enter the plan or diagnostic
JSON. Executor staging cleanup is attempted on every exit path; a cleanup
failure is fatal and never silently publishes a result.

The academic preset defaults to a native three-line table: 1.5 pt top/bottom,
0.75 pt below the header, no vertical/interior body borders, zero cell indent,
direct cell alignment, repeated headers, and row splitting disabled. Merge
declarations are validated all-or-nothing. An invalid declaration preserves
the complete unmerged grid and records one notice after the caption; an
over-page vertical group degrades to an unmerged splittable grid.

M4 formula execution first attempts editable Office Math/WPS math content while
keeping the number shell separate. `EQUATION_INSERT_FAILED` follows exactly one
closed ladder: the declared validated image when available, otherwise readable
source, plus a visible notice at the formula node. A missing or invalid fallback
image produces `FORMULA_FALLBACK_IMAGE_UNAVAILABLE` but never suppresses a valid
native attempt. On the verified macOS WPS 12.1.26055 build, professional
`BuildUp` is a structural no-op for the supported families; the executor detects
that postcondition and honestly uses the marked image/source fallback. It does
not report a linear Type-20 object as native success. Windows uses the same
descriptor and recovery contract; real Windows M4 execution remains the
M5/final cross-platform gate.

Marked recovery is visible and local. Missing citation/reference runs stay
inline in their paragraph; formula/object failures use a nearby block notice;
document issues deduplicate at the reserved `生成质量提示` anchor. Notices contain
only a stable code, controlled label, readable reason, and actual fallback—never
paths, payloads, hashes, bookmark maps, field values, or exception
representations. A document with no optional formula, citation, bibliography,
image, or degradation has an empty reserved anchor but no visible notice or
extra spacing.

Only allowlisted object-local failures such as `EQUATION_INSERT_FAILED`,
`CROSS_REFERENCE_FAILED`, and `BIBLIOGRAPHY_INSERT_FAILED` can recover. Engine
loss (`ENGINE_LOST`), protocol/capability mismatch, unknown native exceptions,
rollback failure, required field/index/repagination failure, staging/hash or
cleanup failure, save/export/validation/publication failure, and failure of the
terminal source/notice fallback are fatal. A recoverable image-rung failure may
still roll back into the declared source notice; an unavailable engine stops
immediately.

M5 exports the staged DOCX to PDF for geometry analysis before publication.
The normalized top-left point-coordinate page model checks blank pages,
boundaries, headings, captions, images, tables, TOC density, fields, and final
page utilization. Only high-confidence findings in the closed repair matrix may
trigger one full relayout; remaining material findings receive at most one
notice-only patch. The hard caps are two generations, one notice patch, and
three PDF exports. `Pillow>=10`, `pypdf>=4`, and `pdfplumber>=0.11` are core
dependencies and are checked before WPS starts.

Recoverable object failures remain visible where the object belongs. Fatal
conditions including `ENGINE_LOST`, protocol/capability mismatch, cleanup,
save/export, validation, and publication failure abort without a public partial
artifact. On macOS WPS 12.1.26055, native formula BuildUp is still honestly
reported through the marked image/source ladder. cross-platform acceptance: COMPLETED
after three consecutive real M5 gates on both macOS and Windows, including the
post-acceptance Windows rerun. 0.8.0 released on 2026-08-24 with the M5
long-form route enabled by default for DOCX/PDF. See `docs/longform-markdown.md` and
`docs/macos-longform-m5-verification.md`.

## Native heading numbering (docx)

Generated DOCX files carry **native Word/WPS multi-level heading numbering**
instead of plain-text number prefixes. Both the formal Chinese hierarchy
(`第一章` → `第一节` → `一、` → `（一）`) and the bid-document hybrid hierarchy
(`第一章` → `1.1` → `1.1.1` → `关键工法01`) are preserved visually and linked
to one native list. Reordering, inserting, or deleting headings in WPS
renumbers the document automatically. The document title and intentionally
unnumbered headings stay unnumbered.

Implemented as a post-generation pass (`scripts/numbering_native.py`,
`apply_native_numbering()`) hooked into `orchestrator.generate()` for the
`docx` format. It strips plain-text prefixes from Heading 1-4 paragraphs,
merges a WPSComposer-owned definition into `word/numbering.xml` without
discarding existing list definitions, and binds the heading hierarchy with
native `numPr` links. It accepts both `Heading 1` and `heading 1` built-in
style names. Idempotent and never blocks a successful generation.

Writer tables have an explicit cell-format contract on both Windows and
macOS: first-line, left, and right paragraph indents are zero; before/after
spacing is zero; cells are vertically centered; and requested per-column
left/center/right alignment is retained. Do not rely only on inherited named
styles for these properties.

Wide PNG diagrams whose actual width/height ratio is at least 1.25 are placed
in their own landscape Writer section; the following content resumes in a
portrait section. The rule is driven by the image header, not by a filename or
project-specific figure number.

Document title handling: the first H1 (the Markdown document title) is
rendered on the cover page only — it does **not** appear in the body, is
not collected by the TOC, and does not take part in numbering. The TOC
title ("目  录") uses a non-outline style so the TOC does not list itself.

Note: the TOC field cache keeps the pre-refresh entries; update fields
(Ctrl+A → F9, or right-click TOC → Update Field) after opening to rebuild it.

## Installation

The repository-level `install.py` installs this bundle through a Codex personal marketplace.
Restart ChatGPT/Codex Desktop after installation, then install or
enable `wps-composer` from the Plugins Directory before invoking `$WPSComposer`.
