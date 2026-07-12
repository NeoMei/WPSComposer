---
name: WPSComposer
description: Generate rich-layout DOCX, PPTX, and XLSX documents via WPS Office COM automation. Use when the user wants to create documents that need real layout control (multi-column, floating text boxes with text wrapping, WordArt, shaded/merged tables, charts, auto-updated TOC and fields) that python-docx or openpyxl cannot produce. Triggers on "WPS", "rich layout document", "排版文档", "用 WPS 生成", or when output quality requires a real layout engine rather than static OOXML. Covers all three WPS apps: Writer (docx), Spreadsheets (xlsx), Presentation (pptx).
---

# WPS Composer

Generate DOCX / PPTX / XLSX with full layout control by driving the WPS Office
COM engine. WPS computes the layout (columns, wrapping, field results, chart
rendering), so output matches what you'd see in the WPS GUI — no hand-rolled
OOXML, no guessing about wrapping or page breaks.

## Quick start -- Markdown to document

The ``generate()`` function is the single entry point for all document generation::

    from orchestrator import generate

    # One line: MD -> beautiful document
    generate("report.md", format="docx", preset="academic")
    generate("slides.md", format="pptx", preset="business")
    generate("data.md",  format="xlsx")
    generate("report.md", format="pdf",  preset="consultant")

Under the hood:
1. ``md_parser.py`` parses Markdown into a ``StructuredDocument``
2. A format-specific renderer drives the WPS Composer
3. Design presets control colours, fonts, and spacing

## Existing and currently open documents

For conversational formatting, inspect first and then patch the returned
element ids. All patches are partial, so omitted properties stay unchanged.

```python
from wps_engine import inspect, edit, attach_active

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
| `writer.py` | WriterComposer | WPS Writer → docx |
| `sheet.py` | SheetComposer | WPS Spreadsheets → xlsx |
| `slide.py` | SlideComposer | WPS Presentation → pptx |
| `pdf.py` | PdfComposer | PDF edit (pypdf/pdfplumber, cross-platform) |

All existing `from wps_engine import ...` imports continue to work.
New code can import directly from the specific modules::

    from writer import WriterComposer
    from sheet import SheetComposer
    from slide import SlideComposer
    from pdf import PdfComposer

### Usage
`python
from wps_engine import WriterComposer, SheetComposer, SlideComposer

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

`scripts/design_presets.py` — 4 colour+font presets for quick styling:
```python
from design_presets import get_preset, list_presets
preset = get_preset("academic")       # or: consultant, business, tech
preset.get_color("primary")           # "#1A3C8B"
preset.get_font("title")              # ("Arial", 40, "#1A3C8B")
```

`scripts/layout_templates.py` — 14 PPT slide layouts + 4 talk-type sequences:
```python
from layout_templates import get_layout, get_talk_preset
layout = get_layout("cover")          # 14 layouts total
talk = get_talk_preset("defense")     # conference/business/defense/school
```

`scripts/quality_checks.py` — slide-level validation against design rules:
```python
from quality_checks import validate_slide, review_deck, REVIEW_DIMENSIONS
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
from wps_engine import PdfComposer
PdfComposer.merge(["a.pdf","b.pdf"], "out.pdf")
PdfComposer.split("in.pdf", "out_dir")            # one-page PDFs
PdfComposer.extract_pages("in.pdf", [1,3], "out.pdf")  # keep pages 1 & 3
PdfComposer.rotate("in.pdf", 90, "out.pdf")
PdfComposer.add_text_watermark("in.pdf", "CONFIDENTIAL", "out.pdf")
print(PdfComposer.extract_text("in.pdf"))           # via pdfplumber
```

## Output discipline

Always export a PDF alongside the native format so layout can be visually
verified (`render_docx.py`-style gate). Deliver only the requested format to the
user unless they ask for intermediates.
