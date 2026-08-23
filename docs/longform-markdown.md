# Long-form Markdown and M5 quality lifecycle

WPSComposer uses the protocol-v2 long-form engine by default for public
DOCX/PDF generation. PPTX and XLSX continue to use their existing renderers.
This guide describes the accepted Markdown, native-object behavior, quality
gate, and failure boundaries that are currently implemented.

## Public routing and runtime budget

Normal callers use the unchanged public entry point:

```python
from skills.WPSComposer import generate

path = generate(
    "report.md",
    format="pdf",
    preset="academic",
    timeout=600,
    overwrite=False,
)
```

DOCX/PDF default to `layout_engine: longform`. The deprecated old Writer path
is available only when the source explicitly declares:

```yaml
---
layout_engine: legacy
---
```

There is no automatic fallback to legacy. One absolute timeout covers native
generation, internal PDF export, analysis, optional relayout, optional notice
patch, validation, publication, and cleanup. The default is 600 seconds. The
closed maximum is two full native generations, one notice-only patch, and three
PDF exports. Only the requested public artifact is returned.

Core quality dependencies are `Pillow>=10`, `pypdf>=4`, and
`pdfplumber>=0.11`. They are checked before WPS starts. Windows generation also
requires `pywin32`; macOS requires WPS Office, Node.js 20+, and the installed
locked JSAPI runtime.

## Frontmatter

Common long-form controls include:

```yaml
---
title: 项目技术报告
author: WPSComposer
design: academic
toc: true
title_page: true
heading_numbering: decimal
caption_numbering: auto
figure_index: true
table_index: true
bibliography_include_uncited: false
layout_engine: longform
---
```

`heading_numbering` accepts the controlled schemes supported by the selected
policy. `caption_numbering` accepts `auto`, `global`, or `chapter`. Under
`auto`, objects before the first numbered H1 are global; later objects use the
chapter sequence, and an unnumbered H1 does not reset it.

## Figures and tables

```markdown
:::figure {#fig:architecture caption="总体架构" width="full" kind="diagram"}
![总体架构](media/architecture.svg)
:::

:::table {#tab:metrics caption="关键指标" style="three-line" repeat_header=true}
| 指标 | 2025 | 2026 |
|---|---:|---:|
| 时延 | 80 ms | 60 ms |
:::
```

Figure attributes include `#id`, `caption`, `width="auto|column|full|Npt"`,
`orientation="portrait|landscape"`, `kind`, `layout="stack|columns"`, and
`columns=2`. Accepted local media are decoded PNG, JPEG, TIFF, BMP, GIF, and
restricted static SVG. Remote resources, WebP, active SVG, corrupt or oversized
media are rejected before native execution.

Table attributes include `#id`, `caption`, `style="three-line|grid"`,
`orientation`, `merges="A2:A3;B2:C2"`, and `repeat_header`. Merge declarations
are all-or-nothing. Invalid merges preserve the complete grid and emit a local
controlled notice. Landscape tables use an explicit landscape section and the
following section returns to portrait.

## Formulas, references, citations, and bibliography

```markdown
:::equation {#eq:energy fallback_image="media/energy.png"}
E = mc^2
:::

参见 {{ref:fig:architecture}}、公式 {{ref:eq:energy}} 和 {{cite:engine}}。

:::bibliography
[engine] WPSComposer. Long-form quality lifecycle.
:::
```

The formula grammar is a bounded restricted-LaTeX subset. It supports Unicode
variables, scripts, fractions, roots, sums/products/integrals, scalable
delimiters, common Greek/operators/relations, matrices, and cases. It rejects
unknown commands, packages, file/URL/shell access, malformed groups, more than
10,000 Unicode code points, or more than 64 brace levels.

`{{ref:id}}` creates a native reference to a controlled bookmark.
`{{cite:id}}` creates deterministic numeric citations in first-visible order.
`:::bibliography` declares `[id] text` entries; the legacy name
`:::references` remains accepted. Figure, table, and formula visible numbers
use controlled native sequences and can be refreshed after edits.

## Pagination coordinates and deterministic quality checks

Native executors return a pagination map keyed by stable node IDs. PDF page
geometry is normalized into a top-left point coordinate system:

```text
[x0, y0, x1, y1] where 0 <= x0 <= x1 <= page width
                         0 <= y0 <= y1 <= page height
```

Page rotation and CropBox/MediaBox offsets are normalized before analysis.
WPS vendor sentinel coordinates are omitted rather than treated as real
bounds. Node identity comes only from the native pagination map; extracted PDF
text is never guessed back into a Markdown node.

The closed checks cover unexpected blank pages, body overflow, headers/footers
and page-number sequence, orphan headings, caption separation, image fit/DPI,
table fit/header/split behavior, TOC density and spacing, field failures,
bibliography spacing, and final-page utilization. Warning/info or medium/low
confidence findings never mutate content. Only allowlisted high-confidence
findings may trigger the one relayout pass.

## Visible degradation and fatal failures

Missing optional content is not an error: if the source contains no image or
formula, normal layout continues without a notice. When an optional object was
requested but its allowlisted native insertion fails, the engine preserves a
readable result and marks the corresponding position. Examples include
`FORMULA_FALLBACK_IMAGE_UNAVAILABLE`, `FORMULA_MALFORMED`, a missing reference,
or a table merge degradation. A notice never silently replaces unrelated body
content.

Fatal boundaries stop immediately and publish no partial result. They include
`ENGINE_LOST`, engine acquisition failure, protocol or capability mismatch,
unknown native exceptions, rollback failure, required field/index/repagination
failure, staging/hash/cleanup failure, save or export failure, validation or
publication failure, and failure of the terminal readable fallback itself.

## Resource privacy

Private source paths, normalized payload bytes/base64, source and payload hashes,
private staging locators, bookmark maps, field values/hashes, and exception
representations do not enter public plans, evidence reports, or document
notices. Each executor copies resources to a private per-generation staging
area and attempts cleanup on every exit. Cleanup failure is fatal.

## Platform and release status

The macOS path is verified on WPS 12.1.26055 with real generation, PDF export,
reopen/refresh, relayout, notice patching, and a 63-page performance fixture.
On that build, professional formula BuildUp is a structural no-op for the
supported families, so the postcondition correctly uses the marked image/source
ladder instead of claiming native success.

Windows shares the semantic plan and recovery contracts, but its final M5 COM
visual/performance evidence is still required. Version 0.8.0 remains unreleased
until that Windows gate passes and any shared fixes are rechecked on macOS.
