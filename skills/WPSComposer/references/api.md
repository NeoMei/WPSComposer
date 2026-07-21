# WPS Composer API Reference

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

# Batch-edit a copy while preserving the original.
result = edit(
    "report.docx",
    output="report-revised.docx",
    export_pdf="report-revised.pdf",
    patches=[
        {"target": "paragraph:1", "font": {"size": 20, "bold": True}},
        {"target": "table:1/cell:2,1", "fill": {"color": "#FFF2CC"}},
    ],
)
```

Common functions:

| Function | Purpose |
|---|---|
| `open_document(path, kind=None, read_only=False, visible=False)` | Open a supported existing file; returns a context-manageable composer |
| `attach_active(kind=None)` | Attach to the user's active Writer/Sheet/Slide without closing it later; auto-detects when omitted |
| `inspect(path=None, kind=None, selection=False, **options)` | Return a JSON-compatible document or selection snapshot |
| `edit(path=None, kind=None, patches=[...], output=None, export_pdf=None)` | Apply ordered patches and save in place or to a copy |
| `supported_formats()` | Return recognized Writer/Sheet/Slide extensions |
| `snapshot_json(snapshot)` | Serialize a snapshot without losing Chinese text |

Every composer also exposes `inspect_document()`, `inspect_selection()`,
`apply_format_patch(target, **patch)`, `save_current()`, and `save(path)`.
Patches are partial: properties omitted by the agent are not reset.

### Address grammar

| Host | Targets |
|---|---|
| Writer | `selection`, `paragraph:N`, `range:START-END`, `table:N/cell:R,C`, `shape:N`, `section:N` |
| Sheet | `selection`, `sheet:N`, `sheet:N/cell:A1`, `sheet:N/range:A1:C20`, `sheet:N/shape:N`, `sheet:N/chart:N` |
| Slide | `selection`, `presentation`, `slide:N`, `slide:N/shape:N`, `.../paragraph:N`, `.../paragraph:N/run:N`, `.../table/cell:R,C` |

All indices are 1-based. Inspection snapshots return these exact ids so agents
do not need to reconstruct selectors manually.

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
| `merge(input_paths, output_path)` | list[str], str | concatenate PDFs |
| `split(input_path, output_dir, stem=None)` | str,str | one-page PDFs |
| `extract_pages(input_path, page_indices, output_path)` | str, list[int], str | 1-based indices |
| `rotate(input_path, angle, output_path)` | str,{90,180,270},str | rotate all |
| `extract_text(input_path, pages=None)` | str, list[int]\|None | pdfplumber |
| `page_count(input_path)` | str | returns int |
| `add_text_watermark(...)` | ... | diagonal text watermark |

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
