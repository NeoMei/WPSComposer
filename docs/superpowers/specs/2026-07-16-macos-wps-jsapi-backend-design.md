# WPSComposer macOS WPS JSAPI Backend Design

Date: 2026-07-16
Status: Approved design
Phase 0: implementation complete 2026-07-20, see [macOS Phase 0](../../macos-phase0.md)

## Summary

WPSComposer turns Markdown into polished DOCX, PPTX, XLSX, and PDF artifacts by
letting WPS perform layout, pagination, font measurement, table sizing, and
final serialization. The macOS implementation must preserve that property. It
must not replace WPS with a pure OOXML generator.

The macOS backend sends a platform-neutral render plan to a JavaScript add-in
running inside Mac WPS. The add-in uses WPS JSAPI to create, format, save, and
export documents. Windows continues to use the existing COM backend.

DOCX, PPTX, XLSX, and PDF are four independent output targets. Generating one
does not require generating any of the others.

## Goals

- Generate polished DOCX, PPTX, XLSX, and PDF artifacts from Markdown on macOS.
- Use the installed Mac WPS layout engine for final layout and output.
- Reuse the existing parser, structured document model, design presets, and
  orchestration wherever they are platform independent.
- Preserve the Windows COM backend.
- Automatically select an available WPS template for DOCX, PPTX, and XLSX,
  while allowing an explicit user override.
- Give standalone PDF output an Obsidian-inspired Markdown reading layout.
- Detect unsupported Mac capabilities explicitly and never silently fall back
  to low-quality direct file generation.
- Record enough metadata to reproduce and diagnose every result.

## Non-goals

- Pixel-identical Windows and macOS output in every case.
- Automating WPS through mouse, keyboard, or macOS Accessibility scripting.
- Replacing WPS with LibreOffice, Chromium printing, or direct OOXML output.
- Requiring a PDF companion when DOCX, PPTX, or XLSX is requested.
- Exposing temporary native files used internally for PDF generation.

## Architecture

```text
Markdown
  -> md_parser / StructuredDocument
  -> ContentProfile
  -> TemplateSelector
  -> platform-neutral RenderPlan
  -> PlatformBackend
       -> WindowsComBackend -> Windows WPS COM
       -> MacWpsJsBackend   -> local JSON-RPC bridge -> WPS JS add-in -> Mac WPS
  -> requested artifact only: DOCX | PPTX | XLSX | PDF
  -> validation + generation manifest
```

Python remains the source of truth for parsing, semantic structure, content
classification, design intent, and orchestration. Platform backends translate
the same render plan into their native WPS object model.

## Core Models

### ContentProfile

`ContentProfile` summarizes information needed for automatic template choice:

- output type, language, and writing direction;
- title and heading depth;
- paragraph and text density;
- image, table, list, task-list, code-block, and callout counts;
- inferred purpose such as report, proposal, article, lesson, product deck, or
  data analysis;
- presentation density and aspect ratio for PPTX;
- frontmatter and explicit template override.

### RenderPlan

`RenderPlan` describes document operations without COM or JSAPI references:

- document settings, page or slide geometry, and locale;
- resolved template and design tokens;
- sections, slides, sheets, and ordered elements;
- paragraphs, spans, lists, tables, images, charts, code blocks, callouts, and
  separators;
- style roles rather than platform-specific identifiers;
- output target and output path;
- validation expectations and capability requirements.

Operations are batched at section, slide, or sheet level so the Mac bridge does
not issue one RPC request per character or cell.

## Backend Contract

Both platform backends implement the same lifecycle:

1. `probe_capabilities`
2. `open_session`
3. `resolve_resources`
4. `create_document`
5. `apply_render_plan`
6. `save_output`
7. `validate_output`
8. `close_session`

The Windows implementation wraps the existing composers. The macOS
implementation maps the lifecycle to WPS JSAPI and the local bridge.

## Mac WPS Bridge

`MacWpsJsBackend` starts or connects to a loopback-only bridge. The WPS add-in
registers after WPS loads. Communication uses JSON-RPC over HTTP or WebSocket on
`127.0.0.1`.

Rules:

- bind only to loopback;
- generate a random token for every session;
- reject requests without the active token;
- canonicalize and allow-list input and output paths;
- expose typed rendering commands, not arbitrary JavaScript execution;
- use request IDs, timeouts, heartbeats, cancellation, and idempotent batches;
- close temporary documents after success or failure;
- leave unrelated user documents untouched.

The backend may launch `/Applications/wpsoffice.app`. WPS must remain available
while rendering. Headless WPS is not a requirement.

## Automatic Template Selection

DOCX, PPTX, and XLSX select templates in this order:

1. explicit user override;
2. frontmatter template or style request;
3. inferred document purpose;
4. deterministic scoring using content structure and density;
5. a stable WPS base template for the requested format.

`TemplateCatalog` records the templates, themes, fonts, and capabilities
available in the current WPS installation. It is cached by platform, WPS
version, account context, and locale and invalidated when those values change.

Only resources reported as available may be selected. Missing required
resources use an explicitly defined WPS-native substitute or fail according to
the render plan. They never trigger a silent non-WPS renderer.

## Output Contracts

### DOCX

- Use WPS Writer.
- Automatically choose a Writer template unless overridden.
- Save and return only DOCX.
- Do not automatically export PDF.

### PPTX

- Use WPS Presentation.
- Automatically choose a presentation template unless overridden.
- Save and return only PPTX.
- Do not automatically export PDF.

### XLSX

- Use WPS Spreadsheet.
- Automatically choose a spreadsheet template unless overridden.
- Save and return only XLSX.
- Do not automatically export PDF.

### PDF

- Treat PDF as a first-class requested output, not a side effect.
- Use the dedicated `obsidian-reading` preset.
- Map the preset to WPS Writer paragraph, character, table, image, and page
  styles, then let WPS export PDF.
- Keep any required native intermediate in the session temporary directory.
- Return only PDF.

## Obsidian-inspired PDF Theme

Standalone PDF should resemble a clean, printable Obsidian note rather than a
formal Word report:

- light background and restrained page margins;
- comfortable readable line length;
- clean sans-serif body text selected from fonts available to WPS;
- clear Markdown heading hierarchy without automatic numbering;
- compact but readable paragraph and list spacing;
- monospace code blocks with a subtle background, border, and padding;
- distinct inline code, links, blockquotes, horizontal rules, and task lists;
- simple tables with clear headers and restrained borders;
- images kept at their natural aspect ratio with predictable spacing;
- Obsidian-style note, info, tip, warning, and danger callouts;
- no automatic cover, formal table of contents, red-header styling, or office
  report decoration unless explicitly requested.

Frontmatter may override the accent color and a small allow-list of reading
layout settings. It may not inject arbitrary CSS or JavaScript.

## Capability Model

Capabilities are tracked separately for Writer, Presentation, Spreadsheet, and
PDF export. Each feature is classified as:

- `native`: directly supported by WPS JSAPI;
- `mapped`: supported through a documented equivalent sequence;
- `unsupported`: unavailable and reported before rendering when possible.

The first macOS milestone must prove:

- WPS JS add-in loading through the current `wpsjs` Mac path;
- creating and closing Writer, Presentation, and Spreadsheet documents;
- inserting and styling text, tables, and images;
- saving DOCX, PPTX, and XLSX;
- exporting an Obsidian-style Writer document to PDF;
- template and font enumeration or reliable WPS-native resolution.

Failure of one component does not disable every output type. The capability
report identifies which targets remain usable.

## Error Handling

Preflight errors include WPS unavailable, unsupported WPS version, disabled
add-in, unavailable component, unresolved template or font, unwritable output
directory, and missing JSAPI capability.

Runtime errors include bridge disconnect, heartbeat timeout, rejected token,
failed render batch, save/export failure, modal dialog, and missing or empty
output.

Failures return structured codes, a readable explanation, affected output
type, and actionable recovery guidance. Partial files are never reported as
successful artifacts.

## Validation and Testing

Platform-independent tests cover parsing, structured documents, content
profiling, deterministic template scoring, RenderPlan schema validation, and
backend contract behavior.

Independent Mac smoke tests cover:

- DOCX with headings, body text, a table, and an image;
- PPTX with title, content, image, and table slides;
- XLSX with data, formulas, formatting, and a chart;
- PDF with headings, code, callout, table, task list, and image.

The layout corpus covers long text, Chinese and mixed-language content, nested
headings and lists, wide tables, embedded images, code, callouts, presentation
content, spreadsheet formulas, and charts.

WPS reopens native outputs and checks missing content, overflow, overlap,
unresolved fonts, broken images, and component errors. PDF is validated
independently and is not a required companion to other formats.

Cross-platform comparison measures structural and visual quality rather than
strict pixel identity. Substitutions and layout differences are recorded.

## Generation Manifest

Every successful run records:

- requested and produced output type;
- backend, platform, WPS version, and locale;
- selected template, theme, and resolved fonts;
- design preset and allowed frontmatter overrides;
- substitutions and warnings;
- input hash, RenderPlan schema version, and output path;
- parse, render, save, export, and validation timings.

For PDF, the manifest records `obsidian-reading` and does not expose temporary
native documents as outputs.

## Delivery Sequence

### Phase 0: Mac feasibility gate

Build the smallest WPS JS add-in and bridge that proves all four output paths on
the installed Mac WPS version. Produce DOCX, PPTX, XLSX, and standalone PDF
smoke artifacts and record exact JSAPI gaps.

**Tested 2026-07-17: NO-GO on WPS 12.1.26035.** The reversible probe loaded
all three JS add-ins and WPS produced structurally valid PPTX and XLSX files,
but Writer `SaveAs2` and standalone PDF export did not produce files. The same
failure occurred with a one-paragraph Writer document, so Phase 1 remains
gated and the public macOS `generate()` path stays disabled. See
[macOS Phase 0 evidence](../../macos-phase0.md) for the exact environment,
capability matrix, artifacts, errors, and rollback verification. This result
does not change the approved architecture; it records that the current tested
WPS version does not satisfy its feasibility prerequisites.

### Phase 1: Shared render plan and core elements

Introduce the backend contract and RenderPlan. Implement headings, paragraphs,
lists, tables, images, core styles, saving, PDF export, and the manifest.

### Phase 2: Template selection and format depth

Add WPS resource discovery, automatic template scoring, presentation layouts,
spreadsheet charts/formulas, Writer numbering/fields, and the full PDF theme.

### Phase 3: Parity and hardening

Complete the capability matrix, expand the layout corpus, add cancellation and
recovery, tune RPC batching, and compare Windows and Mac output.

## Acceptance Criteria

- The same Markdown entrypoint can request DOCX, PPTX, XLSX, or PDF on macOS.
- The requested artifact is produced by Mac WPS, not a pure format writer.
- No unrequested companion format is produced.
- DOCX, PPTX, and XLSX automatically use an available WPS template unless
  explicitly overridden.
- Standalone PDF uses `obsidian-reading` and returns only PDF.
- Core elements render without overlap, clipping, or missing content.
- Unsupported Mac features are explicit in capabilities and the manifest.
- Windows COM remains available behind the same backend contract.
- The Mac smoke suite passes before broader parity work begins.

## Feasibility Evidence

- WPS documents client development support on macOS and provides JS add-ins.
- `wpsjs` 2.2.x contains explicit Mac adaptation, launches
  `/Applications/wpsoffice.app`, and registers add-ins in the Mac container.
- The installed Mac WPS contains `wpsapi` and `jswpsapi` frameworks.
- The codebase already separates parsing, models, renderers, and composers well
  enough to introduce a backend boundary.

Official references:

- https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-client-dev-introduction
- https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/generate-the-first-wps-addin
- https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/Application/obj
- https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Document/obj
