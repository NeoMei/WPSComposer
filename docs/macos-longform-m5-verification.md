# macOS long-form M5 verification

This document records the release-facing macOS truth for the M5 long-form
quality lifecycle. It does not substitute for the separate Windows gate.

## Environment

- Date: 2026-08-23
- Host: Darwin arm64
- WPS: WPS 12.1.26055
- Protocol: 2
- Semantic plan: `longform-1`
- Public version: 0.7.2; 0.8.0 is not released

## Command and evidence

Each run used a fresh pytest base directory and the real WPS JSAPI bridge:

```bash
WPSCOMPOSER_RUN_REAL_WPS=1 ../../.venv/bin/python -m pytest -q \
  tests/longform_m5/test_macos_real_wps_m5.py \
  --basetemp=build/longform-m5/pytest-real-N
```

The post-fix gate passed three consecutive times in `pytest-real-5`,
`pytest-real-6`, and `pytest-real-7`. Each evidence directory contains six
fixture PDFs, one performance PDF, `evidence.json`, SHA-256 values, and
representative screenshots. Evidence JSON contains only relative artifact and
screenshot names.

After the evidence runner was generalized for the pending Windows gate, one
additional full Darwin regression passed in `pytest-real-8`: 63 performance
pages, one generation/export/analysis, zero patch, zero performance issue codes,
and 139.7457 seconds total.

## Stable results

| Fixture | Pages | Expected final issue codes | Native behavior |
|---|---:|---|---|
| academic | 4 | none | native multilevel headings, fields, citations, bibliography |
| toc_dense | 4 | none | one bounded relayout compacts the terminal paragraph/TOC result |
| wide_objects | 5 | none | full-width SVG and landscape table return to portrait |
| degradation | 3 | `FORMULA_MALFORMED`, `FORMULA_FALLBACK_IMAGE_UNAVAILABLE` | readable source and one local marked notice |
| unicode | 2 | `HEADING_ORPHAN` | one relayout followed by one mapped notice patch |
| plain_short | 1 | none | no unnecessary expansion or visible quality anchor |

The synthetic performance document produced 63 pages. In all three consecutive
runs it used one generation, one PDF export, one analysis, zero patches, and
zero issue codes. Total lifecycle time was approximately 139-140 seconds,
comfortably below the 600-second acceptance budget.

Representative cover, TOC, body, landscape-table, degradation, Unicode,
performance-first, performance-middle, and performance-last screenshots were
visually inspected. No unexpected blank page, clipping, duplicated degradation
code, caption/body overlap, or uncentered page number remained.

## Known verified degradation truth

On this WPS build, professional formula `BuildUp` does not create the required
native professional structure for the supported formula families. The executor
checks the postcondition and uses the declared image/source ladder. It does not
claim a linear Type-20 object as native success. The visible issue codes above
are therefore expected fixture results rather than hidden failures.

## Open cross-platform gate

Windows evidence is pending. Before 0.8.0 can be declared complete, Windows
must run the same six fixtures, the 50-100-page performance gate, native heading
numbering and formula/table/reference checks, dedicated-worker ownership and
timeout tests, default/legacy public routing, and PPTX/XLSX regressions. Any
shared-code change found there requires affected macOS tests and real evidence
to be repeated.
