# macOS long-form M5 verification

This document records the release-facing macOS truth for the M5 long-form
quality lifecycle. It does not substitute for the separate Windows gate.

## Environment

- Date: 2026-08-24
- Host: Darwin arm64
- WPS: WPS 12.1.26055
- Protocol: 2
- Semantic plan: `longform-1`
- Public version: 0.8.0; released 2026-08-24 after cross-platform acceptance

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

After the evidence runner was generalized for the Windows gate, one
additional full Darwin regression passed in `pytest-real-8`: 63 performance
pages, one generation/export/analysis, zero patch, zero performance issue codes,
and 139.7457 seconds total.

The required post-Windows re-verification then ran on the Windows-fix HEAD plus
the macOS acceptance fixes. The final three consecutive runs are
`final-postfix-1`, `final-postfix-2`, and `final-postfix-3`. Each contains six
fixture PDFs, one 63-page performance PDF, 22 screenshots, and a privacy-safe
`evidence.json`; all recorded SHA-256 values were independently recomputed.
The fresh platform-independent suite passed **2542 tests with 12 explicit
platform gates skipped**.

## Stable results

| Fixture | Pages | Expected final issue codes | Native behavior |
|---|---:|---|---|
| academic | 4 | none | native multilevel headings, fields, citations, bibliography |
| toc_dense | 4 | none | one bounded relayout compacts the terminal paragraph/TOC result |
| wide_objects | 5 | none | full-width SVG and landscape table return to portrait |
| degradation | 3 | `FORMULA_MALFORMED`, `FORMULA_FALLBACK_IMAGE_UNAVAILABLE` | readable source and one local marked notice |
| unicode | 2 | none | Unicode text plus the two-line code block survive; one generation, no patch |
| plain_short | 1 | none | no unnecessary expansion or visible quality anchor |

The synthetic performance document produced 63 pages. In the final three
post-Windows runs it used one generation, one PDF export, one analysis, zero
patches, and zero issue codes. Total lifecycle times were 158.5137, 144.5114,
and 170.3104 seconds, comfortably below the 600-second acceptance budget.

Representative cover, TOC, body, landscape-table, degradation, Unicode,
performance-first, performance-middle, and performance-last screenshots were
visually inspected. No unexpected blank page, clipping, duplicated degradation
code, caption/body overlap, or uncentered page number remained.

The review found that fenced code blocks had been counted semantically but
silently omitted by the long-form plan builder. The plan now emits compact
monospace paragraphs and the Windows executor forwards the closed paragraph
formatting arguments. The Unicode evidence now visibly contains
`def quality_gate(document):` and `return "visible result"`; restoring those
lines also removed the former false `HEADING_ORPHAN` outcome.

## Known verified degradation truth

On this WPS build, professional formula `BuildUp` does not create the required
native professional structure for the supported formula families. The executor
checks the postcondition and uses the declared image/source ladder. It does not
claim a linear Type-20 object as native success. The visible issue codes above
are therefore expected fixture results rather than hidden failures.

WPS 12.1.26055 on macOS preserves `U+2705` (`✅`) in the PDF text layer but
does not paint its glyph. Probes with Arial Unicode MS, Apple Symbols, Zapf
Dingbats, CJK Symbols Fallback, and Apple Color Emoji produced the same result.
The source character is retained rather than silently substituted with another
code point; this is a documented WPS PDF-rendering limitation.

## Windows gate status: COMPLETED (2026-08-24)

The Windows M5 gate passed on branch `codex/longform-m3` commit `5148b1a`
(WPS Office 12.1.0.26899 zh-CN, Windows 10.0.26200, pywin32 312,
Python 3.14.3): three consecutive real-WPS gate runs green plus a
pooled-code rerun, platform-independent suite 2512 passed / 38 skipped,
public-route DOCX/PDF/PPTX/XLSX generation verified, and a same-process
multi-generation RPC death root-caused and fixed via a suite-app pool.
Full details, evidence numbers, and the bug list live in
`docs/windows-verification.md`.

## macOS re-verification status: COMPLETED (2026-08-24)

The Windows run changed **shared Python**, so the older `pytest-real-*` evidence
was not reused as post-fix evidence. The full suite, focused regressions, PDF
text/visual checks, and three fresh M5 native runs described above completed on
the new code.

Acceptance found and fixed five additional defects: the M4 environment gate
failed instead of skipping in ordinary suites; the documentation test required
the obsolete Windows-pending wording; a reused Windows executor leaked prior
front matter; failed pooled COM construction left the apartment initialized;
and fenced code blocks were omitted from long-form plans.

## Windows post-acceptance rerun status: COMPLETED (2026-08-24)

Because the last two document-generation fixes touched Windows execution and
the shared plan, Windows performed a final clean-checkout post-acceptance rerun
starting from `46a1221`. Commit `45a178f` records 2517 passing tests plus three
consecutive real M5 gates; `7ffbd5b` corrects the WPS host build to
12.1.0.28043. The Unicode code block remained visible and monospace, and all
three 63-page performance runs stayed below 600 seconds without patches or
issues. No shared generation/runtime code changed after the accepted macOS
evidence. 0.8.0 released on 2026-08-24 after the final release gate.

### What Windows changed and why macOS should be unaffected

| Change | Files | macOS impact |
|---|---|---|
| Resource ids derive from base-relative posix paths | `longform/resources.py` | Machine-independent by design; m3 snapshots regenerated and must now match byte-for-byte on macOS too |
| Platform-neutral absolute-path checks in evidence validators | `longform_m3/m4/m5_evidence.py` | Shared; test expectations updated with the code |
| Chapter caption STYLEREF code built at COM time from localized style name | `writer.py` (Windows COM path only) | None — the macOS add-in already resolves the localized name itself (`writer-longform-v2.js` `addNativeNumberShell`) |
| Heading numbering links built-in styles to the list template | `writer.py` (Windows COM path only) | None — mirrors what the macOS add-in already does |
| Cover rendering from stashed front matter | `longform/windows_executor.py` (Windows only) | None |
| Suite-app pool for COM instances | `_dispatch.py`, `_base.py`, `slide.py`, `sheet.py` | Windows-only COM code; the macOS JSAPI bridge is untouched |
| Bullet `writer.add_list` emits required `ordered` arg | `longform/plan.py` | Shared; platform-independent tests updated |
| Various test doubles/expectations aligned | `tests/longform*/test_windows_executor*.py` | Fake-composer tests run cross-platform; they pass on Windows and must pass on macOS unchanged |

### Steps

```bash
git clone https://github.com/NeoMei/WPSComposer.git && cd WPSComposer
git checkout codex/longform-m3
git log -1 --oneline   # must be 5148b1a or later

python3 -m venv .venv
.venv/bin/python -m pip install -e '.[dev,pdf]'   # pdf extra: reportlab is required by pdf-gate tests
(cd macos/wps-jsapi-probe && npm ci)               # one-time; probe fixtures need the wpsjs template

# 1. Full platform-independent suite — expect 0 failures.
#    Skips should be limited to pypdf (if unavailable) and the env-gated
#    real-WPS gates. (Windows reference: 2512 passed / 38 skipped.)
.venv/bin/python -m pytest -q

# 2. Minimum required: one complete macOS M5 real gate. Record git HEAD,
#    use a fresh evidence directory per run, three consecutive runs.
WPSCOMPOSER_RUN_REAL_WPS=1 .venv/bin/python -m pytest -q \
  tests/longform_m5/test_macos_real_wps_m5.py \
  --basetemp=build/longform-m5/postfix-1

# 3. Only if step 2 shows native caption/heading/cover behavior differing
#    from this document's stable-results table, also rerun:
WPSCOMPOSER_RUN_REAL_WPS=1 WPSCOMPOSER_M3_EVIDENCE_DIR=build/longform-m3/postfix-$(date +%s) \
  .venv/bin/python -m pytest -q tests/longform_m3/test_macos_real_wps_m3.py
WPSCOMPOSER_RUN_REAL_WPS=1 .venv/bin/python -m pytest -q tests/longform_m4/test_macos_real_wps_m4.py
```

### Decision tree

- Suite green + three M5 gates green with the stable-results table
  reproduced (pages, issue codes, caps, performance under 600 s) →
  update this document with the new evidence and 0.8.0 may ship.
- Any behavior difference in captions/headings/cover → investigate
  against the macOS add-in source first (the Windows fixes deliberately
  mirror it), rerun all three gates, and record findings here before
  shipping.
- Any failure in shared pure-Python code (plan building, snapshots,
  validators) → that is a cross-platform regression from the Windows
  run; fix it, push, and repeat from a clean checkout of the new HEAD
  on BOTH platforms.

### Notes

- No CRLF handling is needed on macOS; the byte-stable snapshot failure
  mode is Windows `autocrlf=true` only.
- The `wpsjsVersion`/probe mock tests and symlink/chmod tests that skip
  on Windows run normally on POSIX.
- Keep evidence directories out of git (they live under `build/`).
