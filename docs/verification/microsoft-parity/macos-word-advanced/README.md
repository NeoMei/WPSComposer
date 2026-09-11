# Mac Word advanced parity evidence

Task3 native feasibility/implementation acceptance uses real Microsoft Word AppleScript object operations. Python only builds plans, stages resources, reads saved DOCX ZIP/XML and renders the native PDF for inspection; it never edits OOXML. Each native operation uses the existing MacWordAdapter private container staging, WordJobLock, total deadline, exact owned-document cleanup and preexisting-document snapshots. Integrated runs additionally create an unchanged synthetic unsaved sentinel and close that reference only after exact name/path/saved/text verification.

## Evidence status

- `primitive-01`: genuine native equation saved and reopened/PDF exported. Also revealed that appending body text in an equation paragraph can create an unwanted second OMML equation.
- `primitive-02` / `primitive-03`: rejected AppleScript syntax for paragraph indentation retained. Word uses `paragraph format left indent` / `paragraph format right indent`.
- `primitive-04`: native merges and landscape/two-column sections work. Eight-bit RGB setters silently saved black; these artifacts are not correct color acceptance.
- `primitive-05`: sixteen-bit AppleScript RGB (each channel multiplied by 257), document-local custom paragraph style, and an equation in a 1×3 native table passed save/reopen/PDF and XML checks. Following body text remained outside OMML.
- `integrated-01`: fixture import error, unchanged sentinel safely closed.
- `integrated-02` / `integrated-03`: fixture resource paths containing `:` caused native image object access failure. All owned documents cleaned; sentinel preserved.
- `integrated-04`: automated first-pass structural checks were true, but subsequent PDF/XML review found that inserting at a cell's `end` placed images past cell terminators, creating a spurious third cell and cropping the second image. **The report's original `passed: true` is superseded by this finding; this run is not accepted.** It also contains a normalization issue later fixed independently by the parent task.
- `integrated-05` / `integrated-06` / `integrated-07`: native custom-bullet `NumberStyle = bullet` failed with `-10006` across document-local template variants. These failures remain evidence of an unsupported capability, not passes. The compiler rejects custom glyphs before Word startup.
- `integrated-08`: normal core acceptance, 16 structural/PDF/sentinel checks true, 13 pagination nodes, 3 pages. PDF pages were opened and visually reviewed: native equation, following body, landscape merged table and intact outside cell, restored portrait page, two complete side-by-side images, custom colored/underlined/struck style, independent bold/strike inline runs, and text plus PAGE footer. The only issue is the intentionally inserted `REFERENCE_UNRESOLVED` cell fallback. Subsequent formula recovery/width changes require later runs listed below.

Earlier report source hashes were read at the end of each run; the emitted `compiled.applescript` is the authoritative executed script when concurrent edits occurred. The fixture now snapshots source/compiler hashes before execution and records a compiled-source hash separately.

- `final-normal` / `final-normal-02`: all **17** structure/PDF/sentinel checks passed, 13 pagination nodes and 3 PDF pages. Formula middle column receives the available body width minus the two 36-point numbering gutters. Final body remains outside OMML.
- `final-source-fallback`: deliberately failed the build-up command after a real OMML object was created. This correctly exposed that setting cell content retained an OMML wrapper; `source_fallback_retained` was false, so this run failed.
- `final-source-fallback-02`: all **15** checks passed after deleting only the owned formula's precise text range before writing the readable source. The native DOCX contains no OMML after recovery; the numeric field/bookmark, visible `EQUATION_INSERT_FAILED` notice, dynamic execution issue, other content, PDF and unsaved sentinel survive. The PDF was visually inspected.

`final-normal-02` and `final-source-fallback-02` executed compiler SHA-256 `f601d34dd0734d13eefde8e276b1fea8699b1a085e62966358640d93e1268164`. The subsequent compiler changes only add preflight rejection for invalid alignment and mismatched span text; emitted native operations for these accepted fixtures are unchanged.

## Platform-independent regression

- Longform M1–M4 plus Mac compiler/runtime/public-error tests: **1553 passed, 10 skipped**, `/tmp/wpscomposer-parity-word-advanced-tests-final.log`.
- M5 tests: **130 passed, 2 skipped**, `/tmp/wpscomposer-parity-word-m5-tests-final.log`.
- Focused advanced compiler tests: **27 passed** including the actual cropped-column DOCX regression, range-boundary rejection, native fallback records and explicit paragraph rules. The complete related Mac group is **71 passed**.
- Skipped tests require opted-in real WPS acceptance; these are not Microsoft Word native passes.

## Explicit remaining boundaries

- Custom bullet glyphs: native setter failures retained; preflight rejects them. Default native bullets and numbering continue to work.
- Custom character styles and hyperlink spans: no native acceptance yet; preflight rejects them instead of dropping attributes.
- Equation `fallbackResource`: explicit native image fallback is not implemented and is rejected before startup. Source fallback retains visible issue text and numeric bookmark; its native injected-failure evidence is recorded separately.
- SVG pictures remain unsupported by this native Word compiler; the closed PNG/JPEG normalization path is accepted.
- These results concern the generation-plan adapter. They do not establish Windows parity, session editing, or full frozen capability-catalog completion.

## Reproduction

From the `office-description` worktree, use the established Python environment and a new output directory:

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps PYTHONDONTWRITEBYTECODE=1 \
../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_advanced.py \
  --output-dir docs/verification/microsoft-parity/macos-word-advanced/new-native-run
```

Add `--force-math-failure` to replace only this fixture's native equation build command with a deliberate error and verify the M4 readable-source recovery path. Raw scripts/logs, native DOCX/PDF, PDF images, issue records and sentinel cleanup evidence are retained for every run. No global Quit, process kill, Normal template change or macro execution is used.
