# WPSComposer Long-form M2 Page Skeleton & Table of Contents Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox syntax.

**Goal:** Implement the M2 milestone — page skeleton (front matter, sections, page numbering, headers/footers), compact TOC, heading numbering, and body styles — by executing protocol v2 plans on both Windows COM and macOS JSAPI. Produce macOS real-WPS evidence on this machine; write the Windows executor blind-but-symmetric and cover it with COM-object mocks.

**Architecture:** Keep the M1 offline semantic pipeline intact and extend it with deterministic page-skeleton policy. Introduce a shared **executor layer** that consumes a validated protocol v2 `GenerationPlan` and drives platform-native Writer primitives. The executor returns a structured `ExecutionOutcome` with a `pagination_map`. Windows maps operations to `WriterComposer` COM calls inside a dedicated worker; macOS sends the same plan to the WPS JSAPI add-in over the existing loopback bridge. The field-refresh convergence loop (≤3 rounds, optional 4th for `FIELD_REFRESH_UNSTABLE`) is defined once as a platform-independent algorithm and implemented in each executor.

**Tech Stack:** Python 3.9+, pytest, existing document model/parser/plan/recording modules, pywin32 (Windows-only), macOS WPS JSAPI add-in runtime, Pillow, pypdf/pdfplumber (not yet used for layout decisions in M2, only installed and importable), NFC normalization, JSON-only plans.

## Global Constraints

- Implement only M2. Do not implement figure/table captions, cross-references, equations, bibliography, PDF quality gates, default migration, or M3-M5 behavior.
- The semantic plan builder remains platform-independent; no COM/JSAPI objects, PDF coordinates, or runtime `GenerationIssue` data may enter `document_model.py` or the plan builder.
- All visible text in the plan stays Unicode NFC; code spans remain normalization-none.
- Every plan must be deterministic for identical input, closed against unknown fields, and verifiable with `validate_generation_plan` before WPS starts.
- Windows verification is deferred to the end of all milestones, but Windows executor code must be written and unit-tested with mocks; do not wait for Windows hardware.
- macOS real-WPS evidence must be produced on this machine before claiming M2 complete.
- `uv.lock` must remain untouched.
- Follow TDD for every behavior change: write failing tests, observe the expected failure, then implement.

## Task List

### Task 1: Page-skeleton policy and front-matter placement

Files: modify `skills/WPSComposer/scripts/longform/policy.py`; add `skills/WPSComposer/scripts/longform/page_policy.py` and focused tests.

Interfaces: `build_policy` returns a `LongformPolicy` extended with page-role sequence, section numbering rules, header/footer text, and TOC density parameters. `build_page_policy(doc, config, policy) -> PageSkeletonPolicy` returns the ordered list of sections/roles and their page-numbering/header/footer/link-to-previous rules.

- [ ] Write failing tests for title-page omission when title is empty, title-page presence when author/date exist, front-matter section with abstract/keywords/TOC, body section with Arabic restart, temporary landscape section that does not restart numbering, header shortening at 64/32 display units, and TOC density minima.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement deterministic page-role ordering, section page-numbering formats (none for cover, roman for front matter, Arabic-restart for body, continue for landscape), header/footer flags, and `linkToPrevious` rules.
- [ ] Run focused tests.
- [ ] Commit: Resolve page-skeleton policy from semantics.

### Task 2: Semantic model extension for page roles and section metadata

Files: modify `skills/WPSComposer/scripts/document_model.py` and `skills/WPSComposer/scripts/longform/semantic.py`; add tests.

Interfaces: `StructuredDocument` carries an optional `page_roles` sequence and per-section `numbering` / `page_role` metadata. `normalize_longform_document` consumes the title/abstract/keywords/TOC auto-detection and emits `ABSTRACT_CONTENT_DEGRADED` / `PAGE_BREAK_CONTENT_DEGRADED` block-level issues when those blocks contain disallowed children. `Section` exposes `outline_level`, `numbering`, and `numbering_scheme` consistently for H1-H6.

- [ ] Write failing tests for front-matter extraction, abstract content degradation, non-empty page-break degradation, title-page absence preserving a non-cover title, and heading level gaps producing `HEADING_LEVEL_GAP`.
- [ ] Run focused tests and observe missing field/behavior failures.
- [ ] Extend the semantic model without changing legacy constructors or default behavior; map block/inline issue placements from the M1 ledger into the model.
- [ ] Run focused and all current longform M1 tests.
- [ ] Commit: Model page roles and front-matter placement.

### Task 3: Protocol v2 schema extension for M2 operations

Files: modify `skills/WPSComposer/scripts/generation_plan.py`; add tests.

Interfaces: extend `_LONGFORM_OPERATION_ARG_SCHEMAS` with `writer.set_page_role`, `writer.set_page_numbering`, `writer.set_header_footer`, and extend `writer.configure_section` args with `restartPageNumbering`, `pageNumberFormat`, `startPageNumber`, `headerText`, `footerText`, `linkToPreviousHeader`, `linkToPreviousFooter`. Extend `writer.configure_toc_styles` with density fields (`minFontSizePt`, `minSpaceBeforePt`, `minSpaceAfterPt`). Extend `writer.add_heading` with `numbering` and `numberingScheme`. Keep the operation whitelist closed.

- [ ] Write failing tests for unknown-field rejection, missing `nodeId` enforcement on mutating operations, roman/Arabic format validation, and density lower-bound validation.
- [ ] Run focused tests and confirm schema failures.
- [ ] Implement the schemas and update `ALLOWED_OPERATIONS` / `_LONGFORM_WRITER_OPERATIONS`.
- [ ] Run focused tests plus existing generation-plan tests.
- [ ] Commit: Add M2 page-skeleton operation schemas.

### Task 4: Section-aware plan builder

Files: modify `skills/WPSComposer/scripts/longform/plan.py`; extend `skills/WPSComposer/scripts/recording_composers.py`; add tests.

Interfaces: `build_longform_plan` emits a plan whose operations follow the fixed state machine: `reset -> configure document/styles -> configure and render front matter -> [configure one section -> render that section's content nodes]* -> insert indexes -> finalize_fields`. Each section is preceded by exactly one `writer.configure_section` with its role and page-numbering policy. Front-matter pages are grouped before the first body section. The TOC and indexes are inserted in front matter only when policy enables them. Heading operations carry resolved numbering information.

- [ ] Write failing tests for cover page not consuming a roman page number, front-matter roman sequence, body Arabic restart, no empty trailing section, TOC style operation presence, heading numbering scheme emitted as args, and deterministic operation order.
- [ ] Run focused tests and observe missing operation failures.
- [ ] Implement section grouping, front-matter insertion, page-numbering/header/footer operations, and TOC density configuration in the plan builder; update the recording composer to mirror the new operations.
- [ ] Run focused plan tests plus all M1 plan tests.
- [ ] Commit: Build section-aware longform plans.

### Task 5: Shared executor interface and field-refresh convergence loop

Files: create `skills/WPSComposer/scripts/longform/executor.py`; add tests.

Interfaces: `LongformExecutor` protocol defines `execute(plan, resources, deadline) -> ExecutionOutcome(staged_artifact, issues, pagination_map)`. `finalize_fields_with_convergence(executor, max_rounds=3)` runs the fixed 5-step refresh sequence, compares stable-key snapshots, and returns the converged snapshot or emits a `FIELD_REFRESH_UNSTABLE` document-level issue after a deterministic 4th round. Pagination map is a stub in M2 (real mapping arrives in M5) but must be typed and serializable.

- [ ] Write failing tests for executor protocol violations, snapshot comparison across rounds, `FIELD_REFRESH_UNSTABLE` emission after 3 rounds of change, and deterministic ordering of stable keys.
- [ ] Run focused tests and confirm module import failure.
- [ ] Implement the shared executor ABC, the convergence algorithm, and a `RecordingLongformExecutor` that records calls and returns deterministic stub pagination maps.
- [ ] Run focused tests.
- [ ] Commit: Define longform executor and field convergence.

### Task 6: Windows COM executor (blind, mock-tested)

Files: create `skills/WPSComposer/scripts/longform/windows_executor.py`; extend `skills/WPSComposer/scripts/writer.py`; add tests using COM-object mocks.

Interfaces: `WindowsLongformExecutor.execute` validates the plan, creates a dedicated `WriterComposer` via `DispatchEx`, dispatches each protocol v2 operation to the matching COM primitive, and runs the shared convergence loop. New COM primitives: `set_page_role(role)`, `set_page_numbering(format, start, restart)`, `set_header_footer(header, footer, link_to_previous)`, `configure_section(landscape, margins, restart, ... )`, `insert_toc_with_styles(title, density)`, and `add_heading_level_native(text, level, numbering, scheme)`. The executor must never fall back to a shared `Dispatch` instance; on ownership failure raise `WINDOWS_DEDICATED_HOST_UNAVAILABLE`.

- [ ] Write failing tests with mocked `win32com.client.DispatchEx`/`WriterComposer` for operation dispatch, roman/Arabic page numbering, header/footer insertion, section restart, TOC density styles, heading native numbering, and convergence loop.
- [ ] Run focused tests and confirm module import failure on non-Windows (skip with pytest.importorskip).
- [ ] Implement the executor and COM primitives behind the same operation-to-method dispatch table; keep all plan interpretation in Python so the COM side remains primitive-only.
- [ ] Run focused tests on this machine by mocking COM; confirm no real WPS is started.
- [ ] Commit: Implement Windows longform executor with mocks.

### Task 7: macOS JSAPI executor and add-in protocol v2 longform support

Files: modify `skills/WPSComposer/scripts/macos_probe/models.py`, `runtime.py`, `templates.py`, and add generated add-in assets under `macos/wps-jsapi-probe/`; add `skills/WPSComposer/scripts/longform/macos_executor.py` and tests.

Interfaces: `MacOSLongformExecutor.execute` sends the protocol v2 plan to the WPS add-in over the existing `LoopbackBridge`. Extend the add-in protocol with a `generate_longform_document` method that receives `plan` and `resources` and executes the same operation set as the Windows executor. Add-in handlers implement `configureSection`, `setPageNumbering`, `setHeaderFooter`, `insertTocWithStyles`, and `addHeadingNative` using WPS JSAPI. The add-in reports `appliedOperations`, `issueCodes`, and a stub `paginationMap`.

- [ ] Write failing tests for bridge command construction, plan serialization, add-in method routing, and macOS executor outcome validation.
- [ ] Run focused tests and confirm module import failure.
- [ ] Extend the macOS probe runtime to stage and serve the updated add-in; implement the JS-side operation handlers in the generated template; wire the Python `MacOSLongformExecutor`.
- [ ] Run focused unit tests.
- [ ] Produce real-WPS DOCX evidence on this machine for the M2 acceptance fixtures (cover/no-cover, roman/Arabic restart, header/footer, TOC density, heading numbering, temporary landscape). Capture the generated files and any screenshots.
- [ ] Commit: Add macOS longform executor and real-WPS evidence.

### Task 8: Real-WPS acceptance fixtures and snapshots for M2

Files: create `tests/longform_m2/` fixtures and tests; update `tests/conftest.py` if needed.

Interfaces: fixtures exercise the six required acceptance categories restricted to M2 scope. Tests assert that generated DOCX packages contain the expected structure: no page number on cover, roman page labels before body, Arabic page labels restarting in body, centered header text with a paragraph bottom border line, centered footer page number, TOC entries with compact leading/trailing spacing, and native heading numbering in the document XML. Tests run against the recording executor offline; macOS real-WPS tests run only when the bridge is available.

- [ ] Write failing tests for cover/no-cover documents, roman/Arabic page numbering, header/footer presence and shortening, compact TOC density, heading numbering schemes, no blank pages from empty front matter, and temporary landscape section page continuity.
- [ ] Run focused tests and observe missing fixture/executor behavior.
- [ ] Create Markdown fixtures and expected structural snapshots; wire them to the recording executor and, on macOS, to the real-WPS executor.
- [ ] Run focused tests plus the full longform suite.
- [ ] Commit: Add M2 acceptance fixtures and snapshots.

### Task 9: Pipeline-to-executor integration and diagnostic hygiene

Files: modify `skills/WPSComposer/scripts/longform/pipeline.py` and `skills/WPSComposer/scripts/generation_plan.py` if needed; add tests.

Interfaces: `build_longform_generation` continues to return an offline `LongformBuild`. A new optional step `execute_longform_plan(build, executor, ...) -> ExecutionOutcome` binds a validated plan and resource manifest to an executor without changing the public `generate` path. Remove `sourcePath` from diagnostic JSON output; keep it only inside the private staging transport map. Ensure the pipeline never imports platform modules unless an executor is explicitly provided.

- [ ] Write failing tests for diagnostic JSON redaction (no `sourcePath` in `LongformBuild.to_json`), executor binding, and platform module isolation.
- [ ] Run focused tests and observe failures.
- [ ] Implement the executor binding helper and redact absolute paths from the public build JSON.
- [ ] Run focused tests plus the full suite.
- [ ] Commit: Bind executor to pipeline and redact source paths.

### Task 10: Ledger cleanup and M2 follow-up bookkeeping

Files: remove `.superpowers/sdd/task-6-report.md` if it is a historical untracked artifact; update `.superpowers/sdd/progress.md`.

Interfaces: `.superpowers/sdd/progress.md` records each M2 task, commit, review, and which ledger items were consumed or deferred.

- [ ] Verify `.superpowers/sdd/task-6-report.md` is safe to remove; if not, document why it must stay.
- [ ] Remove the file and record the cleanup.
- [ ] Update the progress ledger with M2 consumed/deferred items: consume `block/inline issue->degradation mapping`, `ABSTRACT_CONTENT_DEGRADED/PAGE_BREAK_CONTENT_DEGRADED`, and `sourcePath` redaction; defer `add_cross_reference emission` and `figure/table index front-matter placement` to M3.
- [ ] Run the full test suite to ensure no regressions.
- [ ] Commit: Clean up historical ledger file and record M2 state.

## Scope Boundaries

### In Scope (M2)

- Front matter: title page, abstract, keywords, TOC, figure/table index placeholders (structure only; content population is M3).
- Sections: cover, front-matter section(s), body, and temporary landscape sections.
- Page numbering: none on cover, roman numerals for front matter, Arabic numerals restarting at 1 in body, no restart for temporary landscape.
- Headers/footers: centered body header with bottom border, centered footer page number, no header/footer on cover, link-to-previous for body sections.
- TOC density: TOC 1/2/3 font sizes and paragraph spacing with defined minima for the single allowed compact-retry.
- Heading numbering: native WPS multi-level lists for `chinese-formal`, `decimal`, and `hybrid-bid`; H1-H4 numbering, H5-H6 unnumbered but styled.
- Body styles: Songti/Times New Roman 12pt, 1.5 line spacing, first-line indent, justified alignment.
- Executor layer: shared interface, Windows blind implementation with mocks, macOS real-WPS implementation.
- Field-refresh convergence loop: ≤3 rounds, optional deterministic 4th round on instability, `FIELD_REFRESH_UNSTABLE` issue.
- M2 ledger cleanup and diagnostic path redaction.

### Out of Scope (deferred to M3-M5)

- Native captions, figure/table index population, cross-references, and references (M3).
- Equation objects and bibliography styling (M4).
- PDF quality gate, page-role-aware checks, issue-to-node mapping, re-layout loop, performance gate, and default migration (M5).
- Public `generate()` routing changes: M2 executors are reachable through internal helpers only.
- Windows real-WPS verification: deferred to the final cross-platform acceptance gate.
- SVG support: M0 optional capability gate; do not assume enabled.

## Verification

- All platform-independent tests pass: `.venv/bin/python -m pytest tests/longform/ tests/longform_m2/ -v`.
- All generation-plan and recording-composer regression tests pass: `.venv/bin/python -m pytest tests/test_generation_plan.py tests/test_recording_composers.py -v`.
- Windows executor tests pass with mocked COM on this machine: `.venv/bin/python -m pytest tests/longform_m2/test_windows_executor.py -v`.
- macOS real-WPS executor produces valid DOCX files for all M2 fixtures; structural assertions (page labels, header/footer text, TOC styles, heading numbering) pass against exported OOXML where feasible.
- Field-refresh convergence tests demonstrate ≤3-round convergence for stable fixtures and deterministic `FIELD_REFRESH_UNSTABLE` handling for deliberately unstable fixtures.
- No unexpected blank pages are introduced; temporary landscape sections do not restart Arabic numbering or create empty trailing sections.
- Full suite passes: `.venv/bin/python -m pytest -v`.

## Key Design Decisions

1. **Executor abstraction.** Both platforms consume the same validated `GenerationPlan`. The Python side owns operation ordering, failure-policy interpretation, and the convergence algorithm; the platform code only implements primitive operations. This keeps Windows and macOS symmetric even though Windows cannot be tested on this machine.
2. **Section configured per content group.** Each logical section (cover, front matter, body chapter, landscape insert) is preceded by exactly one `writer.configure_section` operation. The plan builder computes all numbering/header/footer/restart policy before WPS starts; executors do not re-derive it.
3. **Field convergence is an executor primitive.** `writer.finalize_fields` is a single protocol operation whose implementation runs the fixed 5-step refresh sequence and up to 4 rounds. Snapshot comparison uses stable keys (`ownerNodeId + fieldKind + ordinalWithinNode`) and SHA-256 of NFC-normalized results, never customer-visible text.
4. **macOS evidence first.** Because the macOS bridge works here, M2 acceptance is gated on real DOCX/PDF output from this machine. Windows code is written to the same contract and validated with mocks; the final cross-platform gate comes after M5.
5. **TOC density is a one-way compact retry.** The plan builder emits density parameters; the executor applies them once. No automatic second retry is allowed in M2 — that belongs to the M5 PDF quality gate.
6. **Issue placement mapping from the M1 ledger is resolved now.** Document-level issues go to the fixed quality-notice area; block-level issues map to `writer.add_degradation_notice`; inline issues map to `writer.add_inline_degradation`. `ABSTRACT_CONTENT_DEGRADED` and `PAGE_BREAK_CONTENT_DEGRADED` are emitted during semantic normalization.

## Open Questions

1. Should the macOS add-in continue to receive the whole plan as one JSON blob, or should M2 split it into per-operation commands? A single blob matches the existing `generate_writer_document` shape and is simpler; per-operation commands would make partial progress easier but add latency. Recommend keeping a single plan blob for M2.
2. Does WPS macOS JSAPI expose a reliable `LinkToPreviousHeaderFooter` equivalent, or must we copy header/footer ranges between sections? This needs one probe document in Task 7 before finalizing the JS-side implementation.
3. The spec requires the cover page to use no page number and not count toward roman numbering. On Windows COM this is typically done with a cover section whose footer has no page field and whose next section starts roman at 1. Is the same two-section structure the cleanest on macOS JSAPI? Prototype in Task 7.
4. Should `writer.configure_section` carry an explicit `pageRole` enum, or should roles be inferred from operation sequence? Explicit `pageRole` (`cover`, `front_matter`, `body`, `landscape`) makes executor behavior deterministic and simplifies pagination-map stubs. Recommend adding it in Task 3.

