# WPSComposer Long-form M3 Figures, Tables & Cross-references Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task. Steps use checkbox syntax.

**Goal:** Implement the M3 milestone: editable native figure/table/equation numbering, populated figure/table indexes, three-line and constrained-merge tables, deterministic image preparation and sizing, and bookmark-backed cross-references on Windows COM and macOS WPS JSAPI.

**Architecture:** Keep parsing, semantic resolution, image/table policy, bookmark allocation, and plan construction platform-independent. Extend the closed protocol v2 plan with fully resolved caption, media, merge, bookmark, reference-run, and index descriptors; executors translate those descriptors into WPS-native `STYLEREF`, `SEQ`, `REF`, bookmark, table, and inline-shape objects without re-deriving policy. Reuse M2's dedicated executors and bounded field-convergence algorithm, removing only the M3 operations from their deferred-operation maps.

**Tech Stack:** Python 3.9+, pytest, Pillow>=10, existing longform semantic/policy/plan/executor modules, pywin32 COM mocks (Windows implementation), WPS JSAPI loopback add-in (macOS implementation and real-WPS evidence), JSON-only generation protocol v2, read-only OOXML inspection for acceptance evidence.

## Global Constraints

- Implement only M3. Native Office Math construction, bibliography/citation ordering, the general M4 degradation framework, the M5 PDF quality/re-layout gate, public `generate()` default migration, and release/version changes remain out of scope. M3 does implement the native equation-number field shell needed to prove `(1-1)`; equation content remains the existing readable M2 fallback until M4.
- Use WPS APIs to create all headings, captions, fields, bookmarks, indexes, images, and tables. Do not add OOXML mutation or post-processing; OOXML may be read only for structural assertions.
- The semantic model and plan builder remain platform-independent and deterministic. They must not import COM/JSAPI modules, launch WPS, read PDF geometry, or contain runtime `GenerationIssue` page/bbox data.
- Keep generation protocol `protocolVersion: 2`, `semanticVersion: longform-1`, and `resourceManifestVersion: 1`. Plans remain pure JSON, reject unknown fields, enforce size/count/depth bounds, and pass `validate_generation_plan` before WPS starts.
- `caption_numbering: auto` is resolved per object: use chapter numbering only when a numbered H1 precedes that object; otherwise use global numbering. Explicit `global` always stays global; explicit `chapter` also falls back to global before the first numbered H1. Unnumbered H1s never reset sequences.
- Chapter captions use native `STYLEREF` plus `SEQ ... \\s 1`; global captions use only native `SEQ`. Controlled sequence identifiers are `WPSC_FIG`, `WPSC_TAB`, and `WPSC_EQ`. Visible formats are exactly `图 1-1`, `表 1-1`, `(1-1)` for chapter mode and `图 1`, `表 1`, `(1)` for global mode.
- Cross-references use `REF <generated-bookmark> \\h`. Bookmarks wrap only the complete visible number range (chapter component, separator, and sequence number), so references update after target movement without including caption prose. User text never enters a field code or bookmark name.
- Figure captions appear below figures; table captions appear above tables; formula numbers are right-aligned. Figure/image containers and table-caption first rows must use native keep-with-next/keep-together semantics so captions stay with their objects.
- Academic preset three-line tables use 1.5pt top and bottom borders, a 0.75pt header-bottom border, and no vertical or interior body borders. Cell paragraphs have zero first-line indent and use their declared left/center/right alignment.
- Supported raster inputs are PNG, JPEG, TIFF, BMP, and GIF, selected by fully decoded signature rather than filename extension. SVG is allowed only through the already-passed M0 native capability and must pass the static-SVG restrictions in the design spec. WebP is not an M3 allowed media type.
- Enforce 50 MiB per resource, at most 80,000,000 pixels, and at most 32,768 pixels on either side. Pillow decompression-bomb warnings are failures. EXIF transforms, multi-frame GIF first-frame extraction, and multi-page TIFF first-page extraction produce a lossless private PNG payload while preserving usable ICC data.
- Image sizing is deterministic in points and never changes page orientation implicitly. `auto`, `column`, `full`, and explicit `Npt` widths are capped by the available slot; aspect ratio is preserved; paired columns use a fixed 12pt gap. Explicit `orientation="landscape"` remains the only image/table-triggered direction change.
- Reuse the initial `sourceSha256`, normalized `payloadSha256`, and `normalizerId` across generation, reopen, and mutation acceptance. Resource paths and bytes remain outside plans, semantic snapshots, diagnostics, issues, and evidence JSON.
- Operation `failurePolicy` remains closed: configuration, field refresh, index update, save, validation, and publication fail hard; only named object-local errors may degrade. Unknown errors and rollback failures terminate execution.
- `writer.finalize_fields` retains M2's exact order and bound: repaginate/heading+`SEQ`; bookmarks+`REF`; TOC/figure index/table index; repaginate/page fields; stable snapshot. Require two equal adjacent snapshots within three rounds, or insert/deduplicate `FIELD_REFRESH_UNSTABLE`, run one fourth final round, and freeze.
- Windows executor code is written blind and covered with COM-object mocks. Final Windows real-WPS M3 verification remains deferred to the cross-platform gate after all milestones; do not wait for Windows hardware.
- macOS real-WPS DOCX/PDF evidence, reopen, refresh, and move/insert/delete mutation evidence must be produced on this machine before M3 is complete.
- Only the requested public artifact is returned. PDF files and screenshots created for M3 development evidence remain internal acceptance artifacts.
- `uv.lock` remains untouched. Follow RED-GREEN-REFACTOR for every behavior change: write a failing test, observe the expected failure, implement the minimum behavior, and rerun focused tests before each commit.

## Task List

### Task 1: Caption/reference semantic contract and object-local numbering

Files: modify `skills/WPSComposer/scripts/document_model.py`, `skills/WPSComposer/scripts/longform/md_parser_longform.py`, `skills/WPSComposer/scripts/longform/semantic.py`, and `skills/WPSComposer/scripts/longform/bookmark_ids.py`; create `tests/longform_m3/test_semantic_objects.py` and `tests/longform_m3/test_cross_reference_semantics.py`; extend existing parser/bookmark tests.

Interfaces: add immutable `CaptionBinding(mode: str, chapter_node_id: Optional[str], bookmark_name: Optional[str], indexable: bool, referenceable: bool)` and `CrossReferenceRun(node_id: str, target_id: str, target_node_id: Optional[str], target_kind: Optional[str], bookmark_name: Optional[str], fallback_text: str)`. Add `Paragraph.node_id` and `Span.cross_reference: Optional[CrossReferenceRun]`. `FigureBlock` gains `width`, `orientation`, `kind`, `columns`, and `caption_binding`; `SemanticTableBlock` gains `style`, `orientation`, `merge_spec`, `repeat_header`, and `caption_binding`; `FormulaBlock` gains `caption_binding`. `normalize_longform_document` splits `{{ref:...}}` markers into ordered literal/reference spans, assigns deterministic occurrence IDs, and returns resolved targets/bookmarks without changing literal code/math spans.

- [ ] Write failing tests for parser preservation of figure `width/orientation/kind/layout/columns`, table `style/orientation/merges`, deterministic paragraph/reference occurrence IDs, NFC reference text, code-span marker literalness, and multiple references in one paragraph.
- [ ] Write failing semantic tests for per-object `auto` behavior before/after numbered H1, explicit global, explicit chapter fallback before H1, unnumbered-H1 non-reset, independent figure/table/equation bindings, stable bookmark names after section movement, and collision exhaustion producing `BOOKMARK_NAME_COLLISION` plus unresolved inline runs.
- [ ] Write failing tests proving an empty figure/table caption keeps the object but sets `indexable=false` and `referenceable=false`, emits block-level `CAPTION_MISSING`, and converts references to that explicit ID into inline `REFERENCE_UNRESOLVED` fallback text.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_semantic_objects.py tests/longform_m3/test_cross_reference_semantics.py tests/longform/test_parser.py tests/longform/test_bookmark_ids.py -v` and confirm missing-type/behavior failures.
- [ ] Implement parsing, deterministic traversal IDs, object-local binding, safe bookmark assignment, inline reference splitting, and canonical serialization. Remove `_update_caption_numbering`'s current document-wide collapse of `auto`; preserve `LongformConfig.caption_numbering` as the requested mode and store the resolved mode only in each `CaptionBinding`.
- [ ] Rerun the focused tests plus `tests/longform/test_semantic.py` and `tests/longform/test_semantic_m2.py`.
- [ ] Commit: Model native caption and reference bindings.

### Task 2: Image signature preflight, normalization, DPI, and sizing policy

Files: modify `skills/WPSComposer/scripts/longform/resources.py`, `skills/WPSComposer/scripts/longform/policy.py`, `skills/WPSComposer/scripts/longform/pipeline.py`, `skills/WPSComposer/scripts/generation_plan.py`, and `pyproject.toml`; create `skills/WPSComposer/scripts/longform/image_policy.py`, `tests/longform_m3/test_image_preflight.py`, and `tests/longform_m3/test_image_policy.py`; add small generated test assets under `tests/longform_m3/fixtures/media/`.

Interfaces: add `ImageProfile(pixel_width: int, pixel_height: int, dpi_x: Optional[float], dpi_y: Optional[float], frame_count: int, source_format: str, has_icc: bool)`, `FigureImageLayout(resource_id: str, display_width_pt: float, display_height_pt: float, effective_dpi: float, media_type: str, normalizer_id: str)`, and private `PreparedLongformResource(id: str, media_type: str, source_sha256: str, payload_sha256: str, normalizer_id: str, payload_bytes: bytes, image_profile: ImageProfile)`. `PreflightResource` carries immutable image metadata plus private `payload_bytes` excluded from JSON/equality/repr; its manifest entry remains exactly the six allowed keys. `_build_executor_resources(...) -> tuple[PreparedLongformResource, ...]` replaces path-backed `GenerationResource` transport for the longform executor boundary, and `LongformExecutor.execute(plan, resources: tuple[PreparedLongformResource, ...], deadline)` is updated consistently. `resolve_figure_layout(figure, resources, content_width_pt, column_gap_pt=12.0) -> tuple[FigureImageLayout, ...]` is pure and deterministic.

- [ ] Write failing tests that detect PNG/JPEG/TIFF/BMP/GIF by decoded signature despite misleading extensions; reject renamed WebP, corrupt input, 50 MiB overflow, 80M-pixel overflow, 32,768-side overflow, and Pillow decompression-bomb warnings; and preserve the original `sourceSha256` when normalization changes `payloadSha256`.
- [ ] Write failing tests for `none-v1`, `exif-transpose-png-v1`, `gif-first-frame-png-v1`, and `tiff-first-page-png-v1`; assert transformed/multiframe payloads are lossless PNG, dimensions are post-orientation, valid ICC bytes are retained, and `MULTIFRAME_FLATTENED` is emitted without leaking paths.
- [ ] Write failing static-SVG tests for scripts, `foreignObject`, external URLs/fonts, non-data references, `DOCTYPE`/entities, more than 100,000 elements, depth over 256, and embedded-raster resource limits; passing SVG retains `image/svg+xml` and `normalizerId=svg-static-v1`.
- [ ] Write failing sizing tests for `auto` (credible embedded DPI, otherwise 96-DPI natural size), `column`, `full`, explicit point width, 12pt two-column gap, content-width capping, aspect preservation, portrait/landscape slot width, and no ratio-triggered orientation change.
- [ ] Write failing quality-classification tests: photo/scan below 150 DPI is warning, any raster below 96 DPI is block degradation, screenshot/diagram between 96 and 150 DPI is not visibly marked solely for DPI, and label-size checks run only with reliable label metadata.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_image_preflight.py tests/longform_m3/test_image_policy.py tests/longform/test_resources.py -v` and confirm failures.
- [ ] Add Pillow>=10 as a core project dependency, implement one-read in-memory payload preparation, extend the private executor resource binding to carry normalized bytes/metadata, and keep plans/manifests/diagnostics free of bytes and paths.
- [ ] Rerun focused tests and the existing pipeline/diagnostic-hygiene tests.
- [ ] Commit: Preflight and size M3 image resources.

### Task 3: Three-line table policy and all-or-nothing merge validation

Files: modify `skills/WPSComposer/scripts/document_model.py`, `skills/WPSComposer/scripts/longform/md_parser_longform.py`, and `skills/WPSComposer/scripts/longform/policy.py`; create `skills/WPSComposer/scripts/longform/table_policy.py`, `tests/longform_m3/test_table_policy.py`, and `tests/longform_m3/test_table_semantics.py`.

Interfaces: add immutable `TableMerge(top: int, left: int, bottom: int, right: int)` using one-based row/column coordinates and `TablePolicy(style: str, borders: dict[str, float], merges: tuple[TableMerge, ...], repeat_header: bool, allow_row_split: bool, cell_indent_pt: float)`. `resolve_table_policy(table, preset) -> tuple[TablePolicy, tuple[DocumentIssue, ...]]` parses A1 ranges and returns either the complete valid merge tuple or an empty tuple plus one `TABLE_MERGE_INVALID` block issue; it never applies a valid subset of an invalid declaration.

- [ ] Write failing tests for A1 parsing, case normalization, reversed/degenerate/non-rectangular syntax, overlap, bounds, non-empty covered cells, header/body crossing, header vertical merge rejection, header horizontal merge acceptance, and body vertical merge acceptance.
- [ ] Write failing tests proving one invalid range discards all requested merges while preserving every grid cell and returns one stable `TABLE_MERGE_INVALID` issue placed immediately after the caption.
- [ ] Write failing preset tests for academic defaulting to `three-line`, other presets defaulting to `grid`, explicit override, exact 1.5/0.75pt borders, no vertical/interior three-line borders, repeated header, zero cell indent, and declared direct cell alignment.
- [ ] Write failing policy tests for normal rows defaulting to `allow_row_split=false` and vertically merged body groups being marked indivisible; define runtime degradation metadata `TABLE_ROW_FORCED_SPLIT` for a group measured taller than the available page.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_table_policy.py tests/longform_m3/test_table_semantics.py -v` and confirm missing-module/behavior failures.
- [ ] Implement the parser/model/policy changes without measuring fonts or pages in the semantic layer.
- [ ] Rerun focused tests plus existing parser/semantic tests.
- [ ] Commit: Validate M3 table styles and merges.

### Task 4: Closed M3 schemas, native-field descriptors, and plan emission

Files: create `skills/WPSComposer/scripts/longform/native_fields.py`; modify `skills/WPSComposer/scripts/generation_plan.py`, `skills/WPSComposer/scripts/longform/plan.py`, and `skills/WPSComposer/scripts/recording_composers.py`; create `tests/longform_m3/test_native_fields.py`, `tests/longform_m3/test_plan_schema.py`, and `tests/longform_m3/test_plan.py`; extend recording-composer tests.

Interfaces: `caption_numbering_descriptor(kind, binding) -> dict[str, Any]` emits only `{mode, sequenceId, chapterStyleLevel, resetLevel, prefix, suffix}` from controlled values; `cross_reference_descriptor(run) -> dict[str, Any]` emits only trusted target node/kind/bookmark and fixed type-specific prefix/suffix. `writer.add_captioned_figure` gains `numbering`, `bookmarkName`, `widthMode`, `orientation`, `kind`, per-child layout/media metadata, and `keepWithCaption`. `writer.add_semantic_table` gains `numbering`, `bookmarkName`, `style`, `borderSpec`, `merges`, `repeatHeader`, `allowRowSplit`, `cellIndentPt`, `plannedDegradation`, and `keepCaptionWithFirstRow`. `writer.add_equation` gains `numbering` and `bookmarkName` but not native-math args. `writer.add_cross_reference` owns one paragraph and accepts ordered `runs` of strict `text` or `reference` variants. Index operations gain `sequenceId` and `titleStyleId`.

- [ ] Write failing schema tests for required fields, exact enums, bookmark regex `^wpsc_(fig|tab|eq|ref|head|para)_[0-9a-f]{24}$`, one/two figure-child bounds, columns requiring exactly two children, positive finite point dimensions, merge coordinates, rectangular row widths, and strict unknown-field rejection at every nested level.
- [ ] Write failing schema tests rejecting user-supplied field-code strings, raw paths, bytes/base64 payloads, arbitrary sequence identifiers, reference runs without resolved target/bookmark, invalid UTF-16 lengths, and degrade policies without explicit recoverable codes/fallback.
- [ ] Write failing native-field tests for controlled structures equivalent to global `SEQ WPSC_* \\* ARABIC`, chapter `STYLEREF 1 \\s` plus `SEQ WPSC_* \\* ARABIC \\s 1`, and `REF <bookmark> \\h`; assert user captions/IDs never occur in a field descriptor.
- [ ] Write failing plan tests for object-local global/chapter descriptors, formula numbering shell, figure child order and layout, image metadata, three-line/merge metadata, caption-missing omission from indexes/targets, populated index operations in front matter, and `writer.add_cross_reference` emission that preserves literal/reference/literal ordering within one paragraph.
- [ ] Write failing state-machine tests proving index operations remain in front matter before body content, every semantic node is owned exactly once, `finalize_fields` is last, and byte-identical input/preflight data produces byte-identical operation JSON.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_native_fields.py tests/longform_m3/test_plan_schema.py tests/longform_m3/test_plan.py tests/test_recording_composers.py -v` and confirm schema/emission failures.
- [ ] Implement the schema validators, descriptors, plan emission, recording mirrors, and precise policies: figure insertion degrades only `IMAGE_INSERT_FAILED` via `figure-child-stack-then-notice`; table insertion degrades only `TABLE_STYLE_APPLY_FAILED`, `TABLE_MERGE_APPLY_FAILED`, `TABLE_ROW_FORCED_SPLIT`, or `TABLE_INSERT_FAILED` via `grid-then-text`; reference insertion degrades only `CROSS_REFERENCE_FAILED` via its run's inline fallback. Field/index API failures remain fatal.
- [ ] Rerun focused tests plus all M1/M2 plan and generation-plan tests.
- [ ] Commit: Emit closed M3 native-object plans.

### Task 5: Shared field refresh, bookmark, and index contract

Files: modify `skills/WPSComposer/scripts/longform/executor.py`; create `skills/WPSComposer/scripts/longform/field_contract.py` and `tests/longform_m3/test_field_contract.py`; extend `tests/longform/test_executor.py`.

Interfaces: `NativeFieldAdapter` defines `repaginate_and_update_numbering()`, `refresh_bookmarks_and_references()`, `refresh_indexes()`, `repaginate_and_update_page_fields()`, and `snapshot_fields() -> tuple[FieldSnapshot, ...]`. `finalize_native_fields(adapter, max_rounds=3) -> ConvergenceResult` is the single implementation of the five-phase algorithm. Stable field keys are `(ownerNodeId, fieldKind, ordinalWithinNode)`; field kinds include `STYLEREF`, `SEQ_FIG`, `SEQ_TAB`, `SEQ_EQ`, `REF`, `TOC`, `TOF_FIG`, `TOF_TAB`, `PAGE`, and `NUMPAGES`.

- [ ] Write failing adapter-order tests proving the five methods run in the specified order each round, snapshots sort by stable key, visible results are NFC/newline-normalized then SHA-256 hashed, and no visible text/hash/bookmark mapping enters `ExecutionIssue` evidence.
- [ ] Write a failing dispatch test proving each `writer.finalize_fields` operation invokes the convergence algorithm exactly once; neither the platform operation handler nor the outer executor may perform an extra pre-refresh or post-refresh.
- [ ] Write failing convergence tests with caption/reference/index fields for second-round convergence, three changing rounds plus one frozen fourth round, missing/throwing required API as fatal, and `FIELD_REFRESH_UNSTABLE` insertion/deduplication exactly once.
- [ ] Write failing tests that moving, inserting, or deleting target ranges changes the next `SEQ`/`REF` snapshot while the stable bookmark name and owner node ID remain unchanged.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_field_contract.py tests/longform/test_executor.py -v` and confirm missing-contract failures.
- [ ] Implement the adapter contract and make M2's `finalize_fields_with_convergence` delegate to it for backward compatibility; retain existing `ExecutionOutcome` and pagination-map serialization.
- [ ] Rerun focused tests and all M2 executor tests.
- [ ] Commit: Define native caption field convergence.

### Task 6: Windows COM native figures, tables, indexes, and references

Files: modify `skills/WPSComposer/scripts/longform/windows_executor.py` and `skills/WPSComposer/scripts/writer.py`; create `tests/longform_m3/test_windows_executor_m3.py` and COM fakes under `tests/longform_m3/fakes/`; extend existing Windows executor tests.

Interfaces: remove figure/table/cross-reference operations and `writer.add_equation` from `_M2_DEFERRED_OPERATIONS`; bibliography remains deferred. The M3 equation handler inserts the existing readable source fallback plus the native number/bookmark shell, but does not call Office Math. Add `WriterComposer.add_captioned_figure_native(...)`, `add_semantic_table_native(...)`, `add_equation_number_native(...)`, `add_cross_reference_paragraph(...)`, `insert_caption_index_native(...)`, and the `NativeFieldAdapter` methods. `WindowsLongformExecutor.execute` verifies resource IDs/hashes, writes only normalized private payloads, supplies resource locators to composer primitives, and removes them in all success/failure paths.

- [ ] Write failing COM-mock tests for exact field calls: figure/table/equation sequence IDs; chapter/global `STYLEREF`/`SEQ` switches; bookmarks around number ranges; `REF ... \\h`; type prefixes/suffixes; and `TablesOfFigures.Add`/update using the internal `WPSC_FIG` or `WPSC_TAB` sequence label.
- [ ] Write failing tests for figure layout: inline-shape aspect ratio, centered stack, two-column 12pt-gap container, caption below, object/caption keep rules, explicit landscape section only, and one-child failure rolling back only that child before deterministic stack fallback.
- [ ] Write failing tests for table layout: caption above, repeated header, row split disabled, exact three-line borders, grid fallback, zero paragraph indentation, direct alignments, ordered merges, invalid-merge notice after caption, and over-page vertical group rollback to unmerged splittable grid with `TABLE_ROW_FORCED_SPLIT`.
- [ ] Write failing tests for failure boundaries: unknown COM errors, bookmark/field/index/save errors, resource hash mismatch, and rollback failure abort; only the operation's named recoverable codes continue.
- [ ] Write failing tests for the complete native refresh order and snapshots after mocked move/insert/delete mutations. Confirm the executor never calls `Dispatch`, mutates a shared WPS instance, or emits a path in issues.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_windows_executor_m3.py tests/longform/test_windows_executor.py -v` on macOS and confirm it uses mocks without launching WPS.
- [ ] Implement COM primitives and dispatch wiring, retaining `DispatchEx` ownership and M2 lifecycle behavior.
- [ ] Rerun focused tests plus existing writer/COM lifecycle tests.
- [ ] Commit: Render M3 objects through Windows COM.

### Task 7: macOS JSAPI native figures, tables, indexes, and references

Files: modify `skills/WPSComposer/scripts/longform/macos_executor.py`, `skills/WPSComposer/scripts/macos_probe/models.py`, `skills/WPSComposer/scripts/macos_probe/runtime.py`, `skills/WPSComposer/scripts/macos_probe/templates.py`, `macos/wps-jsapi-probe/addin/writer-longform-v2.js`, and relevant generated add-in assets; create `tests/longform_m3/test_macos_executor_m3.py` and extend add-in asset tests.

Interfaces: the bridge still sends one validated plan plus a private resource map. Add-in handlers `addCaptionedFigureNative`, `addSemanticTableNative`, `addEquationNumberNative`, `addCrossReferenceParagraph`, `insertCaptionIndexNative`, and the five `NativeFieldAdapter` phases consume the same descriptors as Windows. Add-in results retain `appliedOperations`, `issueCodes`, `fieldSnapshots`, `childResults`, and `paginationMap` without returning paths, field results, bookmark mappings, or resource hashes.

- [ ] Write failing Python tests for normalized payload staging, manifest/hash checks, bridge request shape, outcome parsing, child-result order, and cleanup on bridge error/timeout.
- [ ] Write failing JavaScript asset tests for exact `Fields.Add` descriptors, bookmark-range creation, native index creation/update, figure/table placement rules, table borders/merges/repeated header, and ordered refresh phases.
- [ ] Write failing JavaScript fake-object tests for two-child local rollback, table style-to-grid-to-text fallback, inline `REF` fallback in the same paragraph, hard failure on field/index/save errors, and `LOCAL_MUTATION_ROLLBACK_FAILED` escalation.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_macos_executor_m3.py tests/longform/test_macos_executor.py tests/longform_m0/test_addin_assets.py tests/macos_probe -v` and confirm missing-handler failures without launching WPS.
- [ ] Implement the add-in handlers, field adapter, resource transport, and Python executor wiring; remove figure/table/cross-reference and equation-shell operations from `LONGFORM_DEFERRED`, while keeping bibliography and native equation content outside M3.
- [ ] Rerun focused tests and regenerate/verify add-in assets through the existing template path.
- [ ] Commit: Render M3 objects through macOS JSAPI.

### Task 8: M3 acceptance fixtures and real macOS WPS mutation evidence

Files: create `tests/longform_m3/fixtures/`, `tests/longform_m3/snapshots/`, `tests/longform_m3/test_acceptance_m3.py`, `tests/longform_m3/test_macos_real_wps_m3.py`, and a redacted evidence harness under `skills/WPSComposer/scripts/longform_m3_evidence.py`; produce internal artifacts under `build/longform-m3/macos-native-<run-id>/`.

Interfaces: fixtures extend the six design-spec categories with global and chapter captions, objects before/after numbered H1, populated indexes, one/two-image figures, PNG/JPEG/TIFF/BMP/GIF/static SVG, three-line/grid tables, valid/invalid merges, a tall vertical group, resolved/unresolved references, and equation numbering shell. The evidence harness returns only capability status, WPS version, relative artifact names, SHA-256 artifact digests, object/field/index counts, refresh rounds, mutation kinds, and screenshot names.

- [ ] Write failing offline acceptance tests for canonical plan snapshots, exact native numbering descriptors, front-matter index population, safe bookmarks, reference runs, all media normalizers, three-line borders, merge degradation, caption-missing behavior, and deterministic failures.
- [ ] Add a real-WPS mutation fixture that saves/reopens a chapter document, moves the second figure/table before the first, inserts a new native captioned object, deletes another, refreshes fields, saves/reopens again, and asserts gap-free `SEQ` results plus updated `REF` and both native indexes. Repeat a smaller fixture for global numbering.
- [ ] Add real-WPS assertions that chapter results are exactly `图 1-1`, `表 1-1`, and `(1-1)`; global results are `图 1`, `表 1`, and `(1)`; unnumbered H1 does not reset; index titles are absent from TOC/TOF results; and generated bookmark names do not change after target movement.
- [ ] Add structural and visual assertions that figure captions are below and on the same page as their image containers, table captions are above and share a page with the first row, three-line borders/merges persist after reopen, and no mutation creates an empty trailing section. Inspect DOCX read-only and export internal PDF/screenshots for representative pages.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/test_acceptance_m3.py -v`, then run the bridge-gated real test `.venv/bin/python -m pytest tests/longform_m3/test_macos_real_wps_m3.py -v` with WPS available. A skip is acceptable in normal CI but is not acceptable for declaring M3 complete on this machine.
- [ ] Record a redacted `platform-evidence.json`; run the existing evidence privacy validator and manually inspect representative figure, table, index, merge-degradation, and reference pages in WPS/PDF.
- [ ] Commit: Add M3 native WPS acceptance evidence.

### Task 9: Pipeline integration, documentation, ledger, and final regression

Files: modify `skills/WPSComposer/scripts/longform/pipeline.py`, `skills/WPSComposer/scripts/longform/__init__.py`, `skills/WPSComposer/SKILL.md`, `skills/WPSComposer/references/api.md`, and `.superpowers/sdd/progress.md`; extend pipeline/import-purity/privacy tests.

Interfaces: `build_longform_generation(markdown, base_dir) -> LongformBuild` remains the public offline-plan entry and `execute_longform_plan(build, executor, deadline) -> ExecutionOutcome` remains the optional native boundary. It passes `tuple[PreparedLongformResource, ...]` exactly as defined in Task 2; private normalized payload references are released and all executor-staged copies are deleted after cleanup. No public `generate()` routing changes in M3.

- [ ] Write failing tests proving semantic/plan imports are platform-pure, no WPS process starts during build, plan/diagnostic JSON contains no payload/path/bookmark map/field hash, normalized payload cleanup runs on success/error/timeout, and M1/M2 plans without M3 objects remain snapshot-compatible.
- [ ] Run `.venv/bin/python -m pytest tests/longform_m3/ tests/longform/test_pipeline.py tests/longform_m2/test_task9_pipeline.py -v` and confirm integration/documentation failures.
- [ ] Wire private payload lifecycle into the execution boundary, export only intended internal M3 types, and document the Markdown attributes, numbering defaults, allowed media, three-line/merge rules, field-refresh behavior, and M3/M4 boundary.
- [ ] Update `.superpowers/sdd/progress.md`: consume `add_cross_reference emission` and `figure/table index content population`; record every M3 task commit/review/evidence path; carry only equation native content, bibliography, general degradation, PDF quality, and default migration forward.
- [ ] Run focused M3 tests, all longform tests, generation-plan/recording tests, evidence privacy tests, and finally `.venv/bin/python -m pytest -v`. Expected: all platform-independent tests pass; macOS bridge tests pass with the recorded real-WPS evidence; Windows native tests remain mock-only.
- [ ] Perform a fresh final self-review against the M3 spec: native semantics, object-local numbering, mutation refresh, caption/object cohesion, table parity, merge fallback, media limits, cross-reference safety, closed schema, deterministic plan, privacy, and scope boundaries. Fix every discrepancy before the final commit.
- [ ] Commit: Complete M3 pipeline and documentation.

## Scope Boundaries

### In Scope (M3)

- Native figure/table caption fields, independent figure/table/equation sequences, per-object chapter/global selection, number-range bookmarks, and referenceable equation-number shell.
- Native populated figure/table indexes in M2's front-matter slots, compact index styles, and refresh/convergence coverage.
- Figure stack/two-column layout, deterministic sizing, supported media preparation, effective-DPI classification, and child-local image fallback.
- Three-line/grid tables, repeated header, zero cell indent, direct alignment, constrained merges, row-split rules, and deterministic table fallback ladder.
- Resolved cross-reference paragraph runs and same-paragraph unresolved/failed inline fallback.
- Windows COM implementation with mocks and macOS JSAPI implementation with real-WPS mutation evidence.

### Out of Scope (M4-M5 and release)

- Office Math/WPS native equation content, restricted-LaTeX conversion, formula-image fallback, bibliography/citation numbering, and general-purpose degradation notice styling (M4).
- PDF-driven DPI/layout decisions, label OCR, page-role quality checks, bbox issue mapping, automatic re-layout, notice-only patch, performance gate, and public default migration (M5).
- Arbitrary figure grids, more than two images, subcaptions, appendix letter numbering, custom caption labels, custom field codes, WebP, network media retrieval, or non-WPS rasterization.
- Real Windows WPS acceptance in this milestone; it remains a required final cross-platform gate after all milestones.
- Version bump, marketplace publication, deployment, or release notes.

## Verification

- Semantic/policy/schema/plan tests: `.venv/bin/python -m pytest tests/longform_m3/test_semantic_objects.py tests/longform_m3/test_cross_reference_semantics.py tests/longform_m3/test_image_preflight.py tests/longform_m3/test_image_policy.py tests/longform_m3/test_table_policy.py tests/longform_m3/test_table_semantics.py tests/longform_m3/test_native_fields.py tests/longform_m3/test_plan_schema.py tests/longform_m3/test_plan.py -v`.
- Executor tests: `.venv/bin/python -m pytest tests/longform_m3/test_field_contract.py tests/longform_m3/test_windows_executor_m3.py tests/longform_m3/test_macos_executor_m3.py -v`.
- Offline acceptance: `.venv/bin/python -m pytest tests/longform_m3/test_acceptance_m3.py -v`.
- macOS real-WPS acceptance: `.venv/bin/python -m pytest tests/longform_m3/test_macos_real_wps_m3.py -v`; must execute, not skip, before completion.
- Regression: `.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/longform_m3 tests/test_generation_plan.py tests/test_recording_composers.py tests/longform_m0/test_addin_assets.py -v`.
- Full suite: `.venv/bin/python -m pytest -v`.
- Manual evidence: reopen the final DOCX in WPS, update all fields, verify move/insert/delete renumbering and references, inspect populated indexes, and inspect representative PDF screenshots for figure-caption/table-caption cohesion and three-line/merge layout.

## Key Design Decisions

1. **Numbering is object-local, not a document-wide resolved flag.** `auto` can legitimately produce global captions before the first numbered chapter and chapter captions afterward. The semantic layer records each object's mode and nearest numbered H1; executors never inspect headings to choose a mode.
2. **Native fields carry all mutable numbering.** Chapter numbers use `STYLEREF 1 \\s`; object counters use independent `SEQ WPSC_FIG|WPSC_TAB|WPSC_EQ`, with `\\s 1` only in chapter mode. No visible number is frozen into plan text.
3. **Bookmarks wrap number ranges only.** A trusted `bookmark_ids.py` result surrounds the chapter/sequence number, and `REF ... \\h` supplies the current value. This keeps cross-reference text concise and stable when captions or objects move.
4. **Reference paragraphs are one closed operation.** `writer.add_cross_reference` accepts an ordered discriminated run list, preserving multiple literal/reference spans in one native paragraph and enabling same-paragraph fallback without inventing platform-specific cursor operations.
5. **Internal ASCII sequence labels are separate from visible Chinese prefixes.** Native indexes bind `WPSC_FIG` and `WPSC_TAB`; caption/index text still displays `图` and `表`. This avoids localized WPS label lookup and user field-code injection.
6. **Three-line styling and merge validity are separate.** Border policy is deterministic. Merge validation is all-or-nothing before WPS; runtime over-page vertical groups take the one specified unmerge/split degradation path instead of heuristic partial merging.
7. **Image normalization is a resource operation, not a layout renderer.** Orientation/frame/page normalization creates a private lossless PNG payload with new payload hash; the original hash remains the semantic identity. WPS still performs native image insertion and document layout.
8. **The plan owns resolved sizes and recovery ladders.** Executors receive point dimensions, merge order, numbering descriptors, and named fallbacks. Platform code maps primitives but does not reinterpret policy.
9. **M3 proves equation numbering without claiming M4 math support.** The existing readable formula content receives a native number/bookmark container so `(1-1)` and formula references can refresh; editable Office Math remains an explicit M4 deliverable.
10. **Real mutation is the acceptance test.** Static OOXML presence is insufficient. macOS WPS must move, insert, and delete targets, update fields/indexes/references, save, reopen, and retain caption/object cohesion.

## Open Questions

1. No blocking design question remains for M3. M0 already proved native `SEQ`, figure/table indexes, bookmarks, `REF`, reopen, and refresh on both WPS platforms; M3 turns those probe primitives into the closed production operation set.
2. Windows WPS 12.0 production parity is intentionally unconfirmed in this milestone. COM mocks must assert the exact native calls now; the final all-milestone Windows run must still prove visible formats, mutation renumbering, index content, merges, and caption cohesion against a real installation.
3. Native Office Math content and formula failure fallback are deliberately not decided here; M3 fixes only the numbering/bookmark container contract, and M4 must implement the restricted-LaTeX-to-native-math path without changing that contract.
