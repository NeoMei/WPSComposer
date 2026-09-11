# Semantic table primitive implementation report

Date: 2026-09-09. Scope: approved parity Task 3, pure construction prerequisite only. Parent task owns native lease, index, integration and independent review.

## Outcome

Added `msoffice/macos_word_semantic_table.py` and its focused tests. No shared session, fields, recovery, quality, facade, capability registry, public API, application or native fixture was changed by this subtask. No commit/push/index or AppleEvents were run. Read-only Word.sdef inspection supplied command vocabulary; this is not AppleScript compilation or native execution proof.

`iter_semantic_table_commands(...)` yields immutable ordered requests with action/args and frozen operation-local error code. `build_semantic_table_commands(...)` materializes the sequence. `render_semantic_table_commands(commands)` produces only the native body for an already guarded boundDoc/boundWindow scope. No success ACK or session execution is fabricated.

Contract preserved on the supported construction boundary:

- Explicit start/end of the bound main-story range; no document-end append, paragraph-boundary normalization, or active-document lookup. The returned native table reference remains authoritative even when inserting before another table.
- Materialize original headers/rows; use `str(value)` for each visited cell without padding ragged rows. Original values are retained for `_content_column_widths`, including its original false/None and multiline/CJK behavior.
- FirstLineIndent=float(cell_indent_pt), Left/RightIndent=0, exact left/center/right mapping. No font, padding, shading, paragraph spacing, vertical alignment or heading-bold defaults added.
- Set row split by truthiness; only write repeat-header=true when requested, including a one-row table. Do not force a false heading property when repeat-header=false.
- Caller supplies available body width from frozen last-section -> document PageSetup lookup. Clamp numeric width to at least 72pt and reuse `_content_column_widths`. None explicitly skips fitting. Preserve AutoFitBehavior(0) best effort, AllowAutoFit=false, preferred width points best effort, SetWidth(...,adjust-none) with Width fallback. Optional script catches rethrow timeout (-1712), disconnect (-609), and cancellation (-128).
- Border order top/left/bottom/right/insideHorizontal/insideVertical, then row-1 bottom override. Zero disables style without writing width; .75pt maps to Word enum6 / line width75 point, 1.5pt maps to enum12 / line width150 point, every other nonzero finite value maps to enum2 / line width25 point.
- Merges retain input order and integer coordinates, without overlap sorting, reverse order or rewriting. Complete all styles/borders before merges; after successful merges move to actual table.Range.End, emit one CR at that explicit range, and move the bound-window selection to table end+1.
- Reuse recovery._literal solely as a pure encoder for quotes, backslashes, Unicode, CR/LF/TAB and control characters such as BEL. No recovery orchestration is reused or changed.

## Error and integration boundaries

This primitive is NOT `add_semantic_table_native` or its fallback. Do not enable those capability rows from these tests.

The iterator exposes consumed command prefixes when a descriptor fails, with TABLE_STYLE_APPLY_FAILED or TABLE_MERGE_APPLY_FAILED according to the frozen phase. Materialization errors before the frozen try remain ordinary errors. Bound coordinates and nonfinite measurements are closed construction inputs. The eager builder validates the whole sequence before any native submission, so its validation timing is not the public direct API's partially-mutating failure timing. The eventual direct integration must explicitly resolve that behavior rather than treating eager normalization as a drop-in method.

Each command carries TABLE_INSERT_FAILED / TABLE_STYLE_APPLY_FAILED / TABLE_MERGE_APPLY_FAILED for a **confirmed operation-local** failure only. The emitted body does not catch native errors into these types. It does not convert timeout, cancellation, lost identity or unknown submission/ACK outcomes to recoverable failures. Session execution must retain existing quarantine/evidence behavior and must not infer a typed recoverable error from this metadata alone.

Required follow-up integration/native work:

1. Obtain exact range and source-bound body-width readback using existing owned/bound-document identity, main-story, deadline and selection guards; validate they remain current immediately before mutation. None width must mean the frozen PageSetup fallback actually failed, not an arbitrary default.
2. Execute under existing topology mutation and submission-state guards. Add source-bound native table identity/count/bounds, cell/format/width/border/merge readback and completion ACK; never identify an inserted middle table by final collection index. Prove selection.Range.End + one CR behavior in nonempty/nonterminal ranges and existing table adjacency.
3. Implement confirmed local failure classification while preserving uncertain transport failures, plus paragraph/table mutation preimage and rollback (including replaced text for noncollapsed ranges). Table/caption/notices checkpoints and controller-owned fallback policy remain above this primitive. Existing append-only recovery does not prove nonterminal replacement rollback.
4. Add postmerge stable-cell lookup and vertical-merge page-overflow detection, then cell degradations. Caption/numbering/issue return contracts, grid/text retry order and typed degradation failures remain unimplemented here.
5. Implement landscape/includePreviousHeading/continuousExit wrapper state and section restoration independently. Then actual Word create/save/close/reopen/XML/PDF and injected create/style/merge/tail failure evidence; preserve synthetic unrelated unsaved document. Dictionary vocabulary alone is insufficient.

## Verification

TDD skills read: test-driven-development and its writing-good-tests reference; verification-before-completion. Frozen oracle loaded directly via AST from `git show 6dd3a00:skills/WPSComposer/scripts/writer.py`: _create_native_table, _apply_native_table_borders, _native_border_width, _fit_native_table_to_body, _current_section_page_setup, _visual_text_width, _content_column_widths. COM-boundary doubles record actual frozen property requests; no pywin32 or native Office dependency is loaded.

- Initial RED: 16 failures because the pure module was absent, all with the explicit missing-builder assertion.
- GREEN: 16 passed after implementation.
- Additional RED: control-text regression failed with TABLE_STYLE_APPLY_FAILED, 16 passed.
- Final GREEN: 232 passed in 1.93s across new semantic-table tests (17), existing numbering/business-object tests and Windows M3/M4 executor suites. Exit code 0.

Command:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_macos_word_semantic_table.py tests/msoffice/test_macos_word_numbering.py tests/msoffice/test_macos_word_business_objects.py tests/longform_m3/test_windows_executor_m3.py tests/longform_m4/test_windows_executor_m4.py -q --tb=short
```

The complete repository suite, AppleScript compilation, native mutation, save/reopen/XML/PDF, failure rollback, independent review and public capability acceptance were not run by this subtask.

## Scoped source hashes

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_semantic_table.py` | `9655e9af00366228bcca63063d2341d9687348fa2cd9e9663632ddee9c2d65ec` |
| `tests/msoffice/test_macos_word_semantic_table.py` | `01282417b44e52c9652dee111c061a6ce7f1b199665b9f65095c48e5290f4c30` |
