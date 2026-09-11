# Semantic table primitive independent review

Date: 2026-09-09. Reviewer: independent read-only subagent. Result: **scoped PASS; no actionable findings** in the two reviewed files. This does not accept native execution, direct-method parity, rollback or capability activation.

## Reviewed snapshot and scope

Worktree HEAD: `1686155b5b16c6709b0182b40dbf5a6ed0f3d89f`.

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_semantic_table.py` | `9655e9af00366228bcca63063d2341d9687348fa2cd9e9663632ddee9c2d65ec` |
| `tests/msoffice/test_macos_word_semantic_table.py` | `01282417b44e52c9652dee111c061a6ce7f1b199665b9f65095c48e5290f4c30` |

Read the local AGENTS.md, regression guardrails, `word-remaining-objects-design.md`, implementation report, and frozen `6dd3a00` Writer table factory/border/width helpers. No `.codegraph/` exists in this worktree. Current helper and pure literal encoder were inspected as dependencies. A source search found no callers/imports of the new primitive outside its module. Concurrent quality/runtime changes were excluded from review.

## Findings and contract assessment

No concrete worthwhile defect was found within the internal command-construction boundary.

- Cell materialization preserves original values for width measurement and applies `str(value)` for text, including None, booleans, multiline CJK/emoji and BEL/VT. Ragged rows remain unpadded. First-line indent, zero left/right indent and exact left/center/right mapping match the frozen factory; no unrelated font, padding, shading or paragraph-spacing defaults are emitted.
- Fill precedes row split/header flags, fitting, ordered table borders, header-bottom override, input-order merges and finish. False repeat-header does not force a property reset. Row-split truthiness matches COM boolean semantics.
- Width uses the existing content algorithm with the frozen 72pt clamp and explicit None skip. Optional AutoFit/preferred-width and SetWidth-with-Width-fallback structure matches the frozen sequence. The optional catches rethrow the explicitly handled timeout, disconnect and cancellation errors.
- Border mapping matches COM `-1..-6`, then header `-3`: zero writes no width; .75 maps to 6, 1.5 to 12, and other finite nonzero widths map to 2. Read-only `/Applications/Microsoft Word.app/Contents/Resources/Word.sdef` inspection confirms emitted width enum names and row/column/border/merge vocabulary. Dictionary inspection is not compiler or execution evidence.
- Creation uses the caller's explicit bound main-story range and retains the returned native table reference. Finish derives the actual table end, emits a CR at an explicit boundDoc range and moves boundWindow selection to end+1. It does not substitute document-end append or a final collection index.
- Confirmed-operation metadata separates insert, style, merge and tail insertion phases. The renderer supplies no fake ACK or blanket conversion into recoverable failures. The module and report correctly disclaim native classification and execution ownership.

## Independently rerun verification

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_macos_word_semantic_table.py tests/msoffice/test_macos_word_numbering.py tests/msoffice/test_macos_word_business_objects.py tests/longform_m3/test_windows_executor_m3.py tests/longform_m4/test_windows_executor_m4.py -q --tb=short
```

Result: **232 passed in 1.99s**, exit 0; includes all 17 new tests.

An additional in-memory script ran **120 deterministic pure differential edge cases**, seed 9143, using the AST-loaded frozen factory. Cases varied 1–6 columns, empty/ragged rows, None/boolean/numeric/control/Unicode cell values, width None/negative/below-minimum/fractional values, truthy and falsy header/split values, numeric-string and negative indents, and zero/negative/standard/nonstandard border widths. Assertions compared cells, widths, flags and borders to frozen requests and checked border and phase ordering. Result: all passed, exit 0. No source or test edits were needed.

## Acceptance boundary

The eager builder's normalization and a cell command's grouped validation do not establish the frozen public API's exact partially mutating failure timing. The implementation report already leaves this unresolved for caller integration; this scoped pass does not waive it.

Native guarded submission/ACK, source-bound table identity and readback, nonterminal replacement and adjacency semantics, save/close/reopen/XML/PDF, confirmed-error classification, table/tail rollback, merged-cell/page-boundary validation, caption/degradation/fallback policy and landscape section restoration remain unverified here. Public semantic-table capability rows must remain pending until their own contracts and evidence pass.

This reviewer did not run AppleEvents, apps, UI, AppleScript compilation, native fixtures, full repository tests, index changes or commits. The only file written was this review report. The reviewed source hashes were rechecked after tests.
