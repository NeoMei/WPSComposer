# Semantic table native fixture preparation review

Date: 2026-09-09. Independent review of fixture SHA-256 `5d8cb94250235e4349a0981fee1d1c166dc9cabc5c82354081f5f73f3cc61f0e` and test SHA-256 `9e5a553bb67d34e81463e931bd832a68b43a133d3fdffa90305fc0b20b2a14c8`.

**Result: one P2 evidence-validation finding. No unsafe ownership or uncertainty-followup defect found. Fix the validation gap and group the known-risk count expressions before using this as an automated native acceptance gate.** A bounded diagnostic native run is not blocked on ownership grounds, but the current automated PASS must not certify complete outside preservation.

## P2: outside preservation silently drops textless structural paragraphs

Location: `fixtures/microsoft_parity/macos_word_semantic_table.py:279` (outside projection in `xml_checks`).

The outside comparison only retains paragraphs for which concatenated `w:t` text is nonempty. Consequently it discards every empty paragraph, including paragraphs with `w:pageBreakBefore`, run-level breaks, spacing or other layout-affecting properties. This is broader than permitting the one explicit tail CR emitted by the primitive. An unintended extra blank paragraph or page break outside the replaced range can change document pagination while all XML checks still pass; original blank paragraphs are also unprotected. The fixture report's claim that unknown added outside paragraphs fail is therefore incomplete.

Independent pure reproduction: load the existing passing `xml_pair()`, insert an empty `w:p` with `w:pPr/w:pageBreakBefore` into the after body immediately before its final section properties, and call `xml_checks(before, mutated_after)`. Observed result:

```text
native_table=True, grid_widths=True, borders=True, merge_topology=True,
cell_paragraphs=True, row_flags=True, outside_preserved=True
```

Preserve all outside structural nodes, including textless paragraphs. Exclude only the precisely identified original replacement and inserted table/tail sequence, with the exact allowed paragraph-count and formatting delta. Add negative tests for an extra empty paragraph, an outside page-break paragraph and mutation/removal of a pre-existing empty paragraph. Do not resolve this by weakening outside acceptance or dropping paragraph comparisons.

## Before-run count expression correction

`mutation_commands` lines 130 and 133 emit `if count sections of boundDoc is not 1` and `if count tables of boundDoc is not 1`. The parent identified the same expression form in an already recorded native quality failure. I checked `word-quality-implementation-report.md:121`: the original ungrouped count failed with -1700 while grouped and atomic forms succeeded. Use `if (count sections of boundDoc) is not 1` and the corresponding grouped table count before this run. This is a known-pattern correction supported by the existing report; I did not execute these two semantic-table guards natively.

The style readback already uses `row 1 of semanticTable` for the heading property, and final table lookup uses native ordinals, so those do not repeat the `repeat with item in collection` alias defect. The aggregate `allow break across pages of row options` getter remains a native hypothesis. If it fails, preserve the failed run and use separately reviewed ordinal row observations; compilation cannot establish its runtime behavior.

## Verified preparation properties

- Explicit `execute=True`, fresh output directory and separately managed Word lease are required. Missing local dictionary fails before native inventory.
- Bound document/window paths use exact Foundation string equality. Before table mutation, selection main story, explicit noncollapsed range/text, page measurements, section and prior table counts are guarded. The table reference returned by the unchanged primitive is retained, then recovered uniquely via bookmarked bounds instead of assuming final table ordinal.
- Native styles are observed before merges. Shape of native evidence, integer versus boolean distinctions, native text/width/indent/alignment/border flags, explicit CR/cursor, prior table text and source paths are checked. Required merge topology and saved content are inspected independently in OOXML. PDF gates check tokens/order, bounds, absence of images and actual 480pt drawing geometry, while manual visual inspection remains explicitly pending.
- Session lifecycle creates private owned documents; no active-document mutation, close-all or application quit was added. Sentinel content/state/hash and unique identity use the existing reviewed helpers. Owned close is independently checked before one exact sentinel close; final inventory is compared with the initial inventory.
- Observation errors, malformed results and unverified native lifecycle calls mark uncertainty and refuse later observations. Cleanup observes that marker and avoids further native work on uncertain completion. Failure/evidence-copy handling retains the primary error. Read-only reopen checks native content/bounds and saved-source digest.
- The fixture remains a portrait primitive gate. It does not activate public semantic methods or claim arbitrary rollback, section restoration, multipage repeat/split behavior, controller fallback or UI edit/Undo acceptance.

## Independent execution evidence

Pure test command from the preparation report was rerun: **51 passed, 5 dependency deprecation warnings in 0.49s**, exit 0. It includes 13 fixture tests, 17 primitive tests and 21 existing safety tests.

At the parent's explicit request, performed syntax compilation only for the generated seed, mutation and result scripts with the session's `_JSON` helpers and local Microsoft Word dictionary. All three `/usr/bin/osacompile` calls returned 0; `/usr/bin/osadecompile` retained the ungrouped count expressions, so successful compilation is not a runtime clearance. No `/usr/bin/osascript`, AppleEvents, Word document open, UI interaction or native fixture execution was performed by this reviewer.

Compile-only artifacts: `/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpsc-semantic-fixture-review-7pfq1p8b/` (three source scripts, compiled scripts and decompiled sources). The XML finding was reproduced in memory without file/source mutation. This review report is the only workspace file written.

## Fix verification and bounded native admission

Reviewed corrected fixture SHA-256 `46a275cda25eec64c664713564f6d9cde2f7cd6d402a697bad30060580e62871` and test SHA-256 `363a02672c9b8a8d61e4290f95e9d1fa0157e064832234f4dc40012c5a59c7cc` on 2026-09-09. **The P2 finding is closed. The corrected fixture is suitable for one root-controlled bounded native run under the exclusive Word lease.** This is fixture admission, not native acceptance or public-method parity.

The XML validator now identifies the unique original REPLACE paragraph, requires the replacement table at that exact body index followed by exactly one empty tail paragraph with matching original paragraph properties, and compares every remaining body node in order. It preserves existing empty paragraphs and their formatting, breaks and section structure. Only the documented volatile identifiers/bookmarks and empty unformatted tail runs are normalized. It does not generally discard textless nodes. The two native count comparisons are explicitly grouped.

Independently reran the same three test files: **66 passed, 5 dependency warnings in 3.15s**, exit 0. The 15 added cases cover extra textless paragraphs, page breaks, section properties, original blank-paragraph mutation/removal, tail position/count/format changes and grouped count conditions. Reran the original independent pageBreakBefore reproduction: `outside_preserved=False`, as required. The corrected generated mutation script also passed compile-only validation with exit 0. No AppleEvents or fixture/native/UI execution was performed.

No new concrete issue found. Existing native hypotheses, source freeze, uncertainty-stop behavior and acceptance limits above remain in force; preserve any first-run failure and require actual output/readback/visual evidence before claiming native success.
