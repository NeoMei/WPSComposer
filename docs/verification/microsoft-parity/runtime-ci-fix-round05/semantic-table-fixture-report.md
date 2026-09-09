# Semantic table native fixture preparation report

Date: 2026-09-09. Scope: independent native fixture preparation for the already reviewed pure table primitive. Implementation and tests are ready for independent review; **no native run or acceptance claim**.

## Files and isolation

New fixture: `fixtures/microsoft_parity/macos_word_semantic_table.py`.
New tests: `tests/msoffice/test_macos_word_semantic_table_fixture.py`.

The production primitive and its 17-test file are unchanged; hashes below match the prior implementation report. No public/session forward, recovery implementation, shared fixture helper, quality code, index, commit or capability record was changed. No AppleEvents, osacompile, apps, UI or Word lease was used. Only pure Python/pytest and local synthetic XML/PDF validator test files were executed.

## Prepared native sequence

1. Explicit execute=True and a new output directory are required before any native work. CLI without --execute returns 2 without making the target directory. The exclusive Word lease must be granted separately by root.
2. Preserve source files/hashes and the installed Word dictionary before execution. Use the existing bound MacWordSession lifecycle, reviewed independent inventory, sentinel preimage and uncertainty-stop helpers. Seed one owned portrait document with page width600/height780 and 60pt margins; exact available width480pt.
3. Place Unicode PREFIX, a noncollapsed REPLACE range, Unicode SUFFIX and a pre-existing trailing 1x1 table. Preserve old table bookmark/content and save before.docx. Create a uniquely named unsaved synthetic sentinel and retain its actual inventory/name/state/content hash.
4. Restore only the bound document's explicit marker selection. Before table mutation, verify exact bound document/window paths, main story, range coordinates/text, one portrait section, page measurements and initial table count. This calls the original pure primitive unchanged.
5. Insert a 4x3 native table. Headers and cells cover CJK, emoji, None/True str conversion, left/center/right alignment and2.5pt first-line indent; inferred widths are135.927094400347/220.8142163128442/123.25868928680886pt (independent frozen6dd3a00 oracle values). Borders cover zero, .75,1.5 and arbitrary .5/2pt requests that freeze to .25pt. Repeat header=true and row splitting=false.
6. Capture native cell/width/border/repeat/split observations **before merges**, since native row/column access after vertical merging is not assumed. Then execute merges in the original order: vertical r2c3:r3c3, horizontal r4c1:r4c2. Preserve the primitive's explicit table-end CR and cursor.
7. Bookmark the actual returned table range. Native readback searches for exactly one table with matching range bounds; it never assumes count tables is the inserted table index. Verify prefix, suffix and the original trailing table remain exact, table text is complete without duplicated required content, two tables exist, one CR follows the new table, and the live bound cursor sits at tableEnd+1.
8. Save after.docx, export after.pdf, prove unsaved sentinel unchanged, close owned document, independently verify its close and the sentinel identity, then close only that synthetic sentinel. Reopen the saved source read-only through the normal session API; verify native table bounds/content/outside text again and source digest unchanged after close. Reopened selection coordinates are recorded but are not required to match the original live cursor; view state persistence is not the primitive contract.
9. Inspect saved XML and PDF; retain XML parts, PDF page PNGs, drawing observations, runtime scripts/logs, source/dictionary hashes, owned identity and final inventory. PASS_PRIMITIVE_AUTOMATED requires every explicit required gate, confirmed cleanup, source preservation and no uncertainty/error. It is not public semantic-method or full parity acceptance.

## Validators and failure behavior

Native style validator checks all20 records with exact row shape/order, strict boolean versus numeric types, cell content/indent/alignment, three widths and all seven border positions. Postmerge result validator requires exact bound path/start, valid increasing coordinates, two tables, actual tail CR/live cursor, preserved outside text and original table location, plus complete unique Unicode/data tokens. Missing, malformed, duplicated or mismatched ACKs are failures.

Saved XML checks require exactly two native tables; no drawing/pict/object/field substitute; exact tblGrid widths; table and header-bottom borders; four row flags; original-order topology represented by one vertical restart/continuation and one gridSpan2; all cells' text and paragraph indentation/alignment; preserved prefix/suffix paragraph structures, old table structure and section properties. Unknown added outside paragraphs also fail. Only volatile revision ids and bookmark nodes are omitted from the outside structural comparison.

PDF gate requires every fixture marker exactly once, all table text between PREFIX and SUFFIX, correct ordering of the old table afterward, in-page text, zero images and an actual480pt-wide table rule of .75/1.5pt inside the table region. This is stronger than text presence; manual PNG visual review remains pending.

The fixture never retries a failed constructor or performs recovery/repair. observe_native verifies session readiness and uses the existing structural execution path; any transport failure or malformed observation quarantines the run and rejects subsequent native observations. Native open/save/export/close failures similarly stop. Existing reviewed cleanup preserves the original error, performs at most one exact sentinel close, and stops inventory/close work after uncertainty. Local runtime-evidence copy failure is recorded separately without replacing the primary native error. Missing local dictionary fails before inventory or Word creation.

## Pure verification

- Initial RED:8 failures because the new fixture was absent.
- Initial GREEN:8 passed.
- Additional meaningful RED: duplicate native text, injected outside XML paragraph, misplaced PDF table text and missing runtime-retention helper produced4 failures/9 passes.
- Final GREEN:51 passed in0.44s across13 new fixture tests,17 unchanged primitive tests and21 existing quality fixture safety tests. Exit0. Five PyMuPDF/SWIG deprecation warnings came from the bundled dependency; no test failures.

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_macos_word_semantic_table_fixture.py tests/msoffice/test_macos_word_semantic_table.py tests/msoffice/test_macos_word_quality_fixture.py -q --tb=short
```

## Next review/native gate

Independent review should inspect constructor/readback syntax and exact noncollapsed table bounds before root runs the fixture. Once reviewed and with the exclusive lease, the proposed command is:

```sh
PYTHONPATH=. CLEAN_PYTHON fixtures/microsoft_parity/macos_word_semantic_table.py --execute --output NEW_ABSOLUTE_RUN_DIRECTORY
```

This report does not authorize or claim that command was run. Native first-run failures must remain raw evidence; do not weaken expectations or alter the frozen primitive solely to turn the fixture green. Confirmed Word behavior may require separately reviewed corrections. Landscape, long multipage repeat/split behavior, merge overflow, arbitrary replacement rollback, caption/notices/controller fallback, public semantic API and UI edit/Undo are outside this prepared portrait primitive gate.

## Source hashes

| File | SHA-256 |
|---|---|
| `fixtures/microsoft_parity/macos_word_semantic_table.py` | `5d8cb94250235e4349a0981fee1d1c166dc9cabc5c82354081f5f73f3cc61f0e` |
| `tests/msoffice/test_macos_word_semantic_table_fixture.py` | `9e5a553bb67d34e81463e931bd832a68b43a133d3fdffa90305fc0b20b2a14c8` |
| `skills/WPSComposer/scripts/msoffice/macos_word_semantic_table.py` | `9655e9af00366228bcca63063d2341d9687348fa2cd9e9663632ddee9c2d65ec` |
| `tests/msoffice/test_macos_word_semantic_table.py` | `01282417b44e52c9652dee111c061a6ce7f1b199665b9f65095c48e5290f4c30` |
