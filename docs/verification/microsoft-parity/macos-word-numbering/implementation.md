# Bound Word numbering slice

## Initial phase — 2026-09-09

Read AGENTS, regression guardrails, approved parity plan/progress, full frozen remaining-object contracts, and bound-selection-run-01. No `.codegraph` directory in this worktree. Production unchanged during initial phase.

23 RED tests recorded in `/tmp/word-numbering-red-01.txt`. Nineteen execute AST-extracted `6dd3a00` Writer methods against a controlled COM port, then compare the Mac orchestration through its native cursor port. Cases include source/fallback coercion, current nonterminal/noncollapsed selection, UTF-16 affixes, all sequence families, localized style fallback, invalid descriptors/missing keys/late bookmark errors, and native field/bookmark/paragraph failures. Additional RED tests cover exact public signature, numbering owner/category/ordinal snapshot resolution and acknowledged checkpoint tracking rollback. These tests are not native evidence.

## Implementation brief

- Add focused `macos_word_numbering.py` with a bound cursor transport and shared number shell. The direct method obtains `selection of boundWindow` and verifies its document identity and main story; never reads application-global selection or seeks document end. Initial text replaces selected content; later steps use explicit insertion coordinates and synchronize only the bound selection.
- Preserve frozen operation order and exception boundary: source + TAB precede descriptor validation; prefix precedes field insertion; bookmark validation follows fields; suffix follows bookmark; only right alignment + KeepTogether before CR. No OMath, table, font, indent, KeepWithNext or degradation notice.
- Validate native field count delta, type, exact allowed code forms, code/result bounds, field-end skip and code bookmark identity. Track each SEQ/STYLEREF independently as category numbering and owner or doc:native. Keep successful earlier fields tracked if a later ordinary validation fails. Unknown completion/ACK is quarantined.
- Extend existing field snapshot, mutation invalidation and append-only checkpoint tracking for numbering identities while preserving REF/INDEX behavior. Existing checkpoint remains plain integer and rejects mutation before its saved append boundary; no arbitrary middle/cell recovery claim.
- Native fixture planned after root serial lease: synthetic active sentinel + separate bound document; middle/noncollapsed selection, global/chapter localized Heading 1, emoji affixes, same-owner repeated fields, field update and number-only bookmark/REF, save/close/reopen/DOCX + development PDF, append rollback and uncertainty injection. New OMath/figures/tables excluded.

Root granted production mutation after initial brief on 2026-09-09. Native execution remains ungranted.

## Implementation checkpoint

Implemented focused helper + matching session forward and numbering-only tracking changes in fields/recovery. Initial 23 RED → GREEN. Expanded native-port ACK/uncertainty/guard tests and compile-only coverage, then guarded fixture 2 RED → GREEN. Current numbering tests: 44 PASS. Combined numbering/session/fields/references/recovery + frozen M3/M4 Writer tests: **363 PASS in 6.99s**, `/tmp/word-numbering-focused-02.txt`. No native run or UI acceptance yet.

`fixtures/microsoft_parity/macos_word_numbering.py` is ready with explicit `--execute`, new-output guard, all package source hashes, separate active synthetic sentinel, noncollapsed middle selection, global/chapter fields, same-owner ordinal readback, number-only REF, confirmed append rollback after late bookmark validation, save/close/reopen, XML native structures and development PDF checks. Unknown native completion retains runtime/sentinel; root owns exact recovery.

Production frozen for heading probe and independent numbering review:

- numbering: `2048cdf59b8f86a50ff7db74f5e9c5ad0c7e8f39d01f31ec11d69966e52ed9ea`
- session: `df56137d38d793c878659144676b7b1d23ae3f16b34adbee2304ee06399171d9`
- fields: `5785fa66ac4aa24276280996f46286b685626f6c1dd4ade7fc840bfde6b052f4`
- recovery: `91d9a954ce7ae9210178998a972023f6bce4928ba4e602207ce63823981b5bdc`

## Review fix wave — frozen for re-review

Independent review reported three P2 issues (see `word-numbering-review.md`): cross-story identity collisions, native-order rather than creation-order semantic ordinals, and malformed ACK bypass. Reproduced original 8 RED in `/tmp/word-numbering-review-confirmed-red.txt`, then expanded tests showed 19 RED. Fixed all: tracked creators bind exact `story:main text/chain:1`; tracked insertion ordinals are allocated separately while native topology remains in document order; nested ACK containers and scalar integer/boolean types are validated before cursor or tracking commit. Existing REF/INDEX controlled native rows were corrected from shorthand `main` to the actual story identity.

Root additionally identified unconditional quarantine for pre-submission failures. Real `_execute` transport tests gave 3 RED (`/tmp/word-numbering-presubmit-red.txt`); the cursor now uses the existing actual-submission marker directly, preserving pending heading and observation on script-write/read-only failure, retaining the established session deadline quarantine, and still quarantining submitted ordinary native errors or uncertain completion. No native process was called by these tests: subprocess is the external boundary double.

Final focused result: **387 PASS in 6.75s**, `/tmp/word-numbering-fixed-final.txt` (68 numbering tests plus session/fields/references/recovery and frozen M3/M4 Writer tests). Native fixture seed/readback commands also compiled successfully without execution. Native fixture remains unrun.

Current frozen production SHA-256:

- `macos_word_numbering.py`: `21307cfa3aeb8e74e0a12139b8b32573a9c984344568488984079bb0948ff31d`
- `macos_word_session.py`: `df56137d38d793c878659144676b7b1d23ae3f16b34adbee2304ee06399171d9`
- `macos_word_fields.py`: `ee51806876fcd859482b97963c1de57e077f5411356ddf85261e10bb09d94b6e`
- `macos_word_recovery.py`: `91d9a954ce7ae9210178998a972023f6bce4928ba4e602207ce63823981b5bdc`

Root requested yielding to free a review slot. Next: independent scoped re-review, then root-granted native execution of `fixtures/microsoft_parity/macos_word_numbering.py --execute --output NEW_PATH`. No further production edits or native work until root resumes this worker.

## Native run-01 — PASS, lease released

After independent re-review reported clean 395 PASS at the frozen production hashes, root granted one exclusive prepared fixture run. Executed exactly once:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python fixtures/microsoft_parity/macos_word_numbering.py --execute --output docs/verification/microsoft-parity/macos-word-numbering/run-01
```

Exit 0; report `run-01/report.json` is PASS with **10/10 true checks**. Native results: original noncollapsed middle `REPLACE` selection replaced by `x=1`; later middle equations use fallback coercion and native localized `STYLEREF "标题 1"` plus SEQ; bookmark visible values exactly `1`, `1-1`, `2` excluding prefix/suffix; all three REF values match after update and reopen. Category is numbering and shared-owner SEQ_EQ ordinals are 0 and 1. Late invalid bookmark occurs after a confirmed extra field; append checkpoint rollback restores previous native readback and exact tracked identities. DOCX save, PDF export, owned close and read-only reopen completed; native/XML evidence has three SEQ, one STYLEREF and three REF with no table, drawing or OMath.

Independent post-run hash verification found all **125** live/retained source pairs equal their manifest hashes, and all **3** recorded artifact hashes (DOCX, PDF, document.xml) correct. Initial and final native inventories are both `[]`; unchanged synthetic unsaved sentinel was closed after owned close; `remaining_sentinel=null`; no error or cleanup failure. Offline `pdftoppm` render of the PDF was viewed: one page, legible native numbering/REFs, expected right-aligned formula paragraphs, emoji and CJK suffix visible, no clipping. Preview is `/tmp/word-numbering-run01-preview.png`.

No production/fixture edits during the run, no retry, no quarantine bypass, no UI interaction, installation, commit or push. Word native lease explicitly released to root. This closes the bounded native numbering fixture only; reverse-position stable ordinals and header collision are proven by controlled tests, not a separate native case. Actual edit/Undo/UI acceptance and broader heading/table/figure/OMath parity remain distinct gates.
