# Independent Word quality position-resolver review

2026-09-09. Read-only implementation review; the reviewer did not implement this helper change. No AppleEvents, dictionary compilation, native Office execution, production/test edits, staging or commit were performed. This report is the sole authored repository file in this review.

## Verdict

**No actionable finding in the bounded resolver correction.** The candidate is suitable for the parent's next source-bound native verification step. This does not establish complete `_target_at` success, field/bookmark native resolver acceptance, quality insertion/recovery, or four-method parity.

## Frozen source and scope

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_word_quality.py` | `724f193175390fbeedcfd32488c590e0785790b563a8a98b0e95f707e89981b9` |
| `tests/msoffice/test_macos_word_quality_native.py` | `d0a60d46649e30e1332eb7a2d8d6f7232306b0eb7f8771d9b2f7b93e41c42810` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `47b40d4eabc360fdfcb29244f072b00aa7b286586e95abf5dd848bfa1b83c59e` (unchanged) |

The initial index-to-working diff changed only the three `repeat-with-in` collection loops inside `_position_gate`: enumerate native table/field/bookmark ordinals and assign each native object before reading its bounds. While the parent prepared full31, the index advanced to the same candidate; the reviewer therefore independently repeated the comparison against the immutable old blob `8dd933a`, whose SHA-256 was checked against the prior helper freeze `992cb4decbd99b63d62815fb5ba537d3838077203f56fbd037483b1fde6b0c2e`. No index alteration was made by this reviewer.

An AST comparison confirms `_position_gate` is the only changed function/class, and all **8 position guard conditions** are unchanged, including the document-size f-string. The table comparison still rejects `start <= point + 1 and end >= point - 1`; the field comparison remains inclusive at the point; the crossing-bookmark test remains strict on each side. No coordinate clamp, new insertion position, weakened neighboring-table admission, reordered write, fallback, snapshot, title/dedup/cursor or transport behavior was introduced.

## Evidence reviewed

Read the retained `target-read-probe-01/report.json`, probe generator and generated native scripts, and the independent `target-read-probe-01-parent-cleanup/report.json`.

- At point 111, the original table iterator plus original compound comparison acknowledged ordinary read error `-2763` at `compound-comparison`.
- The ordinal resolver with that same compound comparison acknowledged `false` for both existing tables, followed by `done`; their measured bounds were 59:77 and 173:191. The production change adopts this resolver distinction, not the probe's separate atomic comparison variant.
- The raw probe remains **FAIL** because its runner wrongly expected the indexed-original mode to fail too. This is not rewritten as a successful complete probe. Its acknowledged successful branch is narrow native evidence for the resolver distinction.
- Parent cleanup is separate: exact owned path, saved state and original body hash were checked, exact-path/saved guarded close was recorded, final inventory was `[]`, and quarantine recovery was true. This review read those records; it did not repeat cleanup or make a current inventory claim.

The table native evidence supports the observed root cause. The field/bookmark changes are the parent-authorized same-pattern correction and match ordinal lookup already used by the snapshot compiler; they have no new native result in this patch and are not presented as independently proven by the table probe.

## Tests and boundary checks

The nine emitted-scalar cases use point 111: actual tables 59:77 and 173:191 are admitted, as are 100:109 and 113:130; 100:110, 100:111, 111:130, 112:130 and 100:130 are rejected. These cover both inclusive one-unit neighbors and overlap without modifying the production expression. They evaluate extracted scalar arithmetic only, not Word getter/reference semantics.

The other new tests protect the native resolver instruction shape, compare every emitted condition against the retained failing production script, and keep the probe FAIL and later cleanup facts distinct. These are useful source/evidence regression guards and do not replace a fresh native runtime test.

Fresh affected suite, using the supplied clean-dev interpreter with `PYTHONPATH=.`, ran the original independent log-failure regression plus Word session, quality, quality-native, degradation, recovery, numbering, fields and references tests. The filter was `-k 'not test_quality_scripts_compile_without_running_word'` to exclude all 3 dictionary compiler cases under the no-AppleEvent constraint.

Result: **503 passed, 3 deselected in 7.75s**. This is one more PASS than the implementation's 502 because the review included the unchanged external P2 regression. Retained output: `/tmp/word-quality-review-Gfitbe/word-position-review-green.log`. `git diff --check` passed for the changed helper/test paths. This was not a full repository run or a fresh compiler/native acceptance run.

## Remaining gates

The next source-bound native run must establish that production target snapshotting proceeds past the repaired guard and that the remaining dictionary properties return valid typed facts. Field/bookmark guard behavior, supported and rejected positions, table insertion/strict ACK, Delete restoration/fallback, surrounding rich formatting, M5 convergence, DOCX/PDF readback and actual UI edit/reopen remain separate unfinished gates. The previously fixed diagnostic-I/O P2 remains covered by the passing original regression and unchanged session hash.
