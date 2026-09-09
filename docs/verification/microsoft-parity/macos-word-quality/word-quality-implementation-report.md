# Word quality implementation candidate — freeze for review

2026-09-09. Worktree: `.worktrees/office-description`, branch `codex/microsoft-parity`.

## Status and ownership

The four frozen Python contracts are implemented in one focused `macos_word_quality.py`, with four lazy `MacWordSession` forwards. This is a **staged implementation candidate**, not full four-method parity or native acceptance. No production files outside the new helper and those four forwards were edited by this worker. Existing fields/recovery/numbering/session changes were preserved. No staging, commit, push, install, Word execution, UI action or native artifact generation occurred.

The existing 39 contract tests were preserved unchanged: observed baseline **32 RED + 7 oracle PASS**, then **39 PASS**. New boundary/compiler tests observed missing-helper RED and subsequent focused RED cases for empty-reservation bookkeeping, native serialization, script syntax, reservation failure quarantine and section/list snapshot identity. Final targeted suite: **409 passed in 7.06s**. `git diff --check` passed. This suite is not the repository-wide acceptance run.

Command (the shared worktree has no `.venv`; root checkout's Python was used):

```sh
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest \
  tests/msoffice/test_macos_word_quality.py \
  tests/msoffice/test_macos_word_quality_native.py \
  tests/msoffice/test_macos_word_degradation.py \
  tests/msoffice/test_macos_word_recovery.py \
  tests/msoffice/test_macos_word_numbering.py \
  tests/msoffice/test_macos_word_fields.py \
  tests/msoffice/test_macos_word_references.py -q
```

## Frozen candidate hashes

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_word_quality.py` | `992cb4decbd99b63d62815fb5ba537d3838077203f56fbd037483b1fde6b0c2e` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `6a52b745e820735ccdc17defdde477f4f4d315d310d97ab4a557aa57912305fc` |
| `tests/msoffice/test_macos_word_quality.py` | `f0e433610ada1542af480166cbbb853fc89d505b84bfff73ab6e052ae8026781` |
| `tests/msoffice/test_macos_word_quality_native.py` | `c27a008d65024d90f7faf241e1d9a22de65aafcd2c70977c21de85520af37c37` |

The session hash includes prior independently owned changes. This worker's session contribution is only the four forwards immediately before `degradation_checkpoint`.

## Implemented boundaries

- Reservation reads the actual `selection of boundWindow` End, verifies main story and exact bound document/window/selection document path (or unsaved name/path with available window id), creates the fixed empty bookmark, and validates empty geometry plus body hash/End preservation. It neither clears pending heading nor invalidates field observations. After a submitted failed reservation, private transport diagnostics establish uncertainty and quarantine; local script-write failures do not.
- Python mapping, error timing, serial effects, first-title behavior, redaction, object/dict distinction, fixed bookmark versus numeric cursor, required signatures, page coercion and ignored `node_id` are differential-tested against the frozen Writer oracle.
- `QualityTarget` is an immutable session-and-binding descriptor with the native resolved position and immutable preimage. Numeric and native bookmark resolution are separate. Explicit bookmark lookup reads its first paragraph End; falsy/default lookup reads bookmark Start. No clamping, document-end substitution or rehydration of Python dedup state occurs.
- Native mutations revalidate binding/position/preimage in the same batch before construction. Exact comparisons use Foundation `NSString.isEqualToString`, with both structured operands serialized by native `jsonRows`; Python JSON spelling is not compared to native spelling.
- The preimage records prefix/suffix text hashes; every paragraph's text and character/paragraph-format hashes; fields' exact code/result hashes, coordinates/type/lock state; old table bounds/cardinality/text/row flags; all bookmark names/bounds/projected text; and a native hash of section page setup, header/footer text and field identities, list-template levels and style descriptions. No private document body is returned in these snapshots.
- Middle success checks the actual ordinal (not total table count), UTF-16 table/cell/display bounds, observed CR/BEL terminators, document delta, style/paragraph/row ACK and unchanged projected preimage. Its geometry is scoped to the measured paragraph-boundary primitive; append ACKs remain untouched.
- Ordinary table creation/fill/style failures hold the newly created table reference in the same batch. Before Delete, the projected unchanged preimage and table geometry must match. Delete touches only that reference; a complete zero-delta snapshot must then match before returning the restoration ACK. Failure with no table identity must still prove the complete unchanged preimage. Unknown completion, unsafe partial state, Delete failure, mismatched restoration or malformed ACK quarantine without another native call.
- A valid restoration ACK permits exactly one range fallback batch. It revalidates the original preimage again, validates styled actual range bounds and projected preservation, and returns the existing immutable degradation DTO. No generic Undo, suffix rewrite, CR repair or append rollback is used.
- Content submissions reuse the already-reviewed `_field_topology_mutation_pending` marker consumed in the real session `_execute` immediately before `subprocess.run`. Before-launch script/deadline/read-only failures preserve observations; all actual content submissions invalidate them. Structural success flags and pending-heading cleanup occur only via acknowledged `_committed_range`, not on malformed/uncertain completion. A restored failure conservatively leaves observations invalidated.

## Explicit staging gates and remaining required rows

These gates are present before the respective write. They are **incomplete required capability rows**, not deletion of frozen requirements:

1. Table insertion currently accepts only a nonterminal main-body paragraph start strictly between zero and `document End−1`, separated by more than one coordinate from any old table. Paragraph interior, nonempty terminal paragraph end, native document End, cell/nested table and adjacent-table/consecutive-upsert cases are closed. Reservation itself uses selection End, including a noncollapsed selection, but production reservation still needs its own source-bound native matrix.
2. Crossing fields and bookmarks are closed. A bookmark beginning exactly at the insertion point is projected with the affinity observed in the feasibility run: a collapsed anchor stays fixed; a nonempty bookmark starting there may contain the inserted table while its original text remains unchanged after projection. Every other differing object is rejected.
3. Main/header drawing objects are closed; full character-format snapshotting caps document End at 10,000 native units. Larger documents, drawing-rich documents and broad performance acceptance remain unverified.
4. Range fallback shares its final paragraph with the original suffix when display does not end in CR. Applying block paragraph style can therefore change the suffix paragraph. The current fallback write is closed unless the suffix paragraph already has the required space-before 0, space-after 3, keep-together true and body outline level. The snapshot clips the original suffix from the combined range and still verifies its formatting after insertion. This does not prove all frozen fallback placements and does not append a separator to make them pass.
5. Native ordinal/case-variant bookmark lookup, explicit spanning bookmark position, default bookmark after upserts, sequential notice cursor behavior, repeated bookmark inserts and M5 title/body/list/field convergence require serialized production-native execution. Passing the Python oracle does not establish those Word behaviors.
6. The full snapshot compiler reads additional Word dictionary properties beyond the narrow original feasibility fixture. Three complete paths (target, table/recovery, range) compile via local `osacompile`; compilation did not run Word and is not evidence that these properties return the expected typed values at runtime. Native preflight may fail closed until source-bound evidence resolves any dictionary/runtime mismatch.
7. Failure matrix still needs actual before-identity, fill, style, Delete and unknown-completion native evidence. The subprocess tests establish invalidation/quarantine and no second submission after malformed/uncertain results; they do not execute the AppleScript recovery branch in Word.
8. Complete DOCX reopen, PDF layout inspection, native style/field/object preservation, actual edit → Undo → explicit Save/discard → close → exact-path reopen and nonempty sentinel inventory belong to the parent's serialized lease. None are claimed here.

## Review request

Candidate is frozen at the hashes above. Request independent implementation review before the parent schedules the narrow owned-document native run. Review should prioritize projection completeness, actual runtime property support, strict state/ACK validation, reservation-only bookkeeping and fallback preflight scope. Do not mark the four direct method capability rows complete based on this report.

## P2 transport diagnostic-boundary revision — current freeze

The independent review's original file and retained failure remain unchanged at `/tmp/word-quality-review-Gfitbe/test_reservation_log_failure.py` and its sibling log. Reproduced **1 RED** against the first freeze: successful subprocess ACK followed by `.log` `OSError` left reservation uncommitted but the session usable.

The parent authorized a narrow shared-transport correction. The new quality production helper remains byte-identical. `macos_word_session.py` now isolates all post-submission result/diagnostic processing, and its shared `_retain` commits in-memory quarantine before attempting fallible persistence. Diagnostic writes return a path only when they succeed; successful native results with failed logs are quarantined instead of returning their ACK. A native nonzero result retains its original error classification, `subprocess.TimeoutExpired` retains the existing public timeout mapping, and original `KeyboardInterrupt`/`SystemExit` instances propagate even when log or quarantine writes fail. A new cancellation during otherwise successful result processing also quarantines and propagates.

Diagnostic/quarantine write failures are recorded only by exception class in memory. Failed log paths are omitted from the raised diagnostic; `_quarantine_location` advertises only an existing record without a known failed write. A partial quarantine file from an earlier failure does not become a successful persisted record merely because a later outer `_retain` sees it. If disk persistence is unavailable, the current session is still isolated, but no successful persistent quarantine is claimed. Cross-process protection cannot be inferred from an unavailable persisted record.

Before-submission script-write/chmod failures remain unquarantined and retain field observations. The existing bound-session deadline rule remains unchanged: `_remaining` itself can quarantine an expired session, but the new result/diagnostic handling does not label that as a submission or invalidate field observations. Bookmark-only reservation uses no content-mutation marker. Content submissions still invalidate field observations at the real subprocess boundary even if logging subsequently fails. Four quality forwards and the reviewed numbering/fields/recovery changes remain preserved.

Added permanent tests in `test_macos_word_session.py` cover success/ordinary-error/native-timeout/transport-timeout/cancellation against log failure, quarantine failure and both failures; secondary diagnostic cancellation preserving the primary native error; partial quarantine re-retention; submission-free script/chmod/deadline failures; and content-topology invalidation on post-submission log failure. The quality-native regression additionally verifies no reservation cursor/title commit, unchanged field/pending-heading state, and no second submission on retry.

Fresh result: **492 passed in 13.24s**, using the original review repro plus `test_macos_word_session.py` and the seven files in the initial command above. `git diff --check` passed. This remains a focused suite, not full repository/native acceptance. No staging, commit, Word execution, UI, publication or installation occurred.

Current source/test hashes (supersede the session and quality-native hashes in the initial freeze table):

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_word_quality.py` | `992cb4decbd99b63d62815fb5ba537d3838077203f56fbd037483b1fde6b0c2e` (unchanged) |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `47b40d4eabc360fdfcb29244f072b00aa7b286586e95abf5dd848bfa1b83c59e` |
| `tests/msoffice/test_macos_word_quality_native.py` | `5cb5175c484d01df062e04e4ead786ec23d9e0832bfeb765f2fc38c07cd35cea` |
| `tests/msoffice/test_macos_word_session.py` | `9de0da3998d56a3b0770b556bec0353469d9be9212f5b1e0d847fcf42c3fbc49` |

Frozen for the original reviewer to re-review. All location/native/recovery/format-preservation limitations remain open. In particular, the initial snapshot is a hash of selected native facts; it must not be described as exhaustive rich-document restoration without the independent native OOXML/PDF evidence identified by review.

## Native target-read failure and ordinal resolver fix — current freeze

The first production-source quality run remains **FAIL** at `first-upsert-public`. Only empty reservation was confirmed. Its retained script `run-01/native-runtime/0efcad4e246b4b5d9c9dcf3ff66d37fc.applescript` failed at character offsets `2723:2739` (line 33), inside the table neighbor guard. Those offsets cover `qualityPoint + 1`; the script had set `qualityPoint` to 111. Its preexisting table positions were 59:77 and 173:191, so the intended adjacency rejection was false. The entire failed target-read script contained no table creation, content assignment, bookmark creation or Delete command. Parent cleanup and independent empty inventory are separate evidence in `run-01-parent-cleanup`.

Under a subsequently granted native lease, `target-read-probe-01` opened a fresh owned **read-only copy** of `run-01/before.docx`, with empty starting inventory and unchanged native baseline End 220/body hash. It recorded these acknowledged read-only outcomes:

| Resolver/comparison | Native outcome |
| --- | --- |
| Original `repeat with qt in tables of boundDoc` plus original compound comparison | Ordinary read error `-2763`, at `compound-comparison` |
| Native ordinal table lookup, atomic bounds/getters/comparisons | Tables 59:77 and 173:191, neighbor conditions false |
| Native ordinal table lookup, original compound comparison preserved | Both table comparisons returned false and completed |

The third outcome disproved the probe runner's overly narrow expectation that the compound expression itself would fail under both resolvers. The runner therefore preserved **FAIL**, stopped before further calls and quarantined the copy. That historical result was not rewritten. Parent cleanup independently verified the exact owned path, saved=true and original body hash `9e47b6c699d1b28685dd568fac6d75cb875b18c85f53ba7e7b3b7f895168a15b`, performed exact-path/saved guarded close, observed final inventory `[]`, and recovered quarantine; see `target-read-probe-01-parent-cleanup/report.json`. The successful ordinal-only comparison is narrow native evidence for the resolver distinction, not a successful complete quality run.

The parent then authorized a resolver-only production fix and explicitly extended the same collection-reference correction to the other two loops in `_position_gate`. The gate now obtains each table, field and bookmark through its native ordinal before evaluating the existing properties. **Every admission/rejection condition is byte-for-byte unchanged** against the retained run-01 script. No fallback, snapshot, style, title, dedup, cursor or transport behavior changed in this revision. Fields and bookmarks had the same `repeat-with-in` item-reference form and the snapshot compiler already used ordinal lookup; they received script-contract coverage, but no new native field/bookmark comparison is claimed.

Observed test progression: original table-resolver regression **1 RED**, then the authorized field/bookmark resolver contracts **2 RED**, followed by GREEN. Nine scalar boundary cases verify the inclusive one-unit table-neighbor rejection, including real seeded table bounds and exact lower/upper threshold neighbors. A retained-evidence regression distinguishes the original read error from the ordinal-only success and keeps the probe's raw FAIL plus parent cleanup separate. A contract check compares all emitted position rejection conditions to the source-bound failed script.

Final focused command is the eight test files from the P2 revision excluding the external `/tmp` repro, with `-k 'not quality_scripts_compile'`. Result: **502 passed, 3 deselected in 9.23s**. The three local dictionary compiler tests were intentionally not run during this no-AppleEvent revision; this is not a fresh compiler or native acceptance claim. `git diff --check` passed. No additional AppleEvent, full quality fixture rerun, staging or commit occurred after the probe lease was returned to the parent.

Current hashes:

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_word_quality.py` | `724f193175390fbeedcfd32488c590e0785790b563a8a98b0e95f707e89981b9` |
| `tests/msoffice/test_macos_word_quality_native.py` | `d0a60d46649e30e1332eb7a2d8d6f7232306b0eb7f8771d9b2f7b93e41c42810` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `47b40d4eabc360fdfcb29244f072b00aa7b286586e95abf5dd848bfa1b83c59e` (unchanged) |

Frozen for review and the parent's full31 snapshot. Broader layout snapshot properties and the remaining placement, insertion/recovery, roundtrip and UI matrices remain unverified. The four quality methods cannot be marked complete from this resolver repair.
