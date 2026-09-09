# Word quality candidate independent implementation review

2026-09-09. Read working files in `.worktrees/office-description`; root index was not changed. This is a review of the staged candidate's implementation, not four-method parity or native acceptance. No production edits, native Word execution, UI interaction, staging, commit, publication or installation occurred. Local `osacompile` was exercised by the existing compiler tests; it did not run Word.

## Verdict

**Request changes: one reproducible P2 finding.** Fix the reservation submission/uncertainty boundary before the parent starts production-source native acceptance. The other explicitly unfinished location/recovery rows remain open. Passing the current suite does not close them.

## P2 — Quarantine a submitted reservation when diagnostic log writing fails

Location: `skills/WPSComposer/scripts/msoffice/macos_word_quality.py:181–189` (especially its `except NativeWordError` restriction). Supporting transport location: `macos_word_session.py:340`.

`_reserve_native` relies on `NativeWordError.diagnostic_path` to determine whether the bookmark-only batch was submitted. But the real `_execute` writes the process result to its `.log` after `subprocess.run` returns, outside its submission exception handler. If that write raises `OSError` (for example, disk full), the successful/failed native result never reaches the ACK parser and is not a `NativeWordError`. The public reservation wraps it as `quality anchor insertion failed`, leaves the Python cursor/title uncommitted, and leaves `_quarantined == False`. The native bookmark may already have been created or replaced. Further calls are still permitted, including re-reservation with a potentially different selection and publication after an unknown reservation result. This violates the candidate's declared rule that submitted unknown completion retains evidence and quarantines, even though reservation must preserve content-field observations.

The root cause is shared `_execute`, not only this new reservation caller. Its result-log write precedes native nonzero-result classification, and log writes in its exception branches can also mask the primary timeout or `KeyboardInterrupt`. Repair this narrow shared transport boundary once, rather than only catching quality `OSError`: use a truthful submission signal for **all** post-submission exceptions, ensure at least in-memory quarantine even when log/quarantine-file writing fails, preserve the original primary failure/cancellation, and permit no subsequent native cleanup. Do not claim diagnostic persistence when writing failed. Preserve current before-launch behavior: local script-write failure must not be treated as a submitted bookmark, and reservation must not invalidate the field-topology cache just to reuse a content-mutation marker. Coordinate the shared transport edit with the parent.

Review-only regression source: `/tmp/word-quality-review-Gfitbe/test_reservation_log_failure.py`.

Retained failing result: `/tmp/word-quality-review-Gfitbe/reservation-log-failure-pytest.log`.

Reproduction from the worktree:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest /tmp/word-quality-review-Gfitbe/test_reservation_log_failure.py -q
```

Observed **1 failed** at `assert session._quarantined`: actual `False`. This uses the real public wrapper and real `_execute`; only `subprocess.run` returns a successful reservation ACK without launching Word, and `.log` writing raises an injected `OSError`. Earlier assertions verify one submission, no committed cursor, and unchanged field observations. It does not claim an actual disk-full incident or native bookmark mutation was reproduced.

## Reviewed behavior and fresh validation

- Confirmed frozen hashes: helper `992cb4decbd99b63d62815fb5ba537d3838077203f56fbd037483b1fde6b0c2e`; session `6a52b745e820735ccdc17defdde477f4f4d315d310d97ab4a557aa57912305fc`. `git diff -- macos_word_session.py` relative to root index contains only the four lazy forwards.
- Read the implementation report, implementation brief, anchor design and M5 regression guardrails. Compared public control flow against the frozen Writer oracle guarded by the AST fingerprint.
- Mapping/dedup order, first-title behavior, ignored object fallback/node_id, serial effects, original error messages, required signatures, and default bookmark versus numeric cursor distinctions are preserved by the reviewed Python code and differential tests. No additional Python-contract finding was established.
- Exact Foundation guards precede the relevant read/write logic, and table activation is followed by document/window/selection ownership checks. The target descriptor includes session/binding identity and immutable preimage; a mutation re-resolves bookmark position and compares the native preimage before construction.
- Nonterminal paragraph-start, table separation, crossing field/bookmark and drawing restrictions are emitted before table construction. The fallback paragraph-format restriction precedes range text assignment. Consecutive upserts and terminal/interior placements remain deliberately unavailable and cannot be called complete.
- Content mutations cross the existing real pre-subprocess invalidation marker; local preflight/script failures preserve observations, while ordinary errors, timeout and malformed completion invalidate them. Valid ACK alone commits structural flags/cursor/dedup. The finding above is specific to bookmark-only reservation, whose separate path does not have equivalent coverage for post-launch non-NativeWordError failures.
- Recovery retains the same-batch new-table reference, checks projected preimage before Delete, performs only one Delete, then requires a zero-delta preimage match. Only a strict restoration ACK selects one range fallback. Unknown/malformed restoration does not select fallback. These are code/transport observations, not proof that the native recovery branches work in Word.

Fresh command used the supplied clean-dev interpreter with `PYTHONPATH=.` and these existing test files: `test_macos_word_quality.py`, `test_macos_word_quality_native.py`, `test_macos_word_degradation.py`, `test_macos_word_recovery.py`, `test_macos_word_numbering.py`, `test_macos_word_fields.py`, and `test_macos_word_references.py`.

Result: **409 passed in 11.06s**, including all three target/table/range `osacompile` tests. This was not the repository-wide suite. The additional review regression above remains RED on the frozen candidate.

## Explicit remaining gaps

1. There is no production-source native run of these four methods. Expanded Word dictionary properties, actual field/list/style/header values, UTF-16 geometry, title-once/multiple notice behavior, native name/ordinal lookup, bookmark affinity and save/reopen semantics remain unverified. Compilation alone establishes syntax, not runtime property availability.
2. Actual middle-table failure before identity, after fill, after style, during Delete and unknown completion remain unexecuted for this helper. Pure tests verify Python selection of fallback and compiler ordering; they do not execute the AppleScript `try`/recovery control flow. A successful primitive fixture is not production implementation acceptance.
3. The snapshot hashes selected formatting properties, not every possible formatting fact. For example, existing table rows include bounds/cardinality/text/split flags but no column widths/cell margins/borders; character rows omit superscript/subscript/spacing, and paragraph rows omit tab stops/page-break-before. Header/footer layout hashing records text/fields but not complete rich formatting. There is no reproduced native corruption finding here, but equality of this schema alone must not be described as exhaustive rich-document restoration. Source-bound native fixtures should seed relevant surrounding formats and independently inspect saved OOXML/PDF; expand the guard or close unsupported document features where preservation cannot be established.
4. Interior/terminal/cell/adjacent-table targets, consecutive distinct upserts, crossing objects, drawing-rich and larger documents remain required unfinished capability rows. The staged position restrictions are safe staging limits, not a replacement for frozen Writer semantics.
5. Native M5 typography/numbering/field convergence, DOCX reopen and PDF inspection, actual edit → Undo → explicit Save/discard → close → exact-path reopen, and nonempty sentinel/final inventory remain the parent's serialized native/UI gate. This review certifies none of those outcomes.

Historical memory was used only to orient the independent native-versus-test acceptance distinction; all implementation claims above come from the current candidate and fresh review tests.

## Independent P2 re-review — 2026-09-09

**Original P2 is resolved at this revised source freeze.** No additional actionable finding was established in the narrow shared diagnostic/submission repair. This supersedes the original request-changes verdict for that P2 only; it does not close any four-method native capability or rich-format restoration gap above.

Verified hashes before re-review:

| File | SHA-256 |
| --- | --- |
| `macos_word_session.py` | `47b40d4eabc360fdfcb29244f072b00aa7b286586e95abf5dd848bfa1b83c59e` |
| `macos_word_quality.py` | `992cb4decbd99b63d62815fb5ba537d3838077203f56fbd037483b1fde6b0c2e` (unchanged) |
| `test_macos_word_session.py` | `9de0da3998d56a3b0770b556bec0353469d9be9212f5b1e0d847fcf42c3fbc49` |
| `test_macos_word_quality_native.py` | `5cb5175c484d01df062e04e4ead786ec23d9e0832bfeb765f2fc38c07cd35cea` |

The original independent regression source `/tmp/word-quality-review-Gfitbe/test_reservation_log_failure.py` was kept unchanged and now passes. The original RED log was preserved separately. Fresh re-review command ran that regression, `test_macos_word_session.py`, and the seven originally reviewed quality/degradation/recovery/numbering/fields/references modules with the supplied clean-dev interpreter and `PYTHONPATH=.`: **492 passed in 18.92s**. Complete output is retained at `/tmp/word-quality-review-Gfitbe/word-transport-rereview-green.log`. This includes three dictionary compiler tests; no Word application execution occurred.

The implementation and meaningful fault-injection tests establish the following bounded conclusions:

- `_retain` sets `_quarantined` and `_retain_evidence` before attempting fallible lock persistence, and records secondary failures without replacing the primary exception. `_quarantine_location` omits a failed/partially written record; a second outer retention cannot clear that failed-write state merely because a partial file exists.
- Successful native execution plus failed log writing yields a quarantined error instead of its ACK. A new cancellation during successful result handling is caught by the post-submission guard, quarantines, and propagates. Ordinary nonzero native errors keep their existing code, and native/transport timeout keeps `NativeWordTimeoutError`, when auxiliary persistence fails.
- When a primary cancellation/timeout already exists, secondary diagnostic/quarantine `SystemExit` or I/O failure cannot replace it. Original process cancellation objects propagate. Failed diagnostic paths are omitted; when a path is advertised, the reviewed write path has completed rather than only been planned.
- The actual content-topology invalidation marker still sits immediately before process submission. Script-write/chmod failures do not submit, quarantine, or invalidate observations. Existing bound-session deadline quarantine remains a separate local rule and preserves observations. Bookmark-only reservation uses no content marker; the revised permanent regression proves no cursor/title/heading/structural commit and no second submission on retry.
- The quarantined `close` branch raises before its native close command and keeps staging evidence. Missing on-disk quarantine still cannot prove cross-process exclusion; the repair claims in-memory isolation only when persistence fails.

Only this review document was changed during re-review. Word, Excel, PowerPoint, tests and index were not edited. The parent's independent Excel/PPT review remains separate; this Word re-review does not certify the reviewer's own Excel/PPT implementation.
