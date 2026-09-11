# Mac Word paragraph rule implementation

Status: implementation prepared; public forwarding and native production acceptance pending. This implementer did not perform independent review, native Office execution, UI interaction, personal installation, commit, or push. Inline horizontal lines and WordArt remain untouched and unsupported by this slice.

## Scope and behavior

New focused implementation `skills/WPSComposer/scripts/msoffice/macos_word_rules.py` provides `add_paragraph_horizontal_line(session)`, returning `None` like frozen `WriterComposer.add_paragraph_horizontal_line(self)`.

The operation binds through the existing session transport and appends at the exact terminal body insertion point. A nonempty terminal paragraph gets a separator so existing text/formatting is not turned into the rule paragraph. It inserts a single space plus paragraph mark, applies a centered native single .75pt C0C0C0 bottom border, and clears only the following paragraph's bottom border. It does not select the active document, invoke GUI/clipboard operations, reset formatting, force Normal style, or zero following paragraph spacing. Native ACK checks inserted range/paragraph count deltas, complete rule/following text, unchanged prefix text, border dimensions and color, clean following border, and following style/indent/spacing/alignment preservation around border application.

The existing `_execute_topology_mutation` invalidates field topology at actual subprocess submission. Per root's explicit scope decision, no new shared submission hook was added: `_pending_heading` is cleared and `_structural_changed` is set only after a valid complete operation ACK. An execution error after potential native mutation retains its original typed error/diagnostic and quarantines the session; an invalid semantic ACK retains a private JSON diagnostic and raises `NativeWordError(NATIVE_WORD_QUARANTINED)`. Thus unacknowledged state cannot continue using old positional targets. Preflight rejection or script-write failure before submission preserves the cached observations. `macos_word_session.py` is unchanged by this implementer.

## Validation

- RED: 37 tests failed because the module/guarded fixture did not exist (`/tmp/word-paragraph-rule-red.log`).
- GREEN: 38 passed, including the subsequently added syntax-only dictionary compilation check (`/tmp/word-paragraph-rule-green.log`).
- Combined regression: **269 passed in 10.77s**, covering the new rule tests, existing feasibility-probe tests, session, fields, references and recovery (`/tmp/word-paragraph-rule-combined.log`).
- Tests exercise the real MacWordSession `_execute` and topology boundary; only the external `subprocess.run` AppleEvent execution is doubled. The syntax test executes `osacompile`, not Word AppleEvents. Coverage includes exact document binding, deadline propagation, empty/nonempty terminal cases, malformed/truncated/duplicate/ill-typed ACKs, prefix/range/paragraph mismatch, following formatting mismatch, pre-submit state preservation, and native failure quarantine retaining the primary diagnostic.

Command:

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_rules.py tests/msoffice/test_macos_word_rules_probe.py tests/msoffice/test_macos_word_session.py tests/msoffice/test_macos_word_fields.py tests/msoffice/test_macos_word_references.py tests/msoffice/test_macos_word_recovery.py -q
```

## Parent integration patch (not applied)

Add only this forward within `MacWordSession`, following the surrounding lazy-import convention:

```python
    def add_paragraph_horizontal_line(self):
        from .macos_word_rules import add_paragraph_horizontal_line
        return add_paragraph_horizontal_line(self)
```

Then run the guarded fixture under an exclusive native lease. The fixture refuses an existing output directory and refuses native execution without `--execute`. It also refuses to proceed to the native session if the public forward is absent. It snapshots/hash-binds implementation, session, writer baseline and helper source; retains raw runtime scripts/errors; creates a nonempty unsaved synthetic sentinel; calls the actual public method on a styled, nonempty existing paragraph; appends following text without resetting formatting; saves DOCX and PDF; closes and reopens read-only; checks source hashes, persisted native border and following style/spacing, prefix paragraph XML, PDF text/rule drawing, sentinel state, and independent before/after inventory. A cleanup failure remains separate from the primary failure and blocks PASS. Quarantine suppresses further native cleanup/inventory.

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_paragraph_rule.py --execute --output docs/verification/microsoft-parity/macos-word-paragraph-rule/run-01
```

## Evidence boundary and open gates

The previously accepted native `macos-word-rules-probe/run-02-paragraph` establishes the primitive border insertion. It is not a run of this new public implementation. No source-bound native evidence for this module exists yet. In particular, following-format readback additions and the nonempty terminal separator require the new guarded native run; dictionary compilation is not runtime API proof. The fixture has not been run against Word. Parent must independently review the implementation, integrate the forward, execute the fixture under a lease, inspect rendered PDF, and perform UI edit/undo/save/reopen acceptance. The full branch suite and final capability/API documentation remain parent-owned gates.

## Source hashes at handoff

- module `e807932b76daf4ac7dd307088c392bcf6136bf72b64fd896a1840155a12af7ec`
- tests `d0b0578709535ab30605b63194559d2f934a6fb3494cc2ec74fa4c99f4d13ea4`
- fixture `225b507f393817ba5910cf10c756e8a05eb25bc9cbe4cb7ac69cddc097651268`
- shared session observed `a6f39d1191a7df557a7c911904e598bb5fa538430314a84afc9aeef711e0d534` (not edited here)
- frozen writer baseline `b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205`

## P2 revision — current-paragraph parity and independent sentinel lifecycle

This addendum supersedes the earlier separate-rule-paragraph behavior and three-paragraph fixture expectations. The original RED/GREEN history above is retained rather than rewritten. Both findings in `word-paragraph-rule-review.md` were implemented under root's explicit frozen-baseline ruling; independent re-review remains required.

The compiler now follows `writer.py:1173`: obtain the current terminal paragraph, center it and apply its bottom border first, append exactly space + CR at the terminal insertion point, then clear only the new following paragraph's bottom border. It does not call `_paragraph_boundary()`. Existing terminal text stays in the ruled current paragraph. ACK requires exactly two new UTF-16 units and one new paragraph for empty and nonempty terminal cases, current-paragraph rule properties, and centered following alignment with preserved inherited style/spacing. Prefix text equality now uses Foundation `NSString.isEqualToString:` and converts the result to a native boolean, preventing AppleScript's case-insensitive equality from certifying a changed prefix.

The fixture now expects two paragraphs: existing prefix plus trailing space in the centered ruled first paragraph, followed by centered `FOLLOWING` with cleared bottom border. Persisted XML checks enforce that layout, exact case-sensitive prefix text, native border color/width, and inherited following style/spacing. The old assertion that original paragraph XML formatting is unchanged has been removed; changing the current paragraph's alignment/border is the intended baseline behavior.

Fixture lifecycle now reuses source-bound `macos_word_recovery.inventory` and `close_sentinel`, plus `macos_word_fields.retain_sources`. It records the full sentinel inventory row and independently verifies its UTF-8 text SHA256, empty path and unsaved state before owned writes. It exits the owned context first, retains runtime after close (including close scripts/logs), compares the exact sentinel row and remaining starting inventory, then independently closes the guarded sentinel. Final inventory must equal the starting inventory; read-only reopen also exits before its inventory equality check. Private runtime retention is enabled before both contexts so post-close evidence remains available. Failure cleanup is outside the owned context, is suppressed for unverified close/quarantine, is attempted at most once, and records cleanup errors separately from the primary error.

Validation for this revision:

- RED: **8 failed, 37 passed** (`/tmp/word-paragraph-rule-p2-red.log`), including baseline nonempty ACK rejection, previous extra-separator ACK acceptance, emitted mutation-order contract, two-paragraph fixture readback, and missing independent post-owned-close cleanup proof.
- First GREEN: **45 passed** (`/tmp/word-paragraph-rule-p2-green-initial.log`).
- Final focused coverage is **46 tests**, adding persisted XML negative checks for wrong width/alignment/style, changed prefix case, spacing loss, or an extra paragraph.
- Combined session/fields/references/recovery/probe regression: **277 passed in 5.42s** (`/tmp/word-paragraph-rule-p2-combined.log`). Syntax-only `osacompile` remains included; no Office AppleEvents or UI were executed.

No shared session forward, native lease run, installation, commit, or push was performed. The earlier forwarding patch is still unapplied. The revised native fixture is prepared but not executed; primitive probe evidence is not being reused as production acceptance.

Revised source hashes:

- module `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006`
- tests `70b502e2731d4eec79813e990f60dc3258eff6d8a2afc9497c8a09b2b8d3d0f7`
- fixture `1a082729d36ec2f63a7838b8c0537f0b559641f2c2f92f94cdbfabe13451d9ee`
- reused recovery fixture `7de1c9bb54f3d6fea6c6f90da049f7b74eb7a1ab8ceedb8846396f4e7de8900f`

## Narrow fixture revision — unsaved inventory basename representation

The remaining P2 in `word-paragraph-rule-fix-rereview.md` is corrected. `sentinel_preimage` selects the unique actual inventory row and requires the exact sentinel name, `saved is False`, a nonempty token, and exact token+CR SHA256. It preserves Word's actual `posix full name` string instead of requiring it to be empty. `close_after_owned` continues comparing the complete actual row unchanged. The existing independent `close_sentinel` still checks the native `path` property is empty immediately before closing.

The cleanup doubles now use `[name, name, False, digest]`. A new preimage case accepts that real unsaved representation and rejects absent/duplicate/name-mismatched/saved/ill-typed/changed-hash rows. RED: **1 failed, 46 passed** (`/tmp/word-paragraph-rule-basename-red.log`). GREEN: **47 passed in 0.54s** (`/tmp/word-paragraph-rule-basename-green.log`). This is a fixture/tests-only fix; production module, session forwarding and native state were untouched.

Frozen hashes after the narrow fix:

- fixture `768c43c0d97e333f0026d32ff9cd50a8afce2de6e97e34def73f42f4a5ad7eb8`
- tests `a0bd103d0f0af00da5dc8fe7058d276a9f5cf9e36fbaff76f3d1f201c25c1591`
- unchanged production module `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006`

No session edit, native Office/UI, installation, commit or push was performed. The fixture is frozen for the parent's independent review/native lease sequence.

## Authorized public forwarding integration — frozen for native scheduling

Following `word-paragraph-rule-final-rereview.md` SCOPED PASS and root's explicit forward-only authorization, `MacWordSession.add_paragraph_horizontal_line(self)` now contains the exact previously proposed lazy import and `return add_paragraph_horizontal_line(self)`. No other session code was changed by this implementer. The fixture can now call the real public entry point.

Two public entry-point cases preserve the frozen Writer signature and exercise successful versus malformed native operation ACK through the actual session transport. RED: **2 failed, 47 passed** (`/tmp/word-paragraph-rule-forward-red.log`). GREEN: focused paragraph-rule plus existing session suite **108 passed in 0.65s** (`/tmp/word-paragraph-rule-forward-green.log`); the paragraph-rule file now has 49 cases.

Frozen integration hashes:

- session `3ab747a06c63deabc9e0dbf1de51857d8b7dfa843d6ea8e0822915e91e44f205`
- focused tests `65ab3431c9b4339b5063654aba0b3b31d118aaaf39d74946f0fd3a7870d26c83`
- unchanged rule module `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006`
- unchanged fixture `768c43c0d97e333f0026d32ff9cd50a8afce2de6e97e34def73f42f4a5ad7eb8`

The shared session and this slice are frozen at handoff. No native Word/UI execution, installation, commit, or push was performed. Parent owns subsequent serial native scheduling and must refresh source hashes if another authorized session edit occurs first.

## Authorized native run-01 — FAIL, lease released

Root granted one exclusive Word native lease after pagination acceptance. All three frozen source hashes matched before execution: session `3ab747a0…`, rule module `0a7fdfbe…`, fixture `768c43c0…`. The unmodified fixture ran once against Word **16.112.3**, output `docs/verification/microsoft-parity/macos-word-paragraph-rule/run-01`. No source fix or retry followed. Native execution is stopped and the lease is released.

Observed result: **FAIL at the fixture's following-style object-equality assertion**. The production public method returned `None`; its retained raw ACK in `native-runtime/dea7cc3adff648c4a9cc1acf3abc674e.log` was:

```json
["paragraph-rule",45,0,47,46,48,1,2,true," \r","\r",true,true,true,true,true,true]
```

That ACK established the current ruled paragraph, +2 UTF-16 units/+1 paragraph, exact prefix, centered single .75pt C0C0C0 bottom border, clean following bottom border and following-format check for this call. The subsequent independent fixture readback in `native-runtime/be582eb62db246dd83fa936ad54c2e61.log` reported correct prefix/current/following text, all border/center checks, and geometry `6,9,12,18`; only `(style of followingRange is Word style (style body text) of boundDoc)` was false. The captured row does not distinguish an actual different style from object-comparison semantics. No unsupported interpretation or workaround was applied.

The fixture stopped before producing final `paragraph-rule.docx`, PDF or reopen evidence. Only `preimage.docx` exists, SHA256 `a9026498d86d7ff45c68780c8995d65c8d99bbc951de79865b42eb79655704dc`. Therefore the slice is not natively accepted despite the production ACK.

Cleanup proof passed: owned context exited; the synthetic unsaved sentinel's entire inventory row and SHA256 remained unchanged; independent guarded sentinel close succeeded; post-sentinel inventory equaled the empty starting inventory. An additional independent no-launch inventory after the failure also returned `[]`, retained as `post-failure-independent-inventory.{applescript,log,json}`. No quarantine, retained sentinel, or live owned document remained in inventory. The raw report is unchanged.

- Native report SHA256: `e574cdfdea3e4105df114b7158fc77a3ac5b0835c1a48d71553902d5a592fe1f`
- Post-failure inventory JSON SHA256: `4f53cda18c2baa0c0354bb5f9a3ecbe5ed12ab4d8e11ba873c2f11161202b945`

No UI, installation, commit or push was performed. Further diagnosis or a revised fixture/native rerun requires the parent's next bounded authorization.

## Bounded native diagnosis-01 — style preserved, fixture representation assumptions isolated

Root separately authorized one native diagnostic lease with no production/current-fixture/test edits. The standalone `diagnosis-01/diagnose_style.py` recreated the owned document plus nonempty unsaved sentinel and captured before setup, after setup, after public rule, and after `FOLLOWING` insertion. It saved native DOCX at all three nonempty stages and retained raw scripts/logs plus document/styles XML. The diagnostic completed; native work stopped and lease released.

Evidence establishes that run-01's style failure is a fixture object-comparison false negative:

- Before setup the paragraph's `name local` is `正文`; the document's built-in Body Text style has `name local` `正文文本`.
- Immediately after assigning Body Text, `style of range is Word style (style body text) of boundDoc` is already **false**, while both native style names are exactly `正文文本`.
- After the public rule, both current and following paragraphs remain `正文文本` and centered. Appending `FOLLOWING` preserves those same styles and alignment.
- Direct `style of range as text` and document Word-style object `as text` each return native coercion error `-1700`. Those errors are retained per property rather than suppressed or reinterpreted. `name local of style` returns the expected names at every stage.
- All saved paragraphs after setup/public insertion/following insertion use `w:pStyle w:val="ae"`. In each saved `styles.xml`, paragraph style ID `ae` resolves to `w:name w:val="Body Text"`. Thus the current XML verifier's literal `BodyText` style-ID requirement is a second representation assumption, not evidence of native style loss. No assertion was removed and no production behavior changed in this diagnostic stage.
- Saved rule XML has one current-paragraph single border, size 6/color C0C0C0, and both paragraphs retain `before=120`, `after=180`, `line=360`/exact, `firstLine=240`, and centered alignment. The following paragraph has no bottom border.

Native DOCX hashes:

- setup `e4c2dca7612f4161919a5eb88307caa3899970da6d20eb2ac6840a2efe10f197`
- after rule `81ad81455371c360c4f459b98d0902f2b8362f28f0e23e907bd8deffaba67629`
- after following text `12c14fc398d1dbb1f5277114bf6c72826a99a648ff28ef13313d6096e93fcc6d`

Cleanup proved owned context close, the sentinel's exact unsaved full-row/hash preservation, guarded independent sentinel close, and final inventory `[]` equal to starting inventory. No quarantine or retained sentinel remains. This diagnosis does not include PDF/UI/reopen acceptance and does not turn failed run-01 into PASS.

Frozen diagnostic evidence:

- report `927c6c3832293944dbe6e4934614dc9015d0ea58245a072f690a1a80a7372380`
- diagnostic source `da1de6c1af5d4990720b0ce5d65e6034dd1941befebf3631e063fa4cb7fcc40f`
- extracted `resolved-style-ae.json` `59cfc1991b00b7430524322e4e95ed04bd83c788c49495ddc1fbbd6def2e9893`

Production/session/fixture/tests remain at their previously frozen hashes. A bounded fixture correction should compare stable native style names exactly and resolve paragraph style IDs through each package's styles.xml, retaining failure for missing/ambiguous/wrong mappings. That correction and a new native run are not performed under this diagnostic authorization.
