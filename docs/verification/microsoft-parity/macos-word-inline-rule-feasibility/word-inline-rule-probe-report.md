# Dedicated Mac Word inline-rule feasibility runner

Current status: **ONE NATIVE RUN FAILED STRICT INLINE SEMANTICS / CLEANUP VERIFIED / LEASE RELEASED**. See the run-01 section. Production and frozen fixture source remained unchanged during the run; no retry or WordArt execution.

Initial preparation checkpoint: **PREPARED / PURE TESTS PASS / NATIVE NOT EXECUTED**. At that checkpoint work was limited to this worktree and three new files; no production modification, Word launch, AppleEvents, WordArt execution, native lease, commit, or push had occurred. Root retained native scheduling and production numbering ownership.

## Frozen scope and single hypothesis

Frozen `6dd3a00` `WriterComposer.add_horizontal_line(self)` calls `self.selection.InlineShapes.AddHorizontalLineStandard()` and implicitly returns `None`. This is a native inline horizontal line, distinct from `add_paragraph_horizontal_line` and its paragraph border. No public method or forwarding is added here.

`fixtures/microsoft_parity/macos_word_inline_rule_feasibility.py` imports the existing `macos_word_inline_document_variant.build_commands()` unchanged. Its sole candidate constructor is:

```applescript
set creationResult to make new standard inline horizontal line at boundDoc with properties {text object:insertionRange}
```

The old `macos_word_rules_probe.run_case/main` runner is never called. Only its fixed inline command/readback validator is reused through the existing candidate. The prior count-delta failure and WordArt connection-loss evidence remain unchanged. No retry, alternate constructor, paragraph-border fallback, global process action, clipboard access, security change, or macro is introduced.

## Guard and retained evidence

- Import and CLI without `--execute` are inert; programmatic `run` requires `execute is True` and rejects an existing output directory before native work. A fresh exclusive root-granted Word lease is still required operationally.
- Exact bound document/window paths guard all session phases. The candidate preserves diagnostics followed by its required native inline count increment, native horizontal-line type, and positive finite width/height. Reopen must repeat the strict readback and exactly match the original typed rows.
- Saves native `before.docx`, `after.docx`, and `after.pdf`; retains **every XML/relationship part verbatim** from each saved package, native scripts/logs including owned close, PDF text/drawings and page PNGs. Prefix text must survive exactly. Native VML horizontal-rule XML is required; plain text, a paragraph border, image, missing inline container, or duplicate rule cannot satisfy it.
- PDF validation requires one opaque horizontal rectangle or solid horizontal stroke matching native width/thickness, within the page; arbitrary drawing count is insufficient. The representation remains a feasibility hypothesis until real PDF inspection. PDF PNG generation does not establish independent visual or UI acceptance.
- An ordinary constructor/readback error stays FAIL. While the session remains usable, one diagnostic count read and native `partial.docx`/`partial.pdf` checkpoint retain the failed state without repairing or repeating the constructor. If a checkpoint fails, its error is separate from the primary error; any already saved package and its complete XML remain retained. Closed/quarantined sessions receive no partial native reads or saves.
- The owned context exits first. Independent inventory must then exactly match both the preexisting document inventory and the synthetic sentinel's original native name/full-name/saved-state/text SHA-256. A separate script rechecks exact case-sensitive name, empty path, exact text and unsaved state before closing only the sentinel. Failed cleanup is not retried. Quarantine retains recovery identities and skips further native inventory.
- Before/after DOCX, PDF, XML, relationship, PNG and dictionary hashes are reported. The runner snapshots all shipped Python sources and fixture Python dependencies and verifies live plus retained bytes against their initial hashes at completion. Dictionary preflight failure produces a retained report without starting native inventory; a later hash-audit failure cannot replace the primary error.

## Verification

Test file: `tests/msoffice/test_macos_word_inline_rule_feasibility.py`.

- Initial RED: 17 expected failures because the dedicated fixture was absent. Additional PDF-path and preflight-retention tests were observed failing before their fixes.
- Final new-file suite: **22 passed**.
- Combined pure suite: **38 passed, 2 deselected** in 0.24 seconds. The two dictionary compilation cases were deliberately excluded; no `osacompile` or native Word execution was performed. Five upstream PyMuPDF SWIG deprecation warnings were observed.
- Tests exercise explicit-execute rejection, unique emitted constructor, strict count/type/geometry and XML negatives, opaque/vector PDF geometry negatives, changed-sentinel refusal, quarantine, source drift, dictionary preflight failure, and simulated complete success/ordinary-failure flows proving owned-close → independent sentinel verification → sentinel-close → final inventory. Simulation is not native acceptance.
- The existing adjacent venv lacks PyMuPDF; verification used the already available `/tmp/wps-spike-validator-audit-deps` dependency directory. No dependencies were installed or modified. Current README/AGENTS clean-dev instructions install PyMuPDF through the dev extra.

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_inline_rule_feasibility.py tests/msoffice/test_macos_word_quality_feasibility.py -k 'not compiles_local_dictionary' -q
```

Prepared source SHA-256:

| File | SHA-256 |
|---|---|
| New guarded runner | `3077559de0062195c5a86472e743ecf64a9d8f6e4e28e258a7ef7ad9c9a78e32` |
| New pure tests | `d5462ceca1a3d5d0a6c0bb8a549c9d4c652081b741d7bd052c062dbe631f6871` |
| Unchanged constructor candidate | `386aa4dc70de7eac48742535c37a2e97e36aa77ca40a21626e4124efa50a2ba3` |

## Single leased native run — run-01

After the root granted one exclusive native run on the corrected `ebf8a048` runner and scoped re-review passed, the exact unchanged constructor ran once against Word **16.112.3**. Result: **FAIL / acknowledged ordinary count-delta error**, not quarantine. No retry, WordArt, fallback, code edit, UI action, or additional native experiment occurred. The Word lease was explicitly released immediately after the runner completed.

Evidence: `docs/verification/microsoft-parity/macos-word-inline-rule-feasibility/run-01/`. The untouched native `report.json` SHA-256 is `2111ba9d0febe8617bf4af0c5faee616d1ec7f18d53d68c3f0168e4caa8a4fe6`.

- Exact constructor log `native-runtime/2f4cf608f62e475dbe7bb220242d7560.log`: `inline-creation-counts` reports before inline `0`, after inline `0`, standard-inline `0`, character count `25`, subtype diagnostic error empty. Returned-object class diagnostic failed `-1728`. The unchanged strict gate then raised `WPSC_RULE_COUNT_DELTA_FAILED (-2700)`.
- The single partial-state observation reports inline shapes `0`, floating shapes `1`. Full saved XML shows VML `rect id="Horizontal Line 1"`, `width:415.3pt;height:.1pt`, absolute positioning and anchor lock, without `o:hr`. Offline strict inline XML validation also returns false. This is evidence of the wrong native object semantics; it does not authorize reclassification or weakening the gate.
- Native `before.docx`, `partial.docx`, `partial.pdf`, every saved XML/relationship part, scripts, logs, and source/dictionary snapshots are retained. The exact original body text remains in partial XML. No accepted inline DOCX/PDF, reopen, visual, or UI success is claimed; the failed candidate never reaches those acceptance phases.
- Owned document closed successfully. Synthetic unsaved sentinel `文档85` retained exact name/full-name/unsaved-state/text SHA-256 `833ca68a10f019049d6b6d9685e0447eb0df0144f4ea342a72d79d72c2b215a6` after owned close; the separate exact guard/close acknowledged success. Independent final inventory is `[]`; `quarantined=false`, no remaining sentinel or uncertain recovery state.
- All **158** current source files match before/after capture and their retained copies, including production numbering `21307cfa…`. Dictionary SHA-256 remains `7cb51b924cab566320cc3019e92c1516302ae11220a81ae6279f1930c26320e7`. Sibling `run-01-source-before.json` and `run-01-source-after-audit.json` preserve the full maps; `run-01-offline-observation.json` records the subsequent file-only diagnostic inspection.

Artifact SHA-256: before DOCX `ba8a65579790f83324fecc6d5c651e386bcf0ead8fc277e74f29f758032944a5`; partial DOCX `c591b559f0fa097e3446996ee87c4a9c04908952d327797ec26343b2540eb241`; partial PDF `4d25dcf5f94d674728a424686d27bf6c728957afa16a9a3b47d5789bedf7f512`.

## Root-only next gate

Review the runner and, only after the production owner is stable and the root grants a fresh exclusive Word lease, execute this dedicated file once with `--execute --output <fresh-directory>`. Keep ordinary failures and quarantine artifacts intact; do not invoke the old multi-case rules runner. Native success, independent PDF visual inspection, production implementation of the frozen public contract, and actual UI edit/undo/save/close/reopen remain separate unfulfilled gates.

## Scoped review correction — independent cleanup uncertainty

The finding in `word-inline-rule-probe-review.md` was reproduced with pure fault injection: 14 failing cases demonstrated the missing runner quarantine after sentinel-close timeout, `-609`, `-1712`, malformed/wrong ACK, and post-owned-close inventory timeout/malformed ACK, each both with and without an earlier constructor failure. The existing review and prior failure evidence remain unchanged.

The fix is confined to this runner and its test file. Independent inventory now has a local guarded wrapper; its failed transport or invalid typed inventory marks completion uncertain. The raw sentinel-close subprocess distinguishes an executable launch failure (`FileNotFoundError`/`PermissionError`, no child submitted), a verified pre-close identity guard refusal, a valid close acknowledgement, and uncertain completion. All uncertain paths record FAIL, set runner quarantine, retain raw/partial diagnostics, and write `native-uncertainty.json` in the run directory. The marker records the stage, exact sentinel native name/preimage, close-script path, and the owned document's real closed flag/path/staging root. It directs the root to reconcile this task before further native work; it does not install a global or shared recovery gate.

Once runner uncertainty is set there is no close retry, reopen, or final inventory AppleEvent. A still-open owned context is quarantined through its existing `_retain` method so its context exit also refuses native cleanup. A previously closed owned session remains `_closed=True` and is not falsely relabeled open or session-quarantined. If the sentinel already had a valid close ACK before a later final-inventory failure, that confirmed sentinel closure remains recorded; the wider inventory observation is what remains uncertain. Submission-before-launch failure and confirmed identity refusal remain distinguishable from uncertainty, and the original constructor failure remains the primary error when cleanup also fails.

Final pure verification:

- Dedicated runner suite: **43 passed**, 5 upstream PyMuPDF SWIG deprecation warnings.
- Dedicated + imported quality pure tests: **59 passed, 2 dictionary-compile cases deselected**, 5 warnings, 0.90 seconds.
- Covers both full success/ordinary-constructor-failure paths for all cleanup faults above, before-submission failure, verified identity refusal, live-owner quarantine, final-inventory uncertainty, no native follow-ups/retries, exact recovery identity, retained partial stdout/stderr, and preservation of verified owned/sentinel close facts.
- No native Word, osascript, osacompile, UI, WordArt, production/shared-helper edit, lease, stage, or commit was performed. Native gates remain unfulfilled; scoped re-review belongs to the root.

Corrected source SHA-256 (supersedes prepared runner/test hashes above):

| File | SHA-256 |
|---|---|
| Guarded runner | `ebf8a0480a25efa504311ae4d513c0e341ced6b36e264655e1788935e714ad82` |
| Pure tests | `8edefe47530909336f90c34983bca844222e00fabccd4c7ddba8dd330b7fe1ca` |
| Unchanged constructor candidate | `386aa4dc70de7eac48742535c37a2e97e36aa77ca40a21626e4124efa50a2ba3` |
