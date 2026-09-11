# Mac Word rules and WordArt — preparation checkpoint

Status: **COMPILE_ONLY_PASS / native pending**. No Word AppleEvents, native mutation, GUI, production code edits, commit, or push were executed by this worker. Parent owns Word lease scheduling.

New files:

- `fixtures/microsoft_parity/macos_word_rules_probe.py`
- `tests/msoffice/test_macos_word_rules_probe.py`

The fixture uses three separate new private MacWordSession documents. It addresses insertion at the bound document's terminal `end - 1` and invokes `_execute_structural` for document mutations, retaining the session's topology/deadline/ACK protections. Each case creates a nonempty unsaved task sentinel and verifies its content and saved state before explicitly discarding only that sentinel. It saves DOCX and PDF, closes the owned document, reopens read-only, repeats object readbacks, checks original DOCX SHA256, inspects persisted XML/PDF, and runs final inventory independently of the closed session. Runtime errors and raw scripts are retained. Quarantine or a retained sentinel stops subsequent cases. Existing output directories are rejected. Import is inert and the native path has a real `if __name__ == '__main__'` guard plus `--execute` gate.

Frozen COM semantics used: distinct standard native inline horizontal line; centered single-space paragraph with bottom single .75pt silver C0C0C0 border and clean following paragraph; native WordArt with Arial 36 regular and default position 200/400, preset 0 (dictionary `wordart format1`). The probe's WordArt case is the default representative only. Nondefault presets, scalar boundaries, public method input preflight/typed semantic handles, pending heading behavior and end-to-end production forwarding remain subsequent implementation/acceptance work.

Validation:

- RED: 9 failing tests because the isolated fixture did not exist; `word-rules-probe-red.log`.
- GREEN: **9 passed**, including all three syntax-only `/usr/bin/osacompile` branches; `word-rules-probe-green.log`.
- New tests reject missing/truncated/duplicate/ill-typed native rows, wrong paragraph XML border/next paragraph state, and plain text masquerading as WordArt.
- `word-rules-compile-01/report.json` retains command results, exact scripts, dictionary SHA256 and source snapshots. Every `executed` field is false.
- Fixture SHA256: `58ddd65df4406bf0652fbbc9beb0e0194c20864f7674f083cab3aa671fdae381`.
- Frozen `writer.py` SHA256: `b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205`.
- Word dictionary SHA256: `7cb51b924cab566320cc3019e92c1516302ae11220a81ae6279f1930c26320e7`.

Native lease continuation, from the worktree, with a fresh output name:

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_rules_probe.py --execute --output docs/verification/microsoft-parity/macos-word-rules-probe/run-01
```

Use `--case inline`, `--case paragraph`, or `--case wordart` to isolate a failure in a new directory. Retain rejected native creation evidence; do not reinterpret dictionary compilation as insertion success. The artifact validator currently understands native VML rule rectangles/textpaths; an alternative persisted representation must be inspected and independently characterized before extending that validator. PDF page PNG is generated for parent visual inspection, not automatically marked as visually accepted. UI edit/undo/save/reopen remains the parent's later gate.

## Native lease result — 2026-09-09 13:25 CST

The parent granted exclusive native Word lease. Three sequential cases ran against Word 16.112.3 with the exact compiled fixture above; no production edits were made. **Native work is now stopped and lease returned.** Summary: `docs/verification/microsoft-parity/macos-word-rules-probe/native-summary-01.json`.

- `run-01-inline`: **FAIL**, acknowledged `WPSC_RULE_COUNT_DELTA_FAILED (-2700)` after make-new returned. The required native inline-shape delta was absent. Task sentinel was explicitly closed; before/after inventory empty. No native-object or artifact success. Static investigation finds the class is declared but cannot resolve whether make-new silently no-ops or uses different insertion/collection semantics; no retry after later WordArt connection failure.
- `run-02-paragraph`: **PASS** for probe scope. Direct and read-only-reopen readbacks report centered `" \r"`, single .75pt silver C0C0C0 bottom border, and clean `FOLLOWING\r` paragraph. DOCX XML and PDF checks passed, source DOCX hash unchanged after reopen, task sentinel verified, final inventory empty. The PDF page PNG was visually inspected: one silver rule between the marker and following text, no extra rule. DOCX SHA256 `efc4d0896b86053d9f3c09a6196e5b9674e947a3638a1cc9152c2b3e135566d4`; PDF SHA256 `279980d1f99a01f09e408516e1bfab7c542dfb85af51d1663deb0844ad29cbc7`. This establishes editable native paragraph-border semantics; public production method and UI edit/undo acceptance remain pending.
- `run-03-wordart`: **FAIL / connection loss, no retry**. The exact `make new word art` command returned `连接无效 (-609)`. Cleanup then returned sentinel object missing `(-1728)`. Both errors are retained in `failure.txt` and `failed-runtime`: primary `2f7f0d86e3274909b0b21a9134cdbb0a.log`, cleanup `4d0662c892d74a31b1ebf4e40b44261f.log`. The final independent inventory and a second no-launch inventory are empty, but **empty inventory does not prove sentinel preservation**. Starting user-document inventory was empty; only the private owned document and the synthetic unsaved sentinel had been created. No user document was present. Read-only process evidence shows Microsoft Error Reporting started 13:25:31 and Word started 13:25:32, consistent with a crash/restart. No macOS DiagnosticReports Word crash report was found; binary Office diagnostic logs were not treated as a readable crash backtrace. The fixture/runtime reports nonquarantined because existing transport does not classify -609 as uncertain; parent notified to route that separate runtime fix. Nondefault WordArt preset was intentionally not attempted after connection loss.

The original per-case report JSON is retained without rewriting its observations. The summary/report above distinguishes WordArt's empty final inventory from the failed sentinel-preservation contract and retains the primary error that the subsequent cleanup exception superseded in the top-level error field. There are no WordArt DOCX/PDF/reopen or UI success claims.

Evidence flag addendum: the immutable paragraph report's generic `pdf_wordart_text: true` means **not applicable**, not an exercised WordArt check; it is explicitly excluded from `exercised_checks` and listed under `not_applicable_checks` in the current summary. Paragraph has nine exercised checks. In WordArt's immutable report, `unrelated_inventory_preserved: true` compares only the empty preexisting inventory; it does not establish preservation of the subsequently created synthetic sentinel, whose preservation failed. Do not map either generic flag into broader capability claims.
