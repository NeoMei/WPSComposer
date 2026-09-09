# Mac Word paragraph rule independent review

Status: **CHANGES REQUIRED — two P2 findings.** Review scope is the frozen one-method implementation only. No production code, public forward, native Office process, UI, installation, commit, or push was changed or run.

## Frozen inputs

- `macos_word_rules.py`: `e807932b76daf4ac7dd307088c392bcf6136bf72b64fd896a1840155a12af7ec`
- `test_macos_word_rules.py`: `d0b0578709535ab30605b63194559d2f934a6fb3494cc2ec74fa4c99f4d13ea4`
- `macos_word_paragraph_rule.py`: `225b507f393817ba5910cf10c756e8a05eb25bc9cbe4cb7ac69cddc097651268`
- observed shared `macos_word_session.py`: `a6f39d1191a7df557a7c911904e598bb5fa538430314a84afc9aeef711e0d534`
- frozen `writer.py` baseline: `b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205`

## P2 — nonempty-terminal semantics drift from the frozen baseline and the prefix proof is not exact

The direct baseline at `writer.py:1173-1188` centers the **current paragraph**, sets its bottom border, types one space, types a paragraph mark, and clears only the new following paragraph's bottom border. On a nonempty terminal paragraph this deliberately keeps the existing text in that current ruled paragraph; the following paragraph inherits the centered paragraph formatting except for the cleared bottom border.

The new compiler instead calls `session._paragraph_boundary()` before creating `ruleRange`. For a nonempty terminal paragraph it inserts an extra paragraph mark, leaves the existing paragraph and its formatting unchanged, creates a separate space-only rule paragraph, and preserves the following paragraph's earlier noncentered alignment. This is a visible paragraph-count, border-target and inheritance difference. It is not authorized by the no-argument baseline. The ACK encodes this divergent behavior by accepting `separator == 1`, while the fixture requires exactly three paragraphs and requires the original prefix paragraph XML to remain byte-equivalent. A native PASS from that fixture would therefore certify the Mac-only behavior rather than parity.

The same ACK overclaims exact prefix preservation. `prefixPreserved` uses AppleScript `(content ... as text) is rulePrefix`; AppleScript text equality is case-insensitive by default. A case-only native mutation can therefore return `true`. The Python validator trusts that boolean and can commit `_structural_changed`/clear `_pending_heading` despite not having proved the claimed exact prefix text.

Required bounded fix:

1. Remove the extra `_paragraph_boundary()` path for this method. Apply center alignment and the native single `.75pt` `C0C0C0` bottom border to the current terminal paragraph, append the space and paragraph mark, then clear only the following paragraph's bottom border.
2. Make the ACK require the baseline paragraph-count/range behavior for both empty and nonempty terminal documents. Read back that the original/current paragraph is the ruled paragraph and the following paragraph inherits center alignment while its bottom border is absent. Do not claim that the original paragraph formatting is preserved.
3. Compare preserved prefix text with Foundation `NSString.isEqualToString:` in this method's generated script. Keep the existing topology submission boundary; no shared hook is needed.
4. Update the fixture's three-paragraph and `prefix_paragraph_xml_preserved` expectations to the ruled-current-paragraph baseline before the source-bound native run.

Minimal source reproduction for the semantic drift is `paragraph_rule_commands(MacWordSession())`: it contains `_paragraph_boundary()` before `ruleStart`, whereas the frozen Writer method assigns alignment/border before `TypeText(" ")` and `TypeParagraph()`. The current pure ACK also accepts a constructed nonempty row with no separator, but that is **not** a defect after the baseline ruling; the defect is that production generation and the fixture demand the extra separator.

## Lifecycle and evidence findings

### P2 — fixture closes the sentinel through the owned session before owned close is proved

The guarded fixture closes its unrelated synthetic sentinel in the inner `finally` by calling `session._execute(...)`, while the owned document context is still open. This repeats the older cleanup pattern that the recovery fixture has since replaced. It cannot separately prove that the owned context closed first while the sentinel remained exact and unsaved, and its `native-runtime` copy occurs before context exit, so it omits the owned close AppleScript/log from the retained runtime evidence.

Required bounded fixture fix: preserve the nonempty sentinel's exact preimage inventory row (including text SHA); exit and verify the owned context first; compare the post-owned-close sentinel row exactly with that preimage; close only that sentinel through an independent guarded helper that validates exact name, empty path, exact token and `saved=false`; then require final inventory to equal the starting inventory. Retain the source-bound runtime only after the owned close result is available so its close script/log is included. Cleanup failure must remain separate from the primary failure and must block PASS.

No additional P1/P2 was found in the scoped lifecycle review:

- Local closed/read-only/quarantined/deadline and script-write failures occur before actual submission and preserve pending-heading/field-topology state.
- `_execute_topology_mutation` invalidates observed field topology at the existing actual submission boundary. Complete semantic ACK is required before `_structural_changed` and `_pending_heading` commit.
- Submitted native failure and malformed semantic ACK quarantine the exact session and retain private diagnostics. There is no global quit, selection, clipboard, or template mutation.
- The guarded fixture refuses missing `--execute` and an existing evidence directory and binds the actual public forward before native work. Its ownership evidence becomes suitable after the sentinel/close-order P2 is corrected.
- The accepted `macos-word-rules-probe/run-02-paragraph` proves the native border primitive only. It does not prove the current public implementation's nonempty-terminal behavior.

## Independent pure verification

Command:

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest -q tests/msoffice/test_macos_word_rules.py
```

Result: **38 passed in 0.29s**. This confirms the frozen tests are green but does not close the P2 because their success and fixture assertions currently encode the divergent three-paragraph behavior. Public forwarding and the source-bound native run remain correctly pending.
