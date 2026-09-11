# Mac Word degradation native-fix scoped review

Date: 2026-09-09  
Verdict: **SCOPED PASS — no new actionable P1/P2 found.** This review covers only the two degradation forwards, the native-backed table activation/geometry repair, the guarded acceptance fixture, and their focused tests. No production file was edited and no Office/UI/native job was started.

## Frozen inputs reviewed

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_degradation.py` | `8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7` |
| `tests/msoffice/test_macos_word_degradation.py` | `07396fb846503c098e132e21dd009f90c62544f702636cbd096063fb750a651b` |
| `fixtures/microsoft_parity/macos_word_degradation.py` | `39e1ab6b0a2fe7a7ed180b0d8c4845a2d76a269e92d8f63058b0ffbfab3fa8b6` |
| observed `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `a6f39d1191a7df557a7c911904e598bb5fa538430314a84afc9aeef711e0d534` |
| `word-degradation-table-fix-report.md` | `875c2b6f2b6da87569baa370f0f46c5244b986bdd031acb2d1448c7de2de27ea` |

The older `word-degradation-forward-fixture-report.md` still names the pre-diagnosis hashes (`be11e...`, `918f...`, `bd092...`). It is historical handoff evidence, not the current freeze. The current freeze and rationale are accurately recorded in `word-degradation-table-fix-report.md`.

## Finding disposition

### Exact bound-window activation — addressed

`table_commands` now performs `activate object boundWindow` immediately inside the table mutation `try` and before constructing `noticeTarget`/`noticeTable`. It neither selects a user document nor activates a guessed active document. `window-diagnosis-01` shows the bounded host behavior: without activation Word 16.112.3 returned `-2710` and left body/table counts unchanged; activating the exact bound window created one 1x1 table. The helper still lets timeout, cancellation and invalid-connection numbers escape rather than treating them as an ordinary fallback case.

### 1x1 serialized text versus native coordinates — addressed

`geometry-diagnosis-01` records both required geometries from the actual helper commands:

- Empty body: checkpoint/table start `0`, table end `24`, display end `22`, document end `25`.
- Nonempty CJK/emoji body: checkpoint `11`, Word-inserted table start `12`, table end `36`, display end `34`, document end `37`.

In both cases full table text is exactly `display + CR/BEL + CR/BEL`, while the two terminator pairs consume two native story coordinates. `_table_ack` now requires that exact serialization, `table_end == table_start + display_units + 2`, a display-only subrange, one-row/one-column geometry, a single table-count delta, exact style rows and `allow break across pages == false`. It accepts only the observed `table_start` of `position` or `position + 1`; it does not strip arbitrary terminators or infer geometry from serialized string length. The returned handle intentionally retains native Start/End together with the actual serialized `.Text`, matching the direct table range surface.

### Ordinary failure rollback ordering — preserved

Only the exact single `['degradation-table-failed', checkpoint]` row authorizes recovery. The helper calls the actual `macos_word_recovery.rollback`; fallback submission occurs only after rollback has returned its native preimage/deletion/postcondition ACK. Malformed failure or success rows quarantine without fallback. Failed/uncertain rollback retains evidence and prevents fallback. The guarded fixture injects a real table plus cell text before an ordinary `-2700`, observes the closed failure ACK, then asserts the restored exact body/table preimage before allowing exactly one block-range fallback. This does not simulate rollback.

### Following formatting and fixture ownership — suitable for native acceptance

The fixture covers inline, successful 1x1 block table and partial-table range fallback with empty and nonempty emoji prefixes. It resolves the following paragraph through its own native paragraph start/end rather than counting CR/BEL serialized body text. Immediate and reopen checks require following text to lack degradation italic/color/shading; OOXML additionally requires exact notice formatting, block paragraph spacing/keep/body outline, no-split row, expected table count and unstyled following paragraph.

Each case records a nonempty unsaved sentinel preimage, exits both owned create/reopen contexts before independent exact sentinel close, checks the sentinel row after each owned close, and requires final inventory to equal the starting inventory. Source snapshots and hashes include helper/session/recovery/privacy/transport dependencies. Failures retain runtime and leave a recovery target rather than guessing cleanup.

The two session methods are exact lazy forwards with baseline signatures:

- `add_inline_degradation(self, code, message, fallback_text)`
- `add_degradation_notice(self, code, message, fallback_text, placement="block")`

They introduce no separate state or altered return value. Any later independent session forwards change the session hash and therefore require the final native run to capture and report that current hash, but do not alter this helper verdict.

## Independent pure verification

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest -q \
  tests/msoffice/test_macos_word_degradation.py \
  tests/msoffice/test_macos_word_recovery.py \
  tests/msoffice/test_macos_word_references.py

169 passed in 3.38s
```

The focused degradation file separately passed **74 tests** during this review. These tests and diagnosis reports justify proceeding to the guarded source-bound native run; they are not themselves final public/native/UI parity acceptance.
