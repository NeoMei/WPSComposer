# Two-method native run-01: partial pass, table creation unresolved

2026-09-09. Root reviewed the fixture, requested default following paragraph, and
granted an exclusive Word native lease for run-01 after pure tests.

## Frozen change and pure gate

Changed only fixture call to `session.add_paragraph(following)` with no explicit
italic/color overrides. Production helper remains reviewed and unchanged. Pure
degradation test module: 71 passed in 1.82s. git diff --check passed.

## Native outcome

Command actually executed:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_degradation.py --execute --output docs/verification/microsoft-parity/macos-word-degradation/run-01
```

Word 16.112.3. Overall run FAILED, with 18 checks passed:

- inline-empty: all nine checks passed.
- inline-nonempty: all nine checks passed.
- block-empty: ordinary AssertionError at fixture line 234, expected native table
  count 1 but found 0. Returned semantic box is a range fallback (table_index None).
- block-nonempty and both controlled partial-table-failure cases: NOT RUN.

The two inline passes prove native UTF-16/range/style facts, unchanged prefix,
default normal following paragraph free of notice style, native DOCX/PDF,
persisted OOXML styles, public reopen with native style/readback, real edit and
acknowledged rollback, unchanged source bytes, and both owned and independent
sentinel closes. They do not prove block/table support.

Block raw table batch returned `[["degradation-table-failed",0]]`. Native checkpoint
state before and after that batch contains zero tables and terminal bound zero;
the real no-op rollback acknowledged exact preimage and then fallback inserted
`[NOTICE_BLOCK_EMPTY] 中文😀` once. No quarantine occurred. Original ordinary
AppleScript error details are suppressed by the production table catch, so the
actual failing command is not yet proven from this evidence alone.

Hypothesis for the next minimal probe: the newly created sentinel is foreground;
existing MacWordSession._table_commands starts with `activate object boundWindow`,
while this helper's table builder does not. Test actual creation error on a
private inactive document, then compare after exact owned-window activation.
This is a hypothesis, not an established root cause or a production fix.

## Exact cleanup and retained failure

The fixture correctly exited the owned context before its exception reached the
report; it retained the synthetic sentinel rather than closing it from that
context. A separate bounded cleanup then used the existing recovery fixture's
inventory and close_sentinel helper:

1. Verified no owned document remained and the sole sentinel name/path/saved flag/
   text digest exactly matched its before record.
2. Closed only the exact unsaved sentinel after verifying its unique synthetic
   token and mandatory final CR.
3. Verified native inventory restored to [] (the original inventory).

Cleanup evidence: `run-01/block-empty-cleanup/report.json` plus source/log files.
Original `run-01/report.json`, block-empty failure.txt, native runtime, and frozen
source snapshots are retained unchanged. The report's retained_sentinel field is
the original failure-time fact; the separate cleanup report supersedes live state.

All designated run source hashes were rechecked after the run and still matched.
No production source change, personal install, commit or push. Native execution
has stopped; awaiting root decision on the minimal diagnostic continuation.

| File | SHA-256 after run |
| --- | --- |
| `fixtures/microsoft_parity/macos_word_degradation.py` | `1e6174293f660138ea0c6e0042922934cd7e379c3ce5e601fe03c581d81cb5c1` |
| `skills/WPSComposer/scripts/msoffice/macos_word_degradation.py` | `be11e3347d3d02a7aa167a8923da82a1646e38eab5616b4f14243cc34f2e7a54` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `a6f39d1191a7df557a7c911904e598bb5fa538430314a84afc9aeef711e0d534` |
| `tests/msoffice/test_macos_word_degradation.py` | `918f67df1087bcdebcc5227176ef67faf725a72419d205c59d6272e7d6389c7d` |

## Bounded native diagnosis follow-up: root cause established

Root subsequently granted a narrow diagnosis lease. Executed task-local
word-degradation-window-diagnosis.py into a new retained evidence directory
`docs/verification/microsoft-parity/macos-word-degradation/window-diagnosis-01`.
Production files remained unchanged. The two comparisons used independent fresh
owned private documents, each with a unique foreground unsaved sentinel:

| Exact owned window state | Native result | Before / after native geometry |
| --- | --- | --- |
| Not activated | error -2710, cannot create class table | text CR, end 1, tables 0 unchanged |
| `activate object boundWindow` first | native table created, range 0:2, 1 row / 1 column | document end 3, tables 1 |

This establishes the missing exact bound-window activation as the table creation
failure cause for run-01. The existing general table builder already uses this
activation primitive. No selection or clipboard action was used. Both comparisons
verified the sentinel's entire name/path/saved-state/content-hash record after
owned close, then independently exact-closed it; final inventories are both [].
Raw private error number/message, source copies, native commands, log results,
close reports, and hashes are retained in window-diagnosis-01/report.json and
its inactive/active subdirectories. Native table full-body serialization contains
extra CR/BEL compared with story offsets; future table success ACK geometry still
needs complete helper readback after the scoped activation fix. Do not relax
its strict success ACK without native evidence.

No production fix has been made by this diagnosis step. Native diagnosis is
finished and the Word lease is released to root for review/next assignment.
