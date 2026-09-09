# Continuation integration review

Date: 2026-09-09

Scope: read-only integration review of the continuation package that combines
the previously scoped-reviewed Mac Word field-topology, checkpoint/recovery,
degradation notice, pagination and paragraph-rule slices, plus the installer
scratch exclusion. I reviewed the current production sources and public
documentation, retained one pure reviewer reproduction, and inspected the
reported native artifacts. I did not edit production/tests/fixtures, launch an
Office process or UI, install the plugin, run the full suite, commit or push.

## Verdict

**CHANGES REQUIRED: one production P2 and two verification P2s. No P1 and no
additional worthwhile P2 were found in this bounded integration review.** The
previous scoped findings remain closed; the three findings below are the exact
bounded fix wave. This is not a full Microsoft parity conclusion, Windows
native acceptance, or coverage of the 14 explicitly unimplemented methods.

## P2 — destructive checkpoint rollback error leaves the Word session usable

`macos_word_recovery.rollback()` submits one AppleScript that may delete fields,
bookmarks, tables and the appended body range before checking the saved
postcondition. At `macos_word_recovery.py:283`, that script is sent through
ordinary `session._execute()`. If Word returns an acknowledged ordinary error,
`MacWordSession._execute()` records `_retain_evidence` but deliberately does not
quarantine for every `-2700` (`macos_word_session.py:341-354`). The broad
`rollback()` exception handler then raises `LOCAL_MUTATION_ROLLBACK_FAILED`
without calling `_retain()` (`macos_word_recovery.py:293-294`). The same session
can consequently send another AppleEvent even though rollback may already have
partially deleted content.

The retained pure reproduction is:

```text
.superpowers/sdd/2026-09-08-microsoft-wps-parity/
continuation-integration-review-repros.py
```

It uses the real `MacWordSession`, real checkpoint/rollback compiler and real
transport envelope, replacing only `subprocess.run`. The rollback call receives
an ordinary `WPSC_CHECKPOINT_POSTCONDITION_FAILED (-2700)` after the generated
script contains `set content of rollbackRange to ""`. Current source raises the
expected typed rollback error and retains evidence, but `_quarantined` remains
false; the reproduction fails there. Its GREEN continuation is explicit: a
subsequent `_execute()` must raise `NATIVE_WORD_QUARANTINED`, and the fake native
runner call count must remain three.

Independent command and result:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python \
  .superpowers/sdd/2026-09-08-microsoft-wps-parity/continuation-integration-review-repros.py

AssertionError: assert session._quarantined is True
```

This is P2 because a normal public recovery failure incorrectly leaves a
possibly mutated native document/session available for later work. No
unrelated-document mutation or proven data loss raises it to P1. The minimal
repair belongs at the real destructive rollback submission boundary. Local
argument/tracking/preimage validation, script creation, and a deadline failure
before native submission must not be classified as a partially committed
rollback. Existing `_execute` timeout, cancellation, disconnect and stale
binding quarantine stays authoritative.

## P2 — PowerPoint deadline tests depend on process uptime

Two tests set `session._deadline = 123` as though it were an opaque sentinel:

- `test_publication_uses_session_deadline` at
  `tests/msoffice/test_macos_powerpoint_session.py:111-118`
- `test_office_artifact_validation_uses_same_deadline` at
  `tests/msoffice/test_macos_powerpoint_session.py:257-264`

`save()` now performs real `snapshot_artifact_state(...,
deadline=self._deadline)` calls before and after publication
(`macos_powerpoint_session.py:871-875`). A deadline is an absolute monotonic
timestamp, so `123` expires once the Python process has run for 123 seconds.
Round 22 reached this module later and both tests failed in the real snapshot
deadline gate; the same two focused tests passed in 0.17 seconds in a fresh
process. Root additionally measured the fresh macOS/Python 3.9 process
monotonic value near 0.081 seconds. Round 21 happened to reach the module before
the fixed timestamp expired.

This is a deterministic duration-dependent test defect, not a production
deadline regression and not evidence of module pollution. Fix the tests with a
controlled current time/future absolute deadline and a meaningful simulated
artifact (or explicitly controlled snapshot seam), while still asserting that
validation, both snapshots and publication share the same original deadline.
Do not bypass the production snapshot deadline. It is P2 because it makes the
full acceptance suite fail based only on test order/runtime and can mask real
deadline propagation regressions.

Evidence: `docs/verification/microsoft-parity/unit-round22.txt` records **4352
passed, 2 failed, 12 skipped** after 293.24 seconds. The preceding reported
round 21 had 4352 passes and 12 skips because it reached these tests earlier.

## P2 — paragraph-rule PDF checker rejects Word's actual filled rule

The guarded fixture currently accepts a PDF rule only when a drawing has a
stroke `color` matching C0C0C0
(`fixtures/microsoft_parity/macos_word_paragraph_rule.py:108-114`). Word
16.112.3 emitted the native paragraph border as a thin filled rectangle instead:
the retained drawing is type `f`, fill RGB approximately
`(0.752941, 0.752941, 0.752941)`, rectangle
`(88.56,96.96)-(506.64,97.68)`, height about 0.72 pt and width about 418.08 pt,
with `color: null`. Therefore run 02 reports only `pdf_rule_drawn=false` despite
all eleven native, reopen, OOXML, ownership/source and PDF-text checks passing.

The checker should accept the narrowly observed filled-rectangle
representation with color tolerance and strict page/width/height/placement
bounds, while retaining support for a valid stroked representation if intended.
Negative controls must reject arbitrary filled shapes, wrong color, substantial
height/width shifts and wrong placement. Revalidate the immutable run-02 PDF
offline; do not rewrite its original FAIL report or rerun Word solely for this
checker repair.

Retained evidence:

- `docs/verification/microsoft-parity/macos-word-paragraph-rule/run-02/report.json`
- `docs/verification/microsoft-parity/macos-word-paragraph-rule/run-02/pdf-drawings-diagnosis.json`
- `docs/verification/microsoft-parity/macos-word-paragraph-rule/run-02/parent-review.json`

This is a fixture P2 because the current native acceptance gate produces a
known false failure for the exact Word output it is meant to validate. It does
not require a production rule-rendering change.

## Reviewed integration paths with no new finding

- Acknowledged field/structural mutations invalidate the field-topology
  observation at the native submission boundary. Result-only snapshots retain
  the strict external-drift comparison. Checkpoint success saves and rollback
  success restores the tracked field/reference prefixes and the prior observed
  topology.
- Checkpoint preflight compares the immutable prefix plus existing table,
  field, bookmark, floating-shape and generic inline-shape facts before any
  deletion. New or changed drawing objects block rollback. The one open issue
  is the post-submission ordinary-error isolation described above.
- Block degradation fallback runs only after the real rollback returns its
  exact deletion/preimage/postcondition acknowledgement. Malformed or uncertain
  table results do not authorize fallback. Inline/block range commit updates
  structural and field observation state after strict ACK.
- Pagination remains read-only, uses native UTF-16 offsets and Word
  repagination, and accepts session ownership only for its concrete degradation
  range type. Paragraph-rule generation follows the ruled-current-paragraph
  baseline and exact Foundation prefix comparison; its remaining issue is the
  PDF fixture gate above.
- The public session forwards preserve the reviewed signatures and lazy module
  boundaries. The API/SKILL candidate wording distinguishes representative
  macOS evidence from incomplete Windows/full-method parity. The installer
  exclusion uses exact basenames `.superpowers` and `.worktrees`, so normal
  resources with matching prefixes remain copied.

## Source and evidence identity

Current reviewed SHA-256 values:

| File | SHA-256 |
|---|---|
| `macos_word_session.py` | `3ab747a06c63deabc9e0dbf1de51857d8b7dfa843d6ea8e0822915e91e44f205` |
| `macos_word_recovery.py` | `5ee6b2e65e4e13c6a963bd68d0d576fc24ab31cf6fe18fc586ff13356da9090f` |
| `macos_word_fields.py` | `446e8c9f75ce66e93f55b1d8080aa80db6bf5abd4886c75148d8264a8cc50aeb` |
| `macos_word_degradation.py` | `8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7` |
| `macos_word_pagination.py` | `69f9a34e6a8b52d53f3367c6467b903a0260bba75055f716ccd85ee874c677d9` |
| `macos_word_rules.py` | `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006` |
| `macos_powerpoint_session.py` | `adb8fc126b08a61eb000248555f1fed78a23497787ee663c52e203dddee75c56` |
| PowerPoint session tests | `d4b2a6e3c413274653b7b8595cc9067d635b350775c0658253356980b1cfb4f1` |
| paragraph-rule fixture | `0c85598233d98d3521ccf9807eb363dd8fe9a5d9475424935c23794d486146f9` |
| reviewer reproduction | `5d4572763e662ff4e856d8a9f3e80bf2fe64212a062ecf81dbbd8dc43a5cc628` |
| frozen review package diff | `344f5696b6f398bb9d5cbd34d3dcf83c49d8e5e5e2af7ebb6d00ed2e78005c9e` |

Reported broader evidence remains separate: round 21 passed 4352 tests with 12
skips and 336 recorded sources unchanged; recovery run 05 passed 13 native/UI
checks, degradation-notice run 03 passed 57, and pagination run 01 passed 11.
Round 22's two failures are fully accounted for by the fixed-deadline tests
above. Paragraph-rule run 02 remains an immutable FAIL pending offline checker
correction even though its actual native/reopen/XML evidence and retained PDF
drawing support the intended rule. These results do not certify Windows native
behavior or complete the frozen parity baseline.
