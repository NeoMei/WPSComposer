# Word quality stage-one native fixture — preparation frozen

2026-09-09. Prepared in `.worktrees/office-description`. This worker changed only
`fixtures/microsoft_parity/macos_word_quality.py`,
`tests/msoffice/test_macos_word_quality_fixture.py`, and this report. No production
edit, staging, commit, installation, settings change, Word launch, AppleEvent,
native compilation, UI action or native artifact generation occurred.

## Bounded prepared scenario

The first native stage uses the production session forwards for empty
`reserve_document_quality_anchor` and one `upsert_document_quality_notice`, then
an identical identity with changed message. It does not claim the other two
methods or four-method parity.

The synthetic owned document is seeded through the prior measured middle-table
fixture: two old tables, two SEQ fields, per-boundary paragraph/character styles,
semantic bookmarks, Chinese/emoji content and a nonempty suffix. The owned
selection is deliberately noncollapsed `[safePoint−3, safePoint]`, where safePoint
is the already measured nonterminal body paragraph start. Another nonempty,
unsaved sentinel is active at `2:2` during the public reservation. Independent
readback checks the exact selection document path, active sentinel name, actual
selection geometry, fixed empty bookmark, unchanged full native snapshot and
saved numeric cursor.

The first title and notice share one box. An independent native table-2 readback
requires exact display, native UTF-16 bounds, document delta, actual middle
location, three total tables, style/color/shading, paragraph spacing/keep/outline,
row non-splitting, fixed empty bookmark and Python cursor. Duplicate verification
requires no new or changed transport scripts, identical full native snapshot and
unchanged cursor. Known table terminator behavior remains owned by the production
ACK validator; this fixture does not weaken it.

Confirmed reservation, first-notice and duplicate steps are recorded separately.
After the first notice, `partial.docx` and `partial.pdf` preserve that confirmed
state before the duplicate is attempted. Final DOCX/PDF, read-only exact-path
reopen, unchanged source hash and full native snapshot, all OOXML parts, PDF text
and page PNGs are prepared. The XML comparison removes only the one new notice
table and volatile revision/bookmark markers, then compares every other element,
field code/result and paragraph property; `styles.xml` must be byte-equal. Original
bookmark snapshots are retained, but this XML check does not claim independent
arbitrary bookmark-affinity acceptance. PDF text/count checks do not claim visual
layout acceptance: visual PDF inspection explicitly remains pending.

The successful automated verdict is `PASS_STAGE_ONE_AUTOMATED`. The report retains
these pending rows even on that verdict: first `add_document_quality_notice`,
explicit bookmark first-paragraph End, default bookmark Start, invalid/unopened
targets zero-write, consecutive upserts, recovery/fallback failure matrix, UI
edit→Undo→save→reopen, M5 numbering/field convergence and PDF visual inspection.
No stage-gate rejection is converted into a supported capability PASS.

## Guarded lifecycle and retained evidence

Both CLI `--execute` and exact Python `execute=True` are required, before creating
the output directory. The directory must not already exist. Root must separately
grant the exclusive Word lease; this preparation does not grant one. Production
review and the reservation post-submission `.log` I/O quarantine fix remain root
gates before running this fixture.

Cleanup uses the repaired `macos_word_inline_rule_feasibility` independent
inventory and sentinel-close guard, never the old raw-close quality feasibility
wrapper. Owned close must be acknowledged before exact sentinel identity/hash
verification and one sentinel close. A close timeout, malformed ACK, unverified
connection outcome or inventory uncertainty marks the runner quarantined, retains
owned/sentinel identity and prevents all native followup or reopen. A confirmed
sentinel close remains `closed-acknowledged` if the final inventory later fails.
Main-flow cleanup is never retried by `finally`.

Factory failures before `new_document`/read-only `open_document` return a session
also quarantine and prohibit final inventory. The factory's own retained lock/
staging evidence remains authoritative when no session object was returned; the
runner records the attempted stage and requested reopen path for reconciliation.
Any failure while an owned session is live conservatively retains it without a
native recovery attempt. A local quarantine-marker failure is separately recorded
and cannot replace the original native error. Confirmed steps and partial
artifacts survive cleanup failures.

Every shipped Python source, every fixture Python source (including transitive
cleanup imports), this test, and `pyproject.toml` are copied under `source/` and
checked against both live and retained SHA-256 at finalization. Word.sdef is copied
and independently hashed. Both runtime staging trees, original scripts/logs,
failure records and all non-report artifacts are retained and hashed. Native
version is recorded by sentinel creation. No failed output directory is reused.

## Pure verification and preserved failures

Observed development sequence, all without Word:

- Initial test-first run: **12 failed**, each reporting `stage-one quality fixture
  missing` before the fixture existed.
- Initial implementation run: **11 passed, 1 failed** because the independent
  hand-written display length was 53 instead of 54 UTF-16 units. The test data was
  corrected to 54 display units and a 56-unit table span; this was a test-data
  arithmetic mistake, not a native observation.
- Added cleanup retry tests exposed **2 failed, 12 passed**: attempted/closed
  sentinel cleanup was incorrectly attempted again. Fixed with explicit prior
  attempt/closed gates; sentinel identity remains present for final uncertainty.
- A run-level marker-write failure exposed **1 failed, 14 passed**: an OSError
  while retaining quarantine masked the original native failure. A local-only
  wrapper now preserves the primary error and blocks all further native calls.
- Factory/open failure tests exposed **2 failed, 19 passed**: a factory exception
  with no returned session allowed final inventory. Explicit attempted-open
  stages now quarantine those paths.
- The root checkout `.venv` combined run reported **56 passed, 1 failed** with
  `ModuleNotFoundError: No module named 'fitz'` in the existing inline PDF test.
  This environment failure was retained; no dependency/settings modification was
  made. Parent supplied the full clean development environment below.
- Final complete-environment run: **82 passed, 5 warnings in 0.98s**. Warnings are
  PyMuPDF SWIG type deprecations, plus its interpreter-shutdown deprecation line.
  The 21 new fixture tests cover pure gates/validators and run-level cleanup faults;
  existing inline and quality feasibility suites cover the reused safety helpers.

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest \
  tests/msoffice/test_macos_word_quality_fixture.py \
  tests/msoffice/test_macos_word_inline_rule_feasibility.py \
  tests/msoffice/test_macos_word_quality_feasibility.py -q
```

This is a focused preparation suite, not the root's full-suite/native/UI gate.
Run-level tests substitute the external Word session boundary to exercise runner
sequencing and evidence preservation; they do not prove Word's behavior.

## Frozen review inputs

| File | SHA-256 |
| --- | --- |
| `fixtures/microsoft_parity/macos_word_quality.py` | `20c5c6b340eb621def10d584a0a10e8d47bf773f3bd14a6c3d5482277306f05a` |
| `tests/msoffice/test_macos_word_quality_fixture.py` | `78e74145792d7d182bbc03ffa5c9d0fec05ab4555721c5111a771986670df13e` |
| Reused `fixtures/microsoft_parity/macos_word_inline_rule_feasibility.py` | `ebf8a0480a25efa504311ae4d513c0e341ced6b36e264655e1788935e714ad82` |
| Read-only production candidate `skills/WPSComposer/scripts/msoffice/macos_word_quality.py` | `992cb4decbd99b63d62815fb5ba537d3838077203f56fbd037483b1fde6b0c2e` |

The production hash is a preparation observation, not approval to execute that
candidate. Root owns the independent production review and any replacement hash.
