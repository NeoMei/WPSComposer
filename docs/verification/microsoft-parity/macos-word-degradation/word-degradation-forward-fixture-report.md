# Reviewed helper: session forwards and guarded fixture ready

2026-09-09. Independent design/pure scoped pass was supplied by root (review
ab509f73). Root granted session.py lease only for two forwards. Native lease has
NOT been granted and no native application or fixture execution occurred.

## Changes

- Added exactly the two reviewed session forwarding methods and tests that invoke
  these public entry points through the real helper and native ACK boundary.
  Other pre-existing session.py changes belong to their original owner.
- Production macos_word_degradation.py remains byte-identical to review:
  be11e3347d3d02a7aa167a8923da82a1646e38eab5616b4f14243cc34f2e7a54.
- Added guarded fixtures/microsoft_parity/macos_word_degradation.py. It requires
  --execute, requires a new output directory, retains its complete designated
  source snapshot, fails if sources change during the run, writes incremental
  report.json, and retains every raw native runtime/failure.
- Six native cases: inline, block, controlled partial-table-failure, each with
  empty and nonempty CJK/emoji terminal text. The failure injection creates and
  fills a real single-cell table before an ordinary caught error. The fixture
  observes the real recovery module rollback returning only after its native
  acknowledgement, checks exact original body restoration, then observes the
  single range fallback submission. It never simulates rollback.
- Each case validates native formatting and UTF-16 handles, unchanged prefix,
  no partial-table remnants, following ordinary paragraph free of notice style,
  native DOCX save/PDF, persisted OOXML notice styles and no-split single row,
  PDF notice occurrence, source-preserving public reopen, native style readback,
  real edit followed by actual checkpoint rollback.
- Each case creates a unique nonempty unsaved sentinel and retains its initial
  name/path/saved-state/content digest. Owned create and reopen contexts must
  close successfully before independent exact close_sentinel is called. The
  fixture reuses the checkpoint fixture's helper, and verifies the original
  document inventory after independent close. On failure, it retains evidence
  and records a still-open sentinel instead of guessing another cleanup route.

## Pure verification

Public forward RED: 2 failed, 64 deselected in 0.44s, both expected missing public
methods. GREEN: 66 passed in 0.94s.

Guard/OOXML fixture RED: 5 expected failures while fixture absent. GREEN including
all two-method tests and fixture pure gates: 71 passed in 2.99s.

Final affected regression:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/msoffice/test_macos_word_degradation.py tests/msoffice/test_macos_word_recovery.py tests/msoffice/test_macos_word_references.py tests/msoffice/test_macos_word_session.py -q --tb=short
225 passed in 11.13s

git diff --check
exit 0
```

The guarded fixture was invoked by one pure CLI test without --execute, which
exited at argument validation without creating its output directory or native
session. OOXML tests use temporary synthetic XML to prove style defects fail the
read-only gate. Neither those tests nor script syntax constitute native proof.

## Pending root review/native lease

Review the fixture and grant serialized native execution explicitly before run.
Proposed invocation (not executed):

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_degradation.py --execute --output docs/verification/microsoft-parity/macos-word-degradation/run-01
```

Native table geometry, shading serialization, following text, PDF, close/reopen,
and sentinel preservation remain unverified for these two helpers. Any mismatch
must retain run-01 as failure evidence. There were no personal installation
writes, commits, or pushes.

## Exact current hashes

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_word_degradation.py` | `be11e3347d3d02a7aa167a8923da82a1646e38eab5616b4f14243cc34f2e7a54` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `a6f39d1191a7df557a7c911904e598bb5fa538430314a84afc9aeef711e0d534` |
| `tests/msoffice/test_macos_word_degradation.py` | `918f67df1087bcdebcc5227176ef67faf725a72419d205c59d6272e7d6389c7d` |
| `fixtures/microsoft_parity/macos_word_degradation.py` | `bd092c5d981fb29dd3bdd83d0b691b17b2cdfe7be1ab30522dcb3c0fbb822afa` |
