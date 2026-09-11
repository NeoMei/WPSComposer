# Continuation fix wave independent re-review

Date: 2026-09-09

Scope: the three P2 findings in `continuation-integration-review.md`, reviewed
against `continuation-fixwave-report.md` and
`continuation-fixwave-freeze.json`. This review changed no production, test or
fixture source, launched no Office/native/UI process, performed no installation,
and made no commit or push.

## Verdict

**SCOPED PASS — all three original P2 findings are ADDRESSED. No remaining or
new worthwhile P1/P2 was found in this fix wave.** This verdict is limited to
the rollback isolation, deterministic PowerPoint deadline tests and paragraph
PDF detector/offline revalidation. It does not certify broader parity, Windows
native behavior, or unimplemented baseline methods.

All eight hashes in `continuation-fixwave-freeze.json` exactly match the current
files. The original review SHA-256 remains
`0a29422b75ac6edbeed44c4aa18401e144c3553e5956e49de4ad8991bda0940c`
and its reviewer reproduction remains
`5d4572763e662ff4e856d8a9f3e80bf2fe64212a062ecf81dbbd8dc43a5cc628`.

## P2 1 — destructive rollback isolation: ADDRESSED

`rollback()` now finishes local argument, session, preimage, tracking and
current-state validation and compiles the complete rollback script before it
sets the existing `_field_topology_mutation_pending` marker. The unchanged real
`MacWordSession._execute()` consumes this marker only after script write/chmod
and its final deadline check, immediately before `subprocess.run`. If an
exception returns after that submission boundary and `_execute()` has not
already quarantined, rollback calls `_retain('Destructive rollback completion
unverified')`. The marker is cleared in `finally`.

This gives the required distinctions:

- ordinary `-2700`, malformed ACK/postcondition, `OSError`, timeout and
  cancellation after native submission cannot leave a usable session;
- subsequent `_execute()` fails locally with `NATIVE_WORD_QUARANTINED`;
- tracking and pending-heading state are not committed on failure, while the
  stale field observation is discarded at actual submission;
- argument/preimage/script-write failures before submission preserve tracking
  and field observations and do not receive the destructive-rollback quarantine;
- a pre-submission expired session deadline retains its established session
  deadline quarantine reason rather than being mislabeled;
- only an exact rollback ACK and exact reconstructed checkpoint state restore
  the checkpoint's field observation/tracking and mark structural change.

The original reviewer reproduction now exits successfully. Its subsequent
send is blocked and its native runner call count remains three. The focused
submission-boundary matrix independently passed **13 tests**.

## P2 2 — PowerPoint fixed deadline tests: ADDRESSED

The two tests no longer use process-uptime-dependent `_deadline = 123`. Their
shared helper fixes the artifact-transport clock at `1_000_000_000`, creates a
single absolute deadline 60 seconds later, and writes actual source artifact
bytes. The real `snapshot_artifact_state` hashes those bytes. The publication
double copies the source, invokes the supplied destination validator and records
the exact passed deadline. Assertions bind, in order, source validation,
source snapshot, publication, destination validation and destination snapshot
to the same original deadline, then verify the published bytes and logical
SHA-256.

This models the absolute monotonic budget contract without sleeping or changing
production PowerPoint/artifact transport. Both tests passed independently.

## P2 3 — Word paragraph PDF detector: ADDRESSED

`pdf_rule_matches()` now accepts the retained Word representation only as an
opaque C0C0C0 filled single rectangle with 0.5–1.0pt height, or as an opaque
solid horizontal C0C0C0 stroke with 0.5–1.0pt width. Both paths require the
first A4 page and narrow expected endpoint, length and Y-position bounds. It
does not accept an arbitrary matching color elsewhere in the PDF.

The retained run-02 PDF passes. Sixteen negative geometry/style/page/path
variants fail, and the intended solid-stroke representation passes. The scoped
detector tests passed **17 tests**.

The offline correction does not overwrite historical evidence. I independently
recomputed all **63** files in the original run-02 tree and found the inventory
exactly equal to `original_file_hashes`; the original `report.json` remains
`FAIL`, retains `pdf_rule_drawn=false`, and has SHA-256
`9d0984d6e4282c1c0078828c361ceb15e73abe53ff789f4ee1a7813b995742de`.
The correction's current checker-source hashes and four copied/revalidated
artifact hashes all match. The retained PDF and DOCX remain respectively
`4e06c33150af61fdd3d4d1b4f84bc81e7a1c2a3fe2916b163c44b55c736efd27`
and `8356e9d3fed2bb88293e7a1dc40a984adfbe316c16c07b5aa150de0711b6108d`.

## Independent verification

Commands used the required Python environment and
`PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps`.

| Check | Result |
|---|---|
| Original reviewer rollback reproduction | exit 0 |
| Rollback real-submission/presubmit focused matrix | 13 passed, 47 deselected |
| PowerPoint deadline pair | 2 passed |
| Retained PDF and detector positive/negative cases | 17 passed, 51 deselected |
| Combined three-file scoped run, excluding compile/native helper cases | 177 passed, 6 deselected |
| Python compilation of all five changed files | pass |

The combined run emitted five existing PyMuPDF/SWIG deprecation warnings. No
test failed. No native acceptance was rerun or inferred from these pure tests.

## Frozen identities

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_recovery.py` | `20721fb271df8aa7baed94036251cef1787ec96ae1372439cb95239602b7b756` |
| `tests/msoffice/test_macos_word_recovery.py` | `0fd2a13b4f3ad57a7d4933162a0cfb9f6444e0891f6019c4630f518a9cacb346` |
| `tests/msoffice/test_macos_powerpoint_session.py` | `a9a6c35f914598b2953dad5eabd7395409d4c39064c9b813d4a01afc96a1ae24` |
| `fixtures/microsoft_parity/macos_word_paragraph_rule.py` | `6e42887ea2c279ae1e2de6c97c9c7285ba75174c1e296834826251745179fef0` |
| `tests/msoffice/test_macos_word_rules.py` | `ea757ec867b69be649d6722df8a63646a54b95f70b251579a9be50f10705c0b9` |
| unchanged `macos_word_session.py` | `3ab747a06c63deabc9e0dbf1de51857d8b7dfa843d6ea8e0822915e91e44f205` |
| unchanged `macos_powerpoint_session.py` | `adb8fc126b08a61eb000248555f1fed78a23497787ee663c52e203dddee75c56` |
| unchanged `macos_word_rules.py` | `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006` |
