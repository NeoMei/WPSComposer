# Continuation integration fix wave

Date: 2026-09-09. Implementer: word_checkpoint_impl. Worktree: office-description; base cb9898f28e42927248a42421858d549edc54bc3b.

**DONE for the authorized three repairs; frozen for independent review.** Native fault/recovery acceptance and the full suite are root-coordinated follow-ups, not results of this wave. No native Office, UI, install, commit/push, subagents, or other-worktree changes were performed.

The final review and reproduction were checked before editing: review SHA-256 `0a29422b75ac6edbeed44c4aa18401e144c3553e5956e49de4ad8991bda0940c`; repro `5d4572763e662ff4e856d8a9f3e80bf2fe64212a062ecf81dbbd8dc43a5cc628`.

## Repairs

1. **Destructive rollback isolation.** Only `macos_word_recovery.py` production behavior changed. Rollback compiles and validates before setting the existing `_field_topology_mutation_pending` marker. Real `_execute` consumes that marker after script creation/chmod and the final deadline gate, immediately before subprocess submission, while invalidating the observed topology. An exception after that boundary quarantines the session if `_execute` has not already done so. The marker is always cleared. No new shared hook or session change was needed. A completed mutation invalidates observations before ACK parsing; saved tracking/topology/heading state is committed only after the exact valid rollback ACK and postcondition.

   Public `LOCAL_MUTATION_ROLLBACK_FAILED` remains unchanged, and raw private stderr stays in the existing log. Timeout/disconnect/malformed-ACK quarantine remains authoritative; KeyboardInterrupt propagates unchanged and retains logs through quarantine. Argument/preimage/script-I/O failures before submission do not cause partial-mutation quarantine or discard cache/tracking. Deadline expiry keeps the existing session-deadline quarantine reason and preserves the cache. A raw `_execute` entry quarantine gate is also tested separately from a submitted mutation.

   New tests use the real session, real checkpoint/rollback compiler, real transport envelope, and patched subprocess only. Cases cover ordinary -2700, bad ACK, bad postcondition, OSError, TimeoutExpired, KeyboardInterrupt, blocked later sends/context close, exact retained stderr, pre-submit failures, and cache invalidation/restoration timing. Unexpected later subprocess calls would return an ordinary ACK rather than automatically quarantining; call count and session state therefore expose forbidden retries. The older fake-transport bad-ACK expectation now correctly requires discarding observed topology while preserving tracking.

2. **PowerPoint deadline tests.** Both reported tests now share a deterministic clock at monotonic 1,000,000,000 and deadline 1,000,000,060. Temporary artifact bytes, the real snapshot implementation, and a publication stub that copies bytes and invokes its validator prove that initial validation, both snapshots, publication, and destination validation all receive the original deadline. The logical destination digest is checked. Production PowerPoint and artifact transport were not changed, and no real wait is used.

3. **Paragraph-rule PDF checker.** The fixture now accepts either one opaque C0C0C0 filled rectangle of height 0.5–1.0pt or one opaque, solid horizontal stroke of width 0.5–1.0pt. Both require the first A4 page, width 414–422pt, endpoints within 2pt of 88.56/506.64 and center Y within 2pt of 97.32. These bounds target the known fixture, not general document-line detection. Negative controls reject wrong color, transparent fill/stroke, thick/short/vertical/misplaced geometry, wrong page/size, triangles/multiple paths, diagonal and dashed strokes. The real retained run-02 PDF is the positive filled-rule regression. Existing native/style/XML/text assertions are unchanged.

## RED / GREEN evidence

All commands below run from this worktree with:

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python
```

| Evidence file in this ledger directory | Command arguments / result |
| --- | --- |
| `continuation-fixwave-recovery-red.txt` | `-m pytest tests/msoffice/test_macos_word_recovery.py -q -k 'real_destructive or real_rollback or failed_rollback_ack'`: 6 failed, 4 passed. Failures expose missing submission invalidation/isolation. |
| `continuation-fixwave-powerpoint-red.txt` | `-m pytest tests/msoffice/test_macos_powerpoint_session.py -q -k 'publication_uses_session_deadline or office_artifact_validation_uses_same_deadline'`: 2 failed under the controlled large clock with the original deadline 123. |
| `continuation-fixwave-pdf-red.txt` | `-m pytest tests/msoffice/test_macos_word_rules.py -q -k retained_run02`: 1 failed on the actual run-02 filled rule. |
| `continuation-fixwave-green.txt` | `-m pytest tests/msoffice/test_macos_word_recovery.py tests/msoffice/test_macos_powerpoint_session.py tests/msoffice/test_macos_word_rules.py -q -k 'not compile and not hash_helper_native_synthetic'`: **177 passed, 6 deselected**. |
| `continuation-fixwave-reviewer-repro-green.txt` | `.superpowers/sdd/2026-09-08-microsoft-wps-parity/continuation-integration-review-repros.py`: exit 0, original reviewer reproduction GREEN. |

Five PyMuPDF/SWIG deprecation warnings are retained in the final test log; there are no test failures. The six native/Foundation/compile-related tests were intentionally excluded. An intermediate recovery attempt (`continuation-fixwave-recovery-first-green-attempt.txt`) had one test assertion failure because it incorrectly demanded `_retain_evidence=True` on KeyboardInterrupt; the established interrupt path retains logs via quarantine instead. The assertion was corrected to verify quarantine/log preservation without changing interrupt production semantics.

## Immutable run-02 offline correction

Final report: `docs/verification/microsoft-parity/macos-word-paragraph-rule/run-02-offline-correction-02/report.json`, status **OFFLINE_CORRECTED_PASS**.

Run command:

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python .superpowers/sdd/2026-09-08-microsoft-wps-parity/continuation-fixwave-offline-revalidate.py
```

Five artifact checks and both retained native/reopen ACK schema checks pass. The script verifies every original reported artifact digest, hashes the complete original run-02 tree before and after, and records `original_tree_unchanged=true`. The original status remains FAIL with its original `pdf_rule_drawn=false`. This separate correction re-evaluates retained evidence; it does not claim a new native execution. The earlier offline-correction-01 remains retained; correction-02 captures the final production-source closure after the final small rollback cleanup. Both use the same final PDF detector.

Original retained document: SHA-256 `8356e9d3fed2bb88293e7a1dc40a984adfbe316c16c07b5aa150de0711b6108d`. Original PDF: `4e06c33150af61fdd3d4d1b4f84bc81e7a1c2a3fe2916b163c44b55c736efd27`. Corrected reports include all source hashes, copied artifact hashes, original checks and original-tree hashes.

## Self-review and freeze

Compared recovery/fixture changes to their frozen run-02 source copies and PowerPoint tests to the tracked diff. Only the five authorized files changed; the existing shared session, PowerPoint production and paragraph-rule production hashes are unchanged. No new API, deadline, ownership, backend/default, lazy-import or Python 3.9 contract was introduced. `continuation-fixwave-freeze.json` is the machine-readable final source identity. No further source edits after this freeze without root coordination.

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_word_recovery.py` | `20721fb271df8aa7baed94036251cef1787ec96ae1372439cb95239602b7b756` |
| `tests/msoffice/test_macos_word_recovery.py` | `0fd2a13b4f3ad57a7d4933162a0cfb9f6444e0891f6019c4630f518a9cacb346` |
| `tests/msoffice/test_macos_powerpoint_session.py` | `a9a6c35f914598b2953dad5eabd7395409d4c39064c9b813d4a01afc96a1ae24` |
| `fixtures/microsoft_parity/macos_word_paragraph_rule.py` | `6e42887ea2c279ae1e2de6c97c9c7285ba75174c1e296834826251745179fef0` |
| `tests/msoffice/test_macos_word_rules.py` | `ea757ec867b69be649d6722df8a63646a54b95f70b251579a9be50f10705c0b9` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `3ab747a06c63deabc9e0dbf1de51857d8b7dfa843d6ea8e0822915e91e44f205` |
| `skills/WPSComposer/scripts/msoffice/macos_powerpoint_session.py` | `adb8fc126b08a61eb000248555f1fed78a23497787ee663c52e203dddee75c56` |
| `skills/WPSComposer/scripts/msoffice/macos_word_rules.py` | `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006` |
