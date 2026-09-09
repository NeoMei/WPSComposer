# Rollback failure fixture: file-owned identity correction

Date: 2026-09-09. **Fixture-only correction prepared and frozen for review; no native run performed.** Only `macos_word_rollback_failure.py` and its new dedicated test file were changed. Production and all prior continuation freeze files remain unchanged.

## Root cause and retained failure

The actual run-01 transport log `native-runtime/339ec55a83634bdfbd055ef07a91eb38.log` returned `["binding", null, exactPrivatePath]`. The production `_bind()` deliberately requires a positive window ID only for attached sessions; file-owned sessions validate their private document path and preserve the native missing ID. `_binding()` subsequently selects the private basename and validates its full path. Thus `_window_id=None` is an accepted real file-owned shape, not evidence of failed binding.

The original fixture incorrectly demanded a positive stored window ID before building its independent diagnostic script. Run-01 therefore stopped after appending the one field/table and observing the cache, before injection or rollback submission. The original report remains `FAIL_RETAINED`; no part of this correction reclassifies that run.

Retained original report SHA-256: `ba7594cb4506b2e4fe9e6af10284ef3100e66a74244fd60d8db1de579c95004c`. Original failure.txt: `ceb0b7638f256f46a34c2e4cd79ebe01680d54958fe4401bd49caf3e8cd312c8`. Root independently owns UI discard, sentinel closure and quarantine-marker recovery; this implementer did not perform cleanup.

## Bounded correction

The independent read-only snapshot now selects the exact private basename and requires Foundation exact equality for both full path and native document name. It enumerates `windows of guardedDoc`, recording each native name/id pair without coercing missing values. It never reads active window, active document or global selection. The first successful read stores `{document_name, windows}` as the diagnostic native identity in `owned_binding`; the post-quarantine read must match that exact identity in addition to the existing path, counts, prefix and table checks.

Word.sdef supports the document's window collection at line 3729, and the window name/id properties at lines 31/34. The inspected dictionary SHA-256 is `7cb51b924cab566320cc3019e92c1516302ae11220a81ae6279f1930c26320e7`. Its declaration of integer IDs does not override run-01's actual missing-value observation. Name and ID may each remain null in the diagnostic metadata; the required identity is still the exact file path plus independently read native document name. Window metadata is retained and compared, never converted into a fictitious identifier.

ACK parsing requires the exact header/table schemas and counts, valid digests and types, and a well-formed window list. Incomplete or duplicate window facts, boolean/zero/string IDs, wrong names/paths and before/after identity drift reject the diagnostic. Existing error injection, actual production transport, quarantine checks, blocked write/context-close proof, sentinel checks, source digest and manual-cleanup boundaries are unchanged.

## Pure verification and freeze

RED: the new real `_bind([['binding', None, privatePath]])` regression failed against the original diagnostic interface, retained in `rollback-failure-prep/identity-red.txt`.

GREEN command:

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_rollback_failure.py -q
```

**21 passed in 0.20s**, recorded in `rollback-failure-prep/identity-green.txt`. No Office or osacompile invocation occurred. The None case exercises the real session binding code and confirms the fixture leaves `_window_id` unchanged. All earlier import/CLI/injection/strict-ACK tests remain included.

| Scope | SHA-256 |
| --- | --- |
| Fixture | `80af0630d881199dbf89fbf136bb50cac63906cbffb93df3b9f7a7abd8303a44` |
| Tests | `e0c2d0b87f6f605b913d5dfc71edc3efa151160b77bfbb4760769ae5d8963a1a` |
| Complete shipped source, unchanged | `e3161f3ce0b894d1b1d00edfac84d63199fd8f74e85bfc1ae4532b1c3eb69361` |

Machine-readable freeze: `rollback-failure-prep/identity-freeze.json`. The earlier preparation report/freeze remain historical evidence of the run-01 candidate; this report supersedes their fixture identity and window-ID assumption.

Pending root review and a fresh exclusive lease, use a new output directory:

```sh
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_rollback_failure.py --execute --output docs/verification/microsoft-parity/macos-word-rollback-failure/run-02
```

Native success is still unproven. The expected stopping state remains a quarantined retained owned document requiring root's exact UI discard and guarded recovery.
