# M3 Task 5 Report — Shared native-field convergence

## Result

Implemented the platform-pure five-phase native-field adapter contract and the
single bounded convergence engine. M2's public
`finalize_fields_with_convergence` facade now delegates exactly once through a
legacy adapter while retaining existing `ConvergenceResult`, `ExecutionIssue`,
`ExecutionOutcome`, pagination-map, and short-hash fixture behavior.

## RED evidence

Command:

```text
../../.venv/bin/python -m pytest tests/longform_m3/test_field_contract.py tests/longform/test_executor.py -q
```

Observed before implementation: collection stopped with one expected error,
`ModuleNotFoundError: No module named
'skills.WPSComposer.scripts.longform.field_contract'`. The new behavior tests
already imported the required five-method protocol, convergence function,
closed field kinds, visible-result snapshot hashing, and fatal contract error.

## GREEN evidence

Focused contract and compatibility command:

```text
../../.venv/bin/python -m pytest tests/longform_m3/test_field_contract.py tests/longform/test_executor.py -q
```

Result: `32 passed`.

M2 executor/platform compatibility command:

```text
../../.venv/bin/python -m pytest tests/longform_m3/test_field_contract.py tests/longform/test_executor.py tests/longform/test_windows_executor.py tests/longform/test_macos_executor.py tests/longform_m2/test_acceptance_m2.py -q
```

Result before the final delegation-test addition: `144 passed, 6 skipped`.
The six skips are the existing real-WPS writer registration gate in M2
acceptance; no Task 5 contract test was skipped.

Broad long-form regression:

```text
../../.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/longform_m3 -q
```

Result: `693 passed, 6 skipped` with the same existing M2 bridge gate.

Static verification:

```text
../../.venv/bin/python -m compileall -q skills/WPSComposer/scripts/longform/executor.py skills/WPSComposer/scripts/longform/field_contract.py
git diff --check
```

Both commands exited successfully.

## Files

- Created `skills/WPSComposer/scripts/longform/field_contract.py`.
- Modified `skills/WPSComposer/scripts/longform/executor.py`.
- Created `tests/longform_m3/test_field_contract.py`.
- Extended `tests/longform/test_executor.py`.

## Decisions and invariants

- Each mutating round has exactly one ordered invocation of numbering,
  bookmarks/references, indexes, and page fields, followed by one snapshot.
- Two equal adjacent canonical snapshot digests converge. Three changing rounds
  run one additional snapshot only; no mutating phase runs in the frozen fourth
  diagnostic round.
- Stable keys accept only the ten controlled field kinds and reject empty
  owners, negative/bool ordinals, duplicates, invalid counts, and non-SHA-256
  M3 hashes.
- `snapshot_visible_field` normalizes CRLF/CR to LF, then NFC-normalizes and
  hashes UTF-8. It never stores or serializes the visible field result.
- Required phase absence or exceptions raise a sanitized fatal
  `NativeFieldContractError`; platform exception text is retained only as the
  Python cause and never copied into public issue evidence.
- `FIELD_REFRESH_UNSTABLE` contains only aggregate counts. It contains no
  visible result, result hash, bookmark mapping, resource identifier, local
  path, or reference target text, and exactly one issue is returned.
- The M2 facade has one delegation call into the shared engine. Its adapter
  preserves legacy `refresh_fields(round_index)` snapshots, including historic
  short test hashes. Native platform adapters in Tasks 6 and 7 must implement
  the strict five-phase protocol directly; no COM/JSAPI method was added here.
- Move/insert/delete simulations prove SEQ/REF result hashes change while owner
  stable keys and private bookmark ownership remain stable.

## Scope checks

- No Windows COM or macOS JSAPI implementation was added.
- No plan/schema/resource/public generation API was changed.
- `uv.lock` is unchanged.
