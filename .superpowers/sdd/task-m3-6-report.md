# M3 Task 6 Report — Windows COM native objects

## Outcome

Implemented the Windows COM mapping for M3 native figures, tables, equation
number shells, cross-references, figure/table indexes, and the shared five-phase
native field adapter.  Verification is COM-free on macOS; no WPS process was
started and no real Windows evidence is claimed by this task.

## RED evidence

The first behavior run was:

```text
pytest tests/longform_m3/test_windows_executor_m3.py -v
collected 5 items
5 failed
```

The failures proved the pre-Task-6 implementation still deferred M3 object
operations, did not validate/stage private normalized resources, did not use the
five-phase adapter, and did not expose controlled field-code helpers.

## Implemented contracts

- M3 object operations now dispatch their already-resolved closed descriptors;
  only bibliography remains in `_M2_DEFERRED_OPERATIONS`.  Legacy M2 object
  shapes retain their compatibility fallback without entering the M3 path.
- Normalized private resources are validated against payload hashes, media type,
  normalizer, logical IDs, and the canonical manifest digest before acquiring
  COM.  Staged locators use opaque temporary names and are removed after success
  or failure.
- Writer emits controlled `STYLEREF 1 \\s`, independent `SEQ WPSC_FIG`,
  `SEQ WPSC_TAB`, and `SEQ WPSC_EQ` fields, plus `REF <bookmark> \\h` fields.
  Bookmarks cover only the complete visible number range.
- Figure rendering uses inline aspect-preserving images, centered/kept image
  paragraphs, stack layout, or a borderless 1x3 native container whose middle
  column is exactly 12pt.  Only explicit landscape descriptors add temporary
  landscape sections.  Named child insertion failures roll back that child and
  retry through the stack path before a visible notice.
- Table rendering applies resolved direct alignments, zero first-line indent,
  repeated headers, disabled initial row splitting, ordered merges, exact
  three-line/grid borders, caption-adjacent notices, and a grid-to-text recovery
  ladder.  A vertical merge group crossing a page rolls back the complete table
  and rebuilds an unmerged splittable grid with `TABLE_ROW_FORCED_SPLIT`.
- Equation handling keeps readable source text and adds only the native
  number/bookmark shell; it does not invoke Office Math.
- Figure/table indexes use `TablesOfFigures.Add` with the internal
  `WPSC_FIG`/`WPSC_TAB` labels.  The Writer composer implements all five
  `NativeFieldAdapter` phases, produces privacy-safe stable snapshots, and the
  executor has one convergence owner.
- Unknown COM failures plus bookmark, field, index, save, resource-integrity,
  and rollback failures are fatal.  Only exact operation allowlist codes
  degrade.  Issue messages do not include native exception text, paths, hashes,
  field results, or bookmark maps.
- Dedicated host creation remains `DispatchEx` only; shared `Dispatch` is never
  used by the long-form executor.

## Verification

```text
focused Windows M3 + existing Windows executor + COM lifecycle:
62 passed

broad longform/generation-plan/recording/lifecycle regression:
922 passed, 6 skipped

full suite:
1686 passed, 6 skipped in 170.49s
```

The six skips are the pre-existing M2 real-bridge acceptance cases reporting
that the WPS writer component did not register within their timeout.  They are
not Windows COM tests and do not weaken the Task-6 mock gate.

## Deferred risk

The COM implementation is intentionally written blind on macOS.  Real WPS
differences in `TablesOfFigures.Add`, table-cell image placement, field result
ranges, and temporary landscape section behavior remain for the final
all-milestone Windows verification gate.
