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

## Independent-review hardening wave

An independent Task-6 review identified boundary cases that the first fake-COM
suite did not exercise.  The follow-up began with behavior tests for raw
pywintypes-like failures, selection/range side effects, section-story fields,
multi-page index ranges, partially created private files, locked deletion,
column child ordering, inline REF recovery, and empty manifests:

```text
focused review RED:
16 failed, 24 passed
```

The implementation now converts raw failures only at the exact native primitive
that owns the stable code: `AddPicture`, table creation/style/merge, and REF
`Fields.Add`.  Each recoverable object mutation captures its explicit start/end
range; unknown failures and rollback failures remain fatal.  A failed column
child atomically removes the entire local container and rebuilds every logical
child in original order as a stack.  The container itself participates in
caption cohesion through `KeepTogether` and `KeepWithNext`.

PAGE/NUMPAGES refresh now covers every header and footer story in every section
as well as tracked native fields.  Field snapshots use the actual start/end page
span of each TOC, figure index, and table index, with one aggregate count shared
by every snapshot.  REF failure writes its fallback in the same paragraph and
returns exactly one controlled inline issue, which the executor preserves.

Private resource files are registered immediately after creation, including
write/flush/close failure paths.  Deletion is attempted before composer close
and locked files are retried after close; a permanent cleanup failure is fatal
without exposing a locator or hash.  Protocol-v2 manifest validation is now
unconditional, including the canonical empty manifest.

Final review-wave verification (still COM-free; no WPS process started):

```text
focused Task 6 + existing Windows executor + COM lifecycle:
77 passed

all M3:
323 passed

longform + Writer renderer + COM lifecycle:
372 passed

full suite:
1701 passed, 6 skipped in 155.54s
```

The six skips remain the pre-existing real macOS WPS bridge acceptance cases.
The remaining risk is the same platform gate described above: these stronger
fakes prove the failure contracts but do not replace real Windows WPS evidence.

## Second-review closure wave

The second independent review narrowed three remaining boundaries.  New tests
were run before implementation and produced the expected RED result:

```text
Task 6 second-review RED:
8 failed, 38 passed
```

Private staging handles now receive exactly one close retry.  A permanent close
failure still attempts deletion of every file created so far and reports only a
safe fatal error.  Executor teardown is nested so composer close and strict
resource cleanup run normally while `_resource_locators` is cleared in an
unskippable inner `finally`, including cleanup-failure paths.

Figure-child error conversion now ends at the `add_image` boundary.  Failures
from that insertion/size primitive are rolled back and converted to
`IMAGE_INSERT_FAILED`; failures after a successful image insertion while
formatting its paragraph or advancing Selection are rolled back over the same
explicit range and re-raised unchanged as fatal unknown errors.

Non-empty native index ranges now require working Range, Duplicate, SetRange,
and Information APIs.  Missing or throwing native APIs propagate fatally.  Page
span checks use a one-point range at `End - 1`, avoiding an extra page caused by
a trailing paragraph marker; only a truly zero-length index receives the
one-page empty-index compatibility result.

Final second-review verification (COM-free; no WPS process started):

```text
focused Task 6 + existing Windows executor + COM lifecycle:
84 passed

all M3:
330 passed

full suite:
1708 passed, 6 skipped in 154.34s
```

The remaining risk continues to be real Windows WPS object-model validation;
no mock run is presented as native platform evidence.
