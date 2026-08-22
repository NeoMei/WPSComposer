# M3 Task 7 Report — macOS JSAPI native objects

## Outcome

Implemented the macOS loopback bridge and Writer JSAPI add-in mapping for the
closed M3 figure, table, equation-number shell, cross-reference, caption-index,
and five-phase field-convergence descriptors.  All verification in this task
uses Python bridge doubles and Node fake-object execution; no real WPS process
was started and no native macOS evidence is claimed here.

## RED evidence

The first Task-7 behavior run collected eight new bridge/add-in/asset cases and
failed all eight.  It proved the pre-Task-7 implementation did not validate or
privately stage normalized resources, still deferred the M3 object handlers,
did not expose the five-phase native adapter, and had no generated-asset drift
gate.

Additional RED boundaries were added before their fixes:

```text
private staging failure boundary: 1 failed, 1 passed
hard native error boundary: 4 failed, 1 passed
two-column fake-object export: 1 failed
required native index Count boundary: 1 failed
Windows page-story/index-span parity boundary: 1 failed
shared visible-result SHA-256 boundary: 1 failed
rollback failure through fail-policy boundary: 1 failed
reviewed fatal/cleanup/SEQ boundaries: 7 failed, 5 passed
private response reference boundary: 2 failed
vertical-overflow grid-to-text boundary: 1 failed
committed asset-manifest boundary: 1 failed
closed issue-code/M3-history review boundary: 2 failed, 2 passed
M2/M3 shape and successful-response completeness boundary: 3 failed
```

These cases prevented staging failures from being downgraded, made save/field/
rollback/resource-integrity failures fatal, exercised exact two-column layout
through fake objects, and rejected a missing required index API.

## Implemented contracts

- The Python executor unconditionally validates the canonical resource manifest,
  payload hash, media type, normalizer, logical resource ID, and figure-child
  references before issuing a bridge command.  Opaque resource files are mode
  `0600`, never use logical IDs in their names, and are removed after success,
  bridge error, timeout, save failure, and rollback failure.  Permanent staging
  or cleanup failure remains a safe fatal error.
- The add-in now consumes the same already-resolved descriptors as Windows.
  Only bibliography remains in `LONGFORM_DEFERRED`; native equation content is
  intentionally not attempted.  Equation rendering retains readable source and
  adds only the native number/bookmark shell.
- Controlled native fields use `STYLEREF 1 \\s`,
  `SEQ WPSC_FIG|TAB|EQ \\* ARABIC` (plus `\\s 1` only for chapter
  numbering), and `REF <controlled bookmark> \\h`.
  Bookmarks cover only complete visible number ranges, and ordered reference
  runs plus their exact fallback remain in one paragraph.
- Figure rendering preserves child order, caption cohesion, stack layout, and a
  two-column native container with an exact 12pt middle gap.  It uses local
  child rollback and atomic column-to-stack recovery; rollback failure aborts.
  A landscape section is created only when the descriptor explicitly requests
  it.
- Table rendering applies caption-above cohesion, repeated headers, initial
  no-row-split, direct resolved alignment, zero first-line indent, exact border
  descriptors, and ordered merges.  Vertical-group overflow rebuilds an
  unmerged splittable grid, while insertion/style/merge failures follow the
  exact grid-to-text recovery ladder without leaving a partial native table.
  The same atomic rollback and text fallback also covers a failed grid rebuild
  after vertical-group overflow.
- The field adapter owns five ordered phases per mutation round: numbering,
  bookmarks/references, indexes, page fields, then snapshot.  Equal adjacent
  snapshots converge immediately.  Three changing rounds receive one read-only
  frozen fourth snapshot and one deduplicated `FIELD_REFRESH_UNSTABLE` issue.
  Header/footer `PAGE`/`NUMPAGES` fields in every section are both updated and
  included in snapshots.  Index counts use actual native range page spans, and
  visible results use NFC-normalized UTF-8 SHA-256.  Missing/throwing required
  field, bookmark, index, repagination, or save APIs are fatal.
- The add-in response is a closed envelope containing only
  `appliedOperations`, controlled `issueCodes`, redacted `fieldSnapshots`,
  ordered `childResults`, `paginationMap`, and the protocol-required output
  locator.  Python ignores the returned locator in favor of its expected staged
  path, requires strict M3 SHA-256 history, binds issue/child/pagination node IDs
  to the validated plan, and accepts only the closed privacy-safe pagination
  stub.  Public issue/child codes are restricted to fixed add-in codes plus
  codes explicitly declared by the validated plan; applied/degraded child
  status and issue-code combinations must agree.  A successful response must
  report exactly every planned operation and exactly the complete ordered set
  of planned figure-child outcomes; omissions, extras, and reordering are
  rejected.  Extra or recursively sensitive response keys are rejected.
- A successful M3-native plan with `writer.finalize_fields` must return a
  nonempty, strictly valid snapshot history.  Native-plan classification
  reuses Task 4's descriptor-shape discriminator (`numbering`, `runs`, or
  `sequenceId`) rather than operation names, so validated legacy M2 object
  shapes retain their explicit missing-history compatibility path.
- Bridge exceptions, timeouts, non-success results, and unknown native errors
  are fatal.  Cleanup still runs on every path; if cleanup also fails, the safe
  primary bridge/save/field/rollback error is retained with
  `cleanup_failed = True` rather than being replaced.
- `writer-longform-v2.js` is verified through a generated SHA-256 asset manifest
  before runtime profile copying.  The manifest was regenerated twice from the
  canonical add-in path and compared byte-for-byte to prove deterministic asset
  generation; no generated file was manually edited after generation.

## Verification

Generated-asset and syntax checks:

```text
../../.venv/bin/python -m skills.WPSComposer.scripts.macos_probe.templates \
  --write-addin-manifest macos/wps-jsapi-probe/addin
second generation + cmp: exit 0
node --check macos/wps-jsapi-probe/addin/writer-longform-v2.js: exit 0
git diff --check: exit 0
```

Fresh focused Task-7/compatibility/assets/probe command:

```text
../../.venv/bin/python -m pytest \
  tests/longform_m3/test_macos_executor_m3.py \
  tests/longform/test_macos_executor.py \
  tests/longform_m0/test_addin_assets.py tests/macos_probe -q
412 passed in 53.42s
```

Broad long-form regression, including the existing M2 native gate:

```text
../../.venv/bin/python -m pytest \
  tests/longform tests/longform_m0 tests/longform_m2 tests/longform_m3 -q
909 passed, 6 skipped in 7.49s
```

Fresh full suite:

```text
1750 passed, 6 skipped in 154.27s
```

The six skips are the existing M2 real-WPS Writer registration gate.  Task 7
forbids starting real WPS, so the skip state is preserved for Task 8 rather than
being represented as fake native evidence.

The same independent reviewer examined the implementation through four rounds.
After the final M2/M3 shape and response-completeness fixes, the reviewer
approved all code boundaries with no remaining code defect.

## Deferred risk

Real macOS WPS JSAPI differences in field/index method signatures, range and
bookmark endpoints, table merge/repeat-header behavior, image sizing, and
section boundaries remain for Task 8's native platform evidence run.  Native
bibliography and native equation content remain explicitly deferred by design.
