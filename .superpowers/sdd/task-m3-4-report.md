# M3 Task 4 Report — Closed native-object plans

## Result

Implemented closed M3 descriptors, nested schemas, deterministic plan emission,
and recording-composer mirrors for native captions, figures, tables, equation
number shells, cross-reference paragraphs, and figure/table indexes.

## RED evidence

Command:

```text
.venv/bin/python -m pytest tests/longform_m3/test_native_fields.py tests/longform_m3/test_plan_schema.py tests/longform_m3/test_plan.py tests/test_recording_composers.py -q
```

Observed before implementation: `22 failed, 12 passed`. Failures were
behavior-specific: the native-fields module did not exist; M3 nested fields were
rejected as unknown by the shallow M2 schemas; reference runs were not accepted;
and the plan builder did not emit native numbering, media, table, or index
descriptors.

## GREEN evidence

Focused command:

```text
.venv/bin/python -m pytest tests/longform_m3/test_native_fields.py tests/longform_m3/test_plan_schema.py tests/longform_m3/test_plan.py tests/test_recording_composers.py -q
```

Result: `36 passed`.

M1/M2 and generation-plan compatibility command:

```text
.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/test_generation_plan.py tests/test_recording_composers.py -q
```

Result: `591 passed, 6 skipped`. The six skips are the existing real bridge
registration gate in `tests/longform_m2/test_acceptance_m2.py`; no Task 4 test
was skipped.

Additional M3 regression command:

```text
.venv/bin/python -m pytest tests/longform_m3 -q
```

Final result: `220 passed`.

## Files

- Created `skills/WPSComposer/scripts/longform/native_fields.py`.
- Modified `skills/WPSComposer/scripts/generation_plan.py`.
- Modified `skills/WPSComposer/scripts/longform/plan.py`.
- Modified `skills/WPSComposer/scripts/recording_composers.py`.
- Created `tests/longform_m3/test_native_fields.py`.
- Created `tests/longform_m3/test_plan_schema.py`.
- Created `tests/longform_m3/test_plan.py`.

## Decisions

- Native field descriptors map only controlled object kinds to fixed sequence
  IDs and visible affixes. External IDs, captions, and fallback text never enter
  a field descriptor.
- M3 shapes are fully closed and strict while legacy M2 operation shapes remain
  accepted for public compatibility. Presence of the M3 discriminator
  (`numbering`, `runs`, or `sequenceId`) selects the strict schema.
- Caption-missing figures/tables render their object but carry no bookmark and
  cannot become index/reference targets.
- Unresolved inline references remain literal fallback text; resolved references
  preserve text/reference/text order inside one paragraph operation.
- Field/index operations use fatal policy; only the exact figure, table, and
  reference recovery allowlists may degrade.
- Equation operations contain readable source plus native numbering/bookmark
  shell only; native math remains M4 scope.

## Independent review fix wave

### Review RED evidence

Added `tests/longform_m3/test_plan_review_fixes.py` and updated the focused
schema harness so every M3 plan includes its required finalizer.

Initial review reproduction:

```text
.venv/bin/python -m pytest tests/longform_m3/test_plan_review_fixes.py tests/longform_m3/test_plan_schema.py -q
```

Result before fixes: `18 failed, 19 passed`. The failures reproduced mixed
M2/M3 native shapes, unsafe legacy resource/reference policy, target descriptor
mismatch, missing finalizer, late indexes, duplicate semantic ownership,
cohesion flags set false, incomplete merge/degradation semantics, loose digest,
invalid media/normalizer pairs, and fatal-policy recovery smuggling.

Two additional RED checks covered `doc:*` semantic-owner collisions and table
degradation consistency with the resolved merge set. A full-suite collection
run also exposed duplicate `test_plan` module names; packaging `longform_m3`
made the complete repository suite collectible.

### Review GREEN evidence

Focused review and Task 4 suites:

```text
.venv/bin/python -m pytest tests/longform_m3/test_plan_review_fixes.py tests/longform_m3/test_native_fields.py tests/longform_m3/test_plan_schema.py tests/longform_m3/test_plan.py tests/test_recording_composers.py -q
```

Result after the final legacy-child exclusivity check: `61 passed`.

All M3 tests:

```text
.venv/bin/python -m pytest tests/longform_m3 -q
```

Final result: `245 passed`.

M1/M2 compatibility:

```text
.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/test_generation_plan.py tests/test_recording_composers.py -q
```

Result: `591 passed, 6 skipped` (the same existing bridge registration gate).

Full suite:

```text
.venv/bin/python -m pytest -q
```

Final result: `1611 passed, 6 skipped in 154.02s`.

### Review decisions

- Native compatibility mode is selected once for the whole plan. Any mix of
  legacy and M3 native-object shapes is rejected before per-operation schema
  selection.
- Legacy mode remains supported but validates logical resource IDs, closed
  layouts/reference kinds, exact safe recovery policy, and target existence.
- M3 references are checked against a plan-level
  `targetNodeId -> (kind, bookmarkName)` map with exact three-field equality.
- M3 plans require exactly one finalizer at the end; native indexes precede all
  native body objects; semantic owner IDs and figure-child IDs are globally
  unique inside the plan.
- Table merge validation now mirrors Task 3: bounds, header boundary, overlap,
  empty covered cells, vertical groups, and degradation-specific recovery
  shapes are closed together.
- Resource manifest digests are canonical `sha256:` values; all node IDs use
  the same UTF-16-unit bound as nested strings.

## Second independent review fix wave

### Second review RED evidence

The lifecycle/index/bookmark/table/safety reproductions were added before the
implementation changes and run with:

```text
../../.venv/bin/python -m pytest tests/longform_m3/test_plan_review_fixes.py -q
```

Observed result: `8 failed, 25 passed`. The eight failures independently proved
that a body-section index, an index without a captioned/indexable target, a
duplicate index, a cross-owner duplicate bookmark, pre-applied table row split,
nonzero initial cell indent, an arbitrary pure-text degradation policy, and a
pure-text plan without a terminal finalizer were all incorrectly accepted.

Plan-builder fallback was then isolated by replacing the page skeleton with an
empty deterministic skeleton. The figure and table variants each failed with
`DID NOT RAISE ValueError` before their corresponding late-index fallback was
removed.

### Second review GREEN evidence

Focused review tests:

```text
../../.venv/bin/python -m pytest tests/longform_m3/test_plan_review_fixes.py -q
```

Result: `33 passed`.

All M3 tests after the review fixture update:

```text
../../.venv/bin/python -m pytest tests/longform_m3 -q
```

Result: `255 passed`. The final plan-only check, including both late-index
fallback branches, was `7 passed`.

M1/M2/M3 compatibility regression:

```text
../../.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/longform_m3 tests/test_generation_plan.py tests/test_recording_composers.py tests/longform_m0/test_addin_assets.py -q
```

Result: `859 passed, 6 skipped`. The skips remain the existing real-WPS writer
registration timeout gate.

Fresh full-suite verification:

```text
../../.venv/bin/python -m pytest -q
```

Result: `1621 passed, 6 skipped in 154.54s`.

### Second review decisions

- Native-object shape compatibility remains a whole-plan M2/M3 decision, while
  every longform-v2 plan now independently shares the terminal-finalizer and
  closed failure-policy lifecycle, including plans containing only text.
- M3 index validation tracks the active `configure_section` role, permits at
  most one index of each kind, and requires at least one matching object whose
  caption is non-empty and whose binding is indexable.
- Recording plans preserve caller order. They no longer relocate indexes to
  hide invalid body placement; the validator reports the invalid plan.
- The normal plan builder emits indexes in its front-matter role. If its page
  skeleton fails to provide that slot, it refuses to append a late index after
  body construction.
- Bookmark names have a plan-global reverse owner map across figures, tables,
  and equations; two different owners cannot reuse one bookmark.
- M3 tables begin only with `allowRowSplit=false` and `cellIndentPt=0.0`;
  forced row splitting remains expressible only through the exact named table
  degradation descriptor.
