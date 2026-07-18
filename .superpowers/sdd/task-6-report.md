# Task 6 report: Presentation renderer seam and public Slide plan

## Status

PASS. The public Slide renderer now records the exact nine-operation Task 2
Presentation catalog and the macOS add-in executes that closed plan against the
writable staged presentation. Public macOS routing and production gates remain
disabled, and no installed-WPS gate was run.

## Scope and behavior

- Added `RecordingSlideComposer` with logical image resources and immutable,
  validated Presentation plans.
- Added `SlideComposer.slide_count`, injected `composer_factory=SlideComposer`,
  and removed renderer access to `p._doc`.
- Preserved Windows pacing for the representative document: title slide,
  section slide, two bullet slides, table slide, and image slide.
- Implemented literal dispatch for reset, size, preset, title, section,
  bullets, blank, image, and table without dynamic method lookup or execution.
- All created slides, shapes, pictures, and tables are resolved from the
  presentation returned by the staged writable `Presentations.Open` call.
- Kept the Task 3 lifecycle: `ReadOnly=false`, native `Save()` exactly once,
  no generation `SaveAs`, `Close()`, and nested alert restoration.
- Applied preset backgrounds, colors, and fonts; title/section/bullet
  typography; table header/body styling; and natural-size image aspect fitting.

## TDD evidence

### Recorder, pacing, and renderer seam RED

Tests were written first for `RecordingSlideComposer`, public `slide_count`,
factory injection, closed-catalog output, resources, and six-slide pacing.

Command:

```text
.venv/bin/python -m pytest tests/test_recording_composers.py tests/test_slide_renderer.py -q
```

Observed RED:

```text
4 failed, 5 passed in 0.05s
```

The failures were the expected missing recorder, missing `slide_count`, and
missing `composer_factory` seam. After the minimal Python implementation:

```text
9 passed in 0.03s
```

### Complete Presentation dispatch and validation RED

Executable Node tests were added before replacing the feasibility-only
Presentation executor. The initial focused run observed:

```text
6 failed, 26 passed, 101 deselected in 0.71s
```

The expected failures proved the literal nine-operation catalog and executor
were absent, `slide.set_size` was unsupported, operation envelopes were not
fully closed, recursive/UTF-8 limits were incomplete, and catalog integrity
was not checked.

Two additional edge cases were driven independently through RED and GREEN:

```text
derived title/bullet boxes: 2 failed, 5 passed
safe fractional geometry:  1 failed, 135 deselected
non-object resources:      1 failed, 136 deselected
```

The fixes added pre-open validation for derived title/section/bullet target
boxes and accepted finite safe-range fractional coordinates while still
rejecting negative, overflowing, or four-times-bound geometry.

The resource-container regression ensures even an empty array is rejected
before opening WPS rather than being treated as an empty resource mapping.
Final Presentation-focused GREEN:

```text
36 passed, 101 deselected
```

## Validation and security review

Before `Presentations.Open`, the add-in now verifies:

- exact `{component,operations}` plans and exact `{op,args}` operations;
- a null-prototype dispatch/rule catalog plus an exact `Set` of all nine names;
- Task 2 required/allowed fields and closed preset/color/font/spacing objects;
- 10,000 operations, 10,000 table cells, 100,000-character strings,
  64-level nesting, 2,000,000 UTF-8 plan bytes, and safe finite numbers;
- ordered planned slide count, reset semantics, and referenced slide indexes;
- positive slide size and finite non-negative explicit and derived geometry,
  including safe endpoints and target boxes bounded by four slide dimensions;
- logical image identifiers, fixed loopback resource URLs, missing resources,
  unsafe resources, and unused resources.

No `eval`, `new Function`, arbitrary property path, arbitrary URL, global
active-presentation ownership, GUI automation, permission automation, or
public routing/gate change was introduced.

## Verification

Required focused suite:

```text
.venv/bin/python -m pytest tests/test_recording_composers.py tests/test_slide_renderer.py tests/macos_probe/test_addin_assets.py -q
146 passed in 2.39s
```

Public parser/API regression suite:

```text
.venv/bin/python -m pytest tests/test_markdown_parser.py tests/test_public_api.py -q
3 passed in 0.02s
```

Complete deterministic suite:

```text
.venv/bin/python -m pytest -q
451 passed in 7.49s
```

The following also exited zero without diagnostics:

```text
node --check macos/wps-jsapi-probe/addin/presentation.js
.venv/bin/python -m compileall skills/WPSComposer/scripts
git diff --check
```

Per the Task 6 boundary, no real WPS gate was run; Task 9 owns it.

## Files

- `skills/WPSComposer/scripts/recording_composers.py`
- `skills/WPSComposer/scripts/slide.py`
- `skills/WPSComposer/scripts/renderers/slide_renderer.py`
- `macos/wps-jsapi-probe/addin/presentation.js`
- `tests/test_recording_composers.py`
- `tests/test_slide_renderer.py`
- `tests/macos_probe/test_addin_assets.py`
- `.superpowers/sdd/task-6-report.md`

## Concerns

No blocking concern is known within Task 6. Native installed-WPS behavior is
intentionally deferred to Task 9 and production macOS generation remains off.
