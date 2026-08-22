# M3 Task 9 report — Pipeline integration and milestone closeout

Date: 2026-08-23

## Result

Integrated the private normalized-resource lifecycle into the offline/native
boundary, added lazy platform-pure M3 exports, closed long-form diagnostic
privacy without changing executable plan compatibility, documented the M3
Markdown/native API and its intended Python audience, and closed the M3 ledger.
The public `generate()` route and `uv.lock` remain unchanged.

## RED–GREEN evidence

Task 9 began with `tests/longform_m3/test_task9_pipeline_integration.py`.
The first focused run produced four behavior-specific failures:

```text
4 failed, 9 passed
```

The failures proved that M3 integration contracts were not exported, diagnostic
JSON exposed the actual source/payload SHA-256 values, and executor error/timeout
tracebacks retained `PreparedLongformResource` objects. Success already released
the private tuple, which served as the control case. The first GREEN integration
run, including M1/M2 pipeline compatibility, returned:

```text
41 passed
```

An initial eager-export implementation then exposed a real circular import in
the hash-seed subprocess test. The failing broad run was `1 failed, 428 passed,
1 skipped`. Replacing the exports with module-level lazy loading restored parser
purity; the focused circular-import/import/pipeline rerun returned `42 passed`.

## Implementation decisions

- `execute_longform_plan()` still validates before executor contact and passes
  exactly `tuple[PreparedLongformResource, ...]`. A `finally` block drops its
  transport tuple on every exit. Completed exception frames are cleared so a
  caller retaining an error cannot retain normalized payload transports through
  executor frame locals; staged-file deletion remains the concrete executor's
  responsibility and cleanup failure is fatal.
- `LongformBuild.to_json()` keeps its deterministic plan unchanged but removes
  semantic bookmark/reference lookup maps, replaces source/payload hash values
  with redaction markers, and exposes only manifest version/resource count in
  diagnostics. The executable plan retains its required canonical manifest
  digest, controlled bookmark descriptors, and opaque resource IDs.
- `longform.__init__` exposes only platform-pure integration contracts and M3
  value types through lazy imports. It does not import COM, JSAPI, WPS bridge,
  or concrete executor modules and it cannot circularly import `md_parser`.
- The Python API is documented for orchestration agents, plugin maintainers,
  and advanced deterministic-plan integrations. Ordinary users continue through
  the existing `generate()` route.
- M2 structural snapshots are asserted byte-for-byte for all six fixtures.

## Documentation and ledger

`SKILL.md` and `references/api.md` now document figure/table/formula/reference
syntax, object-local caption numbering, native indexes and bounded field
convergence, decoded media allowlist/normalizers/limits, deterministic sizing,
three-line borders, all-or-nothing merges, exact fatal versus object-local
fallback boundary, direct engine-unavailable failure, Python API audience, and
the M3/M4/M5 boundary.

The ledger records every M3 task commit range/review, consumes M2's deferred
cross-reference emission and figure/table index population, records macOS WPS
12.1.26055 run 20, and carries forward only native equation content,
bibliography/citations, general degradation, PDF quality/re-layout, and public
default migration.

## Whole-M3 audit

The fresh audit checked the plan, native executors, add-in, acceptance tests,
reports, evidence metadata, and scope diff against all M3 requirements:

- per-object global/chapter numbering and sequence-transparent unnumbered H1;
- controlled `STYLEREF`/`SEQ`/`REF`, number-only stable bookmarks, populated
  native figure/table indexes, and readable equation source plus number shell;
- bounded single-owner mutation refresh and saved-state convergence;
- figure/table caption placement and page cohesion on both executor paths;
- decoded media signatures, private normalizers, limits, deterministic point
  sizing, explicit orientation, and two-column 12 pt gap;
- academic three-line borders, direct alignment/repeated headers, all-or-nothing
  merge validation, and the one forced-row-split recovery;
- closed plan schemas, canonical determinism, diagnostic/evidence privacy,
  M1/M2 compatibility, and platform-pure imports;
- Windows M3 remains covered by COM mocks only; real Windows stays at the final
  all-milestone gate;
- macOS run-20 evidence proves save/reopen, move/insert/delete, gap-free
  numbering, REF/index updates, stable bookmarks, borders/merges, cohesion, and
  no trailing empty section.

No public routing, release/version, native Office Math, bibliography renderer,
general M4 degradation, or M5 PDF quality work entered M3.

## Independent-review fix waves

The independent reviewer found three Important whole-M3 gaps and one Minor
parity gap. Each received a behavior-specific RED before its fix:

- `design: business` and the other non-academic presets still passed
  `"academic"` into table policy, so an implicit table became three-line. The
  plan now derives one controlled document preset and business/consultant/tech/
  proposal default to grid while explicit styles continue to override it.
- Abstract/list references had semantic IDs and resolved targets but were
  flattened to fallback text during plan emission. Abstracts now emit native
  reference paragraphs. Mixed ordered/unordered lists emit every item in order
  with closed `listFormatting`, controlled marker+tab, `List Paragraph`, 24 pt
  hanging indent, and matching Windows/macOS behavior; plain companion items
  retain the same list layout.
- Windows vertical-group overflow retried an unmerged grid without guarding the
  second insertion. A named recoverable grid failure now rolls back that grid,
  writes the deterministic text fallback, and records `TABLE_INSERT_FAILED`,
  matching macOS. Unknown failures and rollback failures remain fatal.
- Documentation now uses the parser's semicolon-separated merge syntax.

The final independent re-review directly reran Task 9, schemas, both executor
paths, recording/generation-plan compatibility, M2 compatibility, acceptance,
asset manifest, and run-20 digest checks. Disposition: `APPROVED`, with no
remaining Critical, Important, or Minor findings.

## Verification

Focused Task 9/pipeline compatibility:

```text
42 passed
```

All long-form, M0/M2/M3, generation-plan, and recording tests:

```text
1113 passed, 7 skipped in 7.74s
```

Fresh full suite:

```text
1785 passed, 7 skipped in 154.43s
```

The six M2 skips are the existing writer-registration gate. The one M3 skip is
the normal-CI form of the real test already executed without skip in run 20
(`1 passed in 62.59s`). No Task 9 test is skipped.

Run-20 evidence privacy/schema validation, Python compileall, add-in asset
manifest verification, and `git diff --check` also exited successfully.
