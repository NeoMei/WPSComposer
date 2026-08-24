# M3 Task 8 report

Date: 2026-08-23

## Scope

Task 8 closes M3's macOS acceptance gate with byte-stable offline fixtures and
snapshots, a privacy-closed native evidence harness, and a real WPS
save/reopen/mutate/refresh/save/reopen run.  The implementation also repairs
the JSAPI behaviors exposed by the native run: localized chapter fields,
native Heading 1 numbering, sequence-transparent unnumbered headings,
cell-local two-column image insertion, borderless column containers, stable
bookmark ranges, and native move/insert/delete behavior.

## RED-GREEN evidence

The offline acceptance tests were added before their fixtures, snapshots, and
planner behavior.  Initial RED failures covered missing canonical fixtures,
missing descriptors/snapshots, absent active-numbered-chapter tracking, and the
missing macOS mutation bridge operation.  GREEN added four fixture families:

- `chapter_native`: chapter-local figure/table/equation sequences, native
  indexes, resolved and unresolved references, grid and three-line tables,
  valid merges, a tall vertical group, and one/two-image layouts.
- `global_media`: global object numbering and every accepted static/raster
  media normalization route.
- `auto_boundary`: objects before and after a numbered H1, followed by an
  explicit unnumbered H1 whose objects remain in the active numbered chapter.
  The parser/planner emits `numbering: false` and
  `sequenceTransparent: true`; this is not patched in by the runtime harness.
- `degradation`: deterministic invalid-merge, image, caption-missing, and
  recovery results.  Missing captions omit the native caption operation.

Canonical generation-plan and deterministic-failure snapshots are compared as
UTF-8 bytes.  Review-driven RED cases then proved exact moved sequence result
sets, exact TOF/TOT ordering, a REF targeting the moved table, exact bookmark
set stability, true horizontal column geometry, border suppression, and a
closed recursive evidence schema.  The privacy RED set explicitly rejects
`bookmarkNames`, `fieldResults`, `payload`, `resourceHash`, `paths`, arbitrary
nested count keys, absolute/traversing artifact names, and absolute screenshot
names.

## Real macOS WPS evidence

- WPS Office version: `12.1.26055`
- Candidate evidence directory:
  `build/longform-m3/macos-native-20260823-20`
- Closed evidence metadata:
  `build/longform-m3/macos-native-20260823-20/platform-evidence.json`
- Internal native PDFs: `chapter-mutated.pdf`, `global-mutated.pdf`
- Representative PNGs: seven files under `screenshots/`

The no-skip bridge test generated chapter/global documents, saved and reopened
them, moved the second figure and second table before the first objects,
inserted a new captioned table, deleted a disposable table, refreshed fields,
and saved/reopened again.  For each numbering mode it asserts the complete,
gap-free figure/table/equation result sets rather than counts alone.  It also
asserts the exact changed REF (`表 1-2` or `表 2`), exact TOF/TOT title order,
and the exact surviving bookmark set after move/insert/delete.  Ten mutation
refresh phases each converged in two rounds.

The final structure checks passed for native caption/index fields, no field
error text, sequence-transparent unnumbered H1 behavior, exact 12pt two-column
gap, different horizontal image coordinates with aligned vertical placement,
fully disabled container borders, table borders/merges, caption/object and
caption/first-row page cohesion, and absence of an empty trailing section.

## PDF and visual inspection

Both PDFs were rendered with Poppler.  All seven final PNGs were inspected at
original detail:

- chapter cover: intentional cover content, no accidental blank page;
- chapter indexes: exact `图 1-1`, `图 1-2`, `表 1-1`, `表 1-2`, `表 1-3`;
- chapter figures: genuinely horizontal two-column group with no visible
  container borders, followed by the full-width second figure;
- chapter tables/references: exact `表 1-2` REF after movement, intact grid,
  three-line table and merge, unnumbered appendix heading, inserted `表 1-3`;
- global indexes: exact moved figure and table ordering with gap-free global
  numbering;
- global figures: both two-column groups are horizontal and borderless;
- global tables/references: exact `表 2` moved-table REF and inserted `表 3`.

No page showed clipping, overlap, black replacement glyphs, leaked diagnostic
markers, unresolved bookmark/style errors, visible empty container borders, or
an unintended section transition.

## Verification

```text
real no-skip:
1 passed in 62.59s

privacy/schema validation of run 20:
passed

tests/longform_m3:
388 passed, 1 gated skip in 0.89s

tests/longform tests/longform_m0 tests/longform_m2 tests/longform_m3:
925 passed, 7 skipped in 7.69s

full suite:
1766 passed, 7 skipped in 154.38s
```

The one offline M3 skip is the same real test exercised separately with
`WPSCOMPOSER_RUN_REAL_WPS=1`; it is not used to claim native success.  The six
remaining broad-suite skips are the pre-existing M2 writer registration gate.
The independent-review disposition is recorded below after its final pass.

## Independent review

The same independent reviewer performed the final pass after the four
review-driven fixes and directly inspected all seven run-20 PNGs at original
detail.  The reviewer independently revalidated artifact digests, both
DOCX/PDF structures, exact sequence/index/reference results, 12pt borderless
column geometry, bookmark delta/stability, privacy attack rejection, the
`auto_boundary` parser/planner snapshot, and this report.  Final disposition:
`APPROVED`, with no Critical, Important, or Minor findings.
