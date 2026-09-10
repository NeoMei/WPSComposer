# Page-story diagnostic result assessment and minimum repair route

Date: 2026-09-10

This assessment is source-only. It does not interpret `STORY_MATRIX_DIAGNOSED` as proof that every requested native property resolved. That status means the diagnostic delivered complete typed rows twice, preserved the measured state and bytes, closed its owned copy, and left inventory empty. A row containing `ordinary-property-error` or `unresolved-property` remains a failed production-coverage observation even when the run has diagnostic status.

## What the current layout snapshot protects

`macos_word_quality._layout_commands()` currently hashes four distinct surfaces:

1. Every section's ordinal, orientation, page width/height, four margins, header/footer distance, gutter, and text-column count.
2. For each section and primary/first/even slot, both header and footer object metadata: header/footer index, `is header`, `link to previous`, content, a zero-floating-shape guard, and every page field's type, code/result content, code/result Start/End, and locked flag.
3. Every native list template's ordinal/name/outline flag and every native list level's index, linked style, number format/style, start/reset, positions, trailing character, and alignment.
4. Every style exposed through Word's native style collection: localized name, description, and automatic-update value.

The surrounding full snapshot separately protects main-story prefix/suffix, paragraph bounds/text and paragraph/list formatting, all seven character values at every native coordinate, main-story fields, tables/row-break flags, and bookmarks. A page-story repair cannot remove, default, or weaken any of those rows. The page-story probe does not test those surfaces.

## What the diagnostic does and does not cover

For each of the six story kinds, the probe attempts concrete class/type/Start/End/content, section count and first-section index, exact field identity, and the complete next-story chain. Every getter result is typed before JSON serialization. Undefined, native missing, empty class, and ordinary errors are distinct rows. Saved/read-only/body/hash/topology/list-template state is checked around each block, and the complete typed output must repeat exactly.

The four fixtures distinguish absent parts, active primary parts, stored but inactive first/even parts, and a three-section explicit/inherited/replaced topology. Their OOXML preflight supplies the exact source truth for evaluating the native chain.

The probe does not read or prove:

- header/footer object `header footer index`, `is header`, or `link to previous` values;
- per-part floating-shape count, shape identity, inline-shape identity, or any shape formatting;
- paragraph/run formatting inside page stories, including the hidden run's `vanish` property;
- page-story content hashing as a separate native value (it retains exact fixture content and repeats its typed JSON instead);
- section page setup or text-column values;
- list-template/list-level values;
- styles or any of the seven main-story character values;
- a native story node's section ownership if both `section-count` and `section-first-index` return typed errors/unresolved values;
- a stored inactive part if Word does not expose it as a concrete story-range node;
- whether linked sections produce one shared node, one node per section, or another chain topology until the multisection native result is cross-checked against preflight.

The existing layout snapshot also does not protect page-story character/paragraph formatting or inline shapes; it only rejects nonzero floating shapes and protects content/fields/metadata. A replacement can preserve the existing guarantee without claiming broader page-story-format fidelity. Any broader parity claim needs a separate fixture and contract.

## Admission checks for the four native reports

Before considering production code, require all four reports to have the frozen fixture/probe hashes, `STORY_MATRIX_DIAGNOSED`, two equal typed snapshots, stable before/after state for every block, private/source byte equality, acknowledged owned close, and final inventory `[]`.

Then inspect the rows rather than the terminal status:

- `absent`: all six roots must be consistently classified as unresolved/error with zero concrete nodes. An empty string, zero bound, or zero fields from a nonconcrete root is not absence proof.
- `primary`: primary header/footer must each be a concrete, fully populated chain node owned by section 1 with the preflight ordinary+hidden content and exact locked PAGE field. The other four roots must remain explicitly absent.
- `inactive-first-even`: all four stored first/even parts must be concrete despite their activation flags being off, with exact content/field data and section 1 ownership. If any is indistinguishable from an absent part, story ranges cannot preserve stored inactive content and the candidate is rejected.
- `linked-multisection`: every native chain must have a deterministic section mapping that distinguishes section 1 ownership, section 2 inheritance, and section 3 replacement. Exact expected node count must be derived from the native result and reconciled with preflight; it must not be guessed. All twelve distinct source parts must be accounted for, or the story-only candidate is rejected.

Any required class/type/bound/content/section/field property reported as ordinary error, native missing, or empty class blocks production use for that fixture. Stable error observations diagnose the API; they do not satisfy fidelity.

## Minimum reliable page-part replacement

If every admission check passes, retain the proven-safe per-section header/footer object reads for shape count, header/footer index, `is header`, and `link to previous`. Remove only `text object of qpart` and all dependent content/field reads. Add one story-chain traversal per story kind for content and page fields, with concrete class/type/bounds and section ownership required before accepting a node.

Normalize the layout hash as two explicit sets:

- `page-part-meta` rows in the existing section/slot/header-then-footer order, retaining exact index/is-header/link values and the zero-shape guard;
- `page-story` and `page-field` rows keyed by story kind, resolved section identity, chain ordinal and field ordinal, retaining exact content, field type, code/result content and bounds, and locked value.

Append an explicit `page-story-absent` row only for a root whose frozen native matrix establishes the same nonconcrete signature as the absent fixture. Do not coerce an unresolved getter to empty text. Require the active, stored-inactive and linked fixtures to prove that the signature does not collapse stored content into absence.

This hybrid is the smallest candidate that retains the old link/index/is-header/shape guard while avoiding the dirty absent-part text-object getter. A story-only replacement would lose link and shape semantics. A metadata-only replacement would lose content and fields.

## Minimum reliable style replacement

Replace the bulk automatic-update getter with the already diagnosed two-tier snapshot:

- `style-state` for every exposed native ordinal: exact localized name, in-use flag and built-in flag;
- `style-definition` for every style whose native state says in-use or non-built-in: exact name, description and automatic-update value, with ordinal/name identity checked before and after.

The nonempty native evidence is encouraging because the modified but unused built-in Heading3 was included in the definition subset. Production admission still requires the completed multisection report and an exact source-bound check that every deliberately modified built-in is selected. The automatic-update value must use Word's native typed result; source `w:autoRedefine` is not interchangeable with it.

This native collection is not exhaustive OOXML style coverage. The linked `WPSC_Custom_Character` definition is absent as an independent native catalog entry, just as linked source definitions were absent in the earlier fixture. Therefore the native subset can replace only the old native-collection observation. A full “all existing OOXML styles preserved” claim requires a separate package-boundary styles/numbering comparison after saving and reopening the private output. If that comparison cannot be made for an attached unsaved session, reject snapshot-dependent structural edits there rather than claiming exhaustive style preservation.

## Implementation and acceptance order

1. Freeze and independently review all four native story reports against `preflight.json`; reject the candidate on any required unresolved/error row or unmapped stored part.
2. Add focused pure tests for the new normalized page metadata/story rows and two-tier style rows. Keep every existing section, list-template, paragraph, seven-value character, main-field, table and bookmark row unchanged.
3. Implement only the two isolated read substitutions in `_layout_commands()`: hybrid page metadata+story traversal and the two-tier native style snapshot. Do not change `_valid_state()` equality to admit missing rows.
4. Run targeted native pre/post snapshots on all story/style fixtures. Require exact repeated state, unchanged list-template counts, and the same package facts after close/reopen.
5. Run the complete quality snapshot and actual table-edit projection in one source-bound session. It must finish inside the unchanged operation budget and preserve every row outside the inserted range, including all seven character values.
6. Save/reopen the private output and compare hidden OOXML styles/numbering and all header/footer relationships/parts against the source, allowing only the intended main-document table edit. Publish only after that gate.

The current 60-second full-snapshot timeout remains independent. A page/style read repair is not releasable merely because its focused probes pass. If the unchanged full snapshot still times out, optimize the seven-value character collection under its separate exact-equivalence fixture or use a validated batch transport; do not raise the timeout or drop coordinates/properties as the repair.
