# Word full-quality snapshot read repair — evidence-bounded handoff

Date: 2026-09-10
Branch / baseline: `codex/microsoft-parity` at `55b27fa`
Scope: Task 4 full-quality snapshot read side effects in `macos_word_quality.py`
Result: **production and permanent tests remain unchanged; no safe complete repair is supported by current native evidence.** A corrected, source-only, parent-owned defined-style probe is prepared at `word-snapshot-defined-style-probe.py`. It was not run against Word by this worker.

## Decision and root cause

The authoritative retained report is `docs/verification/microsoft-parity/macos-word-quality/snapshot-read-repair-report.md` (SHA-256 `972ad08e4c79bb48c236366810d40742886744e9f535181a6de44b441d25c752`). It establishes two distinct native read side effects in the current `_layout_commands()`:

1. On a fresh read-only DOCX with no header/footer package parts or references, `set qr to text object of qpart` for the absent primary header changes Word's native `saved` state from `true` to `false`. Resolving the section/page setup/header object and reading its shape count did not dirty the document. The native body End/hash and collection counts stayed unchanged, so this is a getter side effect rather than a body edit.
2. `get automatically update of every Word style of boundDoc` changes `saved` from `true` to `false` and materializes 13 list templates on the source-bound fixture. Bulk `name local`, `in use`, and `built in` vectors each preserved saved/body/list-template state. Individual `description` and `automatically update` reads for the 20 native styles classified as in-use or non-built-in also preserved that state.

The apparent page-story alternative is not promotable. `get story range` preserves saved state, but an absent primary-header story remains an unresolved/deferred object specifier: explicit reads of class/type/Start/End/content yield no result, while the main-story positive control yields a concrete `story range`, type, bounds and content. Earlier empty-string content was only coercion of an unresolved getter. No evidence yet maps nonempty, stored-but-inactive, linked/unlinked, hidden, and multisection page stories to exact section ownership and bounds.

The style-subset candidate is plausible but also not yet promotable. The previous complete candidate failed because AppleScript treats identifiers case-insensitively: vector `qualityStyleBuiltIn` collided with scalar `qualityStyleBuiltin`, so the second loop iteration attempted `item 2 of true`; the run then ended dirty. The corrected candidate in the retained report was never executed. It must pass independently from page-part reads before any production replacement.

Changing `_layout_commands()` now would therefore either omit existing page-part content/fields, coerce unresolved stories to false emptiness, or promote an unexecuted style classifier. Each option weakens exact equality. This task explicitly forbids that trade.

## Prepared source-only probe

`word-snapshot-defined-style-probe.py` (360 lines; SHA-256 `ab93844c8ee31d4232f4910edc0a81ce69f8236f7de4c0b802de85029fd973d7`) is inert unless the parent supplies `--execute-native --source ... --out ...`. `--emit-plan` performs no native action and prints the candidate and fixture contract.

The candidate uses deliberately distinct identifiers under case folding:

- vectors: `qualityStyleNames`, `qualityStyleInUseFlags`, `qualityStyleBuiltinFlags`;
- scalars: `qualityStyleCurrentName`, `qualityStyleIsInUse`, `qualityStyleIsBuiltin`;
- resolved object/value variables: `qualityCurrentStyle`, `qualityResolvedStyleName`, `qualityStyleDescription`, `qualityStyleAutomaticallyUpdates`.

The probe reads exact ordinal/name/in-use/built-in state for every exposed native style. It reads description and automatically-update only when the corresponding style is in use or non-built-in, resolves the style by the same ordinal, and requires an exact case-sensitive native name match. Its Python ACK validator requires complete ordered ordinals, exact unique names, typed Boolean flags, and exactly one definition row for every and only in-use/non-built-in style. Automatic-update remains an exact typed Boolean or native missing value (`null` in transport JSON); no default is substituted.

Two complete style snapshots must be byte-for-byte equal as typed JSON. Before and after each snapshot, the probe requires `saved=true`, `read_only=true`, unchanged body End/hash, and unchanged list-template/paragraph/field/table/bookmark/section counts. Before close, private DOCX bytes must equal source bytes. The final close/inventory gate remains parent-owned. Even if every check passes, the emitted status is only `STYLE_SUBSET_DIAGNOSED`; it cannot be cited as full-quality, page-part, insertion, save/export, UI, or parity acceptance.

The source contains explicit requirements for three separate source-bound fixtures:

- empty: minimal body, default styles, no custom style, no materialized list template;
- nonempty: applied paragraph/character styles, modified built-in, custom paragraph and character styles, a numbering-linked style, and an unused custom style, with OOXML style/numbering preflight;
- multisection: at least three sections with style use in first/last sections, a table cell, and a field result, while the probe itself performs no page-story reads.

## Required independent page-story probe

The style probe intentionally cannot resolve the header/footer defect. A separate parent-owned probe must use fresh private copies for this matrix:

1. no header/footer parts or relationships;
2. nonempty primary header and footer containing ordinary text and fields;
3. stored but inactive first-page and even-page header/footer parts;
4. at least three sections containing both linked-to-previous and unlinked page parts;
5. hidden text plus field code/result ranges in page stories.

For each of the six story kinds (primary/first/even header and footer), traverse the concrete story chain and require a typed result for story type, section ownership, native Start/End, content hash, shape count, and every field's type/code/result hashes, code/result bounds and locked flag. An unresolved/no-result property is a distinct failure, never an empty story. Cross-check the native chain against source OOXML relationships and section properties. Each read block must preserve saved/read-only/body End/hash and all collection counts, especially list-template count. Repeat the complete page-story snapshot exactly in the same run, close the exact URL-bound owned copy with saving no, compare source/private bytes, and finish with independent inventory `[]`. Any timeout or unacknowledged state returns the native lease to the parent and retains quarantine; it does not authorize another call.

## Precisely unverified `_layout_commands()` fields

The following section values have source-bound evidence that individual reads preserve saved/body state on the single-section empty-page-part fixture: section ordinal, orientation, page width, page height, top/bottom/left/right margins, header distance, footer distance, gutter, and text-column count. Their equality through a complete repeated full snapshot remains unverified.

These page-part fields remain unverified as a fidelity-preserving replacement: header/footer index, `is header`, `link to previous`, content, shape identity beyond the zero-shape fixture, and every page field's type, code content, result content, code Start/End, result Start/End, and locked flag. Section ownership/order for repeated story types is also unverified.

For list templates, only the zero-count fixture and count preservation are established. A nonzero source-bound fixture still must verify template ordinal/name/outline-numbered plus every list level's entry index, linked style, number format, number style, start-at, reset-on-higher, number/text/tab positions, trailing character, and alignment. It must also prove that the style subset itself does not materialize or alter templates.

For styles, bulk localized name/in-use/built-in vectors and individual definition reads are established only on the retained single fixture. The corrected combined/repeated form is unexecuted. Modified built-ins, unused custom styles, character/table/numbering styles, localized-name identity, and multisection usage need the matrix above. The retained source OOXML has 33 materialized styles, including 13 linked character definitions not exposed as independently defined native candidates; that native coverage mismatch remains explicit and cannot be described as exhaustive OOXML style preservation.

## Remaining full `_snapshot_commands()` gaps

The current committed character path resolves font and shading property records once per UTF-16 coordinate and preserves all seven values: name, size, bold, italic, underline, RGB color, and background RGB shading. Native evidence proves exact equality to the prior seven getter form only at coordinates 0, 59 and 110. Full-coordinate expansion, field-code/result coordinates, table terminators, mixed values, non-BMP boundaries, repeated equality, and completion inside the unchanged 60-second budget remain unverified. The last full candidate timed out after 18 character-format append phases, before tables and bookmarks.

Consequently the full snapshot still lacks one same-run, source-bound proof covering all paragraph style/format rows, all seven character values at every native coordinate, main-story fields, table bounds/dimensions/text/row-break flags, bookmarks, the complete layout hash, saved/body/topology/list-template preservation, and a second exact snapshot. Header/footer and style repairs should remain independent until each passes, then be combined for that final full-snapshot gate.

## Pure checks performed by this worker

- `python3 word-snapshot-defined-style-probe.py --emit-plan`: exit 0; no native imports/actions.
- Source-only validator smoke: accepted a hand-derived catalog with one used built-in, one unused built-in and one unused custom style; rejected a definition attached to the wrong ordinal; case-insensitive identifier uniqueness check passed.
- Focused Office-free regression command, excluding the dictionary compiler test: `98 passed, 3 deselected in 0.49s` with the existing clean-dev interpreter.
- `git diff --check` for the probe source: exit 0.

The first attempted focused test command used the absent worktree `.venv/bin/python` and exited 127. A second attempt with system Python found no `pytest`. Neither reached collection. The reported 98-pass result is the fresh successful run with `/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python`.

## RED/GREEN and edit accounting

There is no production RED/GREEN cycle because current native evidence cannot define a safe complete production behavior. Adding a source-text assertion or a test that merely bans the two getters would not prove page-part/style fidelity and would encourage an unsafe implementation. No production file or permanent test was edited.

Current focused hashes remain:

- `skills/WPSComposer/scripts/msoffice/macos_word_quality.py`: `ba0eda21dcb0e4f7314912ab14bcf5a38f5e42a0478bac7ce3aaebc5869edff1`
- `tests/msoffice/test_macos_word_quality_native.py`: `b3713bbfd0e108a8e5025e85329ec1bc9ff9dbb749db62e3879b30d3942d8954`

Pre-existing dirty `docs/verification/microsoft-parity/status.md` and `docs/verification/microsoft-parity/macos-excel-existing-delete-diagnosis/` were not changed by this worker. `skills/WPSComposer/scripts/msoffice/macos_excel_session.py` and `tests/msoffice/test_macos_excel_session.py` became dirty concurrently during the parent-owned Excel work and were also left untouched. No Office/WPS process, AppleScript compiler/runtime, UI, native test, installation, global setting, Git index, commit, or push was used.
