# Word snapshot read repair: native getter diagnosis, no production promotion

2026-09-09. Worktree `.worktrees/office-description`, baseline HEAD `1686155`.

## Outcome

The first dirty-state getter is now isolated: on a fresh read-only private copy of `run-02/before.docx`, `set qr to text object of qpart` for the absent primary header changes native `saved` from true to false. Prefix/suffix hashes, section/page-setup reads, primary header resolution and its shape count all preserve saved=true. The source OOXML contains one section and no header/footer parts or references. This establishes a native side effect for an absent header range getter, not a body edit.

A documented `get story range` alternative preserves saved state for the six empty page-story kinds in a separate fresh copy. It is not yet an accepted replacement: collection enumeration is unsupported, the range identity/bounds projection subsequently failed native JSON serialization, nonempty/hidden/multisection story ownership remains unverified, and the later full candidate timed out in character formatting before completing.

**No production source or test was changed. No quality table insertion, full quality fixture, save/export, UI, global settings, macros or add-in installation was attempted by this worker.** The previous source/test freeze remains current. This work does not establish completion of the three snapshot-dependent quality methods or full Office parity.

## Evidence

### snapshot-read-repair-01

Seven independently acknowledged blocks, each with stderr phase labels and before/after saved, native body hash/End, read-only flag, list-template count, section/paragraph/field/table/bookmark counts.

| Block | saved | List templates |
| --- | --- | --- |
| Prefix/suffix hashes | true → true | 0 → 0 |
| Section reference | true → true | 0 → 0 |
| Page setup reference | true → true | 0 → 0 |
| Page setup values | true → true | 0 → 0 |
| Primary header reference | true → true | 0 → 0 |
| Primary header shape count | true → true | 0 → 0 |
| Primary header text object reference | **true → false** | 0 → 0 |

Native body End/hash stayed `220` / `9e47b6c699d1b28685dd568fac6d75cb875b18c85f53ba7e7b3b7f895168a15b`. The probe stopped before content/metadata reads and before any later page parts once dirty state was acknowledged. Only final integrity/close actions followed. Exact document/window/selection binding guards remained outside the ordinary-read catch and also guarded `close boundDoc saving no`. The private/source DOCX bytes matched, all measured collection counts matched the baseline, and independent final inventory was `[]`. Raw status is `DIAGNOSED`; this is not in-memory preservation success.

### snapshot-read-repair-02

44 independently acknowledged blocks on a new read-only private copy. All keep saved=true and preserve native body hash/End, read-only state and all measured collection counts. Thirty blocks cover primary/first/even header/footer reference, shape count, index, is-header and link-to-previous. Twelve blocks cover their documented get-story command and text-hash/field-count reads, returning empty SHA-256 and zero fields.

The other two blocks preserve their ordinary failures:

- `count story ranges of boundDoc`: -1708, native collection does not understand count.
- `get story type of every story range of boundDoc`: no bound result; subsequent `probeTypes` access gives -2753.

These are not evidence of absent stored hidden parts. The successful get-story ACK plus empty content does not by itself prove a concrete range's typed identity/bounds. This distinction motivated probe03. Exact guarded close with `saving no`, unchanged private/source bytes, final independent inventory `[]`; raw status `DIAGNOSED`.

`native-contract-and-source-facts.json` retains the local SDEF header/footer properties and get-header/get-footer/get-story commands plus the source package facts. Header/footer has no Exists property. Page setup flags cannot substitute for stored-part coverage.

### snapshot-read-repair-03 — retained FAIL and separate parent reconciliation

One new read-only owned copy. The candidate preserves all six section page-part metadata kinds, then traverses the six documented page-story enums using the already-existing field helper's availability/next-story pattern. It attempts typed story identity, bounds, content and field identity. This is a probe-only alternative layout hash; it is not production code and has no accepted nonempty/multisection coverage.

All six chain phases appear in stderr, but the resulting rows fail `jsonRows` outside the per-read ordinary-error catch with **Invalid type in JSON write (NSAppleEventDescriptor), -10000**. The exact descriptor property is not localized by this run. It is incorrect to infer a native API impossibility from this serialization failure, or claim that the candidate resolved story ownership. The final row ACK was not delivered; post-block saved/body state is unknown. No full-snapshot block or repeat block was submitted. The runner stopped immediately, retained its quarantine, and made no subsequent native calls.

Owned path requiring parent reconciliation:

`/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-i1xklyz6/document-cb81bb637ef946c192e2a95469543a9b.docx`

Runtime directory: the owned path's parent. Failing source/log basename: `57727e82e6014861b4f168a7ac028bd8`. Pre-block baseline was saved=true/read-only=true, body End 220 with the hash above, zero list templates. Parent was notified immediately and the native lease returned. Parent subsequently closed the exact URL-matched read-only copy through UI with no save prompt, observed independent inventory `[]`, recovered quarantine, and observed final `[]`. Evidence is `snapshot-read-repair-03-parent-cleanup/report.json`. This establishes cleanup only; in-memory preservation remains unclaimed, and the original probe remains FAIL.

## Identifier repair and implementation boundary

The preceding style prototype's case-insensitive collision is corrected only in the new probe candidate: vectors `qualityStyleBuiltinFlags` / `qualityStyleInUseFlags` are distinct from scalars `qualityStyleIsBuiltin` / `qualityStyleIsInUse`; current name is separate from names vector. The old failing prototype/report remain unchanged. Corrected lines are retained in probe02 `snapshot-candidate-identifier-repair.json` and probe03 `full-candidate-lines.json`; those exact candidates were not executed to completion. A later modified probe05 full candidate was submitted and timed out as recorded below.

Names and in-use/built-in flags continue to cover the full exposed style list; definition properties remain conditional on in-use/custom native styles in the candidate. This is selected native-style evidence only: the 13 linked character definitions in source OOXML not exposed by Word's native collection remain a coverage limitation. No exhaustive OOXML style preservation is claimed.

The implementation task is currently blocked on both a safe, typed, coverage-preserving native snapshot of page-part ranges and a complete snapshot within the unchanged native execution budget. The production getter is proven to dirty an absent header. Probe06 subsequently established that explicit per-property get yields no result for the absent primary-header story, while main-story positive control succeeds. Nonempty/hidden/multisection evidence and reliable missing-versus-real story semantics are still needed before promotion. Full snapshot execution, repeat equality, and independent review remain required. No guard, style coverage, page-part coverage or public schema was silently removed to pass.

No RED/GREEN implementation cycle or regression-suite run is reported because no production repair was promoted. Historical 504-pass evidence is unchanged and is not a fresh result from this worker.

## Unchanged source hashes

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_word_quality.py` | `f3211eae27998d19e123894e54f620a3755c790cdbd23e6df66227151f37ff48` |
| `tests/msoffice/test_macos_word_quality_native.py` | `b3713bbfd0e108a8e5025e85329ec1bc9ff9dbb749db62e3879b30d3942d8954` |
| `skills/WPSComposer/scripts/msoffice/macos_word_session.py` | `47b40d4eabc360fdfcb29244f072b00aa7b286586e95abf5dd848bfa1b83c59e` |

## Evidence hashes

- `snapshot-read-repair-01/probe.py`: `a038e7b20c2df07bee9f4db689d37068230809f61d2829868ae4838f32ca0c71`
- `snapshot-read-repair-01/report.json`: `d6a3b18d3cb76c0375e8493e79a503b8236f9d54db848f78ecd961b7beb468ef`
- `snapshot-read-repair-02/probe.py`: `eefc3c4c4d03faf8d291f21c74621c7b5a85573bb033ba1b15421d542f014460`
- `snapshot-read-repair-02/report.json`: `94ef51634031c6523b1319ce19727ee50f53363c783995d3df518c59bbe52bcd`
- `snapshot-read-repair-03/probe.py`: `0b658048d98bedb987421a802f50f66f1af4c6ab8d28d84a9c5b4f5f6df7b805`
- `snapshot-read-repair-03/report.json`: `e6e43d4dac0c41268effc88fe53ea24a54a43d1dfad6883806bf8bc07c188ded`

## Continuation after parent cleanup

Parent explicitly returned the native lease for further fresh-copy scalar-resolution diagnostics. Probe04 turned the previous serialization failure into an acknowledged ordinary -2763 at the primary-header combined typed story/bounds row. All before/after saved/body/collection facts stayed stable; exact guarded close and final inventory `[]`. No complete snapshot was submitted.

Probe05 separately observes story type as text = empty string, Start/End as integer = ordinary -2763, content as text = empty string, field count as integer = 0. The native `(storyRange is not missing value)` test returns true for all six kinds despite those missing bounds. Therefore it does not establish real story existence and may represent deferred object specifiers. No such result can prove absence of stored hidden parts. The probe-only full snapshot excludes the extra prototype story-range bounds which were not present in the original snapshot. It was submitted before the parent reinforced the real-story/deferred-specifier distinction; no result from it is treated as proof of story identity or production admission.

### Probe05 full candidate timeout and lease return

The full candidate exceeded the unchanged 60-second native operation limit. The last retained stderr phase is the compound character font/shading read (candidate line index 336). The log contains two paragraph-format phase starts and 18 character-format append phase starts. It contains the completed layout-hash append, but no table-state or bookmark-state phase. The last phase does not establish that a particular individual getter hung; total AppleEvent cost may account for the deadline.

The runner stopped immediately with `NativeWordTimeoutError`, retained quarantine, and issued no further AppleEvents. Post-full saved/body/collections are unknown. The entire native lease was immediately returned to the parent for reconciliation. The pre-full state was saved=true/read-only=true, End/hash unchanged, list templates 0, sections 1, paragraphs 13, fields 2, tables 2, bookmarks 9. No repeat snapshot ran.

Owned path: `/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-r14ji_fl/document-25275478f9cd4ab8bba2e6f00bf5022e.docx`. Failing script/log basename: `c6133140e64246338a6cd975fe0b4e8f`. Parent subsequently closed the exact owned URL through UI with no save prompt, independently observed inventory `[]`, recovered quarantine and observed final `[]`; see `snapshot-read-repair-05-parent-cleanup/report.json`. In-memory preservation remains unclaimed and original timeout evidence remains unchanged.

Two gates remain independently unresolved: (1) a concrete/evaluated page-story API that distinguishes an absent story from a deferred object specifier and preserves hidden/multisection content; (2) complete selected-fact snapshot execution within the unchanged native time budget with same-run saved/body/collection preservation and repeated equality. The production implementation remains untouched.

Additional evidence hashes:
- `snapshot-read-repair-04/probe.py`: `dfe2378a78eae1f22fbf1f94c32ac16bb7f5786476be60cebe96efcf62b25687`
- `snapshot-read-repair-04/report.json`: `e2af629d2b7ac1c3eb8a2af0a238c73c9cc48a11e9123142778d72c9a44985fd`
- `snapshot-read-repair-05/probe.py`: `c89bbafd626b797a478becac9a3176bcd4e25d55aeeabbfa54f87fb68ccb692f`
- `snapshot-read-repair-05/report.json`: `88b79dffd402744776af29cbc88e0f0c6fdc7feda48c4b22d4d5ec613f888092`

### Probe06 explicit get and main-story positive control

On a further fresh private copy, both the original and explicitly evaluated `get probeStoryRange` main-text reference return class `story range`, type class `constant`/value `main text story`, integer bounds 0:220, and a real text value whose hash matches the baseline. All saved/body/collection facts remain stable.

The corresponding absent primary-header reference returns an empty class string under text coercion. Crucially, explicit per-property `set probeValue to get ...` followed by inspection of `class of probeValue` raises ordinary -2753 (variable undefined) for type, Start, End **and content**, both with and without the extra `get probeStoryRange`. The previous empty content hash was therefore a coercion of an unresolved/no-result getter, not evidence of a real empty text range. Explicit evaluation alone does not resolve the absent-page-story reference. The main-story positive control validates that these property getters can return concrete values on a real story. Missing-story classification and hidden/multisection coverage still need their own native contract; no inferred empty content will be used for production snapshots.

### Probe07 bounded vector/record feasibility

All 13 blocks preserve native saved=true, body End/hash, read-only flag and every measured collection count. The document's native `characters` collection count is 154 although native End is 220. This collection cannot be assumed equivalent to the original explicit [0,220) native-unit coverage, including hidden field-code ranges. Bulk Start, End, font name, size, bold, italic and underline queries do not produce validated list vectors; color/shading fail the expected nested RGB-vector shape. Ordinary type-shape errors remain in the raw report; no missing/scalar value is accepted as an empty vector. No bulk result certifies coverage or provides a production optimization.

At explicit one-unit native ranges 0:1, 59:60 and 110:111, `get properties of font object` returns useful font values: DengXian 11 regular at the first two, DengXian 14 italic at the third. This narrow observation motivates the final same-coordinate record comparison only. Exact guarded close succeeded, source/private file hashes match, and independent final inventory is `[]`.

### Probe08 final same-coordinate equivalence and lease return

For positions 0, 59 and 110, explicitly coerced `(get properties of font object of range) as record` and shading property records match the old seven-property reads in exact typed native JSON: name, size, bold, italic, underline, RGB color and background RGB shading. All three equality values are true. Per-block wall-clock times are 0.764 / 0.782 / 0.764 seconds; those blocks include both old and new reads plus script startup/binding, so they do not establish a per-unit performance bound or a whole-document speedup. In particular, this report does **not** multiply an independent block's roughly 0.5-second runtime by 220 and call it a strict lower bound. The observed whole-snapshot 60-second timeout remains the relevant failure fact.

All three phases keep saved=true, read-only=true, body End/hash and native collections unchanged. Exact guarded close, source/private-byte equality and independent final inventory `[]` pass. There is no quarantine. The native lease was returned to the parent immediately after this result, for the separate table fixture acceptance. No further native operation will be issued by this worker. A Python string-quoting error before any native action is retained separately as `pre-native-python-syntax-error.log`; it was corrected before the source-bound native run.

## Final conclusion and feasible next implementation directions

No safe complete read repair has been established, so production and tests remain frozen at their initial hashes. No public quality placement/recovery/style/page-part guard is lowered. The nine missing parity methods and the quality acceptance gates remain required. This investigation supplies two independently useful findings: a demonstrated absent-header `text object` side effect, and a demonstrated no-result story reference that text coercion misrepresents as empty content. The current full snapshot also fails its real 60-second budget on this 220-unit source.

A future implementation can evaluate either approach below, with a fresh bounded plan and independent review; neither is enabled here:

1. **Uniform-range property records with recursive subdivision.** Read the existing seven font/shading facts over a native-coordinate range, retain the full interval when the native API proves every property uniform, and split only mixed intervals until exact native-unit coverage is obtained. Before using that representation, a dedicated synthetic mixed-format matrix must prove each property's uniform/mixed/no-result semantics, including Boolean and enum values, RGB colors, field-code/result ranges, table terminators, and UTF-16 boundaries. An aggregate default/empty value cannot be treated as evidence of uniform formatting. Expansion of the compressed result must match the old per-unit facts exactly on source-bound fixtures. Complete snapshot execution must still fit the unchanged 60-second budget, and story coverage remains a separate prerequisite.
2. **A separately validated native batch snapshot transport, such as the previously proposed Office.js route.** It would need to return explicit typed facts and complete native coordinate/story identities in fewer host round trips, preserve body/saved/collection state, and demonstrate hidden/multisection coverage. Specific APIs, runtime availability, installation/permission requirements and supported property semantics remain to be assessed; this report does not claim a ready Office.js implementation or authorize installing one.

Raising diagnostic timeouts might help isolate slow getters, but raising the public timeout alone would not resolve coverage, no-result coercion or the missing performance proof. No further full snapshot was submitted after the parent prohibited repetition.

Final probe hashes:
- `snapshot-read-repair-06/probe.py`: `687728ff49fa851784715ac7aa80d8ca0b1d5db5fb208baa16354060909f0fb3`
- `snapshot-read-repair-06/report.json`: `5a1c9f89e982a5b031b79bc262a8c7f3595fe54da188d24dfe49c37d3c0a19b3`
- `snapshot-read-repair-07/probe.py`: `8b3520f50462cc3fc561f242be32bcbd1bf720c7ada439c11e6d21d0b9f77b53`
- `snapshot-read-repair-07/report.json`: `7fcc8843a1db70ea3dcbff0e71d35d21b02bf669ab53ab48ad149489afc1ded6`
- `snapshot-read-repair-08/probe.py`: `a23c69ce691ad9b3b599f94925219cd7644a23eea0eb7b784a10583fcb0bfaa7`
- `snapshot-read-repair-08/report.json`: `5c16b1ee2bcc4728ba87e247cb2c586a7c20073629cb8370dbd20a3be517652e`
