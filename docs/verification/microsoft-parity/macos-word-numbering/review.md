# Independent Word numbering review — 2026-09-09

Result: **changes requested — three P2 findings**. Read-only production review against `53ad017`; current HEAD became `202641f2e160f5ff7edf68a39a560e276c86ab46` through the unrelated development-setup commit. No production, test, fixture, native document, or UI changes by this reviewer. Only this report was added.

Read AGENTS, regression guardrails, `word-numbering-report.md`, `word-remaining-objects-design.md`, frozen `6dd3a00` Writer methods, and current numbering/session/field/recovery implementations. No `.codegraph` exists in this worktree. Focused tests were executed with the main checkout virtualenv because this worktree has no `.venv`:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/msoffice/test_macos_word_numbering.py tests/msoffice/test_macos_word_fields.py tests/msoffice/test_macos_word_references.py tests/msoffice/test_macos_word_recovery.py -q
219 passed in 6.90s
```

This includes the existing compile-only AppleScript test, not native script execution. Controlled `_execute` transports below exercise real Python orchestration, snapshot parsing and tracking. They do not establish native Word behavior. Native numbering fixture acceptance remains outstanding.

Reusable reviewer regressions: `/tmp/word-numbering-identity-review-red.py` (2 confirmed RED), output `/tmp/word-numbering-identity-review-red.txt`. Run from this worktree with the same Python and `-m pytest -c pyproject.toml /tmp/word-numbering-identity-review-red.py -q`. The implementer also independently retained `/tmp/word-numbering-ack-shape-red.py` and its 6-RED output; this reviewer independently reproduced selection/style/field null-row failures and text/bookmark/format type confusion as detailed below. Reproduction paths were sent to the root and implementer before the fix wave.

## P2 — Bind tracked numbering identities to their actual story

Location: `skills/WPSComposer/scripts/msoffice/macos_word_fields.py:228–230`.

New numbering identities are made from `create range boundDoc ...` under an explicit main-story guard. Snapshot matching searches every story using only kind, field code and story-relative start. A legitimate header STYLEREF with the same code and offset as the owned main-story chapter field therefore produces two matches and raises `NATIVE_WORD_FIELD_IDENTITY_STALE`. No field was changed or retargeted, so field finalization fails on a valid document.

Controlled reproduction returned exactly that error:

```python
code = ' STYLEREF "Heading 1" \\s '
s._tracked_numbering = [
    (NativeIndexHandle('s', 'own', 'eq:one', 'STYLEREF', 'numbering'), code)
]
s._execute = lambda lines: [
    ['stats', 1], ['identity', 'own', 10],
    ['field', 'story:main text/chain:1', 1, 'STYLEREF', code, 10, '1', 0, 0, 1],
    ['field', 'story:primary header/chain:1', 1, 'STYLEREF', code, 10, '1', 0, 0, 1],
]
s.snapshot_fields()  # NATIVE_WORD_FIELD_IDENTITY_STALE
```

The actual `_story_loop` emits `story:main text/chain:1`; chain index is incremented before use. Preserve the full story identity in matching. Existing REF and INDEX creators also create their fields from `boundDoc` main-story ranges and share the older story-less assumption, so a narrowly shared correction can cover all known-main tracked creators. Add numbering/STYLEREF and REF header-collision regressions; do not merely choose the first ambiguous match or discard observational header fields.

## P2 — Preserve insertion-order ordinals for owned fields

Location: numbering integration at `macos_word_fields.py:225–242`, with ordinal allocation at `:253–256`.

Frozen Writer `snapshot_fields()` iterates `_native_fields()` in creation order and assigns `(owner, kind, ordinal)` in that order. The Mac snapshot assigns owned-field ordinals while iterating native document order. This was invisible for append-only insertion, but the new method explicitly permits the current nonterminal selection. Insert a second equation before a previous equation using the same owner (including the default `doc:native`) and kind, and the original field's stable key changes from ordinal 0 to 1. A previously published stable key now refers to another native field.

Controlled reproduction uses tracking order `[first, second]`, valid main-story code-bookmark identities `first=100`, `second=10`, and native rows in document order `[second, first]`. Mac returns the `SECOND` visible-result hash under `('eq:one', 'SEQ_EQ', 0)` and `FIRST` under ordinal 1. Frozen Writer assigns `FIRST` ordinal 0 and `SECOND` ordinal 1.

```python
code = ' SEQ WPSC_EQ \\* ARABIC '
s._tracked_numbering = [
    (NativeIndexHandle('s', 'first', 'eq:one', 'SEQ_EQ', 'numbering'), code),
    (NativeIndexHandle('s', 'second', 'eq:one', 'SEQ_EQ', 'numbering'), code),
]
s._execute = lambda lines: [
    ['stats', 1], ['identity', 'first', 100], ['identity', 'second', 10],
    ['field', 'story:main text/chain:1', 1, 'SEQ_EQ', code, 10, 'SECOND', 0, 0, 6],
    ['field', 'story:main text/chain:1', 2, 'SEQ_EQ', code, 100, 'FIRST', 0, 0, 5],
]
```

Keep native topology collection and its changed-position guards in native story order. Resolve each owned identity against that inventory, but derive semantic ordinals from the tracked insertion order. Add a frozen-Writer differential for snapshots after reverse-position insertion, including the default owner and same-owner repeated SEQ/STYLEREF. Do not reorder the topology inventory to solve semantic ordinal assignment.

## P2 — Malformed numbering ACKs can bypass quarantine

Location: `macos_word_numbering.py:34`, `:100`, `:135`, plus scalar ACK equality at `:89`, `:159`, `:168`.

The constructor, style reader and field reader call `len(rows[0])` before checking that the row is a list. A syntactically valid transport envelope containing `[null]` therefore raises a raw `TypeError` instead of the closed error/quarantine path. This matters most after field creation: `_mutate` has already returned, so its exception handler does not cover the parser failure; the native field may exist, tracking is absent, and the session remains writable. This contradicts the explicit unknown-ACK stopping boundary.

Controlled transport `[ ['selection', 0, 0] ]` followed by `[None]`, calling the real `BoundNumberingCursor.field(...)`, produced:

```text
TypeError object of type 'NoneType' has no len()
quarantined=False tracked=False structural_changed=True
```

`[None]` at selection/style likewise produced unquarantined `TypeError`. Separately, equality-only ACK checks accept booleans as integer bounds and integers as boolean format verification: `[['text', False, True, 'x']]` passes for requested range `[0,1]`, `[['bookmark','b',False,True]]` passes for `[0,1]`, and `[['format',1,1]]` passes. All remained unquarantined in controlled reproduction. The implementer independently flagged these shapes; this review reproduced them without edits.

Validate container shapes before indexing/length operations, enforce exact integer and boolean types, and route every malformed native ACK through `_invalid()` before committing cursor/tracking state or allowing continuation. Cover null/scalar/dict nested rows for the post-mutation field ACK, not only wrong values inside a well-shaped row; retain the current strict field scalar checks.

## Confirmed scope and remaining gates

- The direct public signature, source/fallback coercion, source then TAB before descriptor validation, prefix before fields, late bookmark validation, suffix order, field-end skip, and `{issues: []}` return match frozen Writer in the inspected orchestration and differential tests. The initial saved format coordinate intentionally uses the original selection End, matching frozen `_native_position()` even for a replacement selection.
- Actual scripts address `selection of boundWindow`, check bound-document identity/main story and current selection bounds, use UTF-16 text offsets, and avoid document-end seeking, OMath, table creation and extra font/indent/KeepWithNext changes. Native confirmation of actual selection replacement and paragraph insertion remains a fixture gate.
- Each new field gets a separate code bookmark and numbering owner/category. Native field type/count/position and code/bookmark readback are checked. Ordinary late Python bookmark errors preserve earlier acknowledged fields; uncertain mutation exceptions stop and quarantine. The ACK-shape finding is the remaining bypass identified here.
- Recovery includes numbering tracking in immutable checkpoint prefixes, preflight, exact native bookmark state and acknowledged restoration. No additional recovery regression was found; all focused REF/INDEX/recovery tests passed. This preserves plain-int append-only checkpoints and does not claim middle-selection, cell, shape, OMath or section rollback.
- The new tests do not currently cover cross-story collisions, insertion-order snapshot parity, or malformed nested ACK shapes. Existing 219 PASS cannot close those cases. Native fixture execution, saved/reopened DOCX field/bookmark evidence and development PDF inspection are separate unfinished gates.

## Reviewed production hashes

| File | SHA-256 |
|---|---|
| `macos_word_numbering.py` | `2048cdf59b8f86a50ff7db74f5e9c5ad0c7e8f39d01f31ec11d69966e52ed9ea` |
| `macos_word_session.py` | `df56137d38d793c878659144676b7b1d23ae3f16b34adbee2304ee06399171d9` |
| `macos_word_fields.py` | `5785fa66ac4aa24276280996f46286b685626f6c1dd4ade7fc840bfde6b052f4` |
| `macos_word_recovery.py` | `91d9a954ce7ae9210178998a972023f6bce4928ba4e602207ce63823981b5bdc` |

All paths in the hash table have prefix `skills/WPSComposer/scripts/msoffice/`. Heading, quality RED work and development metadata were excluded.

## Scoped fix re-review — 2026-09-09

**Result: all three original P2 findings are resolved at the hashes below. No remaining code blocker found within this slice before root grants the native fixture lease.** This supersedes the initial changes-requested disposition above while preserving the original failure evidence. Native execution, save/close/reopen acceptance and PDF inspection remain unrun gates; this is permission to proceed with validation, not a native parity completion claim.

Read the updated owner report, changed numbering/field code, relevant test changes and the unchanged session `_execute` submission boundary. Re-ran both original external reproduction files together with the numbering, field, REF, recovery, session and frozen M3/M4 Writer tests:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest -c pyproject.toml /tmp/word-numbering-identity-review-red.py /tmp/word-numbering-ack-shape-red.py tests/msoffice/test_macos_word_numbering.py tests/msoffice/test_macos_word_fields.py tests/msoffice/test_macos_word_references.py tests/msoffice/test_macos_word_recovery.py tests/msoffice/test_macos_word_session.py tests/longform_m3/test_windows_executor_m3.py tests/longform_m4/test_windows_executor_m4.py -q
395 passed in 10.10s
```

This count is the owner's 387-test selection plus the original eight external regressions. No native Word script was executed: the compilation test is compile-only, and transport fault tests replace the subprocess boundary. No production/test/fixture edits were made by this reviewer.

Verified corrections:

1. Tracked numbering/REF/INDEX matching now requires the exact emitted main-story identity `story:main text/chain:1`. The original STYLEREF header collision and the added REF header-collision regression pass. Existing INDEX page-span, growth and stale-identity checks still pass. Changing older controlled rows from `main` to the actual emitted story name is consistent with the existing script protocol.
2. Owned-field ordinal allocation is derived from each tracking list's insertion order and attached to the resolved handle. Native inventory remains in native document order for observation and drift checking. Original reverse-position identity tests and new frozen-Writer differentials for default/explicit owners and SEQ/STYLEREF pass; topology positions remain `[10,100]`. The correction stabilizes semantic keys without weakening or reordering the native topology guard.
3. `_single_row` validates row containers before indexing and length checks; `_exact_ack` distinguishes booleans from integers recursively. The original six ACK repros and expanded malformed-field, bookmark and formatting cases pass. A malformed field ACK reaches `_invalid()`, quarantines, and cannot commit tracking or a new cursor position.
4. Root's additional pre-submission issue is also resolved: `_mutate` calls `_mutation_preflight`, sets the existing pending marker and invokes `_execute` directly. The unchanged `_execute` consumes that marker only after script writing, chmod and its final deadline check, immediately before the subprocess boundary. On pre-submission script-write/read-only failure, observed topology, pending heading and availability are preserved. Final-deadline failure keeps the existing session-deadline quarantine reason without claiming an uncertain mutation. Submitted ordinary native errors clear topology/pending heading and quarantine; malformed ACKs and submitted uncertainty likewise stop. The marker is cleared in `finally`. Real transport boundary tests cover all four cases and pass.

Recovery source remains unchanged from the initial review, including numbering checkpoint prefixes, append-only boundary and acknowledged tracking restoration. The focused recovery tests continue to pass. Frozen source/fallback coercion, selection replacement orchestration, UTF-16 offsets and late-validation tests also remain green.

| Re-reviewed file | SHA-256 |
|---|---|
| `macos_word_numbering.py` | `21307cfa3aeb8e74e0a12139b8b32573a9c984344568488984079bb0948ff31d` |
| `macos_word_fields.py` | `ee51806876fcd859482b97963c1de57e077f5411356ddf85261e10bb09d94b6e` |
| `macos_word_session.py` | `df56137d38d793c878659144676b7b1d23ae3f16b34adbee2304ee06399171d9` |
| `macos_word_recovery.py` | `91d9a954ce7ae9210178998a972023f6bce4928ba4e602207ce63823981b5bdc` |

Hash paths use the same `skills/WPSComposer/scripts/msoffice/` prefix. Root was informed of the clean scoped result before this report append. Native fixture scope and lease remain root-owned.
