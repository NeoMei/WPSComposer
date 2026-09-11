# Independent design/preflight review: Mac Word degradation two-method slice

Date: 2026-09-09  
Scope: `add_inline_degradation` and `add_degradation_notice` only, compared with
the preparation brief, frozen implementation brief, `WriterComposer` direct
baseline, current checkpoint recovery, and actual consumers. No session or
production code was edited and no native application was started.

## Verdict

**DESIGN/PURE SCOPED PASS — no actionable P1/P2 found.** The frozen module
preserves the direct scalar/display contract, validates exact native ACKs before
committing session state, returns the range interface actually consumed by the
baseline, and permits block range fallback only after an acknowledged real
checkpoint rollback.

This is not native acceptance. Word table boundary/text geometry, block style
readback, final-CR behavior and style isolation for following body text remain
mandatory native gates.

## Frozen files reviewed

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_degradation.py` | `be11e3347d3d02a7aa167a8923da82a1646e38eab5616b4f14243cc34f2e7a54` |
| `tests/msoffice/test_macos_word_degradation.py` | `6fcbdfb24f3e59d8a20df3222899a25033feb59ce1e09bf4ee47e058d2bb0ca8` |

The session forwards were intentionally outside this worker's writable scope and
were proposed only in its report. This review does not treat the new module's
presence alone as public-method availability.

## Contract findings

### Display and scalar behavior

The module calls `WriterComposer._degradation_display` directly. It therefore
retains code validation/defaulting, `fallback_text or ""` scalar coercion,
private-text redaction, wrapper de-duplication, and the distinct inline/block
forms. `message` remains inert. Exact `placement == "inline"` delegates;
`None`, empty string, uppercase `INLINE`, numeric/list/mapping values and every
other value retain block behavior.

The added native-safety gate runs after the shared display is formed and before
checkpoint/native calls. It rejects only unpaired surrogates, disallowed control
characters, and more than 1,000,000 UTF-16 units. It does not impose a new plan
schema or narrow safely representable scalar inputs.

### Placement, range, and style

Both paths use the session's exact terminal insertion coordinate and never use
UI selection. Inline inserts exactly the UTF-16 display extent and adds no CR.
Block table creation uses the same collapsed coordinate and adds no Python-side
paragraph boundary. The successful ACK verifies table-count delta, 1x1 geometry,
full table text/bounds, exact display subrange, italic/dark-red/light-red style,
0/3 paragraph spacing, keep-together/body outline and row page-split disabled.

The range fallback applies the same display and block paragraph style to the
exact acknowledged extent after rollback. State (`_pending_heading`, positional
invalidation and field-topology observation) changes only after a valid final
ACK. Malformed or failed insertion does not claim a committed handle.

The module does not reset styling after the inserted range. This matches the
baseline's direct method, which styles only the inserted/table cell range, but
Word may propagate range or table formatting to the terminal insertion point.
Native acceptance must therefore insert ordinary following body text and prove
it does not inherit degradation italic, red font, red shading, keep-together or
body/table boundary artifacts. Pure ACK tests cannot establish that host
behavior.

### Returned handle consumption

`NativeDegradationRange` supplies immutable session identity plus the baseline
attribute surface `Start`, `End`, and `Text`. `NativeDegradationBox` supplies
`.Range` and optional table ordinal. Actual executor degradation dispatch and
media fallback callers ignore these returns. The baseline quality-upsert family,
which is outside this slice, consumes `.Range.End`; the snapshot exposes it.
No current consumer passes these values back for native mutation, so a frozen
snapshot is sufficient for the reviewed two methods and is not represented as a
live COM proxy or durable edit target.

### Table failure and rollback

The native table batch catches ordinary Word construction/style failures and
returns only `['degradation-table-failed', exact_checkpoint_position]`. Timeout,
cancel and invalid-connection codes escape the batch. Python accepts precisely
that single typed row before invoking the real `macos_word_recovery.rollback`.
Malformed success/failure ACKs, transport errors and cancellation quarantine and
cannot trigger rollback/fallback.

Rollback obtains a fresh exact preflight inventory and performs its acknowledged
table/field/range transaction. Only after it returns does the module recheck the
same session gates and append one styled block range at the checkpoint coordinate.
Rollback failure retains evidence, quarantines, and prevents fallback. Fallback
failure is fatal and quarantines; there is no recursive fallback or second
checkpoint. Tracking lists remain owned by recovery and are not fabricated by
the notice module.

## Independent pure verification

The exact focused module passed:

```text
64 passed in 1.06s
```

Coverage includes shared display/redaction/scalars, UTF-16 emoji extents,
immutable range/box handles, every non-inline placement, ACK type/bound/text/style
adversaries, state-before-ACK, ordinary table-failure rollback and fallback,
rollback failure, malformed table-failure ACK, timeout/cancellation, unsafe input
preflight, read-only/closed/quarantined sessions, fallback failure, existing
field/index tracking preservation and checkpoint failure. Log SHA-256:
`67985beb948d4a8a5e3dd1acfc013db16252505cfa0220a89365d08686eed44d`.

`git diff --check` passed. The owner's reported larger combined run is separate
evidence and was not repeated in this bounded review.

## Native acceptance requirements

Before declaring these methods supported on Mac Word, a source-bound fixture
must prove at least:

1. inline display exact text/UTF-16 bounds and no new paragraph;
2. block table actual 1x1/full-range/display-subrange/final-CR geometry and all
   required formatting readbacks;
3. ordinary native table failure creates a partial mutation, returns the closed
   failure ACK, then real rollback restores the exact preimage before the range
   fallback appears once;
4. rollback failure or uncertain completion prevents fallback;
5. ordinary following body text remains independently styled after both inline,
   table-block and block-range fallback paths;
6. saved DOCX/PDF/reopen output, exact source hashes, owned cleanup and unrelated
   unsaved-sentinel preservation.

Those are proof gates for the designed behavior, not new public requirements.
