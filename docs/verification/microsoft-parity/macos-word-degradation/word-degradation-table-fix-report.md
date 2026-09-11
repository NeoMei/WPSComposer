# Native-backed table activation and terminator geometry fix

2026-09-09. Root authorized scoped fixes after run01 and minimal diagnosis.
No session.py changes beyond previous two forwards. Final run02 has NOT run;
awaiting scoped independent review and explicit lease.

## Native evidence

window-diagnosis-01 independently reproduced inactive-window table error -2710,
unchanged native body/end/table count; exact bound-window activation produced 1x1.
Added only `activate object boundWindow` before make-new table in helper.

geometry-diagnosis-01 ran the actual helper's table construction, population and
style emission (including that activation), with extra private readback. Empty
and nonempty CJK/emoji prefixes both produced acknowledged correct style:

- Display `[GEOMETRY_NOTICE] 中文😀` occupies 22 UTF-16 native coordinates.
- Empty: table start 0/end 24, cell start 0/end 23, document end 25.
- Nonempty: checkpoint coordinate 11, table start 12/end 36, cell start 12/end35,
  document end37. Word inserts its own CR before a table after nonempty text.
- Full table Text is display + CR/BEL + CR/BEL. Cell Text is display + CR/BEL.
  Each serialized CR/BEL pair is one native terminator coordinate, not two.
- Native italic=true, dark red=true, light red=true, paragraph 0/3/keep/body=true,
  no-split row=false for allow-break across pages.

Both geometry documents were separate owned sessions; unique foreground sentinels
were unchanged, owned docs closed first, independent exact sentinel closes and
final inventories [] passed. All raw scripts/logs, private errors, copied sources
and reports remain under docs/verification/microsoft-parity/macos-word-degradation.

## Changes and strictness

Parser now accepts exactly the observed table serialization and native bounds.
No arbitrary terminator stripping or relaxed success detection. Malformed ACK
continues to quarantine; range fallback remains distinct from table success.
Handle.Text preserves actual serialization while Start/End preserve native facts.

Fixture normal-tail lookup now resolves the paragraph's native start/end by
exact text match. Counting full-body UTF-16 overcounts table terminator text;
using native paragraph geometry avoids an incorrect block formatting probe.

## Verification

Activation RED: 1 failed, 71 deselected in 0.24s (inactive transport refused table).
Activation GREEN: 72 passed in 0.61s.
Native geometry RED: 2 failed, 72 deselected in 0.19s (observed ACK rejected).
Combined helper/recovery/references GREEN: 169 passed in 3.09s.
After fixture native-paragraph lookup correction: 74 passed in 0.59s.
`git diff --check` clean. No final native acceptance run, personal install,
commit or push occurred. Word lease released after diagnostics with inventory [].

## Frozen hashes for review

- helper: 8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7
- tests: 07396fb846503c098e132e21dd009f90c62544f702636cbd096063fb750a651b
- fixture: 39e1ab6b0a2fe7a7ed180b0d8c4845a2d76a269e92d8f63058b0ffbfab3fa8b6
- session: a6f39d1191a7df557a7c911904e598bb5fa538430314a84afc9aeef711e0d534
