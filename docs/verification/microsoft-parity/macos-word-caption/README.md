# Private native captions and selection-range repair

The private Mac caption primitive inserts native SEQ/STYLEREF numbering, number-only bookmarks and caption text at the bound selection, centers its paragraph, explicitly sets cohesion and inserts a paragraph mark. No complete public figure/equation/table family is enabled.

A shared Writer/Mac bug used the old selection end for the returned/formatted caption range; replacing a longer selection could produce reversed bounds. The repair captures the replacement start, keeps ordered native range checks, and preserves late keep-with-next truthiness after the base formatting acknowledgement. Scoped source review resolves both findings.

Raw `native-run-01` passes nine checks: long Unicode selection replacement, native global/chapter numbering and bookmarks, exact owner/category tracking, original unsaved sentinel, native DOCX/PDF/reopen and cleanup. Root verifies all source/artifact hashes and the one-page PDF, with zero missing glyphs. The initial fixture validator was overly permissive about extra fields, outside text and negative bounds; final run04 below supersedes that validator for current acceptance. Preserve raw results, not a final family acceptance claim.

`ui-01` separately proves real edits to existing TAIL and caption True, Undo, explicit Save, close, exact-path reopen and close. The copy remains byte-identical to the original DOCX, including native field/bookmark structures. Final Word inventory is empty. Inline CUA screenshots are in the root task; no saved UI PNG is claimed.

Run04 closes bounded inherited nondefault formatting and returned range/trailing-paragraph details. Windows/WPS native parity remains pending. No release, merge or personal install is authorized by these results.

Run02 preserves a fixture-only compile failure in the added nondefault paragraph seed before caption invocation. Word dictionary requires `paragraph format left indent`/`paragraph format right indent`; shorthand names fail compilation. Root independently confirms all source hashes, no remaining sentinel and empty final inventory. No new production failure or passing extension is inferred. Independent review also caught three tests depending on ignored build evidence and a weak one-character-range oracle; both are fixed and reviewed in full48.

## Final bounded acceptance

`native-run-03` remains raw **FAIL** because PDF extraction moved emoji before adjacent italic CJK glyphs and inserted whitespace. All ten native/source/ownership checks passed; root verified the visible PDF. A separate offline correction retains the raw failure. The final PDF validator requires exact per-line non-emoji text/order and emoji counts 0/2/1/0, with visual inspection required separately.

`native-run-04` passes **11 checks** with nondefault italic/size/indent/spacing, exact returned handle, one-unit trailing CR and following character, long Unicode replacement, global/chapter native fields and bookmarks, unchanged unsaved sentinel, DOCX/PDF/reopen and empty final Word inventory. Root independently checks all nine source copies and current hashes, all three artifact hashes, identical native reopens and zero PDF missing glyphs.

`ui-04` separately verifies actual replacement of existing `True` with `CAPTION-CURRENT-CHECK`, Undo, explicit Save, close, exact file-chooser path reopen, visual caption inspection and close. Source and UI copy remain byte-identical; final Word inventory is empty. Screenshots are inline in the root task. Independent scoped review closes both final findings; full48 passes5,281/12 with402 unchanged hashes. No full public figure/equation/table family or cross-platform parity is enabled by these results.
