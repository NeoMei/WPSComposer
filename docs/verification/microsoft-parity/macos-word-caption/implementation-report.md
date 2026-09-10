# Private macOS Word native caption primitive implementation report

Date: 2026-09-11 (Asia/Shanghai)

Status: production implementation and focused offline verification are complete and frozen. Native run 01 and its separate UI edit acceptance passed. Its permissive fixture validator was corrected offline without changing the archived raw evidence. Root-owned native run 03 exercised the corrected document and handle validators; all ten native/source/inventory checks passed, while the PDF text extractor check failed because overlapping emoji/CJK glyph coordinates changed extracted character order. Root visual inspection confirmed correct rendered captions and no missing glyphs. A fixture-only extractor correction is frozen for native run 04.

## Scope and resulting behavior

This change adds the private macOS Microsoft Word `_add_native_caption` primitive. It preserves the frozen Writer caption contract except for one explicitly approved bug fix. A noncollapsed selection now uses the original selection Start as the caption range Start. The frozen Writer implementation captured the old selection End, so replacing `4..100` with a shorter caption could attempt to format and return `100..14`.

The implementation:

- operates on the current selection of the bound Word window;
- reuses `BoundNumberingCursor`, `add_number_shell`, native field identity bookmarks, owner/kind/category tracking, strict typed acknowledgements, and session quarantine;
- emits global `SEQ` or chapter `STYLEREF` plus `SEQ` fields with Unicode prefix and suffix text;
- places the caller bookmark around the visible number only;
- appends `" " + str(caption)` only when `caption` is truthy;
- centers the acknowledged caption paragraph and sets KeepTogether true;
- applies explicit KeepWithNext false or true only after the base formatting acknowledgement, preserving the frozen late truthiness and failure order;
- validates ordered UTF-16 bounds, inserts the final paragraph mark after the returned span, and returns immutable acknowledged Start/End coordinates;
- does not reset unrelated paragraph or font formatting;
- leaves the equation-specific right-alignment path unchanged; and
- does not add public figure, table, or equation methods.

## Changed files and final frozen SHA-256

- `skills/WPSComposer/scripts/writer.py` — `3e9265751c5f062816d0ff5abe138dd12062f4f545b16127c4de1534cbb965cd`
  - Captures `selection.Range.Start` before replacement, fixing the approved shared long-selection bug.
- `skills/WPSComposer/scripts/msoffice/macos_word_numbering.py` — `9321b6a763dd327f4d9fb4cd985689de2c77f158203fe8a020df409cb8ba5180`
  - Adds `NativeCaptionRange`, ordered caption formatting acknowledgements, explicit KeepWithNext handling, and `add_native_caption`.
- `skills/WPSComposer/scripts/msoffice/macos_word_session.py` — `db2163572c5d8906076eec3832005f3698a771e4f80b9b64129e7c4002cce092`
  - Adds only the private `_add_native_caption` delegation.
- `tests/msoffice/test_macos_word_numbering.py` — `cfd5c650e7923c947c532ac8b9a3c6fb42571164c1e450a6e3894f01ab5efc60`
  - Adds frozen differential tests, real cursor and range tests, strict ACK and quarantine tests, late truthiness ordering, the Writer long-selection regression, guarded fixture tests, closed native readback validation, and three-row handle/paragraph-bound validation.
- `fixtures/microsoft_parity/macos_word_caption.py` — `63396e9f3b2d17e3c4794e98cd2e3cd18854a0383edce5e80a625317a3deed30`
  - Adds the explicit `--execute` root-owned fixture. It uses an exact unsaved sentinel, a long noncollapsed Unicode selection, global and chapter captions, number-only bookmarks, exact fields/body/bounds, nondefault formatting, exact returned handle text, the terminating paragraph mark, the first following UTF-16 unit, save/PDF/reopen, OOXML inspection, source hashes, and inventory cleanup.

No commit or push was performed. Other dirty worktree files belong to the root task and were not edited as part of this implementation.

## RED/GREEN and review-driven fixes

The first focused RED was:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/msoffice/test_macos_word_numbering.py -q --tb=short
5 failed, 68 passed
```

All five failures came from the missing Mac private caption method. The subsequent differential test exposed the inherited long-selection range bug. The approved repair captures the replacement Start in both Writer and Mac and rejects reversed bounds before any format submission.

Independent review then found two ordering risks. The final code acknowledges center and KeepTogether before evaluating `bool(keep_with_next)`, and it lets a native base-format failure occur before custom truthiness. The focused tests cover both cases.

After native run 01, independent review produced three validator counterexamples: changed or appended body text, an extra DATE field, and a negative bookmark Start all passed the original acceptance validator. New RED tests reproduced all three. The corrected validator now requires the closed nine-row schema, exact whole body, exact three-field sequence, exact bookmark names/results, strict nonnegative integer bounds, strict booleans, adjacent paragraph bounds, exact paragraph text and flags, and exact object counts.

The final three-row range oracle is intentionally bounded:

1. `NativeCaptionRange.Start..End` must contain exactly `图😀(1)尾 说明😀` or `表😀[1-1]尾 True`, with centered, KeepTogether, explicit KeepWithNext, and preserved nondefault indent/spacing/italic/font-size values.
2. `End..End+1` must be the inserted `\r` paragraph mark.
3. `End+1..End+2` must be the original following space for the global case or the next `\r` for the chapter case.

This avoids treating the returned immutable coordinates as a live Word Range and avoids claiming unrelated formatting propagation into the following paragraph.

## Verification performed by this implementation task

Final focused suite:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest -q tests/msoffice/test_macos_word_numbering.py
96 passed in 1.88s
```

Focused fixture subset immediately before the full focused suite:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest -q tests/msoffice/test_macos_word_numbering.py -k 'caption_fixture or caption_'
25 passed, 68 deselected in 0.45s
```

Static checks:

```text
python3 -m py_compile fixtures/microsoft_parity/macos_word_caption.py tests/msoffice/test_macos_word_numbering.py
git diff --check -- fixtures/microsoft_parity/macos_word_caption.py tests/msoffice/test_macos_word_numbering.py skills/WPSComposer/scripts/msoffice/macos_word_numbering.py skills/WPSComposer/scripts/msoffice/macos_word_session.py skills/WPSComposer/scripts/writer.py
```

Both exited 0. A focused regression also compiles the generated seed, nondefault-format, format-readback, and final-readback AppleScripts with `/usr/bin/osacompile` without executing them. That test skips locally on non-macOS platforms because `/usr/bin/osacompile` and the Word dictionary are Mac-only. The full focused log is `build/word-caption-20260911/focused-tests-final.log`.

Earlier compatible checks during implementation were:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/msoffice/test_macos_word_fields.py tests/msoffice/test_macos_word_recovery.py tests/msoffice/test_macos_word_pagination.py -q --tb=short
168 passed in 2.94s

/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/longform_m3/test_windows_executor_m3.py -q --tb=short
51 passed in 0.10s

/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/longform_m3 -q --tb=short
885 passed, 7 skipped in 9.89s
```

This worktree has no `.venv`, so all pytest runs used the main checkout's existing virtual environment. This implementation task made no Office or CUA calls.

## Root-owned repository and native evidence

The root task ran the immutable full repository verification after the reviewed production freeze:

```text
5274 passed, 12 skipped in 218s
```

All three production hashes matched the freeze. The only later live difference was this focused test file while the corrected fixture validation was being added.

Root-owned native run 01 used the earlier fixture/test snapshots, while the three production hashes above were already final. `build/word-caption-20260911/native-run-01/report.json` recorded `PASS` and nine true checks. Its hashes include:

- report: `21bf10530fbb02e307b5dad4b796b8f245cca7487d47739dfc2e2aff24b0bf8f`
- DOCX: `699604034e25df45548d9a0fc89e22c6dca594abd6619c1dff7252e6a9221134`
- PDF: `86674c950a4a50baddacec6df4282974638afbb5cee18bfd6e1d332ff4881662`
- document.xml: `edc58ceae2008d98da0697f3c333105c69860399e1007520717990336fae0043`

The raw native rows contain the exact final body, three expected fields, number-only bookmark results `1` and `1-1`, adjacent centered caption paragraphs, KeepTogether true, explicit KeepWithNext false/true, zero tables/shapes, and empty before/final inventories. These raw artifacts are preserved under `docs/verification/microsoft-parity/macos-word-caption/native-run-01/`.

The corrected validator replay is `build/word-caption-20260911/offline-validator-correction-run-01.json`, SHA-256 `82fe1519e064e13dd125fe3f8cf96928080824aa3712d258e6a380bc156507fa`. It reports PASS for the unchanged actual run 01 rows and rejects each of the three recorded corruption variants. This is offline correction evidence; it does not relabel the original permissive validator as sufficient final acceptance.

Root-owned UI acceptance under `docs/verification/microsoft-parity/macos-word-caption/ui-01/` separately exercised two edits to existing text, Undo, explicit Save, close, exact-path reopen, and close. The copied DOCX remained byte-identical to the original and the final Word inventory was empty.

## Failed root-owned native run 02 preflight

The first run 02 attempt failed before caption insertion because the fixture used invalid Word dictionary properties `left indent` and `right indent`. It raised AppleScript compile error `-2741` in the nondefault-format seed. The exact unsaved sentinel was closed, final inventory was empty, and all retained sources matched. The fixture now uses the existing proven `paragraph format left indent` and `paragraph format right indent` properties. No production change resulted from this preflight failure.

## Root-owned native run 03 result and PDF correction

Run 03 reached all caption operations. Exact native body, fields, bookmarks, handle spans, inserted paragraph marks, following content, inherited nondefault formatting, save/reopen, OOXML, source hashes, sentinel, and inventory checks all passed. The sole failure was the pdfplumber text check: Apple Color Emoji glyph coordinates overlap adjacent DengXian CJK glyphs, so extraction returned `left 😀 图 (1)尾 说 😀 明` and `😀表 [1-1]尾 True`. Root visually inspected page 1 and confirmed the rendered strings `图😀(1)尾 说明😀` and `表😀[1-1]尾 True`, centered and italic, with no missing boxes; PyMuPDF reported zero gid-zero glyphs.

The fixture-only correction requires exact whitespace-insensitive non-emoji text and order on four lines, plus emoji counts `[0, 2, 1, 0]` for heading/global/chapter/TAIL. It rejects missing emoji, an emoji displaced to TAIL, suffix deletion or reordering, and extra lines. Raw run 03 remains FAIL and unchanged. `build/word-caption-20260911/offline-pdf-validator-correction-run-03.json` records the bounded offline replay.

## Root-owned native run 04 command

Run only while holding the root Word lease and while the five frozen files above remain unchanged:

```bash
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python fixtures/microsoft_parity/macos_word_caption.py --execute --output build/word-caption-20260911/native-run-04
```

Run 04 must pass the corrected closed validator and the new exact handle/paragraph/nondefault-format checks before those additions count as native evidence.

## Remaining boundaries

- The corrected PDF extractor validator remains pending native run 04 at the time of this revision; the underlying range/format checks passed natively in run 03.
- The shared Writer long-selection repair has offline regression coverage and compatible Windows executor tests. Real Windows Microsoft Word and WPS verification remains a separate gate.
- `NativeCaptionRange` is an immutable coordinate snapshot, not a live COM Range or persistent edit target.
- The primitive remains private. This work does not establish a complete public figure/table/equation API family, release, merge, install, or Windows/WPS native parity.
