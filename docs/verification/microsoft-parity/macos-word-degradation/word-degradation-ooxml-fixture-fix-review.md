# Mac Word degradation OOXML fixture fix review

Date: 2026-09-09  
Verdict: **SCOPED PASS — no actionable P1/P2 found.** This review covers only the fixture-side effective-shading correction and its pure tests. Production remained unchanged; no native Office/UI job, personal installation, commit, or push was performed.

## Frozen inputs

| File | SHA-256 |
|---|---|
| `fixtures/microsoft_parity/macos_word_degradation.py` | `8125aa442ddae81d7daed4b8421058d5f50de6953d12b63e346dafc525a7c7ed` |
| `tests/msoffice/test_macos_word_degradation.py` | `0657a1bf7ba21fd37549b6454bad728b7fac5ca36f3299e7b8f790b14417d5b6` |
| unchanged `skills/WPSComposer/scripts/msoffice/macos_word_degradation.py` | `8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7` |
| owner report | `57dd14a5265cd1d6aa31f495b87f4c7afa443396cc5fc0101737b0f5f46a4611` |

## Review result

`_effective_shading` follows the actual element parent map for each text run. It checks, in precedence order, that run's `rPr/shd`, its paragraph's `pPr/shd`, its enclosing cell's `tcPr/shd`, and the closest enclosing table's `tblPr/shd`. Traversal stops at that table, so a sibling or otherwise unrelated table cannot satisfy the notice shading requirement.

The first explicit declaration terminates inheritance. `val="nil"` yields no effective fill; `fill="auto"`, a missing fill (represented as `auto`), and any other explicit fill remain their own effective values and cannot expose an ancestor's `FCE8E6`. Native `val="clear" fill="FCE8E6"` remains valid. This closes the original false negative without accepting a run/paragraph/cell override hidden by table shading.

The following paragraph now uses the same effective ancestor-aware check for each run and rejects inherited `FCE8E6`, while retaining direct italic/red/shading checks. Its parent chain in the actual run02 document ends at body rather than the notice table, so the table cannot contaminate or falsely certify following text.

The eight focused cases cover actual cell+table persistence, table-only inheritance, run `auto`, run clear without fill, a different run fill, paragraph `auto`, a different cell fill, and an unrelated table. I additionally exercised `val="nil"`, clear-without-fill and a different cell fill directly; all resolved with the intended inheritance-stopping semantics.

## Actual saved artifact revalidation

I ran the corrected read-only verifier against:

`docs/verification/microsoft-parity/macos-word-degradation/run-02/block-empty/notices.docx`

It returned exactly:

```json
{"block_paragraph_verified": true, "normal_tail": "FOLLOWING-NORMAL-block-empty", "notice_characters": 24, "table_count": 1}
```

The artifact SHA-256 was identical before and after inspection:

`3f0d8bd6f7fa7ecdda88ccbced238acc3c74727418b85a268c23bf9d677fd80c`

The actual XML stores `FCE8E6` at both the notice cell and enclosing table; the notice paragraph has no direct shading, and the following paragraph is outside the table with no effective notice shading. This revalidation fixes the saved-XML checker only. It does not relabel the incomplete run02 as PASS or prove its unexecuted cases.

## Pure verification

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest -q \
  tests/msoffice/test_macos_word_degradation.py

82 passed in 0.56s
```

`git diff --check` passed for the fixture and focused test files. A later complete source-bound native run remains a separate acceptance gate.
