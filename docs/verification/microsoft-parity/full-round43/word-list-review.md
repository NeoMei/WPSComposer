# Word list argument parity — independent review

Reviewed 2026-09-10 by `excel_private_owner`. Scope was source and isolated tests only; no native application was launched and no production or test files were modified during this re-review.

## Result

No unresolved actionable P1/P2 finding in the frozen list argument and formatting-reset repair. The previously reported P2 direct-format inheritance issue is resolved in source: each list paragraph applies List Paragraph, resets font and paragraph formatting, then applies its explicit list format. After each item, the terminal insertion range is normalized to Normal and its font and paragraph formatting are reset. This follows the existing native Word business-paragraph reset pattern and prevents the list's direct hanging indent from being carried into the next body paragraph.

## Frozen files

Parent confirmed these reviewed hashes after the independent run:

| File | SHA-256 |
| --- | --- |
| `skills/WPSComposer/scripts/msoffice/macos_script.py` | `45f81ec2422de455ca5841e116057246b4522236e018e1d20c1d50f2842e2912` |
| `skills/WPSComposer/scripts/longform/windows_executor.py` | `6d2ca0ea1cd27d62c0d8ab57aa172cc67032d83d319167ff7e8d3903868b6c54` |
| `tests/msoffice/test_word_list_argument_parity.py` | `28182167cad74ecc9f6708ee53a4c9a491c7f5b27905daaa4d02bac9dc926614` |
| `tests/msoffice/test_macos_script.py` | `08f8e55c397deef27b54fa9df82d10caeb1b125653c703cd319f091e17838cf1` |
| `tests/msoffice/test_macos_word_parity.py` | `b53a9e58f79e15febefbba4ce49cd7104ce6009136edcd4539efd59178328125` |

## Checks

- Compared the closed `generation_plan.py` schema with both backends: `items` and `ordered` remain required; `glyph` and finite numeric `indent` remain optional. No schema expansion was introduced.
- Compared literal bullet/number prefixes and formatting with `WriterComposer.add_bullet_list` and `add_numbered_list`: default, empty, custom, and script-like glyph data are preserved and quoted; ordered lists ignore glyph; zero, negative, and fractional finite indent values are forwarded without truthiness-based substitution. Nonfinite and Boolean indent values are rejected by plan validation.
- Checked the reset ordering before explicit font/paragraph settings and the trailing Normal reset. The regression test locks these generated statements and their ordering. A following-body transition assertion would strengthen the test oracle, but its absence does not identify another implementation defect.
- Checked `docs/regression-guardrails.md`: this patch does not modify Body Text style definitions, per-level heading sizes, title handling, native heading numbering, or Windows camelCase style translation. Literal body-list prefixes remain separate from native heading numbering.

## Independent verification

From the worktree:

```sh
PYTHONDONTWRITEBYTECODE=1 ../wps-task-session-startup/.venv/bin/python -m pytest \
  tests/msoffice/test_word_list_argument_parity.py \
  tests/msoffice/test_macos_script.py \
  tests/msoffice/test_macos_word_parity.py \
  tests/test_generation_plan.py \
  tests/test_recording_composers.py \
  tests/longform/test_windows_executor.py \
  tests/longform/test_plan.py \
  tests/longform/test_semantic.py -q
```

Result: **322 passed in 0.85s**, exit code 0.

## Remaining acceptance

The parent owns the frozen full-suite run and native fixture acceptance. This review does not establish rendered DOCX/PDF typography or Windows COM behavior. Native checks should inspect custom/default/empty bullet prefixes, ordered prefixes, zero/negative/custom indent handling (including native tab-stop behavior), formatted paragraph → list → Body Text transitions, and retained M5 heading/body layout. The native tab-stop handling for nonpositive positions remains an acceptance question, not a reproduced source-review finding.
