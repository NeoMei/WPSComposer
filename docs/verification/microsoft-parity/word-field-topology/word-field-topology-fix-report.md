# Mac Word field topology mutation fix

## Scope and source freeze

- Production ownership stayed within `macos_word_session.py` and `macos_word_fields.py`; targeted tests are in `tests/msoffice/test_macos_word_fields.py`.
- No commit or push was made.
- Frozen SHA-256:
  - `macos_word_session.py`: `a96f96f99667e04ec64ed3eb018ba45de308db9a308d37b3df9c98bf2537bc6c`
  - `macos_word_fields.py`: `446e8c9f75ce66e93f55b1d8080aa80db6bf5abd4886c75148d8264a8cc50aeb`
  - `test_macos_word_fields.py`: `8a04fda9ade0e031f6178a8eeec61c4800f88a340b9586acfd1beaef8090e4d3`

## Root cause and correction

`_observed_field_topology` was written only by `snapshot()` and survived session-owned content and field mutations. Later snapshots therefore compared legitimate new fields or shifted positions against an obsolete baseline and raised `NATIVE_WORD_FIELD_IDENTITY_STALE`.

The correction invalidates the observed baseline at the native submission boundary for known topology-changing execution. Successful test transports invalidate after acknowledgement. Preflight and script-preparation failures preserve the baseline; once an AppleEvent is submitted, success and uncertain completion remain conservatively invalidated. Field refresh explicitly uses `field_topology_change=False`, retaining the existing exact code/position/TOC-result-growth drift checks. Formatting-only operations retain the baseline; text replacement, structural edits, business content, index/reference creation, and header/footer field-affecting changes invalidate it. Header/footer invalidation is independent from body structural state, preserving pending headings and stable body targets.

Tracked bookmark/code validation remains unchanged. Deleting or altering a tracked handle still fails stale.

## RED and GREEN

- Independent reviewer RED: `word-field-topology-mutation-review-red.log` recorded 3 failures and 2 passing controls: same-session REF creation, TOC creation, and structural insertion failed stale.
- GREEN command:
  `PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest -q tests/msoffice/test_macos_word_session.py tests/msoffice/test_macos_word_business.py tests/msoffice/test_macos_word_fields.py tests/msoffice/test_macos_word_references.py .superpowers/sdd/2026-09-08-microsoft-wps-parity/word-field-topology-mutation-review-repros.py`
- Result: `228 passed in 1.85s`.
- Coverage includes the original 3 REDs and 2 controls, same-paragraph text-length shift, refresh preserving external position/code drift rejection, header/footer scope, and both local and late transport pre-submission failure preserving the baseline even with sticky prior evidence.
- `git diff --check` passed.

## Native evidence

- Passing evidence: `word-field-topology-native-run-04/report.json` (`passed: true`). Its retained production hashes equal the frozen hashes above.
- Real public Microsoft Word session sequence:
  1. initial native REF snapshot: 1 field;
  2. public `add_cross_reference_paragraph`: 2 fields and the original stable key preserved;
  3. public `insert_toc`: 3 semantic snapshot fields and both prior stable keys preserved;
  4. public structural paragraph insert at document start: all three stable keys preserved.
- The saved/reopened native document contained 5 Word fields and 1 TOC. DOCX save, PDF export/read, exact source reopen, and visible `Prefix`/`Owned` text checks passed.
- Artifacts:
  - `topology.docx`: `155632ba4e4a1acb9a7e059a2ce39de97641a75d02682dd4702b2864f34654db`
  - `topology.pdf`: `9a429892bf08478611ea7bdca3b7378f3b7f980683bdf2efc7773c6c90bd1d37`
- Sentinel inventory was byte-identical before and after owned-document work. Final live Word inventory returned `missing value, missing value` (zero open documents).

## Failed native attempts retained

- run01 and run02 failed public bookmark preflight because the probe used invalid bookmark names; no target mutation occurred.
- run03 proved all four topology checks but failed during artifact validation because the top-level probe was re-imported by multiprocessing spawn. This runner defect was corrected for run04 by presenting a synthetic non-file `__main__.__file__`, activating the production spawn bootstrap. All failed reports/runtime evidence remain retained in their original directories.

## Status

Production is frozen at the hashes above. Native Word lease can be released. Independent review remains for the parent task.
