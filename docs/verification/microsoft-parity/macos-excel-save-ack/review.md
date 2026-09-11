# Excel save acknowledgement review

Verdict: **no findings in the scoped patch**.

Reviewed the saved patch snapshot and the current diff for `macos_excel_session.py` and `test_macos_excel_session.py`. The current diff additionally contains the positive save-and-close acknowledgement test that was appended after the snapshot.

The implementation checks the acknowledgement inside `_run` and requires an actual JSON Boolean `true`. Missing, false, numeric, string, and otherwise malformed acknowledgements therefore enter the existing quarantine and diagnostic path before validation or publication. `save()` and `save_current()` request this acknowledgement directly; `save_copy()` inherits it through `save()` and still restores the prior logical binding in `finally`. The source protection, destination conflict checks, logical-target state, exact native binding, attached-session behavior, and publication checks are unchanged.

For non-attached sessions, `close()` now requires a strict `closed: true` result derived from workbook absence after the close command. A workbook that remains open produces the existing `NATIVE_OFFICE_EXECUTION_FAILED` recovery record with `closed: false`, quarantines the component, marks the local session closed, and releases the held lock through the pre-existing `finally` lifecycle. Attached sessions continue to leave their workbook open.

The added tests cover nine unacknowledged-save combinations across `save`, `save_current`, and `save_copy`, three unacknowledged-close replies, and one acknowledged save-plus-close path. I did not rerun the already-reported 109 focused tests because the review found no concrete concern requiring repetition. Fresh review checks completed with `python3 -m py_compile` and `git diff --check`, both exit 0. The root task's native fixture validation remains separate evidence and was not repeated in this review.
