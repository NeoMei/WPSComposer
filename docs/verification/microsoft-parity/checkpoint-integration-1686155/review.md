# Independent checkpoint integration review

Reviewed immutable candidate `1686155b5b16c6709b0182b40dbf5a6ed0f3d89f` against parent `e7c9de2`. Read-only review of shipped runtime, installer/dependency closure, and associated test integration. No production source, tests, index, branch, or application state was changed. This report and its sibling `checkpoint-integration-scratch/` are the only reviewer artifacts.

## Result

No new concrete P1/P2 integration regression found within this scope. This is a scoped no-findings verdict, not completion of Microsoft/WPS parity, native acceptance, release, installation into the user's environment, or an overall merge approval.

## Examined boundaries

- `msoffice/input_validation.py:181-186`: original ZIP member spelling is checked before normalized filename handling; tests create raw backslash/NUL entries and exercise actual ZipInfo normalization rather than accidentally repaired writer inputs.
- `msoffice/macos_runtime.py:50-96`: portable WordJobLock uses the same byte offset for Windows lock/unlock, retries only contention errors against the original deadline, and closes its stream on rejected acquisition.
- `msoffice/macos_word_session.py:303-415`, `macos_excel_session.py:229-297`, and `macos_powerpoint_session.py:430-509`: in-memory quarantine precedes fallible post-submission diagnostic persistence; primary timeout/cancellation outcomes survive secondary diagnostic errors; quarantined sessions do not submit a cleanup close.
- `msoffice/macos_word_numbering.py`, `macos_word_fields.py:225-265`, and `macos_word_recovery.py:172-205,278-310`: numbering registration is committed after identity acknowledgement, appears in field snapshots with stable creation-order ordinals, and participates in checkpoint-prefix validation and confirmed rollback restoration. Intentional frozen late-argument partial effects are documented and tested, not reported as newly invented regressions.
- `msoffice/macos_word_quality.py` and `macos_word_degradation.py:223-229`: confirmed insertion calls the shared committed-range function, clearing pending headings, invalidating field topology, and marking structural changes. Native quality restrictions remain expressly documented; this review does not elevate method presence to acceptance.
- `install.py:150-156,170-216`, `macos/wps-jsapi-probe/package.json`, and `package-lock.json`: the ordinary plugin copy includes the two new Python helper modules and both dependency manifests; macOS installation continues using `npm ci --omit=dev --ignore-scripts`. Hash comparisons against the isolated installer-test output confirmed byte-identical copies of numbering, quality, package.json, and package-lock.json. The qs test resolves the installed dependency independently from both Express and body-parser and exercises a real loopback request.
- `orchestrator.py:165-166`: absolute-before-resolve output handling remains compatible with suffix validation and the existing generation dispatch.

## Verification

Ran the provided clean development Python (Python 3.9) with `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.`, `-p no:cacheprovider`, and an owned scratch `--basetemp`. Selected suites:

`test_input_validation.py`, `test_word_job_lock_portable.py`, `test_macos_office_diagnostic_io.py`, `test_macos_word_session.py`, `test_macos_word_numbering.py`, `test_macos_word_fields.py`, `test_macos_word_rollback_failure.py`, `test_macos_word_quality.py`, `test_installer.py`, `test_qs_dependency.py`, and `test_generation.py`.

Selection excluded `compile` and `native_fixture`: **425 passed, 6 deselected in 108.26 seconds**. No AppleEvents, Office UI, or native compiler tests were run. `git diff --check e7c9de2 1686155 -- . ':!docs/verification'` passed. Tracked working files were unchanged when checked after the test run.

## Explicit limits

This review binds only to `1686155`; it does not review the subsequent quality helper repair. Root reported the already-known run-02 native failure in `_layout_commands` (`count shapes of qpart`, -1700), which remains outside this new-findings count. Six Partial capabilities, nine missing Mac Word methods, and four quality methods without completed native acceptance remain outstanding as directed. Large native evidence collections were not exhaustively scanned. The 425 tests are separate focused evidence and do not replace the root's full-suite, Windows, native, or UI gates.
