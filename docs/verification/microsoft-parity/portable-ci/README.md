# Portable CI repair checkpoint — 2026-09-09

Initial hosted run [34332171915](https://github.com/NeoMei/WPSComposer/actions/runs/34332171915) tests commit `34a0658e38c6d5a73a7066df7f49f73b53aceb39`. Both Windows jobs failed checkout because four historical PNG names contained colons. Both Ubuntu jobs reached pytest and reported 10 failures, 4,396 passes and 32 skips. Raw failures remain unchanged in `run-01/`.

The four PNGs were renamed without changing bytes. The relocation map preserves original paths and exact Git blob hashes; all 240 original files in the affected evidence directory remain byte-identical. Thirty new tests audit shipped Git paths and relocation integrity. The four test-file repairs pin the intended mocked Mac backend, mock the installed WPS version lookup, and restrict only native AppleScript compilation cases to macOS. Independent reviews are retained here. Neither production platform support nor native assertions were weakened.

Full Mac round 29 passes **4,457 tests / 12 skips** in 210.72 seconds. It ran from an independent Git checkout of temporary snapshot `1ae8bd672445207644639df272939f67decc4c89`, tree `d180ef9b3db1a73c59b4ae8ab4c7e931ef4f9bb8`, parent `34a0658`, with the locked npm dependencies and the documented clean Python 3.9 dev environment. All 335 Python/configuration files recorded by this run remained unchanged. Logs, JUnit and before/after hashes are in `full-round29/`. Concurrent unfinished Word implementation and later dependency changes are outside this snapshot; twelve native/bridge skips remain skips. Retained raw logs contain their original whitespace; source/document diff checks exclude raw verification evidence.

This checkpoint repairs portable checkout/test execution. The next hosted Windows/Ubuntu Python 3.9/3.12 run must establish actual hosted acceptance. It does not establish installed Windows WPS/Office, UI interactions, complete Microsoft parity or release readiness.

## Hosted repair run 02

Hosted run `34334821156` at `e7c9de2` completed: both Linux jobs passed (4,435 passed / 34 skipped each); Windows Python 3.12 had 68 failures and Python 3.9 had 72 failures. Both Windows checkouts now succeed. Raw logs and JUnit for all four jobs are retained under `run-02/artifacts/`.

The reviewed follow-up fixes preserve raw ZIP member spelling before normalization, use a native Windows lock primitive for the shared Word lock, and repair escaped-path and manifest-key assumptions in fixtures. The Windows 3.9-only fixes anchor missing relative output paths before resolution and retain the required SystemRoot variable in clean import-test subprocesses. The latter changes pass 101 local tests and seven independent targeted checks. These are local verification results; corrected Windows CI remains required.

Full Mac round 30 records 4,608 passed / 1 failed / 12 skipped with all 366 recorded files unchanged. Its failure was an obsolete positional CheckpointSnapshot construction in a rollback-injection test; the correction names all six fields and preserves the actual injected-error assertions. Full round 31 subsequently passes **4,790 tests / 12 skips** in 281.77 seconds, with all 373 recorded source/configuration files unchanged and equal to the live candidate. It ran from temporary immutable commit `4a518a70b30053399d72090d7a1ffd4c861e58a8`, tree `c7fdfbdac419faeb8baef79be7a2e7e4af1ba29c`, parent `e7c9de2`; evidence is in `../full-round31/`. Native skips remain unexecuted gates.

## Hosted repair run 03

Run `34339669535` binds candidate `1686155b5b16c6709b0182b40dbf5a6ed0f3d89f`. Both Ubuntu Python versions pass **4,758 tests / 44 skips**. Both Windows Python versions now record **4,717 passes / 10 failures / 75 skips**. All original checkout, ZIP spelling, Word-lock, Python 3.9 initialization and default-output fixes passed their exercised checks. The remaining failures expose the separate Excel/PowerPoint generation lock's unconditional fcntl import plus escaped-path assertions; they are under repair. Raw logs, JUnit and run metadata are retained under `run-03/`. This remains a failed hosted run until a later corrected candidate passes.


## Reviewed repair before run04

The Office job lock now uses the platform native lock implementation, and three assertions normalize portable relative paths or AppleScript-escaped paths. Diagnostic persistence no longer replaces native errors or cancellation, and independent review closes two cleanup exception-masking findings (173 scoped tests). Full round32 passes 4,843 tests / 12 skips in 217.11s from immutable commit `415c9d04f38e5d4a744baadce75a39bf76dc8cca`, with all 377 recorded source hashes unchanged and equal to live. Actual hosted run04 and Windows native acceptance remain separate gates. Repair/review records are in `../runtime-ci-fix-round04/`.
