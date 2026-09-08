# Windows second-candidate acceptance, in progress

Native runs use candidate 89b39e8ea9aaf0d3f5d2bdaad6e71c90ffd49f2a plus Windows identity fix a745d2a, merged at 58db799f65e5dfa406adc59ed2850b415d518d86.

smoke-04 ran the official runner with its requested total timeout of 180 seconds. DOCX generation and all six DOCX checks passed. Public convert_to_pdf then reached the remaining deadline; elapsed 179.734 seconds. The untouched error report, native private directory and trace are retained. This is FAIL, not an API-path completion. Representative with timeout 600 is running in a fresh directory with an unsaved sentinel.

Targeted tests: 155 passed, 5 failed. Four newly added Mac-runtime tests import fcntl on Windows. The fifth, test_mac_auto_ignores_wps_locations_not_supported_by_runtime, compares Windows Path strings against POSIX literal paths. No source/test changes were made to mask these failures.

The requested read-only accessibility alternative was tested against CommandBars 1, 2, Ribbon, Menu Bar and Status Bar on empty Word PID 22736. All CommandBar.Application canonical IUnknown identities matched; all IAccessible queries succeeded. Every WindowFromAccessibleObject call failed with E_FAIL (-2147467259). Documents.Count remained zero. No Add, property mutation, Quit or kill was performed. The probe instance is retained. The accepted Caption challenge remains the production binding.

Fresh representative, saved-artifact and UI gates are still pending. Production admission remains NO_GO.


## Independent pagination failure and native diagnosis

representative-02 finished all three public API routes within 600 seconds and the old runner reported PASS. This is superseded by the independent saved-PDF pagination failure: cover 1, TOC ii, body 3/4/5 and TOC references 3. All three sections lack a start attribute. This reproduces the parent review and invalidates the first-round claim of complete native layout acceptance. See pagination-before-fix.json; raw former results are unchanged.

The sentinel in existing Word PID9252 remained byte-for-byte equivalent in text hash/path/saved state, and was closed without saving; that old application remained empty and responsive.

The dedicated native COM probe found two independent causes. Adding PAGE while the new footer is linked propagates it to the cover. Also, assigning Python integer -1 to RestartNumberingAtSection returns no COM error but leaves restart False and StartingNumber 0. The setter probe repeats -1 after a PAGE field exists with the same result; Python True then immediately reads back True, and StartingNumber 5 is accepted. NumberStyle 2 reads back correctly. The fix uses Boolean values, detaches before footer mutations, verifies restart/start/style, and raises meaningful execution errors. Fixed three-page synthetic Word output now has empty/i/1 footers; the fixed probe retains each COM readback. The immediate process-absence check can precede asynchronous Word exit; final PID checks will follow.

Regression RED: 9 failed/1 passed; GREEN:10 passed. The first longform baseline had one add-in manifest mismatch due to Git's CRLF checkout during merge; 23 changed text files were restored to byte-exact committed LF without semantic edits. The before/after hashes are retained. A final candidate deb67ba is now available and will be merged after the active test run; official stronger page-number gates and full native representative rerun remain pending.
