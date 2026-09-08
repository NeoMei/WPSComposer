# Windows second-candidate acceptance, in progress

Native runs use candidate 89b39e8ea9aaf0d3f5d2bdaad6e71c90ffd49f2a plus Windows identity fix a745d2a, merged at 58db799f65e5dfa406adc59ed2850b415d518d86.

smoke-04 ran the official runner with its requested total timeout of 180 seconds. DOCX generation and all six DOCX checks passed. Public convert_to_pdf then reached the remaining deadline; elapsed 179.734 seconds. The untouched error report, native private directory and trace are retained. This is FAIL, not an API-path completion. Representative with timeout 600 is running in a fresh directory with an unsaved sentinel.

Targeted tests: 155 passed, 5 failed. Four newly added Mac-runtime tests import fcntl on Windows. The fifth, test_mac_auto_ignores_wps_locations_not_supported_by_runtime, compares Windows Path strings against POSIX literal paths. No source/test changes were made to mask these failures.

The requested read-only accessibility alternative was tested against CommandBars 1, 2, Ribbon, Menu Bar and Status Bar on empty Word PID 22736. All CommandBar.Application canonical IUnknown identities matched; all IAccessible queries succeeded. Every WindowFromAccessibleObject call failed with E_FAIL (-2147467259). Documents.Count remained zero. No Add, property mutation, Quit or kill was performed. The probe instance is retained. The accepted Caption challenge remains the production binding.

Fresh representative, saved-artifact and UI gates are still pending. Production admission remains NO_GO.
