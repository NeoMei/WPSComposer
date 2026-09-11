# Unit-test isolation incident and recovery

A new startup-failure unit test missed a stub for `_run` while verifying RED and reached the old application-name executor. This opened a disposable empty workbook and triggered a folder-access dialog. The parent interrupted the exact pytest/osascript processes, clicked Cancel on that test-directory permission dialog, verified the synthetic workbook URL and empty content, and closed only that workbook without saving. No permission grant was made.

The retained cleanup script completed with `TEST_CLEANUP_SENTINELS_PRESERVED`: final inventory is exactly the two original unsaved workbooks, each original A1 marker is unchanged, and alerts=true. The unit test now fails explicitly if native `_run` is reached; its corrected RED was recorded separately before implementation. Do not rerun historical cleanup scripts with old names/paths.
