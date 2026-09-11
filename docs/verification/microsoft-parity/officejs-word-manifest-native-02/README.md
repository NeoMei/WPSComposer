# Word manifest activation diagnosis — September 10

Candidate: `5622a1615e6fd22676a3aa410a393503dab57f91`. Result: **BLOCKED; no Office.js execution or parity acceptance**.

This isolates local manifest discovery and attempted activation with the HTTPS service deliberately stopped. The previously reviewed manifest was copied to its exact Word `wef` path. No certificate was generated or trusted, no security/privacy preference was changed, and no Office.js document read or write ran.

The temporary `CEFRuntimeLoggingFile` preference was initially absent. It was set to the owned log filename and Word was restarted before testing, following [Microsoft's runtime logging instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtime-logging). The expected log file was never created. That is an inconclusive diagnostic result, not proof that manifest parsing succeeded or failed.

## UI observations

1. Opened the exact saved synthetic DOCX named in `setup.json`. The Home add-ins menu listed “Word read-only capability probe” under developer add-ins. The accessibility node reported disabled; invoking it did not open a pane.
2. Closed that document, created an untouched blank document through Word's own template UI, and opened the same menu. The developer entry was again listed as disabled. A screenshot-based direct double-click on its icon also left the menu open without a task pane.
3. The separate “More Add-ins” manager opened successfully. Its My Add-ins view showed only the pre-existing unrelated add-in; the diagnostic developer entry was absent there. No unrelated add-in was opened or modified.
4. Closed the manager and owned blank document. Word returned to DocStage with no document window. The saved synthetic DOCX's SHA-256 was unchanged.

These observations narrow the failed step to attempted activation on this host. They do not identify the root cause or exclude TLS, policy, cache, or client defects. The service was stopped, so this run cannot certify network loading. The status-bar coauthoring compatibility label also appeared on Word's own blank document; it is insufficient evidence of a document-format cause.

## Read-only settings check

The queried `com.microsoft.office` defaults had optional connected experiences, connected Office experiences and online-content experiences enabled. Neither `OfficeWebAddinDisableOMEXCatalog` nor `OfficeWebAddinDisableAllCatalogs` was present. The exact limited query results are in `preference-check.json`. Microsoft documents these keys in [add-in preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/mac/preferences-add-ins) and [privacy preferences](https://learn.microsoft.com/en-us/microsoft-365-apps/privacy/mac-privacy-preferences). This is not an exhaustive check of effective cloud or managed policy.

## Cleanup and next gate

Removed only the exact owned source XML after verifying its hash. Deleted the temporary runtime-log preference, verified it absent, and restarted Word back to DocStage. TCP 3443 remained closed. The owned file hash still matches setup. `cleanup.json` records the results; complete cached deregistration remains unverified, and no shared cache was cleared.

Further activation work needs a host diagnostic that distinguishes policy/cache/client rejection, or a working native-host entry route. Repeating certificate trust or treating menu discovery as API execution does not resolve this gate. The prepared read-only probe still does not implement the missing snapshot/command/parity operations.
