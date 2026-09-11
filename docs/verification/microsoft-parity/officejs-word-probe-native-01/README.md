# Word probe activation attempt — September 10, 2026

Result: **BLOCKED at native activation; no Office.js capability PASS**. Candidate a05bfa0053b5f2abe37a16b4566e5ca99ef327fb. This report records root operator UI observations; it is not a raw Office.js response or automated acceptance result.

The user explicitly authorized temporary trust of the previously reviewed localhost certificate and removal after validation. The exact certificate SHA256 3F47F0E7036D8974F3708E90E15A07E65D5F1447FEF3C1D364EBF6865034C872 was trusted only in the current user login keychain for SSL and localhost. Apple verification then passed localhost/SSL and rejected both another hostname and basic policy. The exact reviewed ReadWriteDocument manifest was registered and the reviewed four-asset HTTPS service listened only on 127.0.0.1:3443.

Word was restarted with no document windows open. A disposable unsaved sentinel and the owned WPSComposer-OfficeJS-owned-probe.docx were opened. Home > Add-ins showed “Word read-only capability probe” in its developer section. Attempts to activate that entry did not display the task pane. Its accessibility tile was marked disabled, as was an existing unrelated add-in tile; this alone does not establish an Office policy restriction. No Office.js requirements, Word.run result, selection binding, snapshot or mutation were obtained.

Separate browser navigation to https://localhost:3443/taskpane.html failed with ERR_CERT_AUTHORITY_INVALID despite Apple's successful scoped verification. Default curl verification also failed with issuer lookup. These client trust differences are retained as observations, not a proven cause of Word's activation failure. OfficeWebAddinDisableAllCatalogs and OfficeWebAddinDisableOMEXCatalog had no explicit defaults values; this does not rule out all Office policy or connected-experience restrictions. The Office.js CDN HEAD request returned HTTP200.

The owned saved probe was closed without saving; its SHA256 remains 798a452f6eabf2d921957703f86ad2988b7faf8cde72d87c26a819782972877d, equal to its initial copy. The sentinel's ASCII marker WPSCOMPOSER OFFICEJS SENTINEL 20260910 3F47 was visible after the attempt. Unicode typing did not reliably arrive, and no independent initial in-memory snapshot was retained, so full sentinel preservation is not certified. Only that task-created disposable sentinel was discarded. Word returned to DocStage with no document window.

## Cleanup verified

- Stopped the exact task server process; TCP3443 no longer accepts connections.
- Removed the exact user's certificate trust rule and exact fingerprint from the login keychain; both commands returned0. Subsequent localhost/SSL verification returns CSSMERR_TP_NOT_TRUSTED and the fingerprint is absent from the login keychain.
- Deleted the task private key. Retained public certificate and historical preparation evidence only.
- Removed only the task manifest from Word's Wef directory after verifying its hash. No whole Office cache was cleared; complete cached add-in deregistration is not certified.
- The separate failed browser tab could not be selected for closing because the browser tool rejected its generated data-URL error page. It may remain open as an inert error page; no certificate bypass was used.

Machine results are in trust-and-manifest-install.json and cleanup.json. Preparation evidence remains unchanged in ../officejs-word-probe-bootstrap-01. The prior approval-pending statements describe that earlier checkpoint only. This attempt adds no parity gate: the nine missing Mac Word methods, middle-table support, read-only quality snapshots and final Windows native/UI acceptance remain open. Next diagnosis must distinguish task-pane activation failure from TLS client trust behavior before another native attempt.
