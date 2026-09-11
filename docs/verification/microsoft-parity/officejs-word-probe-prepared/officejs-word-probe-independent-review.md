# Independent offline Office.js Word probe review

Verdict: scoped no-mutation/permission/lifecycle review **PASS**, with one P3 diagnostic-key correction recommended before native use. No P1/P2 found. This is an offline preparation review, not production support, install approval, document binding proof or parity closure.

## Finding — P3: use the frozen equation fallback method name

`probe.js:17` exports `add_equation_fallback` in `capabilities().gaps`, while the actual frozen direct method is `add_equation_native_fallback` (`skills/WPSComposer/scripts/writer.py:2329`, also the gap-design report). Therefore the claimed nine-method mapping contains a nonexistent key and omits the real fallback key. A consumer joining diagnostic JSON to the parity catalog cannot match that row. This does not enable a mutation or falsely mark native acceptance, since all rows remain false.

Direct offline reproduction: `capabilities({}).gaps` lacks `add_equation_native_fallback` and includes `add_equation_fallback`. Current test only verifies length nine and false acceptance flags. Change the key and add an exact-name-set test rather than a count-only assertion; regenerate retained hashes after that small correction. No package file was edited by this reviewer.

## Source and test evidence

Read REPORT.md, README.md, HASHES.json, all core/UI/configuration/manifest files and all tests. All **16/16** listed SHA-256 values match, before and after independent tests. Reviewed `probe.js` SHA-256: `411cafcaa149f46f81457c49edbeb828d73d2c88390d49cf9271e82e70c55734`.

Executed within the package using installed runtimes:

- `node --test tests/*.test.cjs`: **20 passed, 0 failed** (72.30 ms); separate retained log `officejs-word-probe-independent-js.log`.
- `python3 -B -m unittest discover -s tests -p 'test_*.py'`: **6 passed**; `officejs-word-probe-independent-python.log`.
- `node --check probe.js` and `node --check taskpane.js`: success.

All checks are offline stub/structural checks. No browser, server, Office/native process, certificate, catalog, installation, permission change or shared-index action occurred. Logs and this report are outside the frozen package.

## Reviewed behavior

The only submitted Word object-model batch loads `document.saved` and, conditionally on Desktop1.4, selection start/end/storyType, followed by sync. Other code inspects diagnostics, common document mode/URL availability, requirements and method presence. No document text/OOXML read, setters, save, insertion, snapshot call, arbitrary execution, network command channel, telemetry, or persisted identity appears. getFileAsync is checked for presence only. User-triggered export is a local JSON Blob; no raw path/document text is included automatically. Office's CDN SDK remains an external host dependency, explicitly disclosed.

ReadWriteDocument appears in manifest and prominently in README/UI. This accurately separates code-level read-only behavior from granted capabilities: Word-specific APIs require read/write manifest permission even when only reading. [Microsoft permission model](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)

The manifest uses a Word Document host, task-pane add-in-only schema, WordApi1.1 requirement and HTTPS-loopback SourceLocation, with no activation handler. Configuration accepts only exact HTTPS loopback origins and exclusively creates a package-local manifest. Reviewed shape is consistent with the documented manifest model, but no Microsoft schema-validator or native loading claim is made. [Microsoft manifest documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/xml-manifest-overview)

Pane ID is crypto-random memory state and is explicitly page-lifetime only. Wrong host/platform or unknown/false baseline produces not-run; invalid scalar values produce error. A single busy flag prevents overlapping reads. Timeout/error retires the pane; late results check liveness and sequence before an ACK. Retirement on pagehide blocks new reads. The documentation correctly states timeout does not cancel the host operation and a read ACK does not establish nonmutation or durable document identity.

Setup steps are future bounded work: exact manifest/origin, owned fixtures and sentinel, no restart over unsaved user work, certificate/catalog changes separately coordinated. Removal does not bundle broad deletion commands or imply isolated removal is safe. It explains that Office cache cleanup affects other add-ins and confines that future procedure to approved isolated profiles with inventory/restoration checks. This matches current Microsoft guidance warning against deleting individual cached manifests and documenting application-wide sideload removal. [Microsoft cache guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/clear-cache)

Native readiness, actual requirement responses, host-specific identity lifetime, read-only/protected fixtures, before/after preservation, installation/removal and all nine method contracts remain unverified. The parent retains the decision on any future use.
