# Frozen Word probe HTTPS server — prepared for independent review

Scope: only this directory. `server.py` is Python standard-library code; it does not generate certificates, modify trust, install add-ins, call Office, fetch remote content or expose a command endpoint. Parent will review and start the actual service after prerequisites. No production/index/installed-add-in files changed.

The listener is hard-coded to **127.0.0.1**, default port **3443**. Only GET requests for `/taskpane.html`, `/taskpane.css`, `/taskpane.js`, `/probe.js` are served, with explicit content types. All bytes are loaded into an immutable in-memory mapping before binding. Startup verifies both the entire approved probe `HASHES.json` SHA-256 (`53bd0cf03da7a9dd885d69b885f17ef34c60e5441fd62569b371b26f69af0c5c`) and four separately pinned asset hashes, rejects missing/symlinked/oversized files, and fails before listening on mismatch. The server never reopens a requested path, so later on-disk edits cannot change served bytes. Certificate/key paths are never routed.

No directory listing, manifest, report, key/certificate, arbitrary path, percent-encoded alternative path, traversal, absolute URI or fragment is served. Non-GET methods return 405; missing/foreign/duplicate Host headers are rejected for allowed asset paths. Host must be localhost or 127.0.0.1 with the actual port. There are no CORS, document, telemetry, logging or bridge endpoints. Request URLs, headers and opaque query data are not logged or persisted.

**Parent query ruling:** allow an opaque query suffix only on one of those exact four paths and ignore it entirely for selection. The total request target is bounded at 2048 characters. Query values are not decoded, interpreted or used as trust/identity. Microsoft documents host-added opaque `_host_info` in its dialog flow and says application code must not read it; this is not a claim that native pane loading was tested. The parent's explicit compatibility ruling also covers the task-pane `_host_Info` example. [Microsoft query-metadata guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/parent-to-dialog).

TLS requires supplied certificate and private-key files. Minimum TLS 1.2; no generated or auto-trusted certificate and no interactive encrypted-key prompt. The startup code validates that Python can load the supplied pair, not that Word trusts the certificate or the hostname/expiry meets the future native prerequisite. Parent remains responsible for those independent checks.

Accepted sockets have a default 2-second timeout (bounded to >0 and <=10 for the internal API). TLS handshakes run in workers; a stalled handshake does not block accepting another request. At most eight workers run simultaneously. After TLS, a wall-clock timer also closes the connection within the socket timeout; keep-alive is disabled. Shutdown closes the listener and waits for bounded workers. Ctrl-C stops a manually started foreground service. The production CLI rejects port zero; tests use ephemeral port zero through the internal factory.

## Offline/local TLS verification

Initial RED (module missing) is preserved in `evidence/red.log`. Final `evidence/green.log` records **8 tests PASS**, including multiple routing/tampering subcases. Tests cover pinned manifest/assets, tampered/missing/symlinked input rejection before bind, exact MIME/bytes, four-path whitelist, methods/traversal/absolute URI/Host rejection, opaque host query same-byte behavior, malicious query never selecting another file, request-length bound, bytes frozen after disk mutation, TLS-handshake and idle-read timeout recovery, and missing certificate/invalid timeout startup failures.

The real localhost TLS smoke generated one one-day self-signed certificate with installed OpenSSL in a task-only private temporary directory. Python's test client loaded that certificate explicitly via `ssl.create_default_context(cafile=...)`; **no OS/browser/Word trust store was modified**. Every test server was shut down and its thread joined; a subsequent connection must fail. Temporary assets, certificate and key were removed, with a post-cleanup assertion. No native call or server remains running. Test-only certificate generation was the explicitly authorized TLS test step, not installation or trust setup for the user's Office run.

Reproduce, using the installed Python/OpenSSL only:

```sh
python3 -B -m unittest discover -s tests -v
```

Future parent-owned foreground invocation (not executed by this task):

```sh
python3 -B server.py --package '/absolute/path/officejs-word-probe-prepared' --cert '/approved/certificate.pem' --key '/approved/private-key.pem' --port 3443
```

The reviewed probe manifest uses `https://localhost:3443/taskpane.html`; serve the same origin with a suitable, separately approved certificate. No server startup, manifest update, add-in enablement, permission change or Word readiness is implied by these local TLS tests. The server retains the probe's existing ReadWriteDocument manifest permission without expanding the probe itself.

`HASHES.json` records the exact server/test/report/evidence files. Independent review and parent startup are pending; no further features are planned for this package.
