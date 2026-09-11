# Independent review: frozen Word probe HTTPS server

**Scoped PASS; no P1/P2 findings identified.** The prepared service is suitable for the stated four-asset loopback task subject to the parent's separate certificate/host prerequisites. This review does not approve trust-store changes or establish Word readiness, installation, document preservation or parity.

## Exact freeze and independent checks

All five HASHES.json entries independently matched before and after testing. Server SHA-256 `4f804c287fa2c5cacc648ee166e083569df489eb6ebd66a095eb894fac6e86f1`; test SHA-256 `3d5515c41c2595b62ed9c057c1994752c433251ccd74937b2565504e2734ec7a`.

- `python3 -B -m unittest discover -s tests -v`: **8 tests PASS** in 0.595s. Log: `officejs-word-probe-server-independent-tests.log`.
- Separate reviewer raw-HTTP/TLS cases: **2 tests PASS**. Missing and duplicate Host headers return 403 without asset bytes; eight stalled TLS handshakes expire and the next valid client succeeds. Scratch: `checkpoint-integration-scratch/test_officejs_https_review.py`; log: `officejs-word-probe-server-independent-extra.log`.

Tests create only isolated temporary one-day certificates/assets, trust the certificate explicitly through each Python SSLContext's cafile, and remove their temporary directories. Test teardown shuts the server down, joins its thread, and requires subsequent connection failure. No persistent listener remains, and no OS/browser/Office trust setting was changed. No package, production, shared-index, native Office, installation or browser action occurred.

## Reviewed boundaries

The service loads exactly the four separately pinned asset hashes and a pinned package HASHES.json before binding. Requests select only from the resulting immutable mapping; they do not open filesystem paths. Missing, modified, symlinked, over-size and forged-manifest inputs fail before listening, and disk changes after startup do not affect served bytes. The pinned manifest identifies the reviewed probe package; it is not a request route or a private-key route.

Only 127.0.0.1 is bound. GET is required; no directory listing, command handler, dynamic evaluation, registration, snapshot or telemetry endpoint exists. The raw request target must select one exact known path; traversal, encoded aliases, absolute URLs and fragments do not resolve to files. Target length is capped at 2048 characters. Allowed asset requests require exactly one matching localhost/127.0.0.1 Host with the actual port. Opaque query suffixes are ignored without decoding, interpolation, logging or identity claims. Invalid responses are static; access/error logging is suppressed and no traceback or path is returned.

TLS accepts only a supplied cert/key pair and uses TLS1.2 minimum. Loading the pair does not certify trust, hostname or expiry; the report correctly leaves those independent prerequisites explicit. The test cafile is process-local and does not bypass normal hostname/certificate verification. No certificate generation or trust operation exists in server.py.

At most eight request workers run. Accepted sockets receive a bounded timeout before TLS handshake; after TLS a separate wall-clock timer shuts the connection down, so trickled HTTP input cannot continually refresh the handler lifetime. Handshake and post-handshake are separately bounded phases; this is not a single total two-second connection deadline. No keep-alive is permitted. Non-daemon workers are joined on server_close and their I/O is bounded. Request teardown closes the wrapped TLS socket; timers are canceled after normal completion. CLI Ctrl-C exits the serve loop into context-manager cleanup.

## Remaining operational gates

This review is limited to the frozen server and temporary TLS tests. Actual manifest loading, trusted certificate installation if separately approved, Office SDK network behavior, browser/native execution, add-in enable/removal and source-bound document observation remain parent-owned gates. An approved four-file origin and a working Python TLS client do not demonstrate native Word acceptance.
