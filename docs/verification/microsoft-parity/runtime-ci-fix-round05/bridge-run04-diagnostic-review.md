# Run-04 bridge transport investigation

Scope: the sole Windows 3.12 failure in CI run `34343657808`, built from `595268e`. No production files, HTTP authorization/origin policy, shared index, native Office processes, or machine permission settings were changed. Tests used local loopback HTTP only.

## Exact CI evidence

Parsed `portable-ci/run-04/artifacts/*/pytest.xml` with ElementTree rather than inferring counts from progress output. Each job collected 4855 test cases:

| Job | Passed | Skipped | Failed | Time (s) |
| --- | ---: | ---: | ---: | ---: |
| Ubuntu 24.04 / Python 3.9 | 4811 | 44 | 0 | 237.123 |
| Ubuntu 24.04 / Python 3.12 | 4811 | 44 | 0 | 221.742 |
| Windows 2022 / Python 3.9 | 4780 | 75 | 0 | 638.894 |
| Windows 2022 / Python 3.12 | 4779 | 75 | 1 | 653.875 |

The sole failure is `tests.macos_probe.test_bridge.test_bridge_rejects_unlisted_origin`. `urllib` reached `http.client` response status parsing, then socket `recv_into` raised `ConnectionAbortedError [WinError 10053]` before an HTTP status was obtained. The unchanged assertion expects `HTTPError` with code 403. There is no evidence that the handler accepted the forbidden origin.

## Root-cause evidence and limit

`BridgeHandler.do_POST` rejects invalid authorization/origin before consuming the body. `_send` emits the error response but no explicit `Connection: close` header. BaseHTTPRequestHandler defaults to HTTP/1.0. CPython `StreamRequestHandler.finish` flushes/closes its streams; `TCPServer.shutdown_request` then performs `shutdown(SHUT_WR)` immediately followed by full close. Therefore the current stack half-closes but does not continue reading until peer closure or a bounded acknowledgement opportunity.

This ordering is a plausible match to the lost-response risk described by [RFC 9112 section 9.6](https://www.rfc-editor.org/rfc/rfc9112.html#section-9.6): immediate full closure can reset a connection when subsequent client bytes arrive, and the reset can discard unread response bytes. The staged alternative sends the response, half-closes writing, continues reading, then fully closes.

**The exact Windows CI 187-byte request / WinError 10053 was not reproduced here.** There is no direct Windows execution surface in this reviewer task, and local small-request trials all returned the required 403. The evidence below proves a real early-rejection transport failure for an allowed-size large body on macOS, and validates a targeted teardown candidate. It does not prove that this is the only possible cause of the one Windows CI failure.

## Bounded real-socket experiments

Retained executable experiment: `bridge_rejection_transport_experiment.py`. It imports the actual bridge and uses valid credentials plus forbidden origin. The candidate exists only as an in-memory handler subclass; production is unchanged. Each successful response checks both exact HTTP status and exact error code, and checks no component was registered.

Initial baseline probe:

- 187-byte body: 250 ordinary sends, 100 sends with 1 ms header/body delay, 30 with 30 ms delay: **380/380 HTTP 403**.
- 64 KiB body: **50/50 HTTP 403**.
- 1 MiB body (equal to `MAX_BODY_BYTES`): **9 ECONNRESET + 11 EPIPE**, no HTTP success. This first large payload was opaque bytes; the subsequent controlled experiment uses valid JSON.

Controlled valid-JSON comparison, same bridge/auth/client logic, only teardown changed:

| Candidate | Request | Result |
| --- | --- | --- |
| Current production | 202-byte JSON, 100 requests | 100 exact 403 responses |
| Current production | Valid 1 MiB JSON, 20 requests | 2 exact 403; 10 EPIPE; 8 ECONNRESET |
| Staged teardown subclass | Same valid 1 MiB JSON, 20 requests | 20 exact 403 responses |
| Staged teardown subclass | Declared 1 MiB, no body sent, 5 requests | 5 exact 403 responses; maximum response latency 1.5 ms |
| Staged teardown subclass | Missing token, valid 1 MiB JSON, 5 requests | 5 exact 401 responses |

The staged prototype also ran the complete existing `tests/macos_probe/test_bridge.py` suite by substituting only its handler factory: **12 passed in 5.09s**. No status assertions or client exception handling were relaxed. The experiment records resets as observed failure outcomes for comparison; it does not reinterpret them as successful authorization rejection.

## Proposed narrow repair and regression requirements

For early POST 401/403 rejections only, explicitly close the HTTP connection, send and flush the complete rejection response first, then half-close the socket write side. Before final close, discard raw incoming bytes until peer EOF, a fixed byte cap, or an absolute monotonic deadline. The tested prototype uses `MAX_BODY_BYTES + 1`, chunks up to 64 KiB, and a 0.2-second absolute deadline with each read timeout capped by remaining time. It does not parse rejected input, change auth/origin decisions, wait for the body before sending the error, retry a rejected operation, or process pipelined requests.

Before production acceptance, keep/add strict regression coverage for:

1. Full exact 403/ORIGIN_REJECTED for large valid POSTs at the configured size limit, with no register side effect; retain the original small-request 403 assertion unchanged.
2. Exact 401/UNAUTHORIZED for missing/invalid token and equivalent body framing.
3. Response availability before a client supplies its declared body, with bounded final cleanup even if the client never sends or closes.
4. Absolute deadline and byte cap under slow trickle, oversized declared length, invalid length, and peer reset; none may cause unbounded rejected-body reads.
5. Existing authorized bootstrap/registration/command/result paths and lease behavior unchanged.

A corrected Windows 3.12 CI run and targeted repeated socket tests are still required to validate the original platform symptom. A re-run pass without a source change is not presented as a fix. This investigation supports a narrow staged-teardown repair for a separately reproduced transport defect, with an explicit remaining uncertainty about the single CI symptom.
