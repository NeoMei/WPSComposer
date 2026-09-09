# Early rejected POST teardown repair

Status: implementation frozen for independent review, not yet independently approved. No shared-index changes, commit, native Office operation, or permission changes were made.

Base HEAD: `595268ec209df7f8fd533452d30b61609696134f`.

## Behavior and scope

Only `skills/WPSComposer/scripts/macos_probe/bridge.py` and the new `tests/macos_probe/test_bridge_rejection_teardown.py` were changed. The original `test_bridge.py`, including its strict 403 assertion, is unchanged.

Early POST 401/403 responses now explicitly declare `Connection: close`. The finalizer flushes the complete response, calls `shutdown(SHUT_WR)`, then discards raw input until peer EOF, at most `MAX_BODY_BYTES + 1` bytes, or a single absolute 0.2-second deadline. Each read is at most 64 KiB and receives only the remaining timeout. The discard ignores Content-Length entirely, so missing, malformed, negative, and oversized declarations do not enlarge its bounds. A teardown OSError ends the discard; the standard finalizer still closes the streams in a finally clause.

Authorization and origin checks still precede body parsing and business execution. Neither check order nor error payload/status was changed. Ordinary POST responses, GET and OPTIONS do not enter the added drain path. There is no retry, rejected-body JSON parsing, or pipelined request execution.

## Verification

Interpreter: `/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python`, with `PYTHONPATH=.`.

- Before production changes, the initial 18 dedicated cases yielded **9 failed, 9 passed in 3.09s**. Failures proved absent half-close/drain bounds, absent close declaration and strict real-socket error delivery failures. Log: `bridge-rejection-teardown-red.log`.
- Initial implementation plus the unchanged existing bridge tests: **30 passed in 8.15s**.
- Strengthened fault-phase assertions and added immediate EOF case, then ran the entire `tests/macos_probe` group: **488 passed in 97.70s**. Log: `bridge-rejection-teardown-green.log`. Final dedicated file has 19 cases.
- `git diff --check -- skills/WPSComposer/scripts/macos_probe/bridge.py` passed.

The dedicated cases cover exact large valid-body 403/ORIGIN_REJECTED and 401/UNAUTHORIZED delivery through the real loopback server, the session origin-rejection route, no body parsing or registration, response delivery before an absent body, bounded finish while that peer stays open, absolute deadlines under slow trickle, byte caps independent of Content-Length, EOF, flush/shutdown/settimeout/read disconnects, timeout, and unchanged non-early responses. Real-socket errors are not swallowed or counted as successful HTTP responses. Deterministic stream/socket doubles are used only for fault injection and precise resource/time bounds.

## Frozen evidence

- Source SHA-256: `97be582ec55b2807e4cce2e6f4ad13cbe90e780ce3c40ee254dd0f5e59749ef8`
- Dedicated test SHA-256: `0e7da5e19d1414b249fe0c559085d380a19dbb6055ad02394d3088b0c8af114f`
- Patch: `bridge-rejection-teardown-frozen.patch`
- Prior diagnostic report: `bridge-run04-diagnostic-review.md`
- Preserved prototype: `bridge_rejection_transport_experiment.py`

## Limits

The original run-04 Windows 3.12 187-byte request / WinError 10053 was not reproduced locally. The prior controlled experiment establishes a separate real transport defect for allowed-size 1 MiB JSON requests and the tests establish the repaired local behavior. This does not identify the unique cause of that Windows failure or establish a repaired Windows CI run. A subsequent Windows CI run remains necessary.

The bounded discard deliberately closes peers that exceed the byte/time budget. It does not promise response delivery for unlimited or indefinitely slow input. Independent review remains pending before integration.

## Independent acceptance — checkpoint_integration_review

**PASS.** Independently matched source SHA-256 `97be582ec55b2807e4cce2e6f4ad13cbe90e780ce3c40ee254dd0f5e59749ef8` and dedicated-test SHA-256 `0e7da5e19d1414b249fe0c559085d380a19dbb6055ad02394d3088b0c8af114f`, both before and after focused validation. No actionable P1/P2 was found in this bounded teardown repair. These exact two files may be frozen for full33 integration.

Independent focused run of the unchanged `test_bridge.py` and all 19 dedicated rejection-teardown cases: **31 passed in 7.18 seconds**. `git diff --exit-code -- tests/macos_probe/test_bridge.py` and the changed bridge diff whitespace check both passed. This retains the original strict 403 expectation rather than treating transport failure as a successful rejection.

Read-through confirmed normal routes retain the original authorization/origin ordering and business body parsing. Only rejected POST 401/403 responses set the teardown marker and force connection close. Response flush and write-half shutdown precede raw discard; the absolute deadline and byte cap cannot be enlarged through Content-Length or repeated trickle reads. Discard never invokes JSON parsing, authorization retry, registration, result dispatch, or another request handler. Normal POST/GET/OPTIONS paths do not enter the new linger loop. The handler uses the standard buffered input/read1 and unbuffered socket output; standard stream finalization runs in a finally clause even when the added teardown fails.

The source change supports the separately reproduced large-body local transport defect. It does not prove the unique cause or repair of the original small-body Windows WinError 10053; fresh Windows CI remains necessary. No native Office, production/test source, or shared-index changes were made by the reviewer. Only this report was updated; the focused tests used owned scratch basetemp with bytecode and pytest caching disabled. This acceptance supersedes the earlier pending-review wording for the hashes above and does not claim full33 itself or native/UI acceptance has passed.
