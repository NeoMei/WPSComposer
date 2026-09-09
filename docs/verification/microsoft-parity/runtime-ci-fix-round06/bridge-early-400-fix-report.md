# Early Content-Length rejection teardown repair

Status: implementation frozen for independent review, then full35 / CI06 validation. CI05's `370c8c5` snapshot is unchanged. No shared-index operation, commit, native app, or unrelated source edit occurred.

## Narrow behavior change

`_read_json()` now marks an unread rejected POST immediately before raising its existing ValueError for invalid Content-Length conversion or a negative / greater-than-MAX_BODY_BYTES length. A tiny `_mark_rejected_post()` helper sets the existing rejection and connection-close flags; existing POST 401/403 handling uses the same helper. The finalizer, 0.2-second absolute deadline, MAX_BODY_BYTES+1 discard cap, 64 KiB chunk cap, and error responses are unchanged.

The marker is set at the pre-read rejection sites, not for all HTTP400 responses. Auth/origin checks still precede length parsing. Already-consumed invalid JSON, non-object JSON, and failed client/component binding continue to return 400 without the added linger path. The original strict403 tests were not edited.

## RED and GREEN

- Added 15 permanent cases before production editing. **12 failed, 3 passed, 19 deselected in 2.89s** (`bridge-early-400-permanent-red.log`). The three passing controls verify consumed JSON/business errors do not linger.
- After the repair, ran `tests/macos_probe/test_bridge_rejection_teardown.py tests/macos_probe/test_bridge.py`: **46 passed in 9.32s** (`bridge-early-400-green.log`). Interpreter: `/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python`, `PYTHONPATH=.`.
- Re-ran unchanged `bridge_early_400_transport_experiment.py`: **10/10 exact HTTP400**, with all expected INVALID_REQUEST messages and zero registrations. The original eight large requests failed with BrokenPipe/ConnectionReset; the same eight now returned strict400. Original `bridge-early-400-raw.log` remains; fresh result is `bridge-early-400-raw-green.log`.
- `git diff --check` for both modified files passed.

Permanent coverage includes invalid, negative, and MAX+1 Content-Length; actual valid MAX+1 JSON over real TCP with an authenticated allowed origin and no business registration; no pre-rejection body parsing; absolute deadline under trickle; fixed byte cap under continued input; missing body; response flush and SHUT_WR preceding discard; and consumed JSON/object/business errors avoiding linger. Existing 401/403, missing-body, EOF, peer-disconnect, and ordinary request checks remain unchanged.

## Freeze

- Source SHA-256: `5b60a1cdca0a4d330cb64be65793f3f51e2909124b1aed50793c33a2f243266b`
- Dedicated tests SHA-256: `daaa5cf01648ad72a28627accd0a04e07b040709fe47b72693c3555e5f014d39`
- Patch: `bridge-early-400-frozen.patch`
- Prior bounded investigation: `bridge-early-400-investigation-report.md`

This establishes local delivery for the specified finite boundary-size requests. Bounded discard does not guarantee HTTP400 delivery for an unlimited sender, arbitrary-size streams, or indefinitely slow input. No claim is made that the original Windows small-request WinError10053 has a unique proven cause or that the new snapshot's Windows CI has passed. Independent review/full35/CI06 remain pending.

## Independent acceptance — checkpoint_integration_review

**PASS.** Source `5b60a1cdca0a4d330cb64be65793f3f51e2909124b1aed50793c33a2f243266b` and dedicated tests `daaa5cf01648ad72a28627accd0a04e07b040709fe47b72693c3555e5f014d39` were independently matched before and after review. No actionable P1/P2 was found in this bounded early-400 correction. These two hashes may be frozen for full35 / CI06.

Read-through confirms the new marker is applied only immediately before the existing invalid/negative/oversized Content-Length errors, before `rfile.read`. Both callers are POST routes; authorization/origin order is unchanged, including the existing origin-gated session bootstrap. The shared helper preserves the prior 401/403 behavior. Already-consumed invalid JSON/object and business errors do not set the marker. The finalizer and its absolute deadline, byte/chunk caps, response-first half-close and standard finally cleanup are unchanged from the prior independently accepted implementation.

Independent run of all dedicated cases plus unchanged original bridge tests: **46 passed in 9.54 seconds**, including real TCP strict400 delivery for actual MAX+1 valid JSON and zero registration, rejected framing never reaching the body read, bounded slow/missing/over-cap teardown, consumed-error controls, and previous strict401/403 behavior. Original `test_bridge.py` has no diff; changed-file whitespace check passed. The separately retained ten-request experiment was reviewed as implementation evidence, not rerun as a duplicate native/transport gate.

No production/permanent-test/index/native state was changed by this reviewer; only this report and owned scratch test output changed. Full34 / CI05 bind the older snapshot, and the new source still needs full35 / CI06. This acceptance does not establish the unique cause of the earlier Windows WinError10053 or guarantee delivery for unlimited/indefinitely slow senders.
