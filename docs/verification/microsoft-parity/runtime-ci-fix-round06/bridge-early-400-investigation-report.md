# Bounded adjacent-branch investigation: early HTTP 400 teardown

Finding: reproducible P2 candidate, awaiting root decision before any source/test change. HEAD `370c8c5de176ec8bdcfe7c50263a89b227af2599`; bridge SHA-256 `97be582ec55b2807e4cce2e6f4ad13cbe90e780ce3c40ee254dd0f5e59749ef8`.

## Reproduction

Executed exactly ten local HTTP requests through an isolated real `LoopbackBridge`, with an allowed origin and a valid writer session token created from the bridge's own bootstrap credentials. No production monkeypatch was applied. Each socket had a one-second timeout; the entire executable completed in under one second. No Office/WPS/native/UI operation, source edit, or index operation occurred.

| Framing/body | Requests | Result |
| --- | ---: | --- |
| Invalid Content-Length, valid 118-byte JSON | 2 | 2 exact HTTP 400 / INVALID_REQUEST / Invalid Content-Length |
| Content-Length 1048577, actual valid JSON 1048577 bytes (`MAX_BODY_BYTES + 1`) | 4 | 4 BrokenPipeError / errno 32 during client body send; no HTTP status obtained |
| Invalid Content-Length, actual same 1048577-byte valid JSON | 4 | 2 BrokenPipeError / errno 32 and 2 ConnectionResetError / errno 54 during client body send; no HTTP status obtained |

No component was registered in any case. Socket errors were retained as failure outcomes, never accepted as equivalent to HTTP 400. The experiment ends the normal request on send failure; it does not establish that every error response byte was erased from the kernel buffers. It directly establishes that a normal send-then-read HTTP client cannot reliably obtain the intended rejection.

## Code path and bounded recommendation

`_read_json()` converts Content-Length and checks its sign/size before `rfile.read(length)`. Conversion failure or negative/oversized length raises ValueError; `do_POST()` catches it and emits 400 / INVALID_REQUEST. These branches leave the request body unread. The current staged finalizer is flagged only by `_send_error_json()` for POST 401/403, so these early 400 branches immediately finish without the drain opportunity already provided to authentication/origin rejections.

If root authorizes repair, mark the existing staged-close flag and `close_connection` only at the pre-read Content-Length conversion/sign/size rejection sites (or through a tiny shared marking helper), before raising the same ValueError. Retain the existing response code/message, authentication/origin ordering, MAX_BODY_BYTES+1 discard cap and absolute 0.2-second deadline. Do not flag every 400: JSON-decoding/object-shape/business validation errors after the declared body has been read are a different path and need no added delay for this repair.

Meaningful follow-up tests: strict HTTP400 delivery for an authenticated allowed-origin actual MAX_BODY_BYTES+1 request; invalid/negative length with body; body parser/business never reached; bounded missing/slow/over-cap peers under the same finalizer; already-consumed invalid JSON and ordinary 400 remain unflagged. No candidate prototype or patch was executed in this investigation, respecting the ten-request cap and root-decision boundary.

## Preserved evidence

- Executable: `bridge_early_400_transport_experiment.py`
- Raw per-request outcomes and summary: `bridge-early-400-raw.log`
- Command: `PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python .superpowers/sdd/2026-09-08-microsoft-wps-parity/bridge_early_400_transport_experiment.py`

This is a locally reproduced adjacent-path defect. It does not establish a Windows result or the unique cause of the original CI small-request WinError10053. CI05 remains an independent running verification surface.
