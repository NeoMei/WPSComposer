# Strict native Excel save and close acknowledgement

The public macOS Excel session now requires a Boolean `saved: true` before validation/publication and `closed: true` after native close. Missing, false and malformed acknowledgements enter the existing quarantine/recovery path. `save`, `save_current` and delegated `save_copy` share the check. Attached close behavior is unchanged.

The old implementation failed all 12 fault-injection cases in `ack-red.log`. The final focused run passes 109 tests. The initial obsolete close-script assertion and the mistaken test-file invocation remain separately retained. Independent scoped review found no remaining issues. Full41 passes 5,059 tests / 12 skips from immutable snapshot `2775c23cb17a476394910099fa3197c3fe63078b`; all 386 hashes remain unchanged and equal to live.

`native-run-01/report.json` passes three allowed fresh-empty-sheet delete/write/save/reopen cases, exact final sheet names and markers, source preservation, an independent unsaved sentinel and workbook-inventory restoration. The retained runtime includes actual strict save/close acknowledgements. Root independently checked all recorded live source hashes. The first direct-script invocation failed import before any native action; the successful module invocation is separately retained.

Existing nonempty worksheet deletion is still guarded in production. Native save/close acceptance does not establish UI acceptance for every application, full parity, Windows native completion or release approval.
