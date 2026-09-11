# Scoped fixture review: Word checkpoint sentinel cleanup

Date: 2026-09-09

**SCOPED PASS — no P1/P2.** Compared with the run-04 source copy, the current
fixture changes only sentinel cleanup/evidence and its two pure tests. Production
checkpoint/session sources are unchanged.

The primary owned `MacWordSession` context now exits before the sentinel is
closed. The immediate inventory requires the exact unsaved sentinel row to equal
its pre-run name/path/saved/hash tuple and requires every non-sentinel row to
equal the starting inventory. This provides explicit evidence that the owned
document closed while the unrelated sentinel remained open and unchanged.

`close_sentinel` is independent from the closed session and targets the exact
synthetic name. Before `saving no`, it rechecks exact name, empty path, exact
token plus Word's terminal return, and `saved == false`; it then requires an
exact `[["sentinel-closed", name]]` acknowledgement. A second inventory must
equal the starting inventory. Failure leaves `sentinel` populated in the report
for guarded recovery.

The fixture retains the same main recovery assertions and does not widen the
rollback claim. The old run-04 failure remains bound to its old fixture hash and
is not rewritten as passing evidence.

Frozen fixture/test hashes reviewed:

- `fixtures/microsoft_parity/macos_word_recovery.py`:
  `7de1c9bb54f3d6fea6c6f90da049f7b74eb7a1ab8ceedb8846396f4e7de8900f`
- `tests/msoffice/test_macos_word_recovery.py`:
  `3c7710b146341b2c320d5ff7a1ca872a46ecac70511aa4b778341db120fc0051`

Independent focused result: `2 passed, 46 deselected in 0.28s`. No Office
process was started by this review.
