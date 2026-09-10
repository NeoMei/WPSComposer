# Immutable full regression round47

At snapshot `0b5ed7df6e9ca7f1143f674adb646d139e37d0b5` (tree `c1d682cf51ce8a7354b354483c41e474e4161640`, parentce322b29), pytest completed **5,274 passed / 12 skipped** in218.05s. All402 snapshot source hashes stayed unchanged.

The wrapper deliberately exits1 because the live numbering test file changed during independent fixture review. Pytest itself exits0, but `live_equal=false`; this is not exact-final-source acceptance. Production source hashes remain equal. Preserve this run separately from the forthcoming post-review snapshot.
