# Full regression49 — profile shutdown fix

Executed the complete pytest suite in the candidate worktree after the profile
request timeout fix and two new real-TCP fault-injection cases. Parent commit:
6ced108b39ce3f7e75e34f8d1635ae9e99ada528. The run used the working tree, not an
immutable snapshot;380 tracked Python/source-configuration file hashes were
captured before/after and all remain identical. The included hash maps bind this
result to the tested source. Native figure v3 is a separate build-only runner.

Command: clean-dev-venv/bin/python -m pytest -v --junitxml=build/release-followup-20260911/junit.xml

Result: **5,283 passed /12 skipped**, exit0, pytest237.69seconds (wrapper238.42).
The opt-in native/UI gates remain skipped and are not validated by this suite.
A terminal swigvarlink DeprecationWarning appears after the successful summary;
it is preserved in the raw log. This is not a release approval.

Targeted profile-server run:22 passed in20.15seconds. Before the fix, the two
new idle/partial-request fault-injection cases failed in12.09seconds because the
close thread remained blocked. Those failures were observed in this task's tool
output; the full regression log is retained here separately.
