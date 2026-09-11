# Continuation integration fix wave

Base cb9898f28e42927248a42421858d549edc54bc3b, existing worktree only.
Read continuation-integration-review.md and its repros first once finalized.
Root independently reproduced the destructive rollback P2 and inspected both
other checker/test failures. This brief authorizes only the complete review fix
wave; no additional feature work or new native heading implementation.

## Required repairs

1. macos_word_recovery.rollback: ordinary native error after destructive rollback
submission must quarantine/retain the session and prevent any later native write
or context-close retry. Preserve public LOCAL_MUTATION_ROLLBACK_FAILED and exact
private diagnostics. Invalidate observed topology on actual mutation submission,
using existing session infrastructure. Pre-submit argument, preimage, script I/O
or deadline failures that never reached Word must not be mislabeled as partial
native mutation. Timeout, host loss, malformed acknowledgement and cancellation
retain existing quarantine and interrupt semantics. Do not invent a shared hook
without demonstrating it is necessary. Real MacWordSession._execute with patched
subprocess, not a fake that automatically quarantines every exception, must prove
ordinary -2700 error after partial mutation, prevention of later sends, preserved
tracking/cache on pre-submit failure, and cache restoration only after valid ACK.

2. PowerPoint deadline tests: test_publication_uses_session_deadline and
 test_office_artifact_validation_uses_same_deadline currently set absolute 123.
On this macOS Python3.9, monotonic starts near process start. They pass alone but
fail when the suite reaches them after 123s. Use a controlled clock/deadline and
faithful source/publication stubs or temporary artifact bytes so assertions prove
same budget propagation without depending on host/process uptime. Do not weaken
production timeout or introduce test-specific production behavior. Cover a large
mocked monotonic value with no real wait. See unit-round22.txt and root diagnosis.

3. Paragraph-rule PDF verifier: run-02 completed all 11 native/XML/reopen/cleanup
checks, but its PDF detector accepts only stroke color. Real Word exports one
opaque thin filled C0C0C0 rectangle (height~.72pt,width~418pt). Accept real horizontal
stroke or thin filled rectangle representations with relevant geometry, matching
color and opacity; do not count arbitrary gray shapes, vertical lines, transparent
fills or thick boxes. Preserve all text/style/native assertions. Add RED real
run02 drawings and negative geometries/styles, then GREEN. Revalidate the original
run02 PDF/DOCX read-only using the repaired detector, save separate corrected
report/source hashes, and leave original run02 FAIL unchanged. This readonly
checker correction alone does not need a third native run.

## Boundaries / evidence

No native Office until root grants sole lease after independent review. No
personal install, git commit/push, user security changes, other worktrees or
subagents. Python3.9/lazy imports/native backend/engine defaults/ownership and
absolute deadlines stay intact. Production changes are scoped recovery; tests
and fixture changes as listed. Any additional reviewer issue must come from final
review file rather than extending scope speculatively. Run targeted RED/GREEN,
report exact commands/results/hashes and write continuation-fixwave-report.md.
Return DONE/concerns with report path; root coordinates native and full suite.
