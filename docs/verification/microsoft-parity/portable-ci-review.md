# Portable CI scoped review

Decision: **PASS**. No actionable findings in the workflow or README CI paragraph.

## Scope and source binding

Reviewed `.github/workflows/portable-tests.yml`, the README CI paragraph, and the implementation report for context against committed candidate `202641f2e160f5ff7edf68a39a560e276c86ab46`. The workflow is new and the README paragraph is uncommitted at review time. Reviewed file SHA-256 values:

- `.github/workflows/portable-tests.yml`: `f1828637b10b4f46284e1bd802223ca2ac0e4bde6e84c3eed9b156131e907a71`
- `README.md`: `ab6fa8fe5a7c5a52a1c3e0affe1bef49dbf916f9916337366640f63e39a6b1a0`

Concurrent Word feature and RED-test changes were excluded. This review does not certify the mutable worktree or a successful hosted run.

## Assessment

- The four matrix combinations correctly cover Linux/Windows and Python 3.9/3.12. `fail-fast: false` retains results from other combinations after a failure; the maximum parallelism and 30-minute job timeout are bounded.
- Windows disables `core.autocrlf` before checkout, preserving exact evidence bytes. Checkout fetches full history and explicitly verifies the frozen baseline commit before dependency setup or tests. The baseline exists in the reviewed local repository and is used by the capability tests through `git show`.
- Python setup precedes pip upgrade, editable `.[dev]` installation, and `pip check`. Node.js setup precedes locked npm resource preparation with `npm ci --ignore-scripts`. No plugin installation or native acceptance flags are introduced.
- The pytest step runs the complete configured suite. Bash `pipefail` preserves pytest failures through `tee`; the always-running upload targets only the log and JUnit report. Unique matrix artifact names and 14-day retention are correct. Setup failures without reports produce a warning rather than inventing test evidence.
- The source repository plus branch concurrency key separates fork branches with identical names while intentionally sharing push/PR runs from the same source branch. Canceling older runs is explicit. No branch protection changes are introduced.
- Permissions are read-only for repository contents; checkout does not persist credentials. All four actions are pinned to full commit SHAs. The parent separately verified the official tag-to-commit mapping; this review did not repeat that network lookup.
- README accurately describes the triggers, matrix, retention, and native acceptance boundary. Hosted portable results remain separate from actual WPS/Microsoft Office generation, editing, undo, and close/reopen acceptance.

## Independent validation

- Ran official actionlint v1.7.12 against the exact workflow: exit 0, no diagnostics.
- Parsed the workflow with PyYAML BaseLoader and ran `bash -n` on every run block: all four passed.
- Executed the exact pytest block under the Actions-style Bash invocation (`bash --noprofile --norc -eo pipefail`) with isolated synthetic Python commands returning 0, 1, and 2. All three exit codes were preserved, and each retained the log and XML file.
- Checked the parsed matrix, permissions, checkout order, full-history setting, and disabled credential persistence. README whitespace validation passed.
- No full suite rerun, installation, push, or native application action was performed. Hosted Windows/Linux dependency resolution and test outcomes remain the next execution gate after the parent commits and pushes the reviewed candidate.

Only this review report was written.
