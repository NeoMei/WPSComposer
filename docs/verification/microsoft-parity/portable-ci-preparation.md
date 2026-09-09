# Portable CI implementation report

Prepared 2026-09-09 for Task 6. This report covers workflow configuration and local validation only. No remote workflow was dispatched, no commit or push was made, and no native Office acceptance is claimed.

## Files

- `.github/workflows/portable-tests.yml`: new portable pytest workflow.
- `README.md`: one paragraph identifying CI scope, trigger branches, matrix, retained reports, and the separate native acceptance boundary.
- This published preparation record is copied from the task coordination report; hosted results remain a separate gate.

## Configuration

- All PRs plus pushes to `master`, `main`, and `codex/microsoft-parity`; `git remote show origin` confirms the repository currently defaults to `master`.
- Four jobs: `ubuntu-24.04` / `windows-2022` crossed with Python `3.9` / `3.12`. Matrix failure does not cancel the other jobs; maximum parallelism is four and timeout is 30 minutes per job.
- Workflow concurrency groups use the source repository and source branch across PR/push events, canceling older runs. Forks with the same branch name have independent groups. When push and PR events race, the later run may cancel the earlier event's run; this is deliberate duplicate-work suppression. No branch-protection rules are changed.
- Windows disables Git `core.autocrlf` before checkout. Checkout fetches all history and explicitly checks baseline commit `6dd3a00dff226096ad963cc68a65088865541c25` exists. Credentials are not persisted.
- Upgrade pip, install `.[dev]`, run `pip check`, and prepare locked resources using `npm ci --prefix macos/wps-jsapi-probe --ignore-scripts` under Node.js 22. No package caching is enabled. Portable COM tests inject fake modules; no direct pywin32 test import was found, so the native Windows extra is not installed.
- Run the full configured pytest suite without new skips or native acceptance environment flags. UTF-8 is explicit for portable subprocess/log handling.
- Bash pipefail preserves pytest failure through `tee`; only `test-results/pytest.log` and `test-results/pytest.xml` are uploaded with `always()`, retained 14 days, and uniquely named per matrix job. Setup failures with no report produce an upload warning. No environment dumps, credentials, or broad workspace uploads are configured.
- Permissions are `contents: read`; no plugin install, publish, release, merge, or native application action is present.
- Hosted macOS is excluded because dictionary compilation tests require an installed Microsoft Office application. Linux/Windows CI results do not replace real WPS/Microsoft Office generation, editing, undo, and close/reopen acceptance.

## Official action pins reviewed

Release notes, action metadata, and official Git tag resolutions were checked on 2026-09-09. All four actions use the Node.js 24 action runtime; this is separate from the Node.js 22 selected for project npm resources. GitHub-hosted runners supply the action runtime.

| Action | Release | Verified full commit |
| --- | --- | --- |
| actions/checkout | [v7.0.1](https://github.com/actions/checkout/releases/tag/v7.0.1) | `3d3c42e5aac5ba805825da76410c181273ba90b1` |
| actions/setup-python | [v7.0.0](https://github.com/actions/setup-python/releases/tag/v7.0.0) | `5fda3b95a4ea91299a34e894583c3862153e4b97` |
| actions/setup-node | [v7.0.0](https://github.com/actions/setup-node/releases/tag/v7.0.0) | `820762786026740c76f36085b0efc47a31fe5020` |
| actions/upload-artifact | [v7.0.1](https://github.com/actions/upload-artifact/releases/tag/v7.0.1) | `043fb46d1a93c77aae656e7c1c64a875d1fc6a0a` |

`git ls-remote https://github.com/actions/<action>.git refs/tags/<version>` resolved each pin. Reviewed metadata is at `https://raw.githubusercontent.com/actions/<action>/<version>/action.yml`. Release notes were fetched from official GitHub releases; an unauthenticated GitHub API rate limit was resolved using the existing `gh api` connector without exposing credentials.

The official [Python versions manifest](https://github.com/actions/python-versions/blob/main/versions-manifest.json) confirms x64 packages for both requested minor versions on Linux 24.04 and Windows. At this check, matching available patch versions were Linux 3.9.25 / 3.12.14 and Windows 3.9.13 / 3.12.10; the workflow intentionally requests minor versions and will follow available patches.

## Local verification

- Read repository AGENTS.md and `docs/verification/microsoft-parity/development-setup-audit.md` before implementation. There is no `.codegraph` directory in this worktree.
- Built official `github.com/rhysd/actionlint/cmd/actionlint@v1.7.12` into `/tmp/wps-portable-ci-tools.u1wBTg/bin`, with Go caches isolated under the same temporary directory. `/tmp/wps-portable-ci-tools.u1wBTg/bin/actionlint -color .github/workflows/portable-tests.yml` exited 0 with no diagnostics.
- Parsed workflow YAML using PyYAML BaseLoader, then executed `bash -n` on all four run blocks: all passed.
- Extracted the exact pytest run block and executed it with a synthetic Python command returning exit codes 0, 1, and 2. Every case preserved the same exit code and retained both the log and JUnit file. Each case used a separate temporary directory; no production/test sources were modified.
- README diff and new workflow whitespace checks produced no diagnostics.
- No full pytest rerun was performed for this configuration-only change in the mutable implementation worktree, which contains concurrent Word changes and intentional RED tests. Those changes were neither edited nor staged by this task. The parent owns checkpoint testing and remote execution after review/commit/push.

## Remaining gate

The workflow needs parent review and a committed push before its Linux/Windows jobs can be observed. Local actionlint and shell checks establish syntax and pipeline exit handling, not successful hosted dependency installation or a passing remote test suite. Existing portable failures must be fixed at their source; this workflow adds no exclusions to hide them.
