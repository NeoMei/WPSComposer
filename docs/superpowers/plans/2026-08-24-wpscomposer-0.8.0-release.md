# WPSComposer 0.8.0 Release Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Publish the cross-platform accepted long-form engine as WPSComposer `0.8.0` on GitHub `master`, with an annotated tag and GitHub Release.

**Architecture:** Keep the accepted runtime unchanged and make one release-metadata commit on `codex/longform-m3`. Validate that commit, integrate it into `master` without touching the dirty local `master` checkout, validate the exact remote merge result from an isolated checkout, then publish and verify `v0.8.0`.

**Tech Stack:** Python package metadata, Codex plugin JSON manifest, Markdown, pytest, Git, GitHub.

## Global Constraints

- Release version is exactly `0.8.0`.
- Runtime generation, COM, JSAPI, protocol, and evidence code must not change during release preparation.
- Preserve the dirty local `master` checkout and its untracked files.
- Do not commit `.DS_Store`, `uv.lock`, native evidence under `build/`, or customer documents.
- The accepted source branch is `codex/longform-m3` at or after `d65a541`.
- GitHub `master`, annotated tag `v0.8.0`, and GitHub Release must resolve to the same release commit.

---

### Task 1: Synchronize 0.8.0 release metadata

**Files:**
- Modify: `tests/test_plugin_manifest.py`
- Modify: `tests/test_documentation.py`
- Modify: `pyproject.toml`
- Modify: `.codex-plugin/plugin.json`
- Modify: `README.md`
- Modify: `skills/WPSComposer/SKILL.md`
- Modify: `docs/longform-markdown.md`
- Modify: `docs/longform-m0.md`
- Modify: `docs/macos-longform-m5-verification.md`
- Modify: `docs/windows-verification.md`
- Modify: `.superpowers/sdd/progress.md`

**Interfaces:**
- Consumes: accepted `0.7.2` metadata and completed cross-platform M5 evidence.
- Produces: consistent public version `0.8.0` and release-facing documentation.

- [x] **Step 1: Change the release contract tests to require 0.8.0**

Update the plugin manifest assertion to `0.8.0`. Update the public-status documentation contract to require `0.8.0 released` and the README changelog heading `### v0.8.0 (2026-08-24)`.

- [x] **Step 2: Run the changed contract tests and verify RED**

Run:

```bash
uv run --extra dev python -m pytest -q \
  tests/test_plugin_manifest.py tests/test_documentation.py
```

Expected: failure because metadata and release-facing documents still say `0.7.2` or unreleased.

- [x] **Step 3: Update both version sources and release-facing documentation**

Set `project.version` and `.codex-plugin/plugin.json` to `0.8.0`. Add a dated README entry describing the default M5 long-form route, native numbering/TOC/sections, bounded quality lifecycle, visible degradation, and dual-platform evidence. Change only current status text; preserve historical plan and changelog statements.

- [x] **Step 4: Run the release contract tests and verify GREEN**

Run the Step 2 command again.

Expected: all selected tests pass.

- [x] **Step 5: Commit release metadata and this plan**

```bash
git add .codex-plugin/plugin.json pyproject.toml README.md \
  skills/WPSComposer/SKILL.md docs/longform-markdown.md \
  docs/macos-longform-m5-verification.md \
  docs/superpowers/plans/2026-08-24-wpscomposer-0.8.0-release.md \
  .superpowers/sdd/progress.md tests/test_plugin_manifest.py \
  tests/test_documentation.py
git commit -m "Release WPSComposer 0.8.0"
```

### Task 2: Validate the release candidate

**Files:**
- Verify: all tracked files
- Preserve: `.DS_Store`

**Interfaces:**
- Consumes: Task 1 release commit.
- Produces: a fully verified, pushable release candidate.

- [x] **Step 1: Run the complete platform-independent suite**

```bash
uv run --extra dev python -m pytest -q
```

Expected: zero failures; real native gates remain explicitly skipped because their accepted evidence is already recorded.

- [x] **Step 2: Verify metadata, package build, and plugin bundle**

```bash
uv run --with build python -m build
uv run --extra dev python -m pytest -q tests/test_plugin_manifest.py tests/test_installer.py
git diff --check origin/master...HEAD
```

Expected: wheel and sdist build, tests pass, and no whitespace errors.

- [x] **Step 3: Verify release scope**

Confirm that `git status --short` contains only the pre-existing untracked `.DS_Store`, and that the release commit contains no `build/`, `uv.lock`, or customer document.

### Task 3: Integrate the release into GitHub master

**Files:**
- External write: GitHub branch `master`
- Preserve: dirty local `master` checkout

**Interfaces:**
- Consumes: verified release candidate on `codex/longform-m3`.
- Produces: GitHub `master` containing version `0.8.0`.

- [x] **Step 1: Push the release candidate branch**

```bash
git push origin codex/longform-m3
```

- [x] **Step 2: Integrate through a GitHub PR when authenticated**

Create a PR from `codex/longform-m3` to `master`, verify its head/base and mergeability, then merge it. If GitHub PR credentials are unavailable, use a non-force fast-forward push only after proving `origin/master` is an ancestor of the release candidate:

```bash
git merge-base --is-ancestor origin/master HEAD
git push origin HEAD:master
```

- [x] **Step 3: Verify remote master**

```bash
git fetch origin
git rev-parse origin/master
git show origin/master:pyproject.toml
git show origin/master:.codex-plugin/plugin.json
```

Expected: both version sources report `0.8.0` at the remote release commit.

### Task 4: Validate the exact merged result

**Files:**
- Create temporarily: isolated detached worktree for `origin/master`
- Remove after verification: only that temporary worktree

**Interfaces:**
- Consumes: remote `master` release commit.
- Produces: fresh post-merge test evidence from the exact published tree.

- [x] **Step 1: Create an isolated detached worktree at origin/master**

Use `mktemp -d`, add a detached Git worktree under that directory, verify its
HEAD equals `origin/master`, then install the pinned WPS JSAPI template runtime:

```bash
(cd macos/wps-jsapi-probe && npm ci)
```

The template-backed generation tests require this clean-checkout setup; without
it they fail at template staging rather than exercising the intended behavior.

- [x] **Step 2: Run the full suite and release assertions there**

```bash
uv run --extra dev python -m pytest -q
```

Expected: zero failures, with version assertions at `0.8.0`.

- [x] **Step 3: Remove only the temporary worktree**

Remove it through `git worktree remove` from outside that worktree and prune stale registrations. Do not modify or clean the user's main checkout.

### Task 5: Publish and verify v0.8.0

**Files:**
- External write: annotated Git tag `v0.8.0`
- External write: GitHub Release `v0.8.0`

**Interfaces:**
- Consumes: post-merge verified `origin/master` commit.
- Produces: public version tag and GitHub Release resolving to the same commit.

- [ ] **Step 1: Create and push the annotated tag**

```bash
git tag -a v0.8.0 origin/master -m "WPSComposer 0.8.0"
git push origin v0.8.0
```

- [ ] **Step 2: Create the GitHub Release**

Use the `v0.8.0` tag and release notes summarizing the M5 long-form default route, dual-platform native evidence, bounded relayout/notice lifecycle, Unicode/code block preservation, native numbering/TOC/sections, and installer hardening.

- [ ] **Step 3: Reverse-check all release surfaces**

Verify that `origin/master`, `refs/tags/v0.8.0^{}`, package metadata, plugin manifest, and the GitHub Release target all resolve to the same commit. Record the final URLs and test totals.
