# WPSComposer 0.7.2 Release Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Publish the approved WPSComposer layout and numbering fixes as version `0.7.2` on GitHub `master`.

**Architecture:** Keep the existing public API and update only the two authoritative version sources plus the README changelog. Reuse PR #3 for review and merge, then verify the remote `master` commit and release metadata after the merge.

**Tech Stack:** Python package metadata, Codex plugin JSON manifest, Markdown, pytest, Git, GitHub CLI.

## Global Constraints

- Release version is exactly `0.7.2`.
- Keep all existing public APIs compatible.
- Do not commit the untracked `uv.lock`.
- Do not publish customer bid documents or acceptance artifacts.
- Do not create a Git tag or GitHub Release in this task.
- Target branch is GitHub `master` through PR #3.

---

### Task 1: Synchronize release metadata and README

**Files:**
- Modify: `pyproject.toml:7`
- Modify: `.codex-plugin/plugin.json:3`
- Modify: `README.md:365-373`

**Interfaces:**
- Consumes: current version string `0.7.1` in Python and plugin metadata.
- Produces: consistent version string `0.7.2` and a dated README changelog entry.

- [ ] **Step 1: Run the release assertion before editing**

Run:

```bash
.venv/bin/python - <<'PY'
import json, pathlib, tomllib
root = pathlib.Path('.')
pyproject = tomllib.loads((root / 'pyproject.toml').read_text())
plugin = json.loads((root / '.codex-plugin/plugin.json').read_text())
assert pyproject['project']['version'] == '0.7.2'
assert plugin['version'] == '0.7.2'
assert '### v0.7.2 (2026-08-19)' in (root / 'README.md').read_text()
PY
```

Expected: FAIL because the current metadata is `0.7.1` and the README has no `v0.7.2` entry.

- [ ] **Step 2: Update both version sources**

Change `pyproject.toml` to:

```toml
version = "0.7.2"
```

Change `.codex-plugin/plugin.json` to:

```json
"version": "0.7.2"
```

- [ ] **Step 3: Add the README changelog entry**

Insert above `v0.7.1`:

```markdown
### v0.7.2 (2026-08-19)
- 修复 WPS 表格单元格继承正文缩进的问题，表内段落显式使用零缩进。
- 原生多级编号可与既有 `numbering.xml` 安全合并，支持中文章号、十进制节号和关键工法编号联动。
- 使用 WPS 原生目录域，并支持宽图横向分节、混合方向表格宽度与尾部空白页修复。
```

- [ ] **Step 4: Rerun the release assertion**

Run the Step 1 command again.

Expected: PASS with exit code 0.

- [ ] **Step 5: Commit the release metadata**

```bash
git add pyproject.toml .codex-plugin/plugin.json README.md
git commit -m "Release WPSComposer 0.7.2"
```

Expected: one commit containing only release metadata and README. This implementation plan is committed before isolated execution starts.

---

### Task 2: Validate and merge PR #3

**Files:**
- Verify: all tracked files in PR #3
- Preserve untracked: `uv.lock`

**Interfaces:**
- Consumes: Task 1 commit on `codex/wpscomposer-native-numbering-table-layout`.
- Produces: merged PR #3 and remote `master` containing version `0.7.2`.

- [ ] **Step 1: Run the complete test suite**

```bash
.venv/bin/python -m pytest -q
```

Expected: `829 passed` and exit code 0.

- [ ] **Step 2: Verify release scope and whitespace**

```bash
git diff --check origin/master...HEAD
git status -sb
git diff --name-only origin/master...HEAD
```

Expected: no whitespace errors; `uv.lock` remains untracked and absent from the PR diff; no customer workspace files appear.

- [ ] **Step 3: Push and make PR #3 ready**

```bash
git push
gh pr ready 3
gh pr view 3 --json isDraft,state,headRefName,baseRefName,url
```

Expected: PR #3 is open, not draft, with head `codex/wpscomposer-native-numbering-table-layout` and base `master`.

- [ ] **Step 4: Merge PR #3 into master**

```bash
gh pr merge 3 --merge --delete-branch
```

Expected: PR #3 reports `MERGED`; the remote feature branch is deleted.

- [ ] **Step 5: Verify the merged remote baseline**

```bash
git fetch origin
git switch master
git pull --ff-only origin master
gh pr view 3 --json state,mergedAt,mergeCommit,url
.venv/bin/python - <<'PY'
import json, pathlib, tomllib
root = pathlib.Path('.')
pyproject = tomllib.loads((root / 'pyproject.toml').read_text())
plugin = json.loads((root / '.codex-plugin/plugin.json').read_text())
assert pyproject['project']['version'] == '0.7.2'
assert plugin['version'] == '0.7.2'
assert '### v0.7.2 (2026-08-19)' in (root / 'README.md').read_text()
PY
```

Expected: PR #3 state is `MERGED`; local and remote `master` match; all three version/document assertions pass.

- [ ] **Step 6: Run post-merge regression tests**

```bash
.venv/bin/python -m pytest -q
```

Expected: `829 passed` on the merged `master` baseline.
