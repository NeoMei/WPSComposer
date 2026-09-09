# Independent scoped review: installer scratch-directory exclusion

Date: 2026-09-09  
Scope: the two-name `install.py` exclusion change and its two parameterized
regression cases. This review made no production/test edits, performed no
personal installation, started no native application, and made no commit.

## Verdict

**SCOPED PASS — no actionable P1/P2.** The change excludes only directory or
file entries whose basename is exactly `.superpowers` or `.worktrees`, at the
source root or any recursively visited directory. It does not exclude ordinary
plugin skills, the restored operator documents, or resources with similar
prefixes.

## Frozen source identity

| File | SHA-256 |
|---|---|
| `install.py` | `f54c45f544f3cee8a0e184357a787754f2cd5cbf01788739eadf883af40587c2` |
| `tests/test_installer.py` | `d2734dab7dfcf946b1fd1bfbcb31df390d7a055bc551a47d7501ed1f6f19339d` |
| `installer-scratch-fix-report.md` | `9cade0164717ab54d650e0b33b651bda51d85f35c2090f88b4455a920205390c` |

The production diff adds only the exact strings `.superpowers` and
`.worktrees` to the existing `IGNORED_NAMES` set. Copy, staging, runtime
installation, destination replacement, marketplace update, rollback, and
operator-document restoration code are unchanged.

## Matching and retained-content review

`shutil.ignore_patterns()` receives the sorted basename patterns through
`shutil.copytree`. `copytree` calls the ignore function for every recursively
visited source directory, so exact scratch names are excluded at any depth.
Neither new pattern contains a wildcard. An independent direct check returned
exactly these ignored names from the following candidates:

```text
.superpowers
.worktrees
.superpowers-example
.worktrees-example
.superpowers.md
my.worktrees
SKILL.md
docs
```

Therefore `.superpowers-example`, `.worktrees-example`, `.superpowers.md` and
`my.worktrees` remain eligible for copying. Inspection of the current public
`skills`, `docs`, and `macos` trees found no exact `.superpowers` or
`.worktrees` entry that serves as a shipped resource; the only current exact
directory is the repository-root scratch workspace.

The existing broad `docs` exclusion remains intentional and unchanged:
`_copy_plugin` recreates the destination `docs` directory and copies the three
required operator documents explicitly. The regression verifies exact contents
for those three documents together with `.codex-plugin/plugin.json`,
`skills/WPSComposer/SKILL.md`, the API reference, the public facade script, and
both similar-prefix asset directories. It also verifies nested and root scratch
directories are absent from the destination while source scratch bytes remain
untouched.

## Independent verification

Focused command:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest -q tests/test_installer.py
```

Result:

```text
11 passed in 0.24s
```

The module exercises real temporary-directory copy, staging, destination and
marketplace behavior while stubbing only macOS runtime dependency installation.
No user home or personal plugin path is used. `git diff --check` also passed.

This verdict is limited to installer scratch exclusion and regression safety.
It does not constitute a personal-install, packaged-manifest, or release
acceptance result.
