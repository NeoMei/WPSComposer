# Installer scratch directory exclusion

Date: 2026-09-09 (Asia/Shanghai)

## Scope and diagnosis

Bounded Task 6 audit repair in the office-description worktree. Parent audit 08 reported installation copying 536 local scratch files (about 5.4 MB); that count is parent-provided context, not a new measurement in this subtask.

Inspection confirmed `_copy_plugin` uses `shutil.copytree` with explicit recursive ignore names, without consulting Git ignore state. `.superpowers` and `.worktrees` were absent, so both were copied. The repository `.gitignore` already lists `.worktrees/`. No CodeGraph directory exists in this worktree.

## Change

Added only exact `.superpowers` and `.worktrees` entries to the existing ignore set. No broad dot-directory glob or other cleanup was added. Existing copy, operator-document restoration, and installation transaction behavior remain unchanged.

The new parameterized regression runs the public `install_plugin` against a synthetic source entirely in pytest temporary storage. Each name appears both at source root and beneath the public skill directory. It verifies:

- Both local work directory locations are excluded from the installed artifact.
- Manifest, skill entry, API reference, public facade script, and all three required operator documents retain their exact contents.
- Similarly named `.superpowers-example` and `.worktrees-example` asset directories remain included.
- Source scratch files remain intact.

The existing test module stubs macOS runtime dependency installation; file copying, staging, and marketplace creation remain real. No personal installation or native Office application was used.

## Verification

The first attempted command used the worktree-local `.venv/bin/python`, which does not exist (exit 127). All actual tests then used the existing main-checkout virtualenv interpreter below.

RED before production edit:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/test_installer.py -k local_work_directories -q
2 failed, 9 deselected in 0.04s
```

Both failures were the expected assertion at `tests/test_installer.py:131`: the destination scratch directory existed. No setup/import failure occurred.

GREEN after the two-name ignore change:

```text
/Users/neomei/项目/codexprojects/WpsComposer/.venv/bin/python -m pytest tests/test_installer.py -q
11 passed in 0.21s

git diff --check
exit 0, no output
```

Full-suite verification is left to the parent coordinator after concurrent changes settle. No commit or push was performed. Other agents' pre-existing and concurrent changes were not edited.

## Source identity

HEAD at verification: `cb9898f28e42927248a42421858d549edc54bc3b`.

SHA-256 of complete file bytes:

| File | HEAD before this repair | After repair |
| --- | --- | --- |
| `install.py` | `d71dfdc458c5ed55ec2c9a70c40f97514d9a018594d5a5a853711161cdeab961` | `f54c45f544f3cee8a0e184357a787754f2cd5cbf01788739eadf883af40587c2` |
| `tests/test_installer.py` | `10791625c3f669e4c93a0ed7653b363f895adecfcf6294c5ce404627ed78741f` | `d2734dab7dfcf946b1fd1bfbcb31df390d7a055bc551a47d7501ed1f6f19339d` |

Exact scoped patch:

```diff
diff --git a/install.py b/install.py
index 40fd933..7f1f7f7 100644
--- a/install.py
+++ b/install.py
@@ -25,7 +25,9 @@ OPERATOR_DOCS = (
 IGNORED_NAMES = {
     ".git",
     ".pytest_cache",
+    ".superpowers",
     ".venv*",
+    ".worktrees",
     "__pycache__",
     "*.egg-info",
     "build",
diff --git a/tests/test_installer.py b/tests/test_installer.py
index 979727f..542224a 100644
--- a/tests/test_installer.py
+++ b/tests/test_installer.py
@@ -99,6 +99,41 @@ def test_installer_skips_virtualenvs_and_build_metadata(tmp_path):
     assert (result.destination / ".codex-plugin" / "plugin.json").is_file()
 
 
+@pytest.mark.parametrize("scratch_name", [".superpowers", ".worktrees"])
+def test_installer_excludes_local_work_directories_but_retains_plugin_assets(
+    tmp_path, scratch_name
+):
+    source = tmp_path / "source"
+    retained_files = {
+        ".codex-plugin/plugin.json": '{"name": "wps-composer"}',
+        "skills/WPSComposer/SKILL.md": "public skill",
+        "skills/WPSComposer/references/api.md": "public API",
+        "skills/WPSComposer/scripts/wps_engine.py": "# public facade\n",
+        "docs/windows-verification.md": "Windows operator guide",
+        "docs/longform-markdown.md": "long-form operator guide",
+        "docs/macos-longform-m5-verification.md": "macOS operator guide",
+        "assets/.superpowers-example/template.txt": "template asset",
+        "assets/.worktrees-example/template.txt": "template asset",
+    }
+    scratch_files = {
+        f"{scratch_name}/task/private-evidence.txt": "local task evidence",
+        f"skills/WPSComposer/{scratch_name}/task/private-evidence.txt": "nested work",
+    }
+    for relative, content in {**retained_files, **scratch_files}.items():
+        path = source / relative
+        path.parent.mkdir(parents=True, exist_ok=True)
+        path.write_text(content, encoding="utf-8")
+
+    result = install_plugin(source, tmp_path / "home", tmp_path)
+
+    for relative, content in retained_files.items():
+        assert (result.destination / relative).read_text(encoding="utf-8") == content
+    assert not (result.destination / scratch_name).exists()
+    assert not (result.destination / "skills/WPSComposer" / scratch_name).exists()
+    for relative, content in scratch_files.items():
+        assert (source / relative).read_text(encoding="utf-8") == content
+
+
 def test_installer_refuses_existing_destination_without_force(tmp_path):
     codex_home = tmp_path / ".codex"
     destination = codex_home / "plugins" / "wps-composer"
```
