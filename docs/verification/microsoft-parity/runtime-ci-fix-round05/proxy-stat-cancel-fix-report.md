# Proxy secondary metadata cancellation repair

Implementation frozen for independent recheck. This is a post-full33 correction requiring the next full34 snapshot; no full33 export was changed.

## Defect and narrow repair

The integration review reproduced a secondary `KeyboardInterrupt` / `SystemExit` raised by `Path.is_file()` while `_SessionProxy._error()` assembles optional recovery/quarantine paths. The timeout cleanup had already completed, but the secondary signal replaced the promised `NATIVE_WORD_TIMEOUT` / `NATIVE_OFFICE_TIMEOUT` error.

Only the `evidenced()` exception boundary changed from `OSError` to `BaseException`, with a comment explaining why optional metadata cannot replace the primary failure. A failed lookup returns `None` for that path; the other path is still checked independently. Public error construction, error categories, primary cancellation handling, worker ownership, cleanup, and quarantine persistence are unchanged.

Source: `skills/WPSComposer/scripts/msoffice/windows_session_proxy.py`.
Permanent regression: `tests/msoffice/test_windows_session_proxy_diagnostics.py`.

## Evidence

Used the existing complete clean-dev interpreter with `PYTHONPATH=.:tests/msoffice`.

1. Original reviewer scratch unchanged: **12 failed in 1.41s**, log `proxy-stat-cancel-red.log`.
2. New permanent regression before production change: **12 failed, 19 deselected in 0.92s**, log `proxy-stat-cancel-permanent-red.log`.
3. After the narrow repair, command:

```sh
PYTHONPATH=.:tests/msoffice /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest -q tests/msoffice/test_windows_session_proxy.py tests/msoffice/test_windows_session_proxy_diagnostics.py .superpowers/sdd/2026-09-08-microsoft-wps-parity/checkpoint-integration-scratch/test_proxy_stat_cancel_integration.py
```

**86 passed in 8.12s**, log `proxy-stat-cancel-green.log`.

The permanent matrix covers three applications × two secondary signals × two evidence paths. It requires the public timeout class/code, an absent unverified path and retained other path, completed child/lock cleanup, actual recovery/quarantine persistence, subsequent QUARANTINED classification, and no second worker through the quarantined lock. Existing cancellation identity, close/save, and ordinary proxy tests were included unchanged.

`git diff --check` for both changed files passed. No shared-index operation, commit, native Office process, or unrelated source edit was performed.

## Freeze

- Source SHA-256: `dfa9abfba9962db621a63bccf4058d923ad293a2c4f93888aa6ec507d5294b77`
- Permanent diagnostic test SHA-256: `5f8916ce8eaf6e83d25dcd88939fd472b594f34cffbb546a1aee5dfd040020f3`
- Patch: `proxy-stat-cancel-frozen.patch`
- Original untouched scratch: `checkpoint-integration-scratch/test_proxy_stat_cancel_integration.py`

Independent reviewer approval and full34 validation remain pending.

## Independent acceptance — checkpoint_integration_review

**PASS: the full33 integration P2 is closed by this post-full33 candidate.** Independently matched source `dfa9abfba9962db621a63bccf4058d923ad293a2c4f93888aa6ec507d5294b77` and permanent tests `5f8916ce8eaf6e83d25dcd88939fd472b594f34cffbb546a1aee5dfd040020f3`. Only the best-effort evidence lookup catch changed; error category construction, abort/lock/child state, real cancellation handling and request protocol remain unchanged. Returning no path when its metadata cannot be observed preserves truthful locations without replacing the selected timeout.

Independent focused rerun of the 12 permanent cases plus the original 12 reviewer cases: **24 passed, 19 deselected in 2.29 seconds**. The original scratch assertions are unchanged, SHA-256 `56e150948b44396267b2cf76a8f2ca2c36b8b4950a4c5f01849759047899e829`. The matrix confirms all three proxy components, both secondary signal types, both evidence locations, retained typed timeout, real marker persistence, exact Python-child and lock cleanup, later in-memory quarantine, and blocked second worker. Diff whitespace check passed.

No additional actionable P1/P2 was found in this bounded correction. These exact two hashes may be frozen for full34. The root's reported full33 result belongs to the earlier unchanged snapshot and does not certify this later source. Full34, Windows CI, installation and native/UI parity remain separate gates. Reviewer changed only review reports and owned scratch test output, with no production/permanent-test/index/native changes.
