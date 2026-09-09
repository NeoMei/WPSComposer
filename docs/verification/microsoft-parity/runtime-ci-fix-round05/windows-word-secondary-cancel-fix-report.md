# Windows Word secondary diagnostic cancellation follow-up

Scope frozen for re-review. Only _evidenced_path in windows_runtime.py and eight dedicated test cases in test_windows_word_runtime_diagnostics.py changed. Previous patch/report evidence remains unchanged. No COM, native Office, index mutation, or whole-suite run.

Independent review's retained reproduction under checkpoint-integration-scratch/word-runtime/stat-cancel-repro showed timeout diagnostics alongside quarantine code NATIVE_WORD_EXECUTION_FAILED after a second KeyboardInterrupt from diagnostic is_file. The exact Python child was already killed; optional evidence lookup replaced the original timeout and downgraded the quarantine code.

_evidenced_path now catches BaseException while a primary error object already exists, records the safe metadata-failure category on that primary, and continues to fallback evidence. It does not propagate a secondary KeyboardInterrupt/SystemExit from optional metadata lookup. _persist_worker_diagnostic already catches BaseException and required no change. No changes were made to ordinary cancellation outside primary-error handling.

Eight new tests cover primary timeout/cancellation, secondary KeyboardInterrupt/SystemExit, and diagnostic/quarantine path metadata. All eight failed before the code change. After the change, each preserves the original cancellation identity or timeout type/code, records the exact safe secondary category, retains the child termination/lock release, keeps the correct quarantine code, and rejects the next job before launch. A quarantine marker whose metadata cannot be verified is not advertised as an evidenced path.

Evidence:

- windows-word-secondary-cancel-red.log: 8 failed, 21 deselected.
- windows-word-secondary-cancel-green.log: 49 passed (29 dedicated plus 20 existing Word runtime tests).
- git diff --check passes.
- This wave only: windows-word-secondary-cancel-fix.patch.
- Before snapshots: windows-word-runtime-before-secondary-cancel-fix.py and windows-word-runtime-tests-before-secondary-cancel-fix.py.

Command:

```sh
PYTHONPATH=. /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/msoffice/test_windows_word_runtime_diagnostics.py tests/msoffice/test_windows_runtime.py -q --tb=short
```

Frozen SHA-256:

- `06c37541dbac56e118ccbb5c2883193f52b2d5d979a1652b1272fbf248509c45`  `skills/WPSComposer/scripts/msoffice/windows_runtime.py`
- `b81fccbee788e9a42f6dca8f05c9fa59510a9f14f5c0c928ad9c1863e09dd758`  `tests/msoffice/test_windows_word_runtime_diagnostics.py`

## Independent acceptance — checkpoint_integration_review

The runtime `06c375...` and tests `b81fcc...` hashes above were independently matched. The narrowed helper change correctly preserves an already-selected primary exception during optional diagnostic or quarantine metadata lookup, appending safe failure categories instead of promoting the secondary cancellation. All eight new cases cover the original reported hole, including successful durable quarantine with temporarily unobservable metadata and next-job blocking.

Independent run of both Word runtime modules: **49 passed in 6.60 seconds** using the provided clean Python, `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.`, no pytest cache, and an owned scratch basetemp. The prior Word primary-identity/classification P2 is closed. No additional actionable P1/P2 was found in this bounded Word diagnostic/cleanup patch. These exact two hashes may be frozen for full33 integration; full-suite, Windows CI, and native acceptance remain separate root-owned gates.

No production/test/index/native state was changed by this reviewer. The separate Excel/PowerPoint shared runtime log-close finding is tracked in `windows-generation-diagnostics-report.md` and does not invalidate this Word-specific acceptance.
