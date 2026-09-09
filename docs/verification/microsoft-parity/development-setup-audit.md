# Reproducible development setup audit

The delivered checkpoint is commit `53ad017` (tree `3b58f66dc5651cd4a09350ead3e4561ef55bbffd`). The audit uses a Git-exported checkout, separate from the mutable implementation worktree.

- Git ignored native artifacts: 286 Office/PDF/PNG outputs were absent from the earlier evidence commit, including the two inputs read by the retained paragraph-rule regression. They are now tracked in `53ad017`.
- Archive-only round25 retained 52 failures: 45 tests required locked WPS JSAPI npm template resources and seven required the frozen capability baseline from Git history. Preparing those exact resources, without test skips or assertion changes, yielded round26: **4426 passed, 12 skipped**, all 340 recorded files unchanged.
- The documented editable dev install initially failed in a fresh Python 3.9 venv because the bundled pip 21.2.4 lacked pyproject editable-install support. README and AGENTS now explicitly upgrade pip. `PyMuPDF>=1.24` is declared in the dev extra because native-artifact regression helpers import fitz.
- Corrected installation ran only `pip install --upgrade pip`, `pip install -e '.[dev]'`, and `pip check` inside the separate venv; all passed. There was no personal plugin installation or supplemental PYTHONPATH dependency directory.
- Round27 used that clean venv with PYTHONPATH unset. It passed 4425 tests and failed one documentation test that hardcoded the old two-element dev dependency list. The test now parses TOML and requires all three necessary dependencies while allowing additional dev requirements and formatting changes. The original failure remains in `unit-round27.txt`.

Round28 passes **4426 tests / 12 skips** in 239.86 seconds with all 342 recorded source/configuration/test hashes unchanged. Its overlay consists only of pyproject.toml, README.md, AGENTS.md, and tests/test_documentation.py over commit 53ad017. It uses the fresh venv and no supplemental PYTHONPATH. `unit-round28-export.json` records exact hashes and the environment.

The 12 skips are six passive WPS bridge registration checks and six explicitly gated native-platform tests. They do not establish native acceptance. This audit validates development setup and the recorded checkpoint, not later Word feature work or overall Microsoft parity.
