# Development setup scoped review

Decision: **PASS**. No actionable findings in the requested scope.

## Scope

Reviewed the current diff against `53ad017b4c701ba212662bd3043e907aa664a6e2` for `AGENTS.md`, `README.md`, `pyproject.toml`, and `tests/test_documentation.py`, plus `docs/verification/microsoft-parity/development-setup-audit.md` including its final round28 result. Uncommitted heading, numbering, recovery, and quality feature work is outside this review.

## Findings and validation

- The dev extra now declares PyMuPDF, which the retained native-artifact PDF drawing helper imports as `fitz`. The independent fresh Python 3.9.6 environment installed PyMuPDF 1.26.5 with `Requires-Python: >=3.9`; the corrected install log records a successful editable `.[dev]` install and `pip check` with no supplemental dependency directory.
- README and AGENTS include the necessary pip upgrade before editable installation, explicit availability of the frozen Git capability baseline, and the locked npm template preparation command. The template tests consume resources from `macos/wps-jsapi-probe/node_modules/wpsjs`; no personal marketplace installation is required for that preparation.
- The documentation test now parses TOML and verifies the required dependency subset rather than rejecting any additional dev dependency. Its Python 3.9/3.10 fallback is supported by pytest: installed pytest 8.4.2 declares `tomli>=1; python_version < "3.11"`. The retained round27 failure supplies the regression evidence for the formerly hardcoded list.
- Independently ran `env -u PYTHONPATH ../clean-dev-venv/bin/python -m pytest tests/test_documentation.py -q` from `/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/checkout`, after round28 completed: **9 passed in 0.03s**, exit 0.
- Inspected round28 output and export manifest: **4426 passed, 12 skipped in 239.86s**, exit 0, 342 recorded source/configuration/test files unchanged. Independently recomputed the hashes of all four reviewed overlay files in the current worktree; each matches the manifest's tested overlay. The full suite was not rerun by this reviewer.
- The audit correctly separates the delivered checkpoint and development setup result from later mutable feature work and native acceptance. It retains earlier failure evidence and identifies the twelve skips without presenting them as native passes. Commit 53ad017 records the 286 formerly ignored native artifacts.

Only this review report was written. No source edits, native application actions, installation, or commits were performed by this reviewer.
