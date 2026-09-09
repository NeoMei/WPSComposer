# Windows checkout filename portability fix

Prepared 2026-09-09. Scope: four historical image filenames, their relocation record, and portable repository-path regression tests. No production/native code edits, Office operations, commits, pushes, or remote runs were performed by this task.

## Root cause and retained evidence

Portable CI run `34332171915`, source commit `34a0658e38c6d5a73a7066df7f49f73b53aceb39`, failed checkout in both Windows matrix jobs. The first rejected path was `docs/verification/microsoft-parity/macos-word-advanced/integrated-02/runtime/wpsc-rsrc:K9qCW0iohe-tx8yQ.png`. Logs at `docs/verification/microsoft-parity/portable-ci/run-01/windows-python{312,39}-checkout.log` remain unchanged.

The Git-index regression reproduced exactly four invalid paths, all colon-containing retained PNG names. No other invalid names or case collisions appeared. This agrees with [Microsoft's filename rules](https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file). References to the two tokens were found in the original compiled/runtime AppleScripts and the CI failure logs, not in active tests or production references to these retained files.

## Changes

- Four exact Git `R100` renames, replacing only `wpsc-rsrc:` with `wpsc-rsrc-` for tokens `K9qCW0iohe-tx8yQ` and `QNP5Ulby7mHPtP88`, in `integrated-02/runtime/` and `integrated-03/runtime/`.
- New `docs/verification/microsoft-parity/macos-word-advanced/windows-path-relocations.json` records original Git commit, original repository path, original absolute runtime path, stored repository path, SHA-256, size, and unchanged historical script references for all four images.
- New `WINDOWS-PATH-RELOCATIONS.md` in that directory explains why paths changed and how to reconstruct the original references from the map or Git blobs.
- New `tests/test_repository_checkout_paths.py` audits `git ls-files --cached -z`; it ignores mutable untracked output and evaluates staged renames. It checks each directory/file component for Windows-reserved characters/control characters, trailing space/dot, reserved device names including superscript COM/LPT digits, and casefold/file-directory prefix collisions. Unicode and legitimate similarly named files are covered.
- Relocation regression compares all four retained images with their original Git blobs, verifies SHA-256/size, and compares listed historical scripts byte-for-byte with their original Git blobs while confirming the recorded runtime path appears in them.

## RED → GREEN

1. Before any rename or manifest, `pytest tests/test_repository_checkout_paths.py -q` produced **2 failed, 28 passed**. One failure listed exactly the four invalid tracked paths; the second reported the missing relocation manifest. Retained output: `windows-checkout-red.txt` beside this report.
2. Staged only the four authorized `git mv` operations. Git identified every rename as `R100`; no other index changes were made by this task. Parent was notified immediately afterward and resumed index ownership.
3. Added the map/explanation and ran the same focused command: **30 passed in 0.30s**. Retained output: `windows-checkout-green.txt`.
4. Captured and rechecked SHA-256 for all **240** previously tracked files under `macos-word-advanced`, resolving only the four renamed paths through the map. Every hash matched. Original snapshot: `windows-checkout-original-hashes.json` beside this report. No historical script, log, report, or other evidence bytes changed.
5. New text-file whitespace checks and `git diff --cached --check` passed.

Image content hashes (both run directories contain the same respective image bytes):

| Token | SHA-256 |
| --- | --- |
| K9qCW0iohe-tx8yQ | `fafd85981d01cebf83bb414b164f31fe2fa7daddddbcf0c70d178afdd0c45452` |
| QNP5Ulby7mHPtP88 | `712da9de59a1cb8252556ae78d53e41250f8a4ac59b11a5133e6a0e6c8891217` |

## Handoff boundary

The four renames are staged. The new test, JSON manifest, and explanation were left unstaged for parent review. This ignored report and its RED/GREEN/hash evidence are local coordination outputs. Full-suite testing and a new Windows hosted run remain parent-owned gates; the passing local audit does not itself prove remote checkout or native Office acceptance. Full Git history is still required, now also to compare relocated evidence with the original commit recorded above.
