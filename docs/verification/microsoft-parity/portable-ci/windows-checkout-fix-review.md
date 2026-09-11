# Windows checkout fix scoped review

Decision: **PASS**. No actionable findings in the four image renames, relocation metadata/explanation, or new repository-path tests.

## Scope and root cause

Reviewed the four staged `R100` renames under `macos-word-advanced/integrated-{02,03}/runtime`, `tests/test_repository_checkout_paths.py`, `windows-path-relocations.json`, `WINDOWS-PATH-RELOCATIONS.md`, the worker report, and retained checkout failures from CI run `34332171915`. Both Windows Python 3.9 and 3.12 logs identify the same first colon-containing PNG path as invalid. Linux host-assumption test failures and concurrent Word changes are outside scope.

Reviewed new text file SHA-256 values:

- `tests/test_repository_checkout_paths.py`: `4f26641c514d2ceb5a2badb0c993be3e79086a6fabfe9d9d17dbdefc698cab93`
- `docs/verification/microsoft-parity/macos-word-advanced/windows-path-relocations.json`: `337f0d085c7f744f929e4d9e017d04c0179a6944b8850ba92f1d3203b4a45be5`
- `docs/verification/microsoft-parity/macos-word-advanced/WINDOWS-PATH-RELOCATIONS.md`: `ba72170bce518fc31f7ba50732d0b3e76fd4042fb41250da1c4ba6e79925181c`

## Assessment and independent evidence

- The index contains exactly four image `R100` renames for the requested tokens, replacing the colon with a hyphen. No image content changes accompany those staged renames.
- Independently enumerated all 240 original evidence files directly from Git commit `34a0658e38c6d5a73a7066df7f49f73b53aceb39` and compared each original blob byte-for-byte with its current file, applying only the four declared relocations: **240/240 unchanged**. Independently checked all four staged image blobs against the manifest SHA-256 values: all match.
- The map binds original Git commit, original repository path, original absolute runtime path, current stored path, byte length, hash, and unchanged historical references. The explanation correctly preserves old runtime references as historical evidence and directs archive readers to the map or Git blobs. Searches found no active production or preexisting test references to the renamed retained paths requiring an update.
- The namespace check reads the actual Git index with `git ls-files --cached -z`, so it checks shipped/staged names and does not accidentally collect unrelated untracked native output. Component validation covers the actual colon failure, other reserved characters, control characters, trailing dots/spaces, reserved device names, and file/directory or case collisions. Unicode and legitimate similar names have positive coverage.
- Git path capture remains bytes until explicit UTF-8 decoding of NUL-separated output; it does not depend on Windows locale decoding or quoted Git display paths. Original image/script retrieval uses argument-list Git object lookups and binary stdout, so it neither creates the original invalid filenames nor translates blob line endings. Existing CI disables autocrlf before checkout to preserve historical text evidence bytes.
- Relocation tests require all four exact original entries and distinct valid tracked destinations, compare relocated bytes directly to original Git blobs, verify hashes/lengths, and prove the listed historical scripts remain unchanged and contain the recorded runtime reference. This is substantive integrity validation rather than a weakened expectation.
- Independently ran `env -u PYTHONPATH /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest tests/test_repository_checkout_paths.py -q`: **30 passed in 0.31s**, exit 0. The worker retains the prior two failures in the RED log.

## Boundary

The scoped fix is ready for parent integration. These local checks establish the targeted invalid-name repair and preserved evidence; a new Windows hosted run must confirm checkout and subsequent portable tests. This does not establish native Office acceptance. Only this review report was written; no source/evidence edits, staging, installation, native application operations, commits, or pushes were performed by the reviewer.
