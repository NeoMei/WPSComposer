# Historical image path relocation

The first portable Windows CI run (`34332171915`, source commit `34a0658e38c6d5a73a7066df7f49f73b53aceb39`) failed during Git checkout: four retained runtime PNG names contained `:`, which Windows rejects in filenames. The original failure logs remain under `../portable-ci/run-01/`.

Only these four stored image names changed from `wpsc-rsrc:<token>.png` to `wpsc-rsrc-<token>.png`, in `integrated-02/runtime/` and `integrated-03/runtime/`. Their bytes are unchanged. Original AppleScripts, logs, reports, and other historical evidence remain unchanged, including references to the runtime names used during the original runs.

[windows-path-relocations.json](windows-path-relocations.json) records each original repository path and absolute runtime path, its current stored repository path, SHA-256, byte length, and unchanged source references. All repository paths in the map are relative to the repository root. Resolve an old path through this map when inspecting the archive; the original blobs also remain accessible using `git show '<original_git_commit>:<original_repository_path>'` without creating a Windows-invalid file.

`tests/test_repository_checkout_paths.py` audits tracked Git-index names on every platform and verifies relocated images and their source references against the original Git blobs. A passing local audit establishes filename portability and evidence integrity; the next hosted Windows CI run must separately confirm checkout and subsequent test execution.
