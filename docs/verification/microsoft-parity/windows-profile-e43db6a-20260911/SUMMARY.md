# Windows Python 3.14.3 profile-server verification

**PASS for this targeted unit suite: 22 passed in 35.81s, exit 0.** This result does not certify Office/WPS native parity or release readiness.

Exact tested commit: `e43db6a391bbf8a7f78e6d5f2da46c4e68b4ffea`. Independent checkout: `D:/wpsce43`, branch `codex/windows-profile-e43db6a`. Original checkout stayed at `f7e3ed72289c52b2aa8fd6c36129bf6f11f8f501`; both status outputs were empty before/after.

Command, executed with cwd `D:/wpsce43`:

```powershell
D:/wpsc168-env/Scripts/python.exe -m pytest tests/macos_probe/test_profile_server.py -q
```

Actual output:

```text
......................                                                   [100%]
22 passed in 35.81s
```

Started 2026-09-11 08:46:34 +08:00, exited 08:47:11 +08:00. Python 3.14.3. Runtime import verified before and after at `D:/wpsce43/skills/WPSComposer/scripts/macos_probe/profile_server.py`; `_ProfileRequestHandler.timeout == 5.0`.

- Profile server source SHA-256: `1d6b8ae437034a7ff9c22b2e5d10007186886b5db0331ccbedbccbbf33288402`.
- Test source SHA-256: `e356fefa9c8d1bc44496bc52978f88c613fd1fad3cf93423951f4edc85efe028`.
- All 500 materialized tracked files match their exact candidate Git blob IDs and remained unchanged. Historical verification artifacts were not materialized; complete historical object closure is not claimed.

The run includes both new real TCP parameter cases that leave a client open while simulating ineffective connection shutdown. They completed in this Windows process; this is actual execution, not the Mac or hosted CI result reported by another task.

## Preservation and retained errors

All three quarantine marker byte hashes and the Office/WPS process inventory remained identical before/after. No Office API was used by the provenance helper: it imports only the profile module, reads source/marker bytes, and uses a read-only process query. No Office/WPS application was started or closed; no document was saved/closed; no sentinel/native runner was executed. No npm dependencies or plugin were installed. Full pytest remains unrun; its future prerequisite is the locked template installation documented in windows-candidate.md.

The first provenance helper mistakenly referenced `ProfileRequestHandler` instead of `_ProfileRequestHandler` and exited 1 before producing before.json. Its original source and stdout/stderr/command are retained. A separate provenance-v2.py corrected only the helper attribute; its run exited 0. Candidate source was not edited.

Bootstrap exited 0 but emitted a transient `gitattributesMissing` diagnostic while importing the tree before selected blobs. The full stderr is retained. Final per-file Git object hash checks passed before and after the test; this does not erase the bootstrap diagnostic or claim full Git fsck closure.

## Earlier bounded read-only recovery diagnosis

The preceding 08:28 +08:00 diagnosis is included separately under prior-readonly-20260911; it is not a new native test. Excel PID19360 still had the same unsaved workbook and matching content hash at that observation. Original Word PID13308 and its two live objects were absent, so their post-close content/saved state could not be certified. Old sentinel PIDs8004 and7436 had been reused by Node processes. Word table ownership comparisons, Excel exact pre/post Move counts and PowerPoint candidate-window/PID data were absent from old logs and require a later dedicated native probe. No raw document body is published.

The current evidence commit is an independent evidence-only child of e43db6a; it does not merge or change the development branch and does not publish a release. The later bd988445 Organizer/CI document commit was not substituted for the requested code commit.
