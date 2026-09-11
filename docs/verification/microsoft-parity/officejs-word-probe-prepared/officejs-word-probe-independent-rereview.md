# Independent P3 re-review: Office.js probe method key

**PASS. The original P3 is closed; no new scoped findings.** This remains offline preparation only, not permission to install or native/parity acceptance.

All 19 HASHES.json entries independently matched. Supplied freeze values matched exactly:

- probe.js: `8da68b4f4e2420e66262e3b56086a71db3251c3269d54d85ccd7fd8ebdf0f046`
- tests/probe.test.cjs: `0a0c4063b4ea0c55435051124ac022fd78fddd4fe796f247f8fdbc999d0ef494`
- HASHES.json: `53bd0cf03da7a9dd885d69b885f17ef34c60e5441fd62569b371b26f69af0c5c`

Reversing only the corrected `add_equation_native_fallback` string in memory reproduces the original reviewed probe.js SHA-256 `411cafcaa149f46f81457c49edbeb828d73d2c88390d49cf9271e82e70c55734`. This confirms there is no other core behavior change. UI, HTML, README, configuration and generated manifest hashes also match the prior review.

The new test compares the complete sorted nine-name array with the frozen contract, so it detects misspellings, omissions, substitutions and duplicates rather than merely counting entries. The corrected fallback name is the actual Writer method. All existing nativeAccepted=false/snapshot-not-executed, permission, pane identity, timeout and no-mutation behavior remains unchanged.

Independent reruns:

- `node --test tests/*.test.cjs`: **21 passed, 0 failed**, 72.81 ms; log `officejs-word-probe-independent-rereview-js.log`.
- `python3 -B -m unittest discover -s tests -p 'test_*.py'`: **6 passed**; log `officejs-word-probe-independent-rereview-python.log`.

No package/source/index edits or browser/server/native/install actions were performed. The original review and failure evidence remain retained separately. The package may remain frozen as the reviewed offline probe candidate; actual enablement/removal, host execution, preservation and document identity gates are still open.
