# Independent qs dependency candidate review

Date: 2026-09-09. Reviewer: independent candidate-review agent.

## Verdict

Scoped pass for the dependency-only candidate. No source-backed blocking issue in the five candidate files was identified. Both installed consumers resolve qs 6.16.0, the original strict-limit reproducer rejects, the isBuffer round-trip reproducer succeeds, and legitimate Express request parsing remains intact. Do not describe this as proof that every qs array-limit configuration is safe: the concrete non-throwing residual below survives upstream 6.16.0.

Reviewed independently without reading investigator reports or prior test-result logs. No production/test/dependency files were changed, no dependency installation was performed, and no native application was launched. This document is the only reviewer-created artifact.

## Scope and immutable candidate identity

All paths below are relative to `/Users/neomei/项目/codexprojects/WpsComposer/.worktrees/office-description`.

| File | SHA-256 |
| --- | --- |
| `macos/wps-jsapi-probe/package.json` | `4d605b579d068cdd84cbeec1006c695494a28b5412d25f6caf760926330ab3d8` |
| `macos/wps-jsapi-probe/package-lock.json` | `4a0b46e5b7818906573a58fed894bdd3eb3b4d8aa915801cdeabb4d0351e4f70` |
| `tests/macos_probe/qs_dependency_regressions.cjs` | `508de54b044f593a57e4094dd1e9c441f0eee596871b41de6d7dc28011e0a695` |
| `tests/macos_probe/test_qs_dependency.py` | `49815ab614ed336edd0c5f492a63cc6260d9fba4679f58dbada3b01a3328f7a1` |
| `tests/macos_probe/test_runtime.py` | `f0d90c943dd351d2c9128f1014f594a5c031d21f8d063ff943f1c03cec6f41c0` |

The package change adds an exact qs override while preserving tmp 0.2.7. The lock change contains only qs version/resolved/integrity replacement. The existing runtime assertion at line 286 is correctly synchronized to the expanded exact override set. No public runtime API or native Office behavior is changed.

## Boundary reconstruction

The installed chain is `wpsjs 2.2.3 -> express 4.22.2 -> qs 6.16.0`, with `body-parser 1.20.6 -> qs 6.16.0` deduplicated. Both consumer-relative `createRequire` resolutions independently returned the same `node_modules/qs/lib/index.js` and package version 6.16.0. `npm ls qs --all --prefix macos/wps-jsapi-probe` exited successfully.

Ordinary profile serving uses Python `ProfileServer` (`skills/WPSComposer/scripts/macos_probe/runtime.py:947`), whose request path uses `urllib.parse.urlsplit` and `unquote` (`profile_server.py:9,74`), and binds IPv4 loopback (`profile_server.py:168`). The bridge separately uses Python `parse_qs` (`bridge.py:12,272`). These are not qs paths.

Optional upstream wpsjs debug tooling creates Express servers (`node_modules/wpsjs/src/index.js:34-45`, `src/lib/debug_publish.js:58-113`, and `src/lib/debug_xmlplugin.js:235-246`). This can reach Express query parsing; its `app.listen(serverPort)` call does not explicitly restrict the bind address. Express's extended parser invokes qs with `allowPrototypes: true, arrayLimit: 1000` (`node_modules/express/lib/utils.js:288-292`); body-parser's extended parser supplies limits and `allowPrototypes`, but no comma option (`node_modules/body-parser/lib/types/urlencoded.js:133-176`). qs defaults `comma` to false (`node_modules/qs/lib/parse.js:16`). No `comma: true` setup or `qs.stringify` sink was found in the inspected wpsjs source/add-in surface. Thus installed dependency exposure is established; a complete product exploit path for either advisory is not established.

## Advisory checks and residual boundary

[GHSA-x5fp-wj9c-mxmx](https://github.com/advisories/GHSA-x5fp-wj9c-mxmx) names 6.16.0 as patched. In the current installed source, `parseArrayValue` counts commas before splitting when `throwOnLimitExceeded` is enabled (`node_modules/qs/lib/parse.js:39-53`). The candidate tests reject ordinary, bracket, and encoded-bracket inputs with `arrayLimit: 3`. Independent additional nested, indexed, and lowercase-encoded variants also raised the intended RangeError.

Concrete surviving upstream behavior, reproduced with installed 6.16.0:

```js
qs.parse('a[]=1,2,3,4', { comma: true, arrayLimit: 3 })
// { a: [['1', '2', '3', '4']] }
qs.parse('a=1,2,3,4', { comma: true, arrayLimit: 3 })
// { a: { '0': '1', '1': '2', '2': '3', '3': '4' } }
```

Encoded brackets produce the same oversized nested array. The comma check is conditional on throwing, and bracket wrapping at `parse.js:141-142` precedes the subsequent top-level length check at 145-146. This is a directly observed limitation of the upstream package; the candidate does not introduce it, and the inspected product path does not enable comma parsing. Consequently the strict throwing scenario is verified closed, while broad closure of every scenario described by the advisory cannot be claimed.

[GHSA-4mjr-xmp4-gh2g](https://github.com/advisories/GHSA-4mjr-xmp4-gh2g) also names 6.16.0 as patched. `node_modules/qs/lib/utils.js:327-332` now checks that `constructor.isBuffer` is a function before calling it. The original parse-to-stringify payload succeeds for both `plainObjects` and `allowPrototypes`. Independent array-valued, object-valued, and root-constructor variants also serialized without exceptions. A custom executable getter/function would require a different input boundary than the query-string-only advisory. No such surviving query-string bypass was found.

## Tests, compatibility, errors, and portability

On Node v24.18.0, `node --test tests/macos_probe/qs_dependency_regressions.cjs` passed all 5 tests. The Python wrapper was independently executed through `python3 -B` and `runpy` on Python 3.9.6 and passed. A worktree-local `.venv/bin/python` was absent, so no claim is made that pytest itself ran during this review.

The security assertions check RangeError rather than merely any exception and assert the exact round-trip string. Consumer-relative resolution prevents a stale nested qs copy from being hidden by a root-only import. Compatibility checks cover a limit-equal comma array, encoded literal commas, Unicode, nested objects, indexed serialization, real Buffers, and an actual loopback Express POST exercising both query and extended form parsers.

The HTTP check asserts the complete status/body result; normal assertion failures close the server in `finally`, request errors reject, and the Python subprocess has a 30-second bound and captures both output streams. Unexpected listen/response/JSON failures become failing Node execution rather than a success. The test uses a random loopback port and does not write user-home data or use native WPS/Office. Paths are resolved from the test file and passed as argv, avoiding shell quoting and working-directory dependencies. Existing CI provisions Node 22 and npm ci (`.github/workflows/portable-tests.yml:59-69`), and existing add-in tests already require node; the new wrapper does not introduce a new platform tool prerequisite into that configured suite. Native Windows and Linux execution were not performed by this reviewer.

This review verifies the candidate and the stated runtime-reachability classification only. It does not certify unrelated staged work, release/install state, native Office behavior, or a full dependency audit.
