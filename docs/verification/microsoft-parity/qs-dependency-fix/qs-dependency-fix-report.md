# qs dependency remediation

Outcome: **dependency upgrade verified; scoped candidate PASS**. The isBuffer primitive and strict-throw array-limit regression are fixed in the tested configuration. Broad closure of the array-limit advisory remains inconclusive because the upstream non-throwing mode still exhibits overflow. Application exploitability was not demonstrated. No native Office/WPS acceptance is claimed.

## Scope and version delta

Parent authorized the necessary expansion from lock-only to `package.json` and the matching existing exact override assertion. `package.json` adds `overrides.qs = "6.16.0"` alongside the unchanged `tmp = "0.2.7"`. npm generated the lock; its only changed package entry is `node_modules/qs`, with version, tarball URL and integrity updated.

| Package | Before | After |
| --- | --- | --- |
| qs | 6.15.3 | 6.16.0 |
| express | 4.22.2 | 4.22.2 |
| body-parser | 1.20.6 | 1.20.6 |
| wpsjs | 2.2.3 | 2.2.3 |
| wps-jsapi-declare | 2.2.0 | 2.2.0 |

Express and body-parser both declare `qs ~6.15.1`. On this verification date the registry exposes no newer Express 4 release accepting 6.16.0. Body-parser 1.20.8 alone would leave the Express edge on the affected line. The precise override is the smallest npm-supported repair; no Express major upgrade or unrelated dependency churn occurs. `npm ls` resolves 6.16.0 through both consumers without an invalid-dependency result. Compatibility beyond the focused checks below is not inferred from semver alone.

## Findings and reachability

Official advisories checked live on 2026-09-09:

- [GHSA-x5fp-wj9c-mxmx](https://github.com/advisories/GHSA-x5fp-wj9c-mxmx): affected 6.14.2 through 6.15.3, fixed 6.16.0. With explicit comma mode, bracket arrays can exceed configured array limits. A four-element payload with limit three demonstrates the defect without allocating large arrays.
- [GHSA-4mjr-xmp4-gh2g](https://github.com/advisories/GHSA-4mjr-xmp4-gh2g): affected >=2.2.5 and <6.16.0, fixed 6.16.0. Parsed attacker-controlled `constructor.isBuffer` data can cause a TypeError when passed to `qs.stringify`. Process termination depends on the application's error handling.

The tested invariant is rejection of comma arrays beyond the configured limit when `throwOnLimitExceeded: true`, and safe serialization of parser-produced plain data. Existing nested query/form fields, repeated array keys, Unicode, literal commas and Buffer values must retain their behavior.

Repository source evidence:

- `skills/WPSComposer/scripts/macos_probe/runtime.py:945-951` starts Python `ProfileServer`, which binds 127.0.0.1 (`profile_server.py:169`). `bridge.py:12,272` uses Python `urllib.parse.parse_qs`, and `bridge.py:434-436` uses Python `ThreadingHTTPServer` on loopback. Thus the normal runtime request path does not call npm qs.
- Wpsjs supplies official templates (`templates.py:125-137`, `runtime.py:1022-1029`). The existing `test_start_servers_hosts_only_selected_profile_without_wpsjs_debug` rejects Node/wpsjs resolution during profile hosting (`tests/macos_probe/test_runtime.py:777-802`).
- Optional, manually invoked wpsjs debugging creates Express servers (`node_modules/wpsjs/src/lib/debug_xmlplugin.js:16,246`, `debug_publish.js:60,113`). Express uses extended query parsing (`express/lib/application.js:84,151`) and `qs.parse` with `allowPrototypes: true` (`express/lib/utils.js:288-292`). Body-parser's extended form parser also invokes qs (`body-parser/lib/types/urlencoded.js:169-175`). Neither enables comma mode. No qs stringify sink was found in those consumers or project runtime code. Their parsing can preserve the special constructor shape, but an application serialization sink is not established.

Classification: confirmed vulnerable installed primitive; these two specific exploit paths are not shown to be reachable in WPSComposer's production runtime. npm's initial three moderate entries are qs plus two dependency-propagated entries (Express and body-parser), not three independently demonstrated application vulnerabilities.

## Changes and validation

Changed production/configuration files: `macos/wps-jsapi-probe/package.json`, `macos/wps-jsapi-probe/package-lock.json`. Updated the exact reviewed override contract at `tests/macos_probe/test_runtime.py:286`; added `qs_dependency_regressions.cjs` and its pytest wrapper in the same test directory.

Validation on Node v24.18.0 / npm 11.16.0:

1. Syntax/resolution: generated lock changes reviewed; `git diff --check` passed; clean `npm ci --prefix macos/wps-jsapi-probe --ignore-scripts` passed before and after. `npm ls --prefix macos/wps-jsapi-probe qs express body-parser --json` passed and confirms both consumer paths use the fixed version.
2. Security regression: the final identical Node test file was run against an isolated baseline installation, yielding **4 failed / 1 passed** (each consumer fails the two advisory checks; HTTP control passes). Against the fixed installation, **5 passed**. It covers literal and encoded bracket keys, both prototype-preserving parse modes, valid at-limit input, encoded literal commas, Unicode serialization and Buffer serialization.
3. Compatibility: the Node suite makes a real HTTP request to an isolated ephemeral loopback Express server and checks query and urlencoded form results against literal expectations. It does not invoke the wpsjs CLI or native WPS.
4. Existing tests: the full `tests/macos_probe` run yielded **468 passed / 1 failed** in 67.21 s, with the only failure being the outdated exact overrides assertion. After synchronizing that assertion, `pytest -q tests/macos_probe/test_runtime.py tests/macos_probe/test_qs_dependency.py` yielded **61 passed** in 0.81 s. These are separate runs, not a claimed single all-green full-suite run. Root owns whole-repository testing in a separate immutable checkout.
5. Registry audit: `npm audit --prefix macos/wps-jsapi-probe --json` changes from **3 moderate** to **0 reported vulnerabilities**. This is a dated registry result, not proof that the application has no other vulnerabilities.

Raw outputs, command results, version delta and candidate SHA-256 manifest are retained under `docs/verification/microsoft-parity/qs-dependency-fix/`. No staging, commit, push, native application operation or personal plugin installation was performed by this worker.

## Preserved exploratory failures

The first candidate test incorrectly expected percent-encoded commas to split. Source inspection showed splitting precedes decoding; `%2C` is literal data. The expectation was corrected to preserve that legitimate behavior, while the alternate attack uses encoded brackets with literal commas. Both the failed expectation log and the final identical baseline/fixed test logs are retained.

An initial isolated-baseline harness invoked npm with an absolute external prefix from the worktree and failed package/lock validation. It produced no valid security RED evidence. Rerunning `npm ci --ignore-scripts` with the isolated package directory as cwd succeeded; only the latter successful installation and four expected assertion failures are cited as final RED evidence. Both harness failure logs are retained.

## Upstream non-throwing-mode limitation

The independent reviewer found, and this worker reproduced, that qs 6.16.0 still returns a nested four-element array for `a[]=1,2,3,4` with `{comma: true, arrayLimit: 3}` when `throwOnLimitExceeded` is omitted. The strict throwing configuration rejects the payload. The registry lists 6.16.0 as patched, but the bounded regression evidence proves only the strict throwing behavior; complete array-limit enforcement in every mode is **not** claimed. `upstream-nonthrowing-limit-residual.log` preserves this observation. No project or inspected wpsjs/Express/body-parser consumer enables comma mode, so it does not establish an application-exploitable path or require altering production parsing here. Follow-up upstream behavior work is outside this dependency-only change.

Independent pre-patch investigation agreed with the reachability and dependency-constraint analysis. Final independent candidate review is recorded separately in `qs-dependency-fix-review.md`.
