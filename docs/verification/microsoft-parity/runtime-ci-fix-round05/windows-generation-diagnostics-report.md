# Windows generation diagnostic persistence repair

Scope: `_execute` in the Excel/PowerPoint one-shot Windows runtime. A retained regression reproduces recovery JSON write failure skipping persistent quarantine and replacing native timeout/cancellation. A separate original cleanup failure also replaces the primary native error. RED: 3 failed. GREEN: 22 passed across the new dedicated module and existing Windows Office runtime tests.

The repair attempts uncertainty quarantine before fallible recovery output, preserves the original exception and native code, records safe I/O failure categories, and only assigns its new recovery/quarantine paths after successful writes. Existing verified-closed failures do not quarantine. Uncertain staged jobs remain retained. Lock/eligible staging cleanup still runs; secondary cleanup faults preserve a primary error, while a success-path cleanup failure continues to propagate. If quarantine itself cannot be persisted, its path is not claimed and cross-process isolation is not guaranteed.

This patch has no native execution evidence and is not yet independently reviewed. `_run_worker` and the separate Word runtime were not modified; their inner diagnostic/cleanup handling remains a separate review scope. No public API changes, Office termination, shared index changes or release are included in this patch.

## Frozen source SHA-256

- `4857f4989cc2bb336b926b984046d8937b6f9239f68aa9937a20c8028e8d26f2` `skills/WPSComposer/scripts/msoffice/windows_office_runtime.py`
- `6a31a612c851bcc195bd5622dd77b77b9c5657ea00a5db6cee49f34849daad94` `tests/msoffice/test_windows_office_runtime_diagnostics.py`

## Independent review — 2026-09-09

**The `_execute` narrow repair passes independent review, with a separate open P2 in the unmodified worker failure boundary.** This supersedes the earlier statement that the narrow patch is not yet independently reviewed. It does not claim the complete generation failure path is clear of remaining findings.

Independent repository tests: `test_windows_office_runtime.py` plus `test_windows_office_runtime_diagnostics.py`, **22 passed in 0.79s**. Additional independent probes retained in `test_windows_generation_review_edges.py`: **3 passed, 3 failed in 0.07s**. The three passing edges cover simultaneous quarantine/recovery write failure for typed timeout and cancellation, retained uncertain job directories, absent unpublished output, truthful newly assigned paths, and continued propagation of staging-removal failure after successful publication. The original six dedicated regressions also confirm next-job blocking after successful quarantine, no claimed quarantine after failed persistence, verified-close failure without quarantine, and primary-error preservation during lock cleanup. No new regression was found in the changed `_execute` logic.

### Separate P2: worker failure classification and diagnostic location depend on diagnostic writes

The remaining code in `_run_worker` and `_worker_failure` has a reproducible failure boundary:

1. On a Python-child timeout, `_run_worker` kills the exact child and then writes `diagnostics.json` before constructing `NATIVE_OFFICE_TIMEOUT`. An injected ENOSPC for only that diagnostic path replaces the timeout with `OSError`.
2. When `Popen` raises before any child exists, `_run_worker` writes its `cleanup_verified=True` classification to the diagnostic file and reconstructs the error from that file. A diagnostic write failure loses that known safe classification; the outer `_execute` then quarantines a component even though no child or Office operation started. The probe verifies that the quarantine marker is actually created.
3. An abruptly failed Python worker can produce no diagnostic file. `_worker_failure` nevertheless assigns that nonexistent `diagnostics.json` as `diagnostic_path`; `_execute` preserves it with its existing-path-precedence expression even when it successfully writes `recovery.json`. The returned error therefore points at missing evidence instead of the available recovery record. This is an existing inner/outer contract gap, not a new path created by this repair.

The three RED probes use fake `Popen`/child objects only; they do not launch subprocesses or native Office. They isolate the inner diagnostic path failure, leaving outer recovery and quarantine paths writable. Suggested separate repair: classify known launch failure/timeout in memory before optional writes, preserve the typed failure and cleanup knowledge when diagnostic persistence fails, and claim only evidenced diagnostic paths. Preserve uncertain-job quarantine and avoid claiming isolation when its persistence fails.

These are grouped as one P2 worker failure-reporting boundary; no P1 is claimed. Root was sent the probe paths and exact results before this report update. Reviewer changed only scratch tests and this report, with no production, index, or native Office action.


## Inner-worker repair wave

The three independent inner-worker failures were copied into permanent tests and reproduced (3 failed / 6 passed), then repaired. Worker failure classification and known prelaunch cleanup state no longer depend on writing diagnostics. Missing diagnostic JSON falls back only to an existing worker log; non-object diagnostic JSON is treated as unverified worker failure. Launch cancellation retains its original object and verified no-launch state.

Three further regressions reproduced cancellation leaving the exact Python child running and child cleanup replacing typed timeout (3 failed / 14 passed). The wait failure path now stops only its exact Python handle, independently attempts kill/wait, preserves timeout or original cancellation, and retains cleanup failure categories. Outer cleanup and diagnostic failures append to earlier categories. No Office process or process tree is terminated.

Combined repository and unchanged independent scratch validation: **39 passed** in 0.85s. The intermediate permanent-test missing Path import is retained in `windows-generation-worker-intermediate-nameerror.log`; it was fixed before the final run. No native acceptance is claimed. The separate Word generation runtime remains a distinct unmodified scope.

### Final candidate SHA-256 for re-review

- `9bf563fbc5be591b62558d6f3fff8827a6e45c652df0f65920dbcd76e7306b8a` `skills/WPSComposer/scripts/msoffice/windows_office_runtime.py`
- `76950136d2fb8d6eceaf85578ea7b23d4392654df0867b13b996fbbed526216d` `tests/msoffice/test_windows_office_runtime_diagnostics.py`

## Independent inner-worker re-review — 2026-09-09

Both final candidate hashes above were independently verified. The existing repository modules plus all six unchanged reviewer probes passed: **39 passed in 1.03s**. The original scratch hash remains `930dbadb0e9a464303ee4e33e1076445d45dcdc57dc228b5f5326ac299be962b`. The three prior worker-failure reproductions are corrected; in-memory launch classification, cancellation identity, typed timeout through child cleanup failures, diagnostic object validation, and append-only secondary failure categories were reviewed.

**A new P2 prevents approving this exact `9bf563...` candidate for full33:** evidence metadata lookup is fallible before exact-child timeout cleanup. The timeout branch constructs its error using `_worker_diagnostic_path(job)` before entering the `child.kill` / `child.wait` loop. That new helper calls `Path.is_file()` without handling `OSError`. A permission or filesystem metadata failure for either `diagnostics.json` or fallback `worker.log` therefore replaces the timeout and prevents the exact Python child from being stopped.

Independent `test_windows_generation_metadata_review.py`, retained beside this report, injects only the named evidence file's `Path.is_file` failure (`EACCES`) and uses a fake Python-child object. **2 failed in 0.08s**: both cases observed `(public_code=None, child.returncode=None)` instead of `('NATIVE_OFFICE_TIMEOUT', -9)`. This is an actual cleanup-order effect, not merely missing diagnostic annotation, and no real process or Office application was launched.

Suggested narrow correction: make evidence metadata inspection best effort and ensure no optional evidence lookup can prevent exact-child cleanup or replace the primary timeout/cancellation. The later `diagnostic.exists()` check is in the same optional-metadata boundary and should be handled consistently. No changes were made by this reviewer to production code, original regression assertions, or shared index. Root received the concrete reproductions before this report update.


## Metadata follow-up candidate

The two new independent metadata RED cases are now permanent regression tests. Evidence lookup handles inaccessible metadata without claiming the path, and typed timeout construction no longer performs filesystem inspection before exact-child kill/wait. The following diagnostic existence branch also uses the guarded evidence lookup. Combined validation: **43 passed** in 0.92s, including both unchanged independent scratch files. This candidate awaits re-review.

- `7bcfd0012d3d59ac73e2cb1e15b3e35040c90fc1c71db1ec0bb569e6bcb7e612` `skills/WPSComposer/scripts/msoffice/windows_office_runtime.py`
- `81cbe41fb93c927861938ab3920fb7bf467ede16b2ffc9e7646fd9f02d915d7c` `tests/msoffice/test_windows_office_runtime_diagnostics.py`

## Final independent acceptance — 2026-09-09

**PASS: the original inner-worker P2 and subsequent metadata-order P2 are closed. No open actionable P1/P2 remains in this bounded generation patch; the exact `7bcfd001...` runtime / `81cbe41...` tests candidate may be frozen into full33.** Earlier open-finding statements and RED results remain historical evidence and are superseded by this final acceptance.

The independent reviewer matched both final source/test hashes above and re-ran both repository modules plus both unchanged scratch files: **43 passed in 2.07s**. The original six edge probes remain SHA-256 `930dbadb0e9a464303ee4e33e1076445d45dcdc57dc228b5f5326ac299be962b`; the two metadata probes remain SHA-256 `a56db1d48a55ed6fe603011e32c3b0f61ed6ca6f8d8ca601f5b6f6d816ae643e`. No original assertion was weakened or edited.

Review confirmed that typed timeout construction performs no filesystem metadata lookup before independently attempting exact Python-child kill and bounded wait. Optional evidence discovery now handles `OSError` per candidate, and the later unguarded `diagnostic.exists()` branch has been removed. Existing diagnostic persistence failures, launch cancellation, wait cancellation, malformed metadata, next-job quarantine, preserved primary/cleanup categories, retained uncertainty, and success-path cleanup error propagation continue to pass.

The reviewer changed only this report during final re-review; no production source, tests, shared index, actual Python child, or native Office process was changed/launched. This permits full33 integration; actual full33 results, Windows CI, and native Office acceptance are still distinct gates owned by root. The separate Word generation runtime was not included in this review.


## Secondary diagnostic cancellation follow-up

The independent Word sibling review found a secondary KeyboardInterrupt/SystemExit during evidence stat can still replace an already caught timeout/cancellation. The same Office pattern was reproduced in four new permanent tests (4 failed / 19 passed). Evidence lookup now receives the primary error, catches secondary BaseException and appends safe diagnostic categories. All evidence paths remain optional and verified; original cancellation identity and timeout code survive. Combined final validation: **47 passed** in 1.86s, including unchanged original eight reviewer probes. Awaiting independent follow-up review.

- `07c26d23737c2410dab2f8dc25e83ba150c9ba33b1e74c64368346f830b26391` `skills/WPSComposer/scripts/msoffice/windows_office_runtime.py`
- `54c1bfec840ac70e7135e229bd2a778a6ccc6dc723e643109a41dbb0c5bbf2da` `tests/msoffice/test_windows_office_runtime_diagnostics.py`

## Independent follow-up — checkpoint_integration_review

The `07c26d...` runtime and `54c1bf...` test hashes were independently matched. The secondary evidence-stat cancellation correction is sound: lookup handles BaseException only after an existing primary failure, retains its identity/classification, appends safe diagnostic categories, and does not prevent exact-child cleanup. Both repository suites plus the original eight unchanged independent probes passed: **47 passed**.

**A separate reproducible P2 remains in the complete cleanup boundary:** `_run_worker` still wraps `worker.log` in an unguarded `with` at `windows_office_runtime.py:161`. If its close raises an I/O error while propagating an already-selected worker timeout, cancellation, or Popen launch failure, the context-manager exit replaces that primary exception. The timeout code and cancellation object identity are lost. A failed launch also loses `cleanup_verified=True`, so the outer `_execute` can quarantine a component despite no child having started. This is the same log-close boundary already covered by the Word sibling fix, not a defect in the new evidence-stat catch itself.

New independent scratch probes: `checkpoint-integration-scratch/test_office_log_close_review.py`. They use a fake child and a wrapped local worker-log stream whose close raises EIO; no actual child or native Office is launched. Timeout and cancellation cases confirm the child-cleanup path is reached, while the expected primary outcome is replaced by `OSError`. Launch failure likewise returns the secondary close error. Combined run: **3 failed / 47 passed in 3.20 seconds**. Assertions require typed timeout, original cancellation identity, and known-clean launch classification. The scratch stream exposes both ordinary close and context-manager exit so the same assertions can verify an explicit-finally repair.

Suggested correction: retain the selected primary around worker-log close, append close failures to cleanup categories while re-raising that primary, and continue raising standalone log-close failures after otherwise successful work. Reuse the Word sibling's narrow pattern without changing native worker protocol or any Office process handling.

The initial review invocation encountered only a reviewer basetemp-parent setup error; it was rerun after creating the owned scratch parent. That environment mistake is not a product finding. Review modified only this report and the new independent scratch test; no production source, permanent tests, index, or application state changed.

**Freeze assessment:** secondary-cancel fix accepted in scope, but `07c26d...` is not the final full33 generation-diagnostic freeze until the log-close P2 is fixed and the retained probes pass. Earlier acceptance remains valid for its narrower metadata-order scope and historical hashes.


## Log-close and prelaunch final follow-up

The independent three-case log-close finding was reproduced in permanent tests (3 failed / 23 passed). Three preparation failures (request write, log open, log chmod) additionally reproduce false quarantine before Popen; combined RED is 6 failed / 23 passed. `_run_worker` now explicitly owns the log and its primary failure across setup, launch, wait and response classification. Finally-close preserves an existing exception and records a safe cleanup category; with no primary, the close failure propagates. Known no-launch or acknowledged-closed state is retained independently of diagnostic or final log I/O. Uncertain child state never gains a cleanup acknowledgment.

Final scoped regression including all unchanged three independent probe files: **57 passed** in 4.48s. A success-path test verifies log-close failure still propagates after native cleanup acknowledgment. This replaces the previous Office final candidate; independent re-review is pending. Source remains uncommitted/un-staged.

- `2a1317ffeee59ea30dbcdad64b89a1b308a2ccb1684f082773fa1365eae8d904` `skills/WPSComposer/scripts/msoffice/windows_office_runtime.py`
- `7584079d6227714c4ab325b2082b4e967b78f522e696e114a1af331fb35a4e9c` `tests/msoffice/test_windows_office_runtime_diagnostics.py`

## Independent final re-review — checkpoint_integration_review

**PASS.** Independently verified runtime `2a1317ffeee59ea30dbcdad64b89a1b308a2ccb1684f082773fa1365eae8d904` and tests `7584079d6227714c4ab325b2082b4e967b78f522e696e114a1af331fb35a4e9c`. The prior log-close P2 is closed; no additional actionable P1/P2 was found in this bounded runtime/diagnostic patch. These exact two hashes may be frozen for full33 integration.

Re-ran both repository suites and all three unchanged independent scratch files: **57 passed in 2.98 seconds**. Scratch SHA-256 values remain `930dbadb0e9a464303ee4e33e1076445d45dcdc57dc228b5f5326ac299be962b`, `a56db1d48a55ed6fe603011e32c3b0f61ed6ca6f8d8ca601f5b6f6d816ae643e`, and `bbb2da8ea8de53c6603b7fb5fceeb0f111b0160fdc219316dcba8e00af7847b4` respectively; prior RED assertions were not weakened.

Review confirms explicit log finalization preserves primary timeout/cancellation/launch classification, appends secondary close failure categories, and propagates a standalone close error after success. Setup exceptions before a Python child starts and failures after an acknowledged native close retain `cleanup_verified=True` independently of evidence persistence. Uncertain launched state remains unverified and follows the outer quarantine gate. Exact Python-only cleanup, bounded wait, evidence metadata cancellation handling, and the existing success-path deadline gate remain intact.

Only this report was changed during re-review; tests wrote to the owned scratch basetemp with bytecode and pytest cache disabled. No production source, permanent tests, shared index, or native Office state was modified. This supersedes the prior open log-close finding for the final hashes only. Full33 results, Windows CI, and native/UI acceptance remain separate root-owned gates.
