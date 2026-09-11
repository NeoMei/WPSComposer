# Independent review: Mac Word field topology after owned mutation

Date: 2026-09-09  
Scope: independent scoped re-review of the original P2 in `word-field-topology-mutation-review.md` against the owner's frozen three-file repair. I copied the frozen sources and diff before subsequent checkpoint work could modify shared files. I made no production/test edits and started no native application.

## Verdict

**SCOPED PASS — original P2 ADDRESSED; no new P1/P2 in the frozen repair.** The next `snapshot_fields()` now accepts acknowledged session-owned REF insertion, TOC insertion, text replacement and structural position shifts. The unchanged snapshot logic still rejects external field movement/code replacement and still allows exact TOC-result growth to shift a following REF.

## Original finding: ADDRESSED

The repair keeps `_observed_field_topology` as an observational drift guard, but invalidates it for known content/topology mutations. It preserves `_tracked_indexes` and `_tracked_references`, so snapshot still rebinds every owned handle through its bookmark, exact code and unique native field position rather than trusting the new baseline alone.

The mutation paths are covered centrally:

- `_execute_structural()` defaults to a topology-changing submission; field refresh explicitly passes `field_topology_change=False`.
- `_business_commit()` distinguishes structural/content changes and specific header/footer field changes from format-only updates.
- text-bearing `apply_format_patch()` and public structural operations use the same mutation boundary.
- header/footer invalidation does not mark the main body structurally stale, clear `_pending_heading`, or invalidate stable body paragraph targets.

This preserves the exact no-owned-mutation comparison in `snapshot()`. Result-only refresh therefore cannot hide external position or code drift, and tracked TOC result-extent growth continues to account for the exact shift of a following REF.

## Submission/preflight boundary: ADDRESSED

An earlier repair draft considered `_quarantined` or `_retain_evidence` when deciding whether to clear the baseline. Those are sticky session-history flags: an older failure can leave either set, so they cannot prove that the current call submitted an AppleEvent. Using them would let an unrelated later pre-submission deadline or script-preparation failure erase the external-drift guard.

The frozen code instead uses a per-call `_field_topology_mutation_pending` marker. `_mutation_preflight()`, the first `_remaining()`, script construction/write/chmod, and the final `_remaining()` all complete while the prior baseline remains present. `_execute()` invalidates immediately before `subprocess.run` only when that call's marker is pending. Therefore:

- local/preparation/deadline failure before submission preserves the baseline, even if `_retain_evidence` was already sticky;
- a test transport that returns an acknowledgement invalidates in `_execute_topology_mutation()` after success;
- native return errors, timeout, `KeyboardInterrupt`, `SystemExit`, malformed completion and post-return deadline exhaustion remain conservatively invalidated because submission already occurred;
- the per-call marker is cleared in `finally`, so later operations cannot inherit a stale pending state.

The focused tests cover local preflight failure and a late failure immediately before `subprocess.run`. The frozen code has no remaining dependency on sticky `_quarantined`/`_retain_evidence` for this decision.

## Independent pure verification

On the exact frozen source, the combined focused command passed **228 tests in 3.62s**:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest -q \
  tests/msoffice/test_macos_word_session.py \
  tests/msoffice/test_macos_word_business.py \
  tests/msoffice/test_macos_word_fields.py \
  tests/msoffice/test_macos_word_references.py \
  .superpowers/sdd/2026-09-08-microsoft-wps-parity/word-field-topology-mutation-review-repros.py
```

The original reviewer reproduction independently passes **5/5**. Its immutable RED log remains **3 failed / 2 passed** on the old source. The two controls prove external drift still fails and exact TOC-growth/following-REF movement still succeeds.

## Native evidence review and limits

`word-field-topology-native-run-04/report.json` is source-bound to the frozen session and fields hashes and has all seven named checks true. Its native snapshots show field counts `1 -> 2 -> 3`, preserve earlier stable keys after owned REF and TOC creation, and preserve all three after a public paragraph insertion. The saved DOCX contains three main-document instruction nodes (owned REF, TOC, TOC PAGEREF), the expected Prefix/Owned text, and the reopened report records five Word fields plus one TOC. The two-page PDF contains the expected body and contents text. Artifact hashes match the report.

This is actual native evidence, but its runner is not yet a durable reusable fixture: it used a synthetic non-file `__main__.__file__` to trigger the existing multiprocessing spawn bootstrap instead of a normal guarded entry point. The report's `inventory_before` and `inventory_after` prove the one unrelated unsaved sentinel stayed byte-identical. Its `final_inventory` value is the placeholder `session-closed-or-quarantined`, so the report alone does not prove a final zero-document inventory; the parent's separate live-zero observation remains separate evidence.

The report check named `save_reopen_source` proves saved-output reopen counts/text. It does not contain a before/after source-file hash and must not be cited as source-preservation proof. Run 01/02 wording should also be read narrowly: the invalid-bookmark calls caused no further target mutation, but those probes had already seeded a heading, paragraph and REF.

No native rerun was needed for this review because the final report binds the exact frozen production hashes and the parent independently verified both source hashes, both artifact hashes and the two PDF pages. Future checkpoint integration or other edits to these production files reopen current-source native evidence.

## Frozen evidence

| Item | SHA-256 |
|---|---|
| `macos_word_session.py` | `a96f96f99667e04ec64ed3eb018ba45de308db9a308d37b3df9c98bf2537bc6c` |
| `macos_word_fields.py` | `446e8c9f75ce66e93f55b1d8080aa80db6bf5abd4886c75148d8264a8cc50aeb` |
| `test_macos_word_fields.py` | `8a04fda9ade0e031f6178a8eeec61c4800f88a340b9586acfd1beaef8090e4d3` |
| captured repair diff | `98d3c7b02b7115bf6709a1639efc95c987575bd39706b50ec1ce6d4edd55093e` |
| original reviewer repro | `60f89b2b3cd184673e07835e1f9ecdf2ea5a6354df65a474cd3e430d62472429` |
| original RED log | `2e134ced1e811d8ddbb673cd91aeca2280864e2f71ecfb52fea8cbb718a31de6` |
| combined focused GREEN log | `c78d9043a2f61b954ae0d7ef904037c0e9cc0072066621d0d89d8dbed6daad35` |
| reviewer repro GREEN log | `3f2ea59ea3ff3b1f8ed8689f38597c7b6868acf5bbab9f1e2bf4e322510ff5ce` |
| offline artifact inspection | `cfbfc548e4082ed7b28f0883eb8caf99ba4df9cfe3676ad123552d0c2663ca59` |
| native run-04 report | `7c0865188158c1e6e1bd8426ef5e3acbc428e1e00c30976b0814f8a68d832cc7` |
| native DOCX | `155632ba4e4a1acb9a7e059a2ce39de97641a75d02682dd4702b2864f34654db` |
| native PDF | `9a429892bf08478611ea7bdca3b7378f3b7f980683bdf2efc7773c6c90bd1d37` |

This is a scoped repair verdict. It does not implement or review checkpoint/rollback, certify all Word field stories/arguments, or complete the 628-row parity baseline.
