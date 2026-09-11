# Review: Mac Word field topology after owned mutation

Date: 2026-09-09  
Scope: pure review of `MacWordSession.snapshot_fields()` topology ownership across public, same-session structural mutations. No production or repository-test edits, native Office/WPS calls, checkpoint implementation, or delegation were performed.

## Verdict

**CHANGES REQUIRED — one P2 confirmed.** A legal same-session structural mutation leaves the prior observational field topology active. The next `snapshot_fields()` therefore reports `NATIVE_WORD_FIELD_IDENTITY_STALE` even though this session itself made and acknowledged the change.

## P2 — owned field/structure changes are misclassified as external topology drift

`macos_word_fields.snapshot()` writes `_observed_field_topology` and later requires the same field count, order, kind/code, and start positions adjusted only by each already-observed field's own result-extent growth. This is the correct conservative rule when no known structural mutation has occurred: it detects deleted/reordered/retyped/moved external fields while allowing TOC/result growth to shift following fields.

The baseline is never invalidated by `_execute_structural()` or `_business_commit(..., structural=True)`. Consequently all of these supported sequences fail on the current source:

1. `snapshot_fields()` → acknowledged `add_cross_reference_paragraph()` → `snapshot_fields()`; the observed field count legally grows.
2. `snapshot_fields()` → acknowledged `insert_toc()` → `snapshot_fields()`; the observed field count legally grows.
3. `snapshot_fields()` → acknowledged paragraph insertion before an existing field → `snapshot_fields()`; the existing field legally shifts even though its identity/code/order are intact.

This is P2 because it blocks a normal public call sequence and field-convergence/controller work after a legitimate mutation. No source corruption or unrelated-document mutation is demonstrated.

## Independent RED

Reviewer reproduction: `word-field-topology-mutation-review-repros.py`.

Result on the frozen sources: **3 failed / 2 passed**.

- The three failures are the same-session REF addition, TOC addition, and preceding paragraph insertion above. Each incorrectly raises `NATIVE_WORD_FIELD_IDENTITY_STALE`.
- The first passing control changes an existing field position without any owned mutation and confirms the external drift is still rejected.
- The second passing control grows a TOC result and shifts the following REF by exactly the extent delta, confirming the proven growth rule remains accepted.

Retained output: `word-field-topology-mutation-review-red.log`.

## Bounded repair direction

Invalidate only the **observational baseline** before a same-session structural native batch can mutate the document. The next snapshot may then establish a new baseline from exact native readback. Preserve `_tracked_indexes` and `_tracked_references`: their bookmark, code, native start and unique-field matching still prove that previously and newly owned semantic handles resolve to the expected fields.

The invalidation should be centralized with the existing structural-mutation state transition so it covers public paragraph/list/table/image/section/removal operations and field/reference insertions, including a native batch that applies a prefix before failing. Clearing only in `insert_toc` or reference helpers would leave other legitimate shifts broken.

Do not clear the baseline for a field refresh/update alone. That would weaken the existing exact result-extent shift validation that permits TOC growth while detecting unexpected reordering, code changes, field insertion/deletion, or movement. Snapshot itself must continue to compare and then replace the baseline.

Checkpoint/rollback will need the same ownership rule later: rollback should restore or deliberately rebase observational topology according to its verified post-rollback document, while retaining valid old handles and accepting newly created handles only after their acknowledgement. This review does not define or implement that pending contract.

## Frozen hashes

| File | SHA-256 |
|---|---|
| `macos_word_session.py` | `532550b1fa7dbb10de1e9025307f406bfc69e6409cf5a056da1266808488a467` |
| `macos_word_fields.py` | `c390342553fa4f9c30f77dc751da99e97ffd0abc1ff0e20822f9c5c974fa776d` |
| `macos_word_references.py` | `f2243a6ce644f6a196db087756b8dba51d6d9bb713c586a9f0175819511200f6` |
| reviewer reproduction | `60f89b2b3cd184673e07835e1f9ecdf2ea5a6354df65a474cd3e430d62472429` |
| RED log | `2e134ced1e811d8ddbb673cd91aeca2280864e2f71ecfb52fea8cbb718a31de6` |

The reproduction is pure and models native rows; it proves Python state and acknowledgement handling, not native Word's enumeration behavior for every field/story. Existing run-12 and reference run-04 evidence remain the bounded source of native field/topology behavior. A repair requires focused pure regression and affected native acceptance before it can be called current-source native proof.
