# Mac Word nonterminal table insertion interim guard

Implemented a narrow content-protection guard after the parent's native constructor experiments confirmed that nonterminal table requests can create a table at the wrong location while reporting success. This prevents the known false-success path. It does **not** implement precise table insertion or complete the required structural-edit baseline; middle/before/after/index insertion remains open pending a reliable native implementation and acceptance.

## Change and boundaries

- `edit_preflight.validate_mac_word_table_positions` rejects an `insert` + `table` operation whose position is neither omitted/`None` nor `end`. It is deliberately table-only and does not validate or alter other best-effort operations.
- `validate_word_structural` reuses that helper. Existing auto-engine whole-request capability selection therefore rejects an unsupported Mac Word batch before native open, retaining its existing `EngineUnavailableError` behavior. Direct `MacWordSession.apply_structural_op` also rejects before position/range commands or native execution. Its existing `_writable()` check is unchanged and only checks local read-only/closed flags; no session source modification was necessary.
- `document_api.edit` invokes the same guard after engine resolution and before opening/attaching, filesystem publication preparation or mutation, only for Mac + selected Microsoft + writer. Explicit `msoffice` now raises `ValueError` for this unsafe operation even under `atomic=False`; other explicit best-effort behavior stays in place. `ops` has already been normalized to a tuple, including omitted/`None` values.
- `document_api.apply_ops` materializes the operations once and invokes the same table-only guard before dispatching the first operation, only for an existing Mac Microsoft writer session. A later unsafe table operation therefore rejects the entire batch before earlier writes, for both `atomic=True` and `False`. This API receives an already-open composer; it cannot undo a caller's earlier open, but makes no new native call for the rejected batch.
- Windows and WPS routes are outside both new public guards. Existing Mac WPS automatic routing continues to reject unsupported Word structural editing; it does not receive an invented fallback. Default/end table append remains available. Heading, paragraph, image, quality/degradation builders and generation constructors are unchanged.

The parent explicitly authorized extending the implementation scope to the two narrow `document_api` call sites after review established that shared validation alone would not protect explicit engine and already-open batch calls.

## TDD and verification

The new permanent test file first ran against unmodified production: **20 failed, 10 passed in 2.68s**. The eight public edit failures reached the forbidden native opener; the eight existing-session batch cases and four direct cases failed to raise. These are the intended missing-protection failures, covering `start`, `before`, `after`, and `index`. Ten compatibility guards passed. Preserved log: `mac-word-table-position-guard-red.log`.

After the minimal production change:

- New 30 tests + existing edit-preflight suite: **98 passed in 6.21s**, `mac-word-table-position-guard-green.log`.
- Existing `tests/test_document_api.py`, `tests/msoffice/test_windows_document_api.py`, and `tests/msoffice/test_macos_word_session.py`, excluding the native dictionary compile case: **291 passed, 1 deselected in 27.25s**, `mac-word-table-position-guard-regression.log`.
- Added three final regression guards for explicit patches-only, `ops=None`, and empty-generator requests continuing to the expected opener. Complete new file: **33 passed**, `mac-word-table-position-guard-final.log`.

The new tests retain actual routing, validation and operation dispatch. Native open/transport boundaries use existing COM/AE-free doubles. Successful append requests execute once, Windows capability predicates still accept the prior forms, and non-Mac-Microsoft composer batches retain their prior dispatch. Existing nested-generator and explicit best-effort tests also pass. No native Office/WPS UI, AppleEvents, dictionary compilation, Git/index, commit or push operation occurred. Full36/CI06 snapshots were not modified and cannot establish acceptance of this later repair.

All runs used `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.:tests/msoffice`, `/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest -q -p no:cacheprovider`, owned scratch basetemps, and `--tb=short`. The regression run selected `-k 'not native_dictionary_compiles'`. Changed-file `git diff --check` passed.

## Independent review handoff

Production scope is two files only: `document_api.py` (11 added lines) and `msoffice/edit_preflight.py` (14 added lines). The new dedicated test is `tests/msoffice/test_mac_word_table_position_guard.py`. `macos_word_session.py` was read but not modified.

Frozen SHA-256:

- `document_api.py`: `e7e9e57b7b5daaca03e01091701c345ce1493728c4ad81dfd97d7f27d65d66e0`
- `edit_preflight.py`: `5df8f8689306e73f4516723d79e14efe86f022320e5a68a975f7494709c5e09f`
- Dedicated test: `a0651f705ad33897374eedb69c9c82b2282a125772346610113046690e417859`
- Unchanged session: `47b40d4eabc360fdfcb29244f072b00aa7b286586e95abf5dd848bfa1b83c59e`

Requires a different agent's independent review before full37 freeze. Full37 and hosted CI for this candidate remain separate gates; no overall parity completion or native-position acceptance is claimed.

## Follow-up: patches override bypass repair

Independent review found that `patches` entries can carry an explicit `op`, overriding the existing default in `{"op": "set", **patch}`. The first guard inspected only `ops`, although routing/execution also included normalized patches. Consequently an explicit Microsoft edit with a nonterminal table supplied through patches could reach source/output preparation and the opener before the existing apply_ops guard rejected the operation. Original finding and evidence remain unchanged in `mac-word-table-position-independent-review.md` and `checkpoint-integration-scratch/test_table_position_patch_review.py`.

The approved narrow follow-up changes only `document_api.py` and the dedicated guard test. `edit` now constructs one `combined` tuple from normalized patches followed by ops at the existing routing boundary. Engine selection, table-position validation, and later execution share that tuple. The later duplicate patch normalization was removed. Explicit patch `op` override semantics and outer iterator materialization order remain unchanged; the auto Mac Word ops-row materialization is unchanged. The shared validator and session files were not changed by this follow-up.

Tests-first evidence:

- Four new permanent cases (atomic True/False × lone/mixed patch request) reached the forbidden source-preparation boundary before the fix. The two original reviewer cases still reached their forbidden opener. Two positive patch-override/ops-generator ordering guards passed. Combined RED: **6 failed, 2 passed, 33 deselected in 1.18s**; `mac-word-table-position-patch-red.log`.
- After correction, dedicated guard tests, existing edit-preflight, public document API, Windows document API, and unchanged independent scratch: **320 passed in 7.15s**; `mac-word-table-position-patch-green.log`. Both automatic and explicit engine success cases retain patch-before-ops ordering, execute each yielded structural operation once, and preserve the caller's explicit insert verb. Patches-only/None/empty and other prior guard cases remain covered.

Commands retained the full clean-dev interpreter, `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.:tests/msoffice`, `pytest -q -p no:cacheprovider`, and separate `table-patch-red` / `table-patch-green` owned basetemps. RED selected `patch_override or patch_supplied_structural`; GREEN ran the five files named above without deselection. Changed-file whitespace check passed. No native, index, or other source-file changes were made.

Updated SHA-256 for re-review (supersedes the earlier document/test freeze above):

- `document_api.py`: `1f08cc759fcaac395fd5f766f4347f59635d7890ee3aa9a88d5a8de8c1d8df81`
- `test_mac_word_table_position_guard.py`: `20bb7d68dcb067ec36946bade64ef2e809a8f13ae044fd5768d82fd95fff278e`
- Unchanged `edit_preflight.py`: `5df8f8689306e73f4516723d79e14efe86f022320e5a68a975f7494709c5e09f`
- Unchanged independent scratch: `dc639dc1d8b0de39203207870ffec5f876fc2fbcbc9767404b1dd7d369db6a17`

The repair is ready for the original independent reviewer to recheck. This implementation report does not self-approve the finding or claim full37/CI/native acceptance. Precise nonterminal insertion remains an open requirement.
