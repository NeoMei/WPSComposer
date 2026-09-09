# Independent Mac Word table-position guard review

Verdict: one P2 remains before freeze; the ordinary ops path is correctly protected, but edit's pre-open guard does not inspect the same normalized batch as routing/execution.

## P2: inspect normalized patches as well as ops before native open

`document_api.py:988` invokes `validate_mac_word_table_positions(ops)` only. Existing public normalization uses `{"op": "set", **patch}` both for routing (`:982`) and the eventual combined execution (`:996`), allowing an explicit `op` in a patches entry to override the default. Thus `edit(engine="msoffice", patches=[{"op":"insert", "type":"table", "position":"start", "props":{"rows":1,"cols":1,"data":[["keep"]]}}], output="/other.docx", atomic=False)` reaches the native opener instead of rejecting the unsafe request first. The same holds with atomic=True and a single-use patches iterator.

Independent pure repro: `checkpoint-integration-scratch/test_table_position_patch_review.py`, **2 failed in 0.51s**, log `mac-word-table-position-independent-red.log`. The opener is a test double that records invocation and raises; no native application or document is touched. A separate one-shot reproduction also reached the opener.

The existing apply_ops guard should reject this structural operation before batch writes once the document is opened. This finding does not claim that the known misplaced-table write escapes that second guard. It does show the promised before-open whole-request protection is incomplete, including the earlier publication/source-preparation stages. Auto-engine routing already checks the normalized complete operations and does not share this bypass.

Recommended narrow correction: derive one normalized combined batch and use that identical batch for engine selection, the Mac Microsoft pre-open table-only guard, and execution. Preserve existing patches op normalization semantics rather than introducing unrelated input changes. Add regression for patches iterator containing an explicit structural op, atomic True/False, with opener never reached.

## Verified aspects

Matched all three supplied SHA-256 values:

- document_api.py: e7e9e57b7b5daaca03e01091701c345ce1493728c4ad81dfd97d7f27d65d66e0
- edit_preflight.py: 5df8f8689306e73f4516723d79e14efe86f022320e5a68a975f7494709c5e09f
- test_mac_word_table_position_guard.py: a0651f705ad33897374eedb69c9c82b2282a125772346610113046690e417859

Independent existing/dedicated test run: **221 passed in 2.47s**, log `mac-word-table-position-independent-green.log`. Command used complete clean-dev Python, `PYTHONDONTWRITEBYTECODE=1`, `PYTHONPATH=.:tests/msoffice`, `pytest -q -p no:cacheprovider tests/msoffice/test_mac_word_table_position_guard.py tests/msoffice/test_edit_preflight.py tests/test_document_api.py -k 'not native_dictionary_compiles'`.

Code paths confirm default/None/end inserts remain available; whole existing-session ops generators are materialized before any dispatch; mixed unsafe ops reject before earlier format writes for either atomic value; direct session validates before position/AE commands; automatic routing retains EngineUnavailableError while explicit guard uses ValueError. Windows supports predicates avoid Mac structural validation; WPS/Windows apply_ops bypass the new Mac-Microsoft gate. Ordinary patches-only/None/empty-generator edits remain compatible. The interim ValueError is a deliberate safety preflight exception rather than per-operation PatchError/best-effort reporting; it is consistent with the stated guard contract.

Changed-file whitespace check passed. No source/permanent-test/index/native edits were performed. Only reviewer scratch, logs and this report were added. Full middle insertion, native semantic replacement and quality snapshot acceptance remain open; this containment does not close baseline support.
