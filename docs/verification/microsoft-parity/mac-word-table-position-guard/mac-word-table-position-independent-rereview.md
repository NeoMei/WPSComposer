# Independent re-review: Mac Word table-position patches bypass

Verdict: **PASS for this narrow containment patch; the original P2 is closed. No new P1/P2 identified in this bounded re-review.** Candidate may be frozen for full37 integration. This does not implement or accept general middle table insertion.

## Exact source identity

Independently matched all supplied SHA-256 values:

- document_api.py: `1f08cc759fcaac395fd5f766f4347f59635d7890ee3aa9a88d5a8de8c1d8df81`
- edit_preflight.py: `5df8f8689306e73f4516723d79e14efe86f022320e5a68a975f7494709c5e09f`
- test_mac_word_table_position_guard.py: `20bb7d68dcb067ec36946bade64ef2e809a8f13ae044fd5768d82fd95fff278e`
- Original independent scratch: `dc639dc1d8b0de39203207870ffec5f876fc2fbcbc9767404b1dd7d369db6a17` (unchanged).

## Why the finding is closed

`edit()` constructs the normalized patches-then-ops tuple once before engine resolution. The exact `combined` variable is passed to engine selection, the Mac Microsoft writer pre-open position guard, attached atomic validation, and operation execution. Therefore a patches entry whose explicit op overrides the default set is inspected before source/publication preparation and native opening, just like an ops entry. The later duplicate normalization was removed.

Reviewed all combined consumers: they iterate, index/count or pass the tuple; none requires the previous mutable list. Patches override semantics, iterator consumption order and auto Mac Word ops-row materialization remain intact. The original table-only validator and existing-session whole-batch guard are unchanged. Windows/WPS routes and default/None/end append remain outside the nonterminal rejection. The explicit ValueError and automatic EngineUnavailableError safety-preflight contract remains as previously reviewed.

## Independent verification

Fresh command using complete clean-dev Python:

```sh
PYTHONDONTWRITEBYTECODE=1 PYTHONPATH=.:tests/msoffice /var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-git-export-round25-lgf9jekp/clean-dev-venv/bin/python -m pytest -q -p no:cacheprovider tests/msoffice/test_mac_word_table_position_guard.py tests/msoffice/test_edit_preflight.py tests/test_document_api.py tests/msoffice/test_windows_document_api.py .superpowers/sdd/2026-09-08-microsoft-wps-parity/checkpoint-integration-scratch/test_table_position_patch_review.py
```

**320 passed in 5.68s**, including the original two RED reproductions unchanged, four new lone/mixed patches × atomic-mode pre-open guards, and explicit/automatic engine generator ordering guards. Log: `mac-word-table-position-independent-rereview.log`.

Changed-file `git diff --check` passed. Original finding report, scratch and RED log remain separate and unchanged. No production/test/index changes, native Office commands, installation or UI actions were performed by this reviewer.

Full37 and candidate-specific hosted/native validation are separate gates. Reported CI06 green belongs to the prior runtime candidate, not proof of this later guard. General nonterminal table positioning, semantic replacement and quality snapshot acceptance remain required and open.
