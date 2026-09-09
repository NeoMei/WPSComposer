# Mac Word paragraph rule final P2 re-review

Date: 2026-09-09  
Verdict: **SCOPED PASS — both original P2 findings are addressed; no new P1/P2 found.** This final re-review covers only the narrow sentinel-preimage refinement and the disposition of the two findings in `word-paragraph-rule-review.md`. No production/session forward, native Office/UI work, installation, commit, or push was performed.

## Frozen inputs

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_rules.py` | `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006` |
| `fixtures/microsoft_parity/macos_word_paragraph_rule.py` | `768c43c0d97e333f0026d32ff9cd50a8afce2de6e97e34def73f42f4a5ad7eb8` |
| `tests/msoffice/test_macos_word_rules.py` | `a0bd103d0f0af00da5dc8fe7058d276a9f5cf9e36fbaff76f3d1f201c25c1591` |

## Original P2 disposition

1. **Current-paragraph baseline and exact prefix proof: ADDRESSED.** Production remains unchanged from the prior successful re-review: it formats the current terminal paragraph before appending space plus CR, does not add a separator, requires centered following inheritance with only its border cleared, and uses Foundation `NSString.isEqualToString:` for exact prefix text. Its semantic ACK rejects the old three-paragraph behavior.
2. **Owned/sentinel lifecycle: ADDRESSED.** `sentinel_preimage` now selects exactly one actual inventory row by sentinel name, requires a nonempty token, a four-column row, a textual full-name representation, `saved is False`, and the exact SHA-256 of token plus terminal CR. It deliberately retains Word's actual second-column representation, including the native unsaved `name/name/false/hash` shape, and later requires the entire tuple to remain equal after owned close. The separate guarded `close_sentinel` continues to verify the true native `path of document == ""`, exact name/token and unsaved state before independently closing only that sentinel. Final inventory must equal the starting inventory.

The new pure regression accepts the realistic basename representation and rejects missing/duplicate identities, saved state, wrong hash, wrong name and non-boolean saved state. Thus the fixture no longer false-fails by conflating `posix full name` serialization with the native path property.

The fixture is now suitable for the source-bound native run after the exact public session forward is integrated. That future run must bind its current session/helper/fixture hashes and remains separate from this pure scoped verdict.

## Independent verification

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest -q \
  tests/msoffice/test_macos_word_rules.py

47 passed in 0.55s
```

`git diff --check` passed for the fixture and focused test files.
