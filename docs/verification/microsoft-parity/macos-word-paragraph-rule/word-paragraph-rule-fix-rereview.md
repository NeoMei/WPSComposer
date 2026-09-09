# Mac Word paragraph rule P2 fix re-review

Date: 2026-09-09  
Verdict: **ONE ADDRESSED / ONE STILL OPEN.** This is a bounded re-review of the two P2 findings in `word-paragraph-rule-review.md`. No new scope, production edit, public forward, native Office/UI job, installation, commit, or push was performed.

## Frozen inputs

| File | SHA-256 |
|---|---|
| `skills/WPSComposer/scripts/msoffice/macos_word_rules.py` | `0a7fdfbe3130e748b64a84c18e24552ca09c607e053c315abbcea5140e827006` |
| `tests/msoffice/test_macos_word_rules.py` | `70b502e2731d4eec79813e990f60dc3258eff6d8a2afc9497c8a09b2b8d3d0f7` |
| `fixtures/microsoft_parity/macos_word_paragraph_rule.py` | `1a082729d36ec2f63a7838b8c0537f0b559641f2c2f92f94cdbfabe13451d9ee` |
| observed shared `macos_word_session.py` | `f8b2e89a4aefddb27158fefa958da967754e2e52c8e7d02491a05fa35faa4170` |
| amended owner report | `1375bcda826134404b46085aef17eb54bf826d9a728058ccc14908771f99fb2a` |

## P2 1 — current-paragraph parity and exact prefix proof: ADDRESSED

The generated commands no longer call `_paragraph_boundary()`. They bind the current terminal paragraph, set its center alignment and native single `.75pt` `C0C0C0` bottom border before insertion, append exactly space plus CR at the prior terminal point, rebind the same ruled paragraph, and clear only the new following paragraph's bottom border. The following ACK requires inherited centered alignment and unchanged inherited style/spacing. This now matches frozen `WriterComposer.add_paragraph_horizontal_line` for empty and nonempty terminal paragraphs; existing terminal text remains in the ruled paragraph.

The native prefix check now uses Foundation `NSString.isEqualToString:` rather than AppleScript's case-insensitive `is`. ACK geometry requires exactly two added native units, one added paragraph, exact inserted/following text, current paragraph bounds and all border/alignment flags. The old extra-separator ACK is rejected. The fixture and OOXML verifier now require two paragraphs, the existing prefix plus space in the ruled first paragraph, and centered unbordered following text with inherited style/spacing. They no longer claim the original paragraph formatting is preserved.

## P2 2 — independent sentinel lifecycle: STILL OPEN (narrow fixture blocker)

The structural cleanup repair is otherwise correct: the owned context exits first; its runtime is copied after close; `close_after_owned` requires a verified nonquarantined owned close and exact preimage row before invoking the independent guarded `close_sentinel`; final inventory must equal starting inventory; reopen also exits before its inventory check; failure cleanup is attempted at most once and cannot replace the primary failure.

However, fixture line 142 constructs the preimage as:

```python
expected_sentinel = [sentinel_name, '', False, sha256(token + CR)]
```

The shared `inventory()` records `posix full name of inventoryDoc`, and native Word may return the unsaved document's basename rather than an empty string. Existing native diagnosis rows demonstrate that shape (`[name, name, false, digest]`). Therefore this assertion can false-fail before the public method is called on an otherwise valid unsaved sentinel. An empty path is correctly checked later by the independent `close_sentinel` through `path of sentinelDocument`; it must not also be inferred from inventory's `posix full name` representation.

Required bounded fix: select the unique actual sentinel row from `created_inventory`; require exact name, `saved is False`, and the exact nonempty token digest; store that actual row as `sentinel_before`. Continue requiring byte-for-byte row equality after owned close. Keep the independent `close_sentinel` path-empty/name/token/saved guard unchanged. Add one pure fixture case where inventory returns the basename in column two, and retain the existing changed-hash/owned-not-closed/close-failure cases.

No other P1/P2 was found within the two original findings. The public session forward remains correctly pending, and no source-bound native acceptance should start until the narrow fixture preimage assumption is fixed.

## Independent pure verification

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps \
../wps-task-session-startup/.venv/bin/python -m pytest -q \
  tests/msoffice/test_macos_word_rules.py

46 passed in 0.36s
```

The green suite establishes the production semantic repair and current fixture helper branches, but its sentinel doubles all use an empty second column and therefore do not close the remaining native-representation blocker.
