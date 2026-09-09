# Matrix03 parent reconciliation — separate from the failed invocation

The original Matrix03 A invocation remains an unacknowledged -1708 failure, and its B case remains unexecuted. Parent's subsequent exact-owned/sentinel read-only reconciliation found an actual mutation despite that error:

- New table ordinal 1: start 12, end 23, **1 row × 3 columns**, with REPLACE in its cell text.
- Original table ordinal 2: start 46, end 65, 1×1, original text retained.
- Body end 67, table count 2; exact sentinel unchanged.

Evidence: `constructor-matrix-03-parent-cleanup/readonly-reconciliation.log`. Parent also reported visual confirmation of the 1×3 table, then precisely discarded the owned document and deleted sentinel 文档94 after matching its complete synthetic token. `constructor-matrix-03-parent-cleanup/report.json` records after-close inventory [], recovered quarantine true and final inventory []. These are separate parent reconciliation/cleanup results, not a retroactive ACK or acceptance for the original constructor.

The -1708 message cannot be interpreted as absence of mutation or universal lack of conversion support. Its reference to selection.text object motivates one bounded hypothesis about a range reference becoming invalid during conversion; that explanation remains unproven. Requested 4×3 was not observed (actual 1×3), so range placement, returned-object ACK, row count and column count remain separate unresolved contracts.
