# Native Word list argument acceptance at 433a199

Run-02 passes 26 checks for default/custom/empty glyphs, ordered glyph ignoring, 24/30/27.5/33 pt indents, List Paragraph formatting resets and following Body Text restoration. Actual saved OOXML, native computed paragraph readback, one-page Word PDF, read-only reopen, exact close and source hashes are verified. Native before/after document inventories are empty; this run is not an unrelated open-document sentinel exercise. Independent source/artifact review reports no blocker.

Run-01 remains FAIL: the checker compared Chinese native built-in style names with English literals. Its frozen source and output remain unchanged. Run-02 reads localized built-in identities from the bound document before comparing actual paragraph styles. The production change is identical in both runs.

UI-01 edits the existing custom-list paragraph in Word, observes the replacement, uses Undo, explicitly saves/closes, and reopens the exact copy. The reopened window/path is AX-confirmed; the attempted reopened screenshot returned noWindowsAvailable and is not claimed. Independent final native inventory is empty; both source and saved UI copy are byte-identical to their preimages. The earlier successful screenshots and UI steps remain in the parent task tool record.

Source runners below retain their original build paths. To replay, restore source/ under build/word-list-argument-20260910/source and use a fresh output directory; SOURCE-FREEZE.json must exactly match the runner and implementation. This is scoped Mac list acceptance, not Windows list native evidence, all Word direct methods, or release approval.
