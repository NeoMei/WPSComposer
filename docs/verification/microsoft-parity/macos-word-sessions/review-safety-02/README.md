# Independent review fixes — native acceptance

`report.json`: PASS. Microsoft Word 16.112.3, macOS. Only the exact owned native document was edited and closed. No GUI, global Quit, security or clipboard changes.

Verified with original Word-generated DOCX/PDF artifacts:

- Negative font size rejects the entire mixed text/font patch before an AppleEvent; text remains unchanged.
- Public new document → two paragraphs with a nonempty final paragraph → clone first paragraph to end → move middle paragraph to end. Exact nonempty paragraph sequence is `First`, `First`, `Terminal`.
- Original-source `save_current`, then `close(save_changes=True)`, publish the intended text. Native reopen confirms `Explicit close saved`, `First`, `Terminal`; PDF retained.
- Replacing the synthetic source externally causes close-save to fail while retaining the same session lock and owned document. Explicit discard then closes only that document and releases the lock; external bytes survive.
- Existing unrelated document bindings and saved states remain unchanged.

`source.docx` is the original native fixture; `saved-current.docx` is the verified edited artifact. `current.docx` intentionally contains the simulated external replacement after the conflict check. `reopened.json` reflects the saved edited artifact before that replacement.

Production changes also invalidate disk-based paragraph IDs when text replacement can change paragraph boundaries, preflight numeric domains, and bind attached native read-only state plus source digest. Attached behavior remains **unverified natively** on this host because Word returns no trustworthy native window ID; portable tests cover the conditional path.

Separate `native-10` passed eight established session checks; `new-public-05-reviewed` passed public new-document/PDF/reopen checks after the paragraph-boundary fix. Targeted suite: 138 passing tests across Word sessions, PowerPoint sessions and document routing.

`review-safety-01` is retained as a failed fixture attempt. Its standalone script lacked a `__main__` guard, so the spawned artifact validator reloaded the script and waited for its parent's Word lock. Only that exact validator child was terminated; the fixture recorded failure and verified native close of its owned document. The corrected guarded script was then run independently in this new directory. No original artifact was modified to fabricate acceptance.
