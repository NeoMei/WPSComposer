# Release readiness — September 11, 2026

**Not ready for a Microsoft/WPS parity release.** Runtime checkpoint `9e09bbe0` (parent `ce322b29`), PR9 remains draft. The frozen 628 capability rows and 1,256 platform gates are unchanged. This continuation adds a private caption primitive and a shared Writer caption-range bug fix, plus native diagnostics and evidence reconciliation. It does not enable a complete new public method family.

| Gate | Current evidence | Remaining work |
| --- | --- | --- |
| Source regression | [Full49](full-round49/README.md):5,283 passed/12 skipped after profile timeout fix;380 before/after source hashes equal | Current hosted CI and native acceptance remain separate; full48 retained as earlier immutable evidence |
| Hosted portability | [CI15](portable-ci/run-15/README.md) passed four jobs at a6bd0140: Linux each5,248 passed/45 skipped; Windows each5,194 passed/99 skipped | Hosted tests do not execute opt-in native Office scenarios |
| Mac installed public paths | Current Word isolated install at a6bd0140:216 identical source files and six public checks. Earlier Excel/PowerPoint installed smoke retained separately | Complete argument, format, lifecycle and capability matrix; current all-app installation |
| Mac Word direct API | 71/80 declared method names; declaration is not native certification | Nine absent methods below; complete native acceptance for implemented methods |
| Word full-quality snapshot | Defined-style and page-story diagnostics preserve state on bounded cases | Complete linked story, character/style/layout coverage and repeated full snapshots without read side effects |
| Windows desktop | GitHub evidence d653ad0e recovered: candidate70d3d86 startup blocked by three retained quarantines; UI/native not run; full pytest interrupted | Resolve retained exact-owned state, three older native defects, then current candidate Office/WPS native/UI/installed execution |
| Final acceptance | Historical scoped reviews and native/UI evidence retained | Per-capability certification and final whole-branch review on the finished candidate |

Nine absent methods: `add_captioned_figure_fallback`, `add_captioned_figure_native`, `add_equation_native`, `add_equation_native_fallback`, `add_heading_level_native`, `add_horizontal_line`, `add_semantic_table_fallback`, `add_semantic_table_native`, `add_wordart`. The existing paragraph-border method does not satisfy native inline-line semantics; equation numbering does not satisfy native equation creation.

The latest [heading diagnostics](macos-word-heading-import/README.md) now retain native donor import v5 with **96 passing operation/style-independence checks**. It uses same-text source and detached controls, complete named style materialization, 65 font observations, 19 effective tabs, a source-plus-reciprocal-linked-style challenge, full baseline restoration, native DOCX/PDF and read-only reopen. Earlier failures remain immutable and separately explained.

**The rendered heading artifact fails visual acceptance.** Word canvas and PDF both display seven missing-glyph boxes for each Arabic word. The installed Arial Bold Italic faces lack those glyphs; PDF font analysis proves 21 painted `.notdef` glyphs. A read-only negative guard rejects this output; five pure guard tests and 19 v5-validator tests pass. The exact saved UI copy equals the original DOCX bytes; its explicit close, source preservation and empty final Word inventory are confirmed. A separate fresh `ui-05/fresh-ui-report.json` now establishes the complete UI edit/Undo/Save/close/same-path reopen/close sequence, unchanged bytes and empty final inventory; the earlier close-only report is retained. These results do not enable the public heading method or settle font fallback, arbitrary inherited styles, numbering or selection behavior.

The first italic-only trial saved native false by removing exactly two source/linked italic leaves, then stopped at a strict false-serialization assertion before PDF export. Its input/source bytes and final empty inventory are preserved. The distinct continuation completes26 checks: nonitalic source Arabic renders without missing glyphs, unchanged italic controls retain seven missing glyphs each, and restoring source italic reproduces the defect. Full style/theme bytes restore exactly, native reopen and cleanup pass. These are hypothesis tests, not production font substitutions. The reviewed private native-caption implementation and shared Writer range repair pass final bounded native run04 (11 checks) and fresh UI04 edit/Undo/Save/reopen with byte-identical copy. The strengthened validator and independent review are complete; full48 passes5,281/12 with402 hashes equal. All complete figure/equation/table families remain incomplete.

The horizontal-line document-container variant had already failed with a floating rectangle instead of the required inline object. Its stale compile-only plan has been corrected to the existing native evidence; the same constructor will not be repeated as a new test.

Temporary localhost certificate trust, test private key and listener were removed after the earlier Office.js activation attempt. No functional Office.js path was accepted. No merge, parity tag/release or personal-plugin installation has been performed during this continuation.

The production-shaped heading-style setup now has two preserved native failures before import. The second run uses exact named style resolution and real typed readbacks; the custom parent still bases on Heading1 after requesting Normal. No style-refresh or custom-parent capability follows from these runs. A separate supported style-copy mechanism must be verified; no further setter normalization is assumed. [Current installed Word](installed-word-caption-checkpoint/README.md) passes six public checks with216 exact installed/source hashes and empty before/after document inventories.

Current [figure rollback feasibility](macos-word-figure-rollback/README.md) v3 passes ten bounded native checks including one-image and second-image-failure rollback, complete body/style/theme preservation after save/reopen and empty final inventory. V2 scalar PASS missed changed no-proofing run properties; that raw failure and new regression oracle are retained. V3 restores the original native proofing value. The whole figure family and arbitrary rich-text rollback remain unaccepted.

Execution availability: independent task-review agents stopped at the account usage limit. The alternate provider rejected the encrypted agent task before reading it. New build-only figure diagnostics have root review and pure/compile checks, but their required independent review remains open. The Windows task still exposes its older interrupted turn, but [a GitHub evidence branch](windows-70d3d86-20260910/SUMMARY.md) now supplies a September10 checkpoint for older70d3d86. Root verified23 manifest hashes. It is explicitly BLOCKED, not a current native pass. It also supplies the profile-server shutdown stack and older Word handle/Excel move/PowerPoint window-binding failures. These constraints do not remove any frozen parity requirements.

Next ordered work: (1) verify the documented single-style `organizer copy` command using only owned Word-saved DOCX donors and the exact private recipient; check live memory and saved definitions separately, then same-name overwrite and reopen; (2) repair figure-local rollback using exact object/preimage identities and verify restored bookmark/format graphs; (3) finish the remaining nine Word direct-method families and complete quality/argument/format rows; (4) run current Windows Office/WPS native/UI/installed acceptance and final independent whole-branch certification.


September11 follow-up: the Windows profile-server shutdown failure was reproduced
locally with real idle/partial-request TCP clients and ineffective-shutdown fault
injection (two RED cases). The request handler now has a five-second socket
read/write timeout, retaining joinable handlers and normal shutdown. All22 profile
server tests pass. [Full49](full-round49/README.md) passes5,283/12 with380 unchanged source hashes; independent review and Windows
retest remain pending. Missing node template dependencies in the older isolated
Windows checkout are a separate setup failure, not proof of a generation defect.
