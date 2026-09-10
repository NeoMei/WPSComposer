# Release readiness — September 11, 2026

**Not ready for a Microsoft/WPS parity release.** Candidate `73a00e1f0a79d038d8dcf3402057700ada5f27e9`, PR9 remains draft. The frozen 628 capability rows and 1,256 platform gates are unchanged. This checkpoint adds native diagnostics and evidence reconciliation; it does not enable a production capability.

| Gate | Current evidence | Remaining work |
| --- | --- | --- |
| Source regression | [Full46](full-round46/README.md): 5,253 passed / 12 skipped at exact73a00e1; all401 before/after/live hashes equal | Run affected checks after implementation changes; do not transfer results to changed sources |
| Hosted portability | CI13 passed four jobs at73a00e1: Linux each5,221 passed/44 skipped; Windows each5,167 passed/98 skipped | Hosted tests do not execute opt-in native Office scenarios |
| Mac installed public paths | Isolated installer: 214 identical source files; Word/Excel/PowerPoint each six representative public checks | Complete argument, format, lifecycle and capability matrix |
| Mac Word direct API | 71/80 declared method names; declaration is not native certification | Nine absent methods below; complete native acceptance for implemented methods |
| Word full-quality snapshot | Defined-style and page-story diagnostics preserve state on bounded cases | Complete linked story, character/style/layout coverage and repeated full snapshots without read side effects |
| Windows desktop | Remote continuation accepted; latest refresh exposes old interrupted turn and notLoaded state | Fresh candidate Office/WPS native, UI and isolated-installed execution |
| Final acceptance | Historical scoped reviews and native/UI evidence retained | Per-capability certification and final whole-branch review on the finished candidate |

Nine absent methods: `add_captioned_figure_fallback`, `add_captioned_figure_native`, `add_equation_native`, `add_equation_native_fallback`, `add_heading_level_native`, `add_horizontal_line`, `add_semantic_table_fallback`, `add_semantic_table_native`, `add_wordart`. The existing paragraph-border method does not satisfy native inline-line semantics; equation numbering does not satisfy native equation creation.

The latest [heading input diagnostics](macos-word-heading-import/README.md) preserve two failures: XML converter open could not establish owned identity; the direct-file variant failed while reading an undefined native return variable. Separate empty inventories, input hashes and guarded quarantine recovery close cleanup only. XML retries are stopped. The distinct DOCX input-materialization probe now passes23 native checks, including actual Word serialization, complete named style properties and reopen. Recipient import subsequently passes exact replacement, complete style/XML and effective-tab checks, then fails the fresh mixed-script paragraph/font assertion. Source-independence, final PDF and reopen remain unaccepted; this does not establish complete heading parity.

A read-only font-anchor diagnostic additionally finds direct Arabic and emoji font overrides on the source seed. Source style, whole paragraph, and single-character getters are therefore distinct observations; four empty mixed-range names cannot be accepted as complete appearance equality. The controlled fresh-order trial A saves output then fails strict style preservation: Default Paragraph Font loses semiHidden; B is not executed. A distinct no-reset empty-carrier B completes23 checks and preserves all named definitions, but Word adds an emoji family override and a CJK hint, so the strict style-only verdict remains false. A same-text source-style control, source-independence challenge, final PDF and reopen remain pending.

The horizontal-line document-container variant had already failed with a floating rectangle instead of the required inline object. Its stale compile-only plan has been corrected to the existing native evidence; the same constructor will not be repeated as a new test.

Temporary localhost certificate trust, test private key and listener were removed after the earlier Office.js activation attempt. No functional Office.js path was accepted. No merge, parity tag/release or personal-plugin installation has been performed during this continuation.
