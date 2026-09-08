# Native Word production acceptance — 2026-09-08

Release status: **PENDING Windows production acceptance and final integration checks**. No release tag or package publication has been made. Existing 0.8.1 metadata is unchanged.

| Gate | Evidence | Result |
|---|---|---|
| macOS Microsoft Word public DOCX / conversion / direct PDF | macos-word/report.json and original artifacts | PASS, 5 pages |
| macOS WPS existing public flow | macos-wps/report.json and original artifacts | PASS, 4 pages |
| Real edit → Undo → explicit save → close → reopen | each platform ui-report.json | PASS |
| Word unsaved document preservation | macos-word/preservation.json | PASS |
| Word native images, captions, REF / SEQ | macos-word/media-report.json, transparent-report.json | PASS |
| Word timeout, quarantine, refused unsafe recovery and successful explicit recovery | macos-word/timeout-recovery.json | PASS |
| Legacy DOC conversion with source preservation | macos-word/legacy-doc-report.json | PASS |
| Portable integrated suite | build/msoffice-production/pytest-native-integrated.log | 2795 passed, 12 native-gated skips |
| Windows production public API, native UI and ownership | remote production evidence required | PENDING |
| WPS missing bootstrap recovery tabs | investigation of six visible missing-file tabs | PENDING |

The runner records the Git base and source SHA256 values for the actual dirty working tree used; these source digests, not only the base commit, bind each native report to its implementation. Artifacts are unmodified native outputs. macOS Word theme colors and pagination may differ from WPS; the shared typography, numbering, content and A4 guards all pass.

The UI reports describe observed CUA interactions. WPS records one undo step per typed character, so all inserted characters were undone before saving. The saved WPS copy changes package metadata, but its document text and native typography checks match the original. Word saved copy is byte-identical.

Reproduce with `python fixtures/verify_msoffice_production.py --output-root NEW_DIRECTORY --fixture representative --engine msoffice --timeout 600`; use `--engine wps` for the regression. The runner does not certify native UI or Windows by itself. Operator support boundaries and recovery instructions: [native-word.md](../../../skills/WPSComposer/references/native-word.md).
