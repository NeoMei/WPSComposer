# Native Word release candidate

Status: unreleased. Version metadata remains 0.8.1 until Windows section-numbering correction and final merged native acceptance pass. This document is a release preparation record, not a release announcement.

## User-visible changes

`generate()` and `convert_to_pdf()` accept keyword-only `engine="wps"`, `"msoffice"`, or `"auto"`. The default is WPS. `auto` chooses an installed compatible engine before work starts and keeps that choice through native layout, PDF quality checks, relayout and presentation. Explicit WPS operations no longer silently select Microsoft Office on Windows.

The Microsoft backend drives desktop Word through COM on Windows and AppleScript on macOS. It supports DOCX/PDF generation and DOC/DOCX conversion to PDF through the shared M5 pipeline. XLSX/PPTX and active-document editing continue to use their existing WPS routes. DOCX presentation opens the selected application.

The native adapters preserve existing user documents, use private conversion copies, enforce total task deadlines and retain uncertain task files. macOS quarantine requires explicit verified recovery. Windows verifies the new WINWORD process and binds its application before document changes; a restored temporary Caption challenge supports versions without Application.Hwnd. Safe error codes identify unsupported capabilities and recovery locations without exposing document text or scripts.

WPS native built-in styles now receive the planned six heading sizes. Cover Title explicitly uses body outline membership so it does not reappear in the TOC. Windows style handling accepts the same outline-level keys. WPS cleanup keeps files when host exit cannot be verified, and a failed startup handshake closes only its identified task fixture. Shared-session inspection/editing retains uncertain files without turning a successful operation into a false failure.

## Verification available

- Candidate deb67ba: 2836 passed, 12 native-gated skips; isolated plugin installation, public signatures and installed source hashes passed.
- macOS Word and WPS: native public DOCX generation, conversion and direct PDF; cover, six heading levels, native numbering, repeated-header long table, source preservation, overwrite refusal, actual section page numbers and visible PDF review.
- Both local clients: actual keyboard edit, Undo, explicit save, close and reopen on identified test copies.
- macOS Word: native images, captions, REF/SEQ and indexes, unsaved-document preservation, legacy DOC conversion, timeout/quarantine/recovery, explicit rejection of unsupported equations.
- Windows: initial real API, UI, unsaved sentinel and process identity runs passed their original gates. Stronger parent review found section-numbering defects; correction and final native acceptance are still required. Initial PASS does not certify release.

Full evidence and raw failures: [production acceptance](../verification/msoffice-production/README.md).

## Capability limits

The first Microsoft release covers Word documents. It does not add Microsoft Excel or PowerPoint automation, Microsoft active-document editing, or cloud Office support. macOS Word explicitly rejects unsupported advanced operations such as native equations, merged/semantic-cell tables, landscape media, multi-column figures and unsupported custom style/page combinations. Native pagination and unspecified theme colors may vary by client; explicit semantic and typography requirements must still hold.

Installation, engine selection and failure recovery: [native Word guide](../../skills/WPSComposer/references/native-word.md).
