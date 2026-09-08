# MS Office native Word/PDF feasibility spike

Approved scope: user accepted the preceding architecture assessment and asked to continue the Word/PDF dual-platform prototype. This branch is a feasibility probe, not public engine integration or a release.

Evaluate Microsoft Word on Windows through explicit Word.Application COM and on macOS through native AppleScript. Reuse a common acceptance intent: Chinese body/font/indent, hierarchical headings and native numbering, TOC/fields, multipage table, save/reopen, native PDF export, and document ownership cleanup. Produce machine-readable observations and inspect native DOCX and PDF. Do not label dictionary presence, compilation, mocks or prepared Windows scripts as native acceptance.

Use only task-created documents. Preserve pre-existing open/unsaved user documents; do not quit existing apps, change Normal templates, global macro security, install add-ins, or use WPS for any MS Office acceptance output. Mac VBA is a fallback feasibility option only, not authorized global macro configuration. No public API or production source changes. Python 3.9+ for reusable scripts. Retain raw errors and mark unsupported capabilities explicitly. Fresh output directories prevent stale pass artifacts.

Deliverables: portable prototype scripts under fixtures/msoffice_spike; an evidence report under docs; local raw artifacts under ignored build/msoffice-spike. A Windows native result is conditional on an accessible Windows desktop with Word; absence must be stated. Recommend native adapter versus Office.js from observed gaps, not presumed parity.
