# Offline Word Office.js probe package report

Status: **PREPARED FOR INDEPENDENT REVIEW — NOT INSTALLED OR NATIVELY VERIFIED**.

Scope is confined to this directory. No production/index changes, app writes, server/browser/native launch, certificate trust, security changes, add-in enablement, snapshot invocation or document mutation were performed. The package intentionally requests ReadWriteDocument because Word-specific reads require it; source-level read-only behavior is not a permission restriction.

Read sources: existing `officejs-word-gap-feasibility.md` and `docs/superpowers/specs/2026-09-08-microsoft-parity-officejs-supplement.md`. Current official Microsoft sources were checked through the web tool on 2026-09-09; direct links and specific applicability are in README.md. This creates only the capability/binding candidate pane; bridge, registry and snapshots from the broader supplement are excluded.

Included: Word-only manifest template and concrete HTTPS-loopback rendering; closed local configuration renderer; minimal HTML/CSS/UI/core JS; pane-scoped random identity and lifecycle retirement; actual diagnostic/requirement reporting for all nine gaps plus snapshot sets; optional gated saved/selection read; manual JSON copy/local export; precise future Mac/Windows enable and coordinated removal steps.

Evidence: initial RED failures retained. Twenty Node stub/UI tests pass and six Python manifest/source-constraint tests pass. JavaScript syntax checks pass. No dependency installation. Host/API tests are stubs; manifest tests are bounded structural checks, not Microsoft full-schema certification. No blanket success or support declaration is emitted. All nine gaps remain nativeAccepted=false and snapshot executed=false.

Reviewer focus: ReadWriteDocument wording in manifest/UI/docs; pane ID is page-lifetime only; runtime values never manufactured; host/platform and feature gates; one pending read and timeout retirement/late-ACK rejection; command-free/mutation-free code; no snapshot calls/transport/persistence; export does not include document text/path; explicit setup/removal implications; missing native proof remains visible.

Exact retained files and SHA-256 values are in HASHES.json. Frozen implementation includes probe.js, taskpane.js, taskpane.html, taskpane.css, manifest.word.xml.template, manifest.word.xml and configure.py. Tests and all RED/GREEN evidence are also hashed. Parent decides any retained production home or further installation approval.

## P3 independent-review correction

Corrected the diagnostic key `add_equation_fallback` to the frozen `add_equation_native_fallback`. Added an exact sorted nine-name regression before the fix: new RED log records 20 passing / 1 failing test; after the single-key change, 21 Node tests and 6 Python tests pass. The old RED/GREEN logs remain unchanged; new evidence is retained under `evidence/review-key-*`. No other implementation behavior or native action changed. Re-review is pending.
