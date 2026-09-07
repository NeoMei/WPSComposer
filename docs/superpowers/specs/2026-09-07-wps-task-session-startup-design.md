# WPS task session startup design

Approved scope: user approved the first-stage proposal in the current task on 2026-09-07 and requested implementation. This changes macOS generation startup and completion UX; it does not implement a persistent daemon or dual-format publication.

## Outcomes

A Word operation starts only its required local profile service and one owned isolated WPS host. Runtime startup must not invoke wpsjs debug (which launches WPS itself on macOS). M5 generation opens a task-owned blank DOCX, claims that exact document after registration, and renders into it instead of opening/closing wpsDemo.docx and creating another document. Native output remains staged, validated and atomically published. Interactive use opens only the final successful output after all owned runtime cleanup.

## Boundaries

- Python >=3.9; package remains importable without WPS/pywin32 on non-Windows hosts.
- Keep public generate()/convert_to_pdf() return values and overwrite/atomic-publication rules. Add an optional keyword-only open_result=False for explicit desktop presentation; interactive skill guidance uses True when the user expects to see the result. Existing unattended callers stay unchanged.
- Only explicit open_result opens a final artifact, after all native cleanup; opening failure must not report a published artifact as generation failure. Surface a warning while still returning its path.
- Keep the existing 600-second generate deadline, capability authentication, registration snapshot/restore, fixed component ports/origins and PID identity checks. Do not reuse an arbitrary user WPS process for generation/conversion.
- No broad process termination, no modifications to unrelated user documents, no deletion of pre-existing staging sessions or history entries.
- Only requested artifact formats are publicly returned. The separate both-formats feature is deferred.
- The known numbered-heading style mismatch is outside this change. Preserve source-derived fonts/numbering/style behavior and record it during native acceptance.

## Profile service

Introduce a focused Python static-profile service, binding only 127.0.0.1 on the selected component port. It serves only the selected generated profile directory, with correct HTML/JS/JSON MIME, no directory listing, no traversal/symlink escape, no writes, no request data logging and no-store caching. It launches no WPS processes and changes no WPS registration. Runtime remains responsible for profiles, registration and activation. Services must be joined/closed on all exits. ProbeRuntime accepts an optional set of components (default all for backward compatibility), validates it before side effects and runs preflight/profiles/registration/services only for those components. Real Word callers pass writer; other generation/conversion paths pass their requested component. Keep templates dependency until all pinned template consumers are migrated; do not claim npm dependencies removed.

## Owned task document

Ship a small native WPS-created blank seed as a versioned resource with pinned SHA-256. It contains no visible text, fields, macros or external relationships. Root controller creates the seed through WPS and verifies it before task implementation. Per-session private copies have task-related safe filenames, never wpsDemo.docx. A new activation_document parameter can select the private seed copy. It must resolve inside the active private staging root, be regular, belong to the selected component and be registered before launch. Activation remains once per component per runtime.

After authentication, bridge bootstrap distinguishes an activation document retained for generation from a disposable legacy activation fixture. Retained document identity is resolved by exact canonical path and ownership, never ActiveDocument alone. M5 command may request the authorized activation document for its first generation; it is rejected on mismatch/missing ownership, never falls back to the user's active document. Other commands keep their existing behavior. The real generated document is saved to its private output and closed. A full relayout within the same task uses a fresh owned document, without starting another WPS host. Disposable fixture cleanup enumerates documents by exact path so a focus change cannot cause a missed close; it never closes another document.

## Display

Presentation is a separate final step after atomic publication and adapter/runtime close. open_result=False remains the Python default. open_result=True uses the platform default application (macOS launch can prefer background startup during generation as a best effort; no promise of zero window flashes). Skill docs distinguish library automation from interactive delivery. Test API compatibility and no opening on any pre-publication failure.

## Verification

Platform-independent tests cover loopback service serving/denials, component scoping, one activation, seed/path ownership, retained document claim, exact fixture cleanup after focus changes, relayout, final display ordering and publication failure. Run full pytest with pinned npm templates installed. Native acceptance runs sequentially: generate without final opening; repeat generate; DOCX and standalone PDF conversion; final-output opening; coexistence with one disposable unrelated unsaved WPS document. Verify target artifacts, native numbering/content, absence of wpsDemo windows and new demo history entries, closed task services and restored registration, and preservation of the unrelated document. Screenshots and artifacts go under build/task-session-startup; native success is separate from tests and release. No merge, push or installation is included in this approval.

Writer activation uses the new native blank seed in every route, including standalone conversion; only M5 generation retains it. Other components keep their existing seeds in this Word-focused stage.
