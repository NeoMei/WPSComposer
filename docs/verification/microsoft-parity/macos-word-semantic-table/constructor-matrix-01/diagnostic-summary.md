# Constructor matrix-01: three confirmed failures, active case ACK fault

2026-09-09. Four authorized diagnostic copies, each opened privately from run-01/before.docx and each given exactly one target constructor. No production file was changed. All work ran from an isolated clone; source hashes remained unchanged. No retry or extra native request was issued after the unknown fourth result.

| Case | Range | Requested table | Target window | Native diagnostic result |
|---|---|---|---|---|
| A |12..19 noncollapsed|1x1|inactive, exact sentinel active|Ordinary -2710 caught in-script; exact body unchanged, table delta0; owned close and sentinel preservation confirmed|
| B |12..12 collapsed|4x3|inactive, exact sentinel active|Ordinary -2710 caught in-script; exact body unchanged, table delta0; owned close and sentinel preservation confirmed|
| C |12..19 noncollapsed|4x3|inactive, exact sentinel active|Ordinary -2710 caught in-script; exact body unchanged, table delta0; owned close and sentinel preservation confirmed|
| D |12..19 noncollapsed|4x3|active boundWindow|No final ACK; diagnostic script fault -2753 while assembling the attempt row; retained/quarantined|

The first three observations rule out table dimensions or range collapse as sufficient explanations of the inactive-window failure. Local source review also found that existing successful table builders activate boundWindow before make, while the semantic primitive itself does not. Window activation is therefore the leading hypothesis; no production change or native acceptance follows from that inference.

D raw log: `variable diagCode is not defined (-2753)` at offsets5886:5894, in `set end of nativeRows to {"attempt",diagResult,diagCode,diagMessage}`. The diagnostic script used error-handler parameters with the same names as preinitialized result variables. This exposed an AppleScript handler-binding/scope issue on the non-error path. The failure offset is after the constructor try and subsequent bound-document/sentinel guards, but no table count/body/geometry ACK was returned. D must remain unconfirmed until root reconciles its actual owned document. It must not be silently converted to a successful case or rerun as part of this four-case allowance.

Owned D path: `/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-g6k8gzpf/document-798995b924bf4696abfc488e45e139b6.docx`. Synthetic sentinel:文档92; exact token, saved-state and text-hash preimage are in report.json. Both remain retained; quarantine marker copied to native-word-lock.quarantine.json. Word lease was explicitly returned to root immediately after the unknown result. Root must reconcile before another native run.

A/B/C retain native before/attempt/after/table observations, saved DOCX, extracted XML and exact runtime scripts/logs. D retains its original source-stage DOCX and failed scripts/logs, not an asserted saved post-constructor artifact. Root-level source snapshots, dictionary, isolation manifest, runner log and all original reports remain intact. The separate retention-manifest.json records local evidence hashes and confirms no post-unknown native followup.
