# Semantic table run-01: stopped at native constructor

2026-09-09. One authorized bounded native execution; no retry, repair, post-failure native inventory or close was sent by this agent.

Preflight: four frozen source hashes matched; independent Word inventory was []; no quarantine marker existed. Execution used a separately cloned Git checkout from full32 (actual HEAD415c9d04f38e5d4a744baadce75a39bf76dc8cca), with only the reviewed fixture/test overlay. See isolation.json and preflight/.

Confirmed native ACKs: seed(range12..19, one original table), sentinel-create(name文档91, Word16.112.3), select-middle. The semantic-table batch then failed with Word error-2710: cannot make class table. Native log offsets3078:3191 identify line34, the original primitive's make-new-table command for4rows/3columns at the noncollapsed12..19 range. There is no style/merge/final table ACK. Failure does not prove that no native mutation occurred.

The fixture stopped and quarantined immediately. Owned document remains unclosed/unverified at `/Users/neomei/Library/Containers/com.microsoft.Word/Data/tmp/wpscomposer-session-1qyy66n5/document-22111abb5db84f10a3267bdfcf23e892.docx`. Synthetic unsaved sentinel文档91 is retained. Its exact preimage and token are in report.json. The active lock quarantine was copied verbatim to native-word-lock.quarantine.json. Root received the Word lease with an explicit reconciliation requirement.

Saved before.docx, its retained private counterpart, all runtime AppleScripts/logs, exact source files, local dictionary, original report, uncertainty marker and isolation/preflight data are retained. The before artifact, live-versus-retained sources and dictionary hashes remained unchanged. after.docx/PDF, native reopen, XML/PDF acceptance, sentinel cleanup and final inventory were not reached. No native acceptance or parity claim is made.
