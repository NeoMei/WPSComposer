# Mac Word inline and block degradation implementation brief

2026-09-09. Approved parity plan Tasks 3/5; scope narrowed by root to exactly two
helpers. No session.py writes, native calls, personal installation, commit or push.
Read approved parity plan, word-degradation-brief.md, checkpoint brief, current
recovery/session/fields/references modules, WriterComposer implementations and
longform executor consumers, and regression guardrails.

## API and consumption

Module functions `add_inline_degradation(session, code, message, fallback_text)`
and `add_degradation_notice(session, code, message, fallback_text, placement="block")`.
The session forwarding methods are proposed as report text only. Exact equality
with "inline" delegates; every other placement retains block behavior. Message
is inert. Call the shared WriterComposer._degradation_display, including scalar
coercion, code normalization, redaction, no duplicate wrappers, and inline/block
wrapping. Reject only unsafe text (unpaired surrogate, control except tab/CR/LF,
more than 1,000,000 UTF-16 units) before native mutation.

Current semantic insertion is Word's terminal content end minus one, as used by
all existing append session methods. Inline adds no boundary paragraph or suffix.
Block also uses the exact terminal insertion point; it does not introduce a new
boundary paragraph in Python, matching the baseline's collapsed selection range.

Consumers: WindowsLongformExecutor and writer media fallback callers ignore these
returns. WriterComposer quality upsert reads the box's Range.End; bookmark notice
returns the box. Provide immutable session-tagged range snapshots with Start, End,
Text, plus box handles with Range and optional native table index. They are native
acknowledged semantic snapshots, not COM proxies, live ranges, or durable edit IDs;
no future mutation method will accept these snapshots in this slice. Table Range
uses native table bounds and actual text, while inline Range is the display extent.

## Native sequence and acknowledgement

Reuse existing session exact document binding/deadline via _execute, with explicit
_mutation_preflight before submission. Do not call _execute_structural or
_business_commit: both set state before content acknowledgement. On accepted final
ack only, clear pending heading, mark structural change, invalidate observed field
topology, and allocate the return snapshot's session identity. Errors retain
private evidence; uncertain transport, malformed ack and cancellation quarantine.
Tracking arrays are untouched by notice insertion.

Inline: append shared display at terminal native point with UTF-16 extent; exact
text and range bounds plus italic/#9C0006/#FCE8E6 readback are mandatory. No inserted
CR and no UI selection/reset of surrounding paragraph style.

Block: obtain real recovery checkpoint, construct one table at its exact collapsed
coordinate, fill one cell, apply the shared style and paragraph 0/3 spacing,
keep-together/body outline, disable row page split; acknowledge native table count,
1x1 geometry, complete table range, exact display subrange, and style readback.
Catch ordinary native construction/style errors within the batch and return a
closed table-failed ACK. Never include private error text in public values.
Malformed success/failure ACK, escaped transport error, timeout, stale binding,
or quarantine cannot trigger fallback.

Only a valid table-failed ACK invokes the existing real rollback(session, token).
That method validates exact native preimage and acknowledges restoration before it
returns None. Check quarantine/deadline again, then append one display-only range
with block style. No second checkpoint needed: failure in fallback is fatal and
quarantines, allowing no layered recovery. Rollback's own acknowledged state update
is legitimate; failed or malformed insertion ACK cannot otherwise update state.

## Test and native boundary

Write tests first against public module entry points with controlled native ACK
transport; use real MacWordSession preflight and state. Exercise shared display,
UTF-16, returned handles, strict ACK types/bounds/style/text, failures/cancellation,
no state changes before ACK, exact placement, unsafe inputs before execution, real
recovery checkpoint/preflight/post-rollback ACK sequence for table failure, and
rollback failure preventing fallback. Native script construction has no execution
permission in this slice. A pending probe script may be written for root, with
single-cell geometry/formatting/final-CR and subsequent text style readback listed
as unresolved until native review and serialization permit actual execution.
