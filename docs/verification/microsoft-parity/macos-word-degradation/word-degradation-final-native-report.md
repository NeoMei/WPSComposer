# Two-method direct Mac Word native acceptance: run03 PASS

2026-09-09. Root granted unique Word native lease after scoped helper/fixture
reviews. Executed only the frozen fixture; no production or test source edits.

## Verified native result

`docs/verification/microsoft-parity/macos-word-degradation/run-03/report.json`
reports all 6 cases passed and all 57 checks true. Word 16.112.3.

| Case | Passed checks | Saved DOCX SHA-256 |
| --- | --- | --- |
| inline-empty | 9 | `4a1fe22d3257344c652bfa8662ec8a7b53698b453b1b5946738a36d84946a03b` |
| inline-nonempty | 9 | `150d66fa31592f6167bdbe4bfbf93556667cad05b0842feec34dc4f26670680d` |
| block-empty | 9 | `cc4cf9c61d3b15219e134ab1d94717a55eeb61e2fc01fbde706921d9c91fa2e5` |
| block-nonempty | 9 | `58e68641d490d3b4495f746b0c5ca7ea2a8f990f3294066ace627eb7db4ac1bc` |
| table-failure-empty | 10 | `500117505ac690ea89e6beba4bc034fc1b85c94712aec980a26c2ef2f65cb39a` |
| table-failure-nonempty | 10 | `fb4d5839f3f56d9c4244ee88476359a5bb2294e17a574624c2edbef6928cbb22` |

The remaining 57th check verifies all designated source files unchanged during
the run. Each case's source, scripts/logs, runtime, native DOCX/PDF and private
readbacks are retained. artifact-manifest.json contains all six DOCX and six PDF
hashes and sizes. Rechecked those DOCX hashes against the fixture's post-reopen
hashes; all match.

Each case proved native UTF-16 handle/text/style, no unexpected partial content,
default following paragraph free of notice style, actual native DOCX save and
PDF export, saved OOXML formatting and row non-splitting where applicable,
PDF notice exactly once with normal tail, source-preserving public reopen,
reopened native style/readback, real edit followed by actual checkpoint rollback,
owned create/reopen contexts closed, unchanged unique unsaved sentinel, and
independent exact sentinel close. Every final per-case inventory is [].

Both failure cases created and populated an actual native single-cell table,
then emitted controlled ordinary failure. The exact observed event sequence was:

```text
partial-table-submitted
table-failed-ack
rollback-ack
fallback-submitted
```

The observer called real recovery module rollback and verified exact original
body/end/paragraph/table restoration before recording rollback-ack. No rollback
simulation or fallback-on-uncertain-remnants was used. Final artifacts contain
one notice, no partial-table marker, and zero residual tables for these cases.

## Frozen source evidence

Launch-time verified and post-run reverified against current source AND retained
copies, all 15 designated source files identical:

- helper: 8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7
- fixture: 8125aa442ddae81d7daed4b8421058d5f50de6953d12b63e346dafc525a7c7ed
- session: 3ab747a06c63deabc9e0dbf1de51857d8b7dfa843d6ea8e0822915e91e44f205

Command actually executed:

```text
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python fixtures/microsoft_parity/macos_word_degradation.py --execute --output docs/verification/microsoft-parity/macos-word-degradation/run-03
```

Exit 0, passed=true, error=null. Word native lease released to root after all
final inventories and hashes were checked. No source modifications, personal
installation writes, commits, or pushes. The root's concurrent full suite is a
separate gate and is not claimed by this worker report.

## Scope and retained limitations

This accepts the two methods add_inline_degradation and add_degradation_notice
through MacWordSession, including acknowledged real block fallback. It does not
complete the four quality-anchor methods, full Word parity, installed-plugin
validation, independent GUI edit/undo/save acceptance for these new artifacts,
or a release. Handles remain immutable native snapshots, not live COM proxies.

Earlier run01/run02 failures and their exact cleanup reports remain retained.
run01 proved required exact bound-window activation; geometry diagnosis proved
Word table terminator serialization. run02 exposed the fixture's run-only
shading check; fixture-only effective ancestor shading correction was reviewed
before this run. Those failed runs are not relabeled as successes.
