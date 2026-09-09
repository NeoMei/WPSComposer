# Native notices run-02: table passes native, OOXML fixture false negative

2026-09-09. Scoped review 6ef75927 and root exclusive Word lease authorized run02.
Before launch checked exact frozen hashes:

- helper: 8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7
- tests: 07396fb846503c098e132e21dd009f90c62544f702636cbd096063fb750a651b
- fixture: 39e1ab6b0a2fe7a7ed180b0d8c4845a2d76a269e92d8f63058b0ffbfab3fa8b6
- current session: f8b2e89a4aefddb27158fefa958da967754e2e52c8e7d02491a05fa35faa4170

Fixture saved complete designated source copies/hashes in run02; all were rechecked
after execution and match current bytes. No source changed during this run.

## Actual outcome

Executed the guarded fixture --execute into new
`docs/verification/microsoft-parity/macos-word-degradation/run-02`.
Overall FAILED, 20 recorded checks passed:

- inline-empty and inline-nonempty passed all nine checks each.
- block-empty native table creation succeeded with returned table_index=1,
  native table range 0:27, actual Text ending in both CR/BEL terminators.
- Native display range 0:25 readback exactly matches `[NOTICE_BLOCK_EMPTY] 中文😀`;
  italic, dark red and light red shading all true.
- Default following paragraph readback at native 28:56 has italic=false,
  notice-red=false and notice-background=false. No explicit override was used.
- Native DOCX and PDF were produced. Read-only saved OOXML check raised
  `AssertionError: Notice shading missing` before reopen/PDF-content gates.
- Remaining three cases were NOT RUN because fixture fails fast.

## Proven fixture defect

Saved block-empty/notices.docx has both:

```xml
<w:tblPr><w:shd w:val="clear" w:color="auto" w:fill="FCE8E6"/></w:tblPr>
<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FCE8E6"/></w:tcPr>
```

Notice runs have expected italic and 9C0006 text but no redundant run shading.
The fixture currently insists on w:rPr/w:shd for every notice character. Native
range shading applied to the entire single cell serializes as table/cell shading,
so its assertion rejects correct persisted formatting. Production helper was not
modified. Needed correction is a read-only effective shading lookup respecting
explicit run/paragraph/cell overrides, not merely checking any ancestor color.

## Failure cleanup

No quarantine. The owned context exited and its private document was absent.
Separate cleanup verified the sole retained sentinel 文档26 against initial
name/path/unsaved/content-hash facts and exact unique token, then closed only that
sentinel using the proven helper. Final native inventory []. Evidence is
run-02/block-empty-cleanup/report.json plus raw inventory/close scripts/logs.
Original failure report, failure.txt, copied sources, native runtime and artifacts
are unchanged. Live cleanup state is recorded separately from failure-time
retained_sentinel metadata.

Native lease is released. Waiting for root authorization/review of fixture-only
fix and a new run03; no rerun, production edits, personal install, commit or push.
