# Mac Word direct references and bibliography evidence

2026-09-09. Five-method slice implementation and bounded native acceptance on Microsoft Word 16.112.3 for macOS. **Run-04 passed.** No complete parity, Windows acceptance, controller checkpoint recovery, package installation, or release is claimed.

## Implemented surface

Exact WriterComposer signatures are exposed on MacWordSession for `add_bibliography_native`, `add_bibliography_legacy`, `add_cross_reference_paragraph`, `add_citation_paragraph`, and `add_cross_reference_fallback`. Implementation uses native Word commands and owned-session ranges. The public direct-method baseline remains 6dd3a00 and writer.py SHA-256 b9dbcc140eb38ffa9b1a77dc0462645356eda720a184c1e4013f88fee9878205.

Bibliography preserves caller number ordering, scalar display conversion and arbitrary bounded paragraph geometry; legacy strings are literal and are never numbered again. Plan-only metadata and style compatibility inputs remain inert. Reference bookmarks use the baseline helper's direct syntax; arbitrary safe affixes/list indents remain accepted. Citation text stays static. Each planned unresolved occurrence returns its own issue; controller fallback returns None without a duplicate issue.

## Evidence sequence

| Run | Result | Proven result or failure |
|---|---|---|
| [run-01/report.json](run-01/report.json) | Failed, retained | Both bibliography variants native-acknowledged. Empty-prefix readback rejected Word's native `missing value`. Exact-path read-only diagnostic confirmed the cause; root performed exact owned/sentinel cleanup and guarded quarantine recovery. |
| [run-02/report.json](run-02/report.json) | Failed, retained | Empty affixes and three-field native REF paragraph passed. Implicit tab collection iteration failed -1708. Read-only count/index diagnostics proved indexed access and the custom 24 pt tab. Root exact recovery is retained. |
| [run-03/report.json](run-03/report.json) | Failed, retained | Bibliography, true REF update, both list formats, local partial-mutation rollback and occurrence styles passed. Exact static suffix readback caught inherited red italic shading. Guarded cleanup completed; independent root native document-count evidence retained. |
| [run-04/report.json](run-04/report.json) | Passed | Deferred exact-range styling preserved static prefix/suffix and following normal paragraph. All five direct methods passed the bounded native fixture, saved DOCX/PDF, exact native reopen checks, unchanged source digest and zero leftover native inventory. |

All original run directories and report contents were preserved. These are byte-identical copies from the task's private SDD evidence tree, including failed scripts/logs, strictly read-only diagnostics, recovery records, per-run dependency source copies and raw artifact reports. `evidence-copy-hashes.json` binds all 278 copied original files. Original report locations/names inside recorded payloads have not been rewritten.

## Run-04 acceptance

- Two structured bibliography entries `[3] 中文😀 已引` and `[9] 未引`, with native 21.5 pt left indent, -12.25 pt first-line indent, 0/4.5 pt spacing and keep-together; legacy already-numbered text and 7/False/None remain literal.
- Three native SEQ-result bookmark targets for figure/table/equation and four true REF fields. Controlled target-number change refreshes results to 8/11/13/8. Emoji/CJK and suffixes remain outside field results.
- Bullet indent 24 pt and ordered indent 31.5 pt, exact negative first-line indents, native custom tabs, one-and-a-half line spacing and 0/3 pt spacing before/after.
- Two separate unresolved occurrences with distinct owner nodes and exact already-coded text; each precise substring is italic #9C0006 with #FCE8E6 shading.
- Static fallback prefix/suffix remain unstyled, fallback range is styled, and the next native paragraph is plain. The latter also has independent read-only DOCX run-property evidence in `run-04/post-fallback-paragraph-ooxml.json`.
- Injected native partial append plus field-creation error is deleted exactly once before one local fallback and one issue; invalid direct bookmark causes zero document mutation.
- Actual save, close, native Word reopen, native REF/list/style reinspection, no source digest change, exact unrelated/unsaved-sentinel preservation and post-close full native inventory equality.
- PDF text read and one-page raster inspection completed. The rendered acceptance page shows readable references/emoji, no clipping or overlap, correct indentation, two styled unresolved literals and bounded static fallback shading. This short fixture does not prove multi-line wrapping across pages.

The real WindowsLongformExecutor recovery chain is **NOT RUN**: Mac degradation checkpoint/rollback methods are a separate eight-method follow-up. The five-method fixture does not fake them. Controller propagation unit tests are not evidence of complete executor recovery.

## Source freeze and tests

The final targeted command is:

```bash
PYTHONPATH=.:/tmp/wps-spike-validator-audit-deps ../wps-task-session-startup/.venv/bin/python -m pytest tests/msoffice/test_macos_word_references.py tests/msoffice/test_macos_word_fields.py tests/msoffice/test_macos_word_business_followup.py tests/msoffice/test_macos_word_session.py -q
```

**217 passed**, with all new builder branches compiled by osacompile as syntax-only evidence. Initial and defect-specific RED logs are included. The final focused test module has 47 tests. Full-branch testing and independent review are root-coordinated gates, separate from these results.

| Final source | SHA-256 |
|---|---|
| macos_word_references.py | f2243a6ce644f6a196db087756b8dba51d6d9bb713c586a9f0175819511200f6 |
| macos_word_session.py | 39129d22315eb5701dde8f1f5cfe305de404d07ce538d10d44b2fa1f562d502a |
| macos_word_fields.py | c390342553fa4f9c30f77dc751da99e97ffd0abc1ff0e20822f9c5c974fa776d |
| test_macos_word_references.py | b3afaab28d4d7288ddda3cb7073dac1d9de5a74153e54af82e65c8c08bbf0d6c |
| guarded fixture macos_word_references.py | 6b429c1caf7be8305706a054d9312fbfc241311794c04f1d925c459ffa63a1b0 |

At handoff all twelve files in run-04's source manifest match the working tree, including artifact_transport, runtime, compiler, errors and executor dependencies. Each source copy in the report's `source/` tree matches its recorded digest. The report is source-bound slice evidence, not a substitute for the final whole-branch gate.

Artifacts: DOCX SHA-256 `2ae124e28bd81c0ec4b453e77630052efc29270ac514df235d804f5c483c25d9`; PDF SHA-256 `f40e2be8bcf50aeafa72f3b7f38ccd71e6d17b840316e67cee6f8d33ff774aa2`.
