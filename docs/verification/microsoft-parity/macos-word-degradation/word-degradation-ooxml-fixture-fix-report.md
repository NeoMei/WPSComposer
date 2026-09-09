# Fixture-only effective shading correction ready for review

2026-09-09. Root independently inspected run02 XML and authorized only the fixture
and tests to change. Production helper remains unchanged at
8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7.

The read-only OOXML gate now resolves shading in order: current text run, its
paragraph, its enclosing cell, its enclosing table. It uses the actual ancestor
chain and stops at the closest table. Explicit declarations terminate inheritance,
including missing/auto/other fill or nil. Native val=clear with explicit FCE8E6 is
valid; clear without explicit fill cannot inherit an ancestor's FCE8E6. The normal
following paragraph also checks its effective shading rather than only direct
properties. An unrelated table cannot supply notice shading.

RED added 8 cases: real native cell/table form, table-only inheritance, run auto,
run clear-without-fill, other run fill, paragraph auto, cell other fill, unrelated
table. Observed 2 failed (valid native forms), 6 passed, 74 deselected in 0.60s.
GREEN: complete degradation module 82 passed in 1.22s. git diff --check clean.

Read-only reinspection of the actual failed-run DOCX now passes the corrected gate:
1 table, 24 notice Unicode characters, expected normal tail, block paragraph check.
Artifact bytes unchanged before/after:
3f0d8bd6f7fa7ecdda88ccbced238acc3c74727418b85a268c23bf9d677fd80c.
This does not relabel run02 as passed or substitute for remaining native cases.

Frozen review hashes:

- fixture: 8125aa442ddae81d7daed4b8421058d5f50de6953d12b63e346dafc525a7c7ed
- tests: 0657a1bf7ba21fd37549b6454bad728b7fac5ca36f3299e7b8f790b14417d5b6
- helper: 8594af7ba2c3f78bf126f17a984a52ce19374491646a43f289595ac037b6fcb7

No native run03 has started. Native lease remains released with verified empty
inventory after run02 cleanup. Await root review and explicit rerun lease.
