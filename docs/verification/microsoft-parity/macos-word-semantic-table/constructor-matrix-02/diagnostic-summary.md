# Semantic table constructor diagnostics — Matrix02

**Diagnostic complete; semantic range replacement failed in both variants. No production fix or new public capability is accepted.**

The two authorized independent saved-seed copies each received exactly one target constructor. Both selected the owned main-story range 12:19 after activating the owned window. Exact document, window, active-document, selection-document, and synthetic sentinel guards preceded the constructor; every result had a complete typed ACK. Only the make location differed.

| Case | Constructor location | New 4×3 bounds | REPLACE removed | Old 1×1 bounds | Outcome |
|---|---|---|---|---|---|
| A | boundDoc | 62:78 | no | 42:61 | appended after the old table |
| B | diagRange | 20:36 | no | 59:78 | inserted after REPLACE paragraph, before SUFFIX |

The native before-range was exactly 12:19 in both cases. Saved DOCX/XML confirm the body order (`xml-body-order.json` retains every body node, including empty paragraphs). REPLACE remains outside the new table. Prefix and suffix text and old-table text survive; these observations are not a full formatting-preservation assertion. Reselecting after activation does not fix the document-container insertion. Passing a range as the make location changes the placement but does not implement range replacement.

The local Word.sdef gives table.text object read-only access, but rows/columns are also read-only properties accepted at construction. Read-only metadata alone therefore does not establish why the initializer ignores text object. Previously successful generation/session examples construct at boundDoc with a document-end range; they proved append behavior, not arbitrary-range insertion. A later hypothesis should use an API whose explicit argument is the operative text range, rather than assume initializer/location semantics.

The frozen public session `insert` branch for type=table uses the same `make new table at boundDoc with properties {text object:insertionRange,...}` command shape at macos_word_session.py:609. This is a potential impact on public insertion positions; the public method itself was not executed in this diagnostic and its middle-position behavior has not yet been independently native-probed. Keep middle insertion and replacement in the future support gate. No public method or frozen semantic primitive was changed.

The diagnostic handler bug from Matrix01 is fixed only in the new diagnostic source: success-path diagCode/diagMessage are separate from diagCaughtCode/diagCaughtMessage error bindings. Both Matrix02 success paths emitted `created,0,""` in complete ACKs. Matrix01 original ACK failure evidence remains unchanged; root separately reconciled its appended table and performed owned cleanup.

Preparation: 52 tests passed (7 diagnostic + 17 primitive + 28 fixture); 5 pre-existing PyMuPDF/SWIG deprecation warnings. Both generated scripts compiled with osacompile exit 0. Tests cover appended success versus placement, marker replacement distinction, ordinary errors, collapsed interpretation, separate error bindings, closed variants and the one-line location change. A new-variant RED run failed 2 tests before the two-case implementation; GREEN passed 7 diagnostic tests. Full suite/CI are parent-owned independent gates.

Execution: Word 16.112.3, isolated checkout `/var/folders/0n/49qgdd8x7kgcvh719fw743mh0000gn/T/wpscomposer-semantic-matrix02-rfwf2v_5/checkout`, baseline 415c9d04f38e5d4a744baadce75a39bf76dc8cca plus retained reviewed overlays. Both private output copies were saved and exactly closed; sentinel 文档93 was checked and closed. Initial/final inventory are []; source seed, retained source files and SDEF hashes remained unchanged. No unknown result/quarantine occurred. Lease was returned to parent immediately after completion. No UI/PDF/reopen/full semantic acceptance is claimed by this constructor-only experiment.

Frozen matrix02 source SHA256: bc944340eafdeb60a086a3435d8543adac4c9b0e1601cfa2125b6c28a11bff6c

Frozen diagnostic test SHA256: ef69bf101854d27363b0b52ca7e22d7509ad2affcb0ea4c5d7f31de7a3a327e3

Artifacts: `report.json`, retained `source/`, each case's raw runtime script/log and `after.docx`/XML, `localWord.sdef`, `isolation.json`, `preflight.json`, `compile.json`, `pytest.log`, `execution-runner.log`, `xml-body-order.json`, and SHA256 `retention-manifest.json`. Compile-only prepared scripts and the full prepare logs remain in sibling `constructor-matrix-02-prepared/`.
