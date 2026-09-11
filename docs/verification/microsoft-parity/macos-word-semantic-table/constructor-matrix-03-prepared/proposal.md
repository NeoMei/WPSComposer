# Explicit text-range conversion proposal — not executed

Propose at most two new owned copies of the unchanged run-01 `before.docx` seed. Both activate the exact owned window, select the main-story range, verify selection/document/window/active-document identity and synthetic sentinel, and perform exactly one target command:

`set diagTable to convert to table diagRange number of rows 4 number of columns 3`

| Case | Operative text range | Specific hypothesis |
|---|---|---|
| A | noncollapsed 12:19 containing REPLACE | Conversion operates on the selected characters, creates a table starting at 12, and places REPLACE inside a new cell. Cell text can be overwritten by the later semantic table writer; retaining REPLACE inside this table is acceptable diagnostic evidence. |
| B | collapsed 12:12 | The same text-range command may create a table at the exact insertion boundary without consuming the following REPLACE text. |

The local Word.sdef `convert to table` (code sTXTwCtT, around line 8682) explicitly describes conversion of text within a text range, declares its direct parameter `text range`, accepts explicit row/column counts, and returns a `table`. The dictionary does not state a nonempty-range precondition. This makes collapsed behavior a bounded open question, not documented support. No separator, styling, autofit, selection-order, or constructor-property variants are introduced. Both variants compile; neither has run.

Reused Matrix02 command construction retains its exact owner/sentinel guards, known -2710-only catch, initialized success variables with separate caught-error bindings, full before/after body and every table's bounds/dimensions/text. New observations also capture the returned table's bounds/dimensions/text, exact body text before/after that returned table, and post-conversion prefix/suffix/old-table bookmark text. Evidence separately assesses start-at-12, marker-inside-new-table and marker-outside-new-table. No successful constructor can silently become a preservation or semantic-parity pass. In A the desired marker placement is inside; in B the existing marker should remain outside.

The preparation validator accepts complete typed diagnostic evidence, including evidence showing incorrect behavior. It rejects missing surroundings, bool-as-integer bounds, foreign bound paths, mismatched result status and unrecognized errors. Native ordinary -2710 is the only catch currently authorized by the reused diagnostic contract; all other native errors, timeout, missing/mistyped ACK, or quarantine stop without another mutation or cleanup/recovery attempt. Any eventual runner must use a new isolated checkout, fresh owned copy per case, source/SDEF hashes, retained raw scripts/logs/DOCX/XML, per-step ACK, exact guarded save/close/sentinel cleanup and parent-owned lease. This preparation intentionally adds no runner or public forward.

12 conversion preparation tests plus 7 Matrix02 tests passed (19 total). The initial missing-preparation RED was 6 failures before implementation. Both full generated scripts compiled with osacompile exit 0. Compile is syntax evidence only. `preparation.json` freezes source/script/SDEF hashes; source snapshots, scripts, logs, tests and the verbatim relevant SDEF command are retained here. Root must review this proposal and explicitly authorize any native run.

No production changes. The frozen pure semantic primitive and main semantic fixture remain unchanged. Full width/font/border/header/split/merge/terminal-CR/PDF/reopen/UI and public middle-insertion gates remain outstanding beyond this constructor diagnostic.
