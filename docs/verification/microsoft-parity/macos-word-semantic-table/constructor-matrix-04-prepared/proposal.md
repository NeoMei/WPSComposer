# Single explicit-range conversion proposal — offline, not executed

One new owned copy of the unchanged seed; one noncollapsed 12:19 conversion attempt only. Change exactly one command from the approved Matrix03 A script:

`set diagRange to text object of diagSelection`

becomes:

`set diagRange to create range boundDoc start 12 end 19`

The mutation remains `set diagTable to convert to table diagRange number of rows 4 number of columns 3`. Selection, activation order, exact owned/window/active/selection/sentinel guards, main story and 12:19 range assertions, marker assertion, error catch, result observation, and preservation checks remain identical. A byte comparison of the full generated scripts verifies this single-line difference.

The specific hypothesis is that an explicit document text range avoids the selection-derived reference named in Matrix03's -1708 error after the native conversion mutated the document. The dictionary's text-range direct parameter and the independently observed 1×3 table justify testing the reference representation, but do not prove why the previous call failed. This is not an unsupported-command diagnosis, not a retry of the dirty Matrix03 document and not the old unexecuted B case.

Read and validate the returned table and a unique matching document-table ordinal using exact bounds, actual dimensions and text. Retain table-count delta and actual returned dimensions separately from the requested 4×3 result. A 1×3 result can be complete diagnostic evidence but must keep requested_dimensions_observed=false. Also retain marker inside/outside the new table, prefix/suffix/old-table bookmark text and every document-table record. Existing strict ACK gates reject foreign/mistyped/missing observations, unmatched returned objects or changed surrounding content.

One fresh isolated source checkout and one fresh owned seed copy are required for eventual execution. Only known -2710 is an ordinary diagnostic catch; -1708 or any other error, timeout, missing ACK or quarantine stops immediately with retained source/scripts/logs and hands the lease back. Even if the same error repeats, no additional variants or automatic retry are authorized. Save/close only the task copy after complete ACK and preserve the exact sentinel through cleanup.

Preparation: 29 tests passed (4 new explicit-range, 18 conversion, 7 Matrix02), following an initial 4-failure RED. Full generated script osacompile exit 0; compile only, no AppleEvents or native execution. No public/production/index changes or runner were added. The original Matrix03 report/manifest remain unchanged; `../constructor-matrix-03-parent-finding.md` separately points to parent's actual mutation and cleanup evidence.

Frozen hashes and full source/script/test snapshots are in preparation.json and retention-manifest.json. Root review and new lease authorization are pending.
