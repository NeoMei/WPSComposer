# Mac Word structural table position guard

Native constructor diagnostics showed that a requested nonterminal table can be appended or inserted after the wrong paragraph. Until a constructor fulfills the frozen position contract, Mac Microsoft Word rejects structural table insertion with any position other than omitted, `None`, or `end`.

The public edit guard runs before opening/attaching or output preparation; the existing-session guard checks the complete batch before its first write. The shared direct-operation validator applies the same rule. Windows and WPS behavior remain outside this guard. Precise insertion is still required for parity.

Independent review found one initial P2: explicit operation overrides supplied through `patches` were absent from the before-open check. The repair normalizes patches plus ops once and shares the batch across routing, guard and execution. Both original reviewer reproductions are unchanged and now pass, along with 318 related cases (320 total).

Implementation, initial RED results, original review, repair RED/GREEN results and independent re-review are retained separately in this directory. Full37 validates the reviewed source overlay; hosted CI06 validates the earlier runtime candidate and must not be cited as guard validation. No native middle-table success is claimed.
