# Final Excel delete-selection and fixture regression

Full40 passes **5,046 tests / 12 skips** in 215.26 seconds from immutable snapshot `83c39eff217530f1b089c5583dda599c0b440499`, tree `e05443f98ff6effeb89b0cd6ca59ccd1b7f57c04`, parent `5622a1615e6fd22676a3aa410a393503dab57f91`.

All 386 source/configuration hashes are unchanged before/after and equal to live at completion. This includes the reviewed Excel selection repair, exact sheet-sequence assertions, marker-guarded sentinel recovery and two portable failure-injection cases. Independent fixture re-review has no remaining scoped findings; strengthened native run02 passes all three delete/write/save/reopen cases with parent OOXML checks and owned cleanup. Separate UI evidence binds the unchanged production implementation.

Raw full39 evidence remains historical and unchanged except its explanatory scope note. Hosted CI08 applies to the parent; a hosted run on the next candidate is separate. The 12 skipped gates do not certify native acceptance. Full Microsoft/WPS parity remains incomplete.
