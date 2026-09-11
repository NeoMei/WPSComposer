# Excel delete-selection regression

Full39 passes **5,044 tests / 12 skips** in 220.26 seconds from immutable snapshot `ed0da48783daaa47eb9389bfd722b86fd5f2d0ae`, tree `33cc7fc30bfb2ed819563aee8cac3e49b5e064fb`, parent `5622a1615e6fd22676a3aa410a393503dab57f91`.

All 385 source/configuration hashes are unchanged before/after and equal to the live candidate at run completion. The snapshot includes the Excel selection repair, four regression cases and initial opt-in native fixture. Subsequent fixture review repairs and their two new fault-injection tests are outside this snapshot; full40 covers those changes. The independent scope review, three native write/save/reopen cases and real UI editing/Undo evidence are in `../macos-excel-delete-selection`.

The later `.gitattributes` additions cover only retained evidence byte/whitespace handling and are outside the runtime/test hash set. Hosted CI08 certifies the parent, not this new repair. No native skip is a pass, and full platform parity remains incomplete.
