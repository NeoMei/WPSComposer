# PID-bound OSAKit transport native checks — September 10

This checkpoint verifies the new internal `macos_osa_transport.py` at SHA-256 `35b85f24b97ee29917966efa2f6295c4198c696630c39e05bb4127db3a893b1e`. It is not integrated into the public Excel session lifecycle yet. Native launch and transport substitution remain test scaffolding. Existing-sheet deletion remains guarded in public code.

The read-only test binds the original Excel PID by birth token, resolved executable and bundle, routes native Apple events, and serializes Foundation JSON. Both original unsaved sentinel workbooks are observed unchanged.

Run 04 reached populated-sheet deletion, native save and reopen but failed a fixture assertion: it wrote `PRIVATE-JSON-ROUNDTRIP` and expected a suffixed marker. The raw failed report is unchanged. Parent cleanup separately verified the exact saved workbook path, two sheet names, actual marker and alerts before native close/quit. PID 16859 disappeared and source bytes stayed equal. This is explicit diagnostic cleanup, not automatic lifecycle acceptance.

Run 05 corrects only the marker expectation/write and output directory. The new transport completes secure open, internal populated-sheet deletion, alert restoration, strict save/close ACKs, independently verified native XLSX publication and reopen with exactly Data/Results plus the expected marker. Private PID 17830 closes all owned books and quits; it is absent. Original PID 9212 retains both unsaved books, original markers, and alerts=true. Parent independently checked ZIP integrity, exact sheet order, marker, output hash, unchanged transport source and process inventory.

Production transport callback errors, private launch/readiness, quarantine/recovery and full session wiring still require review and acceptance. These native tests do not close complete Excel parity, Word, Windows native/UI, or release gates. Historical scripts contain run-specific PIDs/paths and must not be rerun unchanged.
