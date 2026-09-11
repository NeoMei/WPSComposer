# Character-format property record reads — September 10, 2026

The production change resolves the font and shading property records once per native UTF-16 coordinate, then projects the same seven typed values in the same order. Coordinate iteration, hash input, admission gates, paragraph/section/style coverage, and the 60-second native operation limit are unchanged. This is a scoped read optimization, not complete snapshot or quality acceptance.

Run09 compares the record prototype with existing getters at all 220 positions in the retained native quality fixture. All comparisons agree. Two complete record reads agree and finish in 42.146s and 42.870s, with saved state, body and measured object counts preserved and the owned document closed. These timings are observations under that run's conditions, not a controlled speedup ratio.

Run10 is retained FAIL: an XML serializer renamed namespace declarations while mc:Ignorable still referenced the original names. The synthetic fixture failed Word binding. Parent independently observed inventory[], retained the private runtime, and completed guarded quarantine recovery. An earlier Python assertion syntax error was caught before native launch.

Run11 preserves the original package namespaces and changes only the first exact run's formatting. It starts on a new read-only owned copy, executes the **actual production-generated character loop first**, repeats it, then compares every coordinate with the legacy getter. The first reads finish in 4.415s and 4.191s and agree; all 220 positions match later legacy observations. Coverage includes explicit font/size/bold/italic/underline/color/shading, supplementary Unicode, fields and table terminators.

Run11's original report remains FAIL because the fixture expected 8-bit RGB. Word returns native 16-bit channels, consistent with existing color conversion code: #112233 is [4369,8738,13107], #EEE8AA is [61166,59624,43690]. No production normalization changed. The separately retained assess.py and assessment.json perform a strict, file-only reassessment: exact typed equality, all coordinates once, explicit native color values, before/after saved/body/measured-count state, owned close and empty final inventory. That assessment passes only the character-read scope and binds the original report SHA256.

Raw source hashes refer to original build paths at execution time. All original reports and sources are preserved; each run's HASHES.json binds retained files. The native runner's DIAGNOSED status alone is not an acceptance gate. Complete in-memory formatting neutrality beyond observed fields is not claimed.

The absent-header getter still dirties native state, all-style definition reads can materialize list definitions, and the full snapshot has not passed within budget. Nine missing Mac Word methods, middle-table construction and final Windows native/UI acceptance remain open. No parity baseline row was removed and no release is authorized by this evidence.
