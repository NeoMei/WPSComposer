# Character-format record getter regression

Full38 passes **5,040 tests / 12 skips** in 259.68 seconds from immutable snapshot `54ce7e624e6fbc51f76c6e83f8e385b5b6af4dc2`, tree `4d0b2070390295f0a1631f1f61bec04edeb393b4`, parent `a05bfa0053b5f2abe37a16b4566e5ca99ef327fb`. All 384 recorded source/configuration hashes are unchanged and match the live candidate. The 140 targeted quality checks and 30 staged-checkout path checks also pass.

Native character-read evidence and independent review are retained in ../macos-word-quality/character-record-read. This verifies a narrow font/shading record getter optimization; full quality snapshots and final platform acceptance remain incomplete. No native skip is counted as a pass.

The later .gitattributes addition applies only to newly retained character-read evidence, preserving raw bytes on Windows and allowing raw log trailing blank lines. It is outside the 384 runtime/test/configuration hash set; git check-attr and staged diff checks verify that documentation-only rule separately.
