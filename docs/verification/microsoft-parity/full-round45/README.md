# Final native-import oracle regression — round 45

Immutable snapshot `e039a732c62d9512226b14c14501cfc3c65ffa44` passes **5,253 tests / 12 skips** in 221.28 seconds. All 401 source/configuration hashes remain unchanged and match the candidate. Production is unchanged from 433a199; the additions are an opt-in native heading import fixture and source-only tests, including exact native preimage/UTF-16 replacement expectations. These tests correctly leave failed native import semantics unaccepted.

CI12 and installed/native list acceptance apply to the identical production files. The new fixture tests have local full-suite evidence here; they require their own next hosted run. Full Word parity, Windows Office/WPS native/UI and per-capability certification remain incomplete.
