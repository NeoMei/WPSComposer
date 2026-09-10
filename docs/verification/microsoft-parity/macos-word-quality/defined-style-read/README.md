# Defined-style read diagnosis

The corrected source-only style probe passes twice on the retained `run-02/before.docx` fixture. Both reads return 373 catalog entries and 20 defined-style rows in 11.835 / 11.924 seconds, with identical typed JSON hashes. Saved/read-only state, body End/hash, list-template count (zero), paragraph/field/table/bookmark/section counts and private/source bytes remain unchanged. The owned copy closes and independent Word inventory is empty. Root independently verified all recorded probe source hashes.

The source is a nonempty, one-section document with two tables, two fields and nine bookmarks. It has 33 OOXML style definitions and no header/footer parts or numbering part. This is **STYLE_SUBSET_DIAGNOSED**, not full style coverage: default-empty, custom/modified/unused built-in, nonzero numbering and multisection fixtures still require separate runs. The native catalog and OOXML style definitions have different coverage and must not be conflated.

The production quality snapshot is unchanged. Absent-header reads, complete page-story/field/shape preservation and complete repeated quality-snapshot acceptance remain unresolved. The accompanying source-only worker report records the probe design and historical limitations; later character-read evidence remains in the adjacent `character-record-read` directory.
