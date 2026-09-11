# Native single-style Organizer copy feasibility

This is a build-only, owned-document probe, not an enabled heading capability.
Two controlled DOCX inputs add one paragraph style to an existing Word-saved
fixture. The input package is not the output implementation: Microsoft Word
opens and saves both donors before its native Organizer copies the named style
into an exact private recipient. No global template or security setting changes.

Native03 passes19 checks: donor styles materialize as17pt/23pt; a first named
copy appears in live memory, a second same-name copy changes the same style;
complete other styles, document XML/body/bookmarks and theme are preserved;
selection10..21 remains the exact original Unicode text; a separate unsaved
sentinel is preserved; native save/read-only reopen and input preservation pass.
Native Word inventory is empty afterward. Root verifies7 retained source hashes
and9 artifact hashes. Four pure/compile checks pass. Independent review remains
pending because the review agents cannot currently execute.

The disk/live distinction is verified independently: before first explicit save
the private disk file has no target style; after save it has17pt. After copying
the23pt donor, the live getter returns23pt while the private disk still has17pt;
a native save changes disk to23pt. Both final target definitions equal their
native donor definitions, ignoring only each style's edit-session rsid leaf.
Full non-target styles and theme comparisons remain exact.

Native01 failed at input-open binding before any Organizer call. Its input
serialization dropped namespace declarations referenced by mc:Ignorable. The
new negative test catches that defect; corrected input construction retains all
original styles bytes/declarations and inserts only the generated custom-style
node. The raw failed input/runtime and failure remain. Word inventory and CUA
DocStage confirmed no open test document. Native02 records a separate guard
rejection before open because an initial recovery lookup used the wrong marker
suffix. Recovery then used WordJobLock.quarantine_path, checked the exact owned
staging identity and the guarded recovery helper; its receipt and original
marker are retained. Native03 runs only after verified recovery.

Remaining: integrate the private style installer into the public native heading
method, independently review it, verify real existing/fresh styled paragraphs,
source inheritance and defaults beyond this closed fixture, numbering/section
continuity, Arabic/font rendering, complete native/UI/regression acceptance.
A19-check fixture pass does not close these requirements. Original source is
build/word-figure-rollback-20260911/native-v3-01/figure-rollback-source.docx; the
frozen fixture source hashes and all exact inputs are retained in each run.
