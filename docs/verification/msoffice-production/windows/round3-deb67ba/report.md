# Final-candidate Windows native acceptance, in progress

Base source deb67ba plus identity fix and pagination repair was merged as c0604a85d1e193f1bdc13f50aa188d8385656a98. Exact Git objects were uploaded through GitHub API after repeated Git transport failures: 68 blobs, 51 trees and three commits matched their local SHAs; the branch was updated without force and re-read. First-section WPS compatibility and two genuinely POSIX-only recovery test markers are commit c2394c2.

## Completed native gates

- Word smoke-05: PASS, timeout600, 306.266 seconds. Public generate(docx), convert_to_pdf and generate(pdf), requested paths, source preservation and overwrite refusal all pass.
- Word representative-03: PASS with strict section policy. DOCX ten checks pass; PDF and direct PDF have cover blank, TOC i, body1/2/3. The independent eleven checks confirm TOC logical references1, actual four numbered levels, FangSong12pt, 24.026pt first-line indent, six heading sizes, complete36-record table, repeated headers and page bounds. All five pages visually reviewed.
- Sentinel-03: original Word PID9252 unsaved 文档7, text SHA4d68aa90aee5e79223ded3195b39cb344afdce67d72183ad55ea9311d4d0d8d8 stayed unchanged in text/path/saved state. Only sentinel Close(SaveChanges=0); old Word remained responsive and empty.
- Numbering/media-03: ordinary explicitly unnumbered H1 remains unnumbered after later numbered H1. Shared sequenceTransparent boundary permits native figure1-1 then1-2; native table1-1 and three REF results all pass. Two actual image drawing elements are present. The small synthetic raster deliberately exercises retained IMAGE_LOW_DPI notices; the notices are not suppressed.
- Windows WPS representative-02: PASS,49.265 seconds, all three public routes, strict DOCX section policy and PDF footer blank/i/1/2. WPS COM reports Name Microsoft Word for compatibility; actual app.Path is the installed WPS office6 path, Version12.0, generated DOCX Application explicitly identifies WPS Office12.1.0.28505 and PDF Creator is WPS文字. Word was not substituted. Static engine_executable detection returned None despite working explicit COM activation; that detection limitation is recorded separately.

## WPS failure and compatible correction

The first explicit WPS public run failed. Dedicated owned-document probes showed strict first-header LinkToPrevious readback remains True until the empty range is realized, even though the first section has no predecessor. Footer/PageNumbers Boolean restart, starting1 and Roman style2 all work. The correction omits predecessor-link setting/checking on the first section only; later sections still detach before content and require readback. No registry or global WPS settings were changed. WPS probe documents were saved and closed individually without Quit. The corrected native WPS probe and full public representative pass.

## Windows tests

First full run before final candidate:2778 passed/6 known platform failures/39 skips. On c0604a8 the full run (one cwd-file-set test deliberately isolated from native artifact writes) completed2812 passed/2 failed/43 skipped/1 deselected in690.67s. Both failures were new Mac recovery tests acquiring actual fcntl locks on Windows; matching the existing lock tests, they now explicitly require POSIX. Final affected groups, including the previously isolated cwd test, passed1021 with16 skips. The one new WPS first-section regression test is included. Raw RED/GREEN and full logs are retained; this is not described as a single clean full-suite command.

## Outstanding structural edit gate

The newly requested insert/delete-heading test exposed a real Word multilevel list defect: inserting H1 changes the old H1 to2 but H2-H4 remain1.1/1.1.1/1.1.1.1. The saved original styles reference separate numId and abstractNum definitions. This was not covered by a single-chapter static fixture. The full edit evidence and a repair are in progress. Production admission remains NO_GO until full H1-H4 automatic renumbering and changed-page TOC refresh, save/reopen, UI and final PID checks finish. Earlier static native passes are not claimed to satisfy this edit gate.
