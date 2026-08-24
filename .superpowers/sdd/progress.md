# Longform M1 progress
Task 1: complete (bce8396 + ade679e, approved by review_task1_v2, 980 tests pass)
Task 2: complete (7b9b639, approved by review_task2_kimi, 1002 tests pass)
Task 3: complete (f04e5cc + df93a2e + 54e2a5c, approved by review_task3_v2, 1033 tests pass)
Task 4: complete (20c0cfc + ed66fc3, approved by review_task4, 1052 tests pass)
Task 5: complete (5a1beeb + a0d205f, approved by review_task5, 1080 tests pass)
Task 6: complete (857f5ad, approved by review_task6, 1102+12env tests)
Task 7: complete (uncommitted, approved by review_task7 after 2 fix rounds)
Task 8: complete (uncommitted, approved by review_task8 after 1 fix round)
Final review: approved pending commit (review_final_m1); spec issue-code divergence fixed (DUPLICATE_FRONT_MATTER_BLOCK)
M2 follow-ups: add_cross_reference emission, block/inline issue->degradation mapping, ABSTRACT_CONTENT_DEGRADED/PAGE_BREAK_CONTENT_DEGRADED codes, figure/table index front-matter placement, sourcePath in diagnostic JSON, .superpowers/sdd/task-6-report.md historical tracked file cleanup
M1 COMPLETE: 15 commits f17691d..9c4e3fd pushed to origin/codex/longform-m1, full suite 1144 passed 2026-08-22
M2 plan written: docs/superpowers/plans/2026-08-22-wpscomposer-longform-m2.md (10 tasks, uncommitted)
M2 Task 1: complete (uncommitted, approved by review_m2_task1 after 1 fix round, 217 longform tests)
M2 Task 2: complete (uncommitted, approved by review_m2_task2 after 2 fix rounds, 233 longform tests)
M2 Task 3: complete (uncommitted, approved by review_m2_task3, 422 tests)
M2 Task 4: complete (uncommitted, approved by review_m2_task4 after 1 fix round, 443 tests)
M2 Task 5: complete (uncommitted, approved by review_m2_task5, 290 longform tests)
M2 Task 6: complete (uncommitted, approved by review_m2_task6 after 2 fix rounds, 311 longform tests)
M2 Task 7: complete (uncommitted, approved by review_m2_task7 after 2 fix rounds, 326 longform tests; real-WPS evidence deferred - sandbox socket block)
M2 Task 8: complete (uncommitted, approved by review_m2_task8_v2 after 1 fix round + regex fix, 402+6skip)
M2 Task 9: complete (uncommitted, approved by review_m2_task9 after 1 fix round, 414+6skip)
M2 Task 10: ledger cleanup - historical task-6-report.md already untracked in 9c4e3fd
M2 ledger: CONSUMED block/inline issue->degradation mapping (Task2), ABSTRACT/PAGE_BREAK degradation codes (Task2), sourcePath redaction (Task9); DEFERRED to M3: add_cross_reference emission, figure/table index content population
M2 COMPLETE: 11 commits bc2ddfe..bc7c9a5 pushed; 1358 tests pass; macOS real-WPS evidence build/longform-m2/macos-native-20260822-1 (6/6 fixtures passed, WPS 12.1.26055, academic.pdf converted)
M3 plan: committed 18798d8, 9 tasks, preflight scan clean
M3 Task 1: complete (commits 3214d0b..2e99595, review approved after 3 fix waves)
M3 Task 2: complete (commits 330da6b..12b053f, review approved after 2 fix waves)
M3 Task 3: complete (commits f966d2c..1e776d0, review approved after 2 fix waves, 1557 passed + 6 gated skips)
M3 Task 4: complete (commits c4910d1..c36e190, review approved after 3 review rounds, 1621 passed + 6 gated skips)
M3 Task 5: complete (commits c07a2ba..60198fb, review approved after 3 review rounds, 1661 passed + 6 gated skips)
M3 Task 6: complete (commits 52a7608..ad0cd35, review approved after 5 review rounds, 1710 passed + 6 gated skips; real Windows deferred to final gate)
M3 Task 7: complete (commit 85362d6, internal independent review approved after 4 rounds, 1750 passed + 6 gated skips; real macOS deferred to Task 8)
M3 Task 8: complete (commit 9731ebc, review approved after real-WPS fix runs 7-20; macOS WPS 12.1.26055; evidence build/longform-m3/macos-native-20260823-20; real gate 1 passed no skip; full 1766 passed + 7 gated skips)
M3 Task 9: complete (this integration commit, `Integrate M3 longform pipeline and documentation`; independent whole-M3 review approved after preset/list-ref/Windows-overflow/list-format parity fix waves; 1785 passed + 7 gated skips)
M3 ledger: CONSUMED `add_cross_reference emission` and `figure/table index content population`; native cross-reference paragraphs and populated front-matter indexes are covered by Windows mocks and macOS run-20 evidence.
M3 evidence: macOS WPS 12.1.26055, `build/longform-m3/macos-native-20260823-20/platform-evidence.json`, chapter/global DOCX+PDF mutation artifacts, seven inspected screenshots, ten mutation refresh phases converged in two rounds.
M3 COMPLETE boundary: Windows M3 remains mock-only until the final all-milestone Windows gate. Carry forward only native equation content and bibliography/citations (M4), general degradation (M4), PDF quality/re-layout (M5), and public `generate()` default migration (M5). No release/version work was performed.
M4 plan: commits e2e235b + de977e4, preflight reviewed clean against the M3 boundary.
M4 Task 1: complete (commits 2d8ea91, 248a8ce, 25b85e3, 5fb7785, 36ae4aa; independent review approved after four fix waves; restricted grammar/conversion and legacy compatibility gates green).
M4 Task 2: complete (commits a69f06a, c63ede3, 5a0a9fa; independent review approved after semantic/page-policy fixes; formula resources, citations, bibliography ordering, and M1-M3 regressions green).
M4 Task 3: complete (commits 18aa80c, 987a230, 0e661f3, d54d905, 5f4469e, c9194a5, 963fd1d; independent review approved after closed-schema, ownership, table-citation, and privacy fix waves).
M4 Task 4: complete (commits 24512c3, 4add866, 0774c92, ae3b41c, 61e095f, c5710d8, 89dd275; independent review approved after controller, rollback/preflight, privacy-parity, and reference-locality fixes).
M4 Task 5: complete (commits bac5195, e92bd97, 0770df1, 8f83b2f, 0109171, 14d601b; independent review approved; Windows COM behavior is mock-verified and deliberately awaits the M5/final real-Windows gate).
M4 Task 6: complete (initial 6e156d1; remediation b1c02f7..897cbc5 plus e13f10b, 7cec9a2, 4a58fcb; repeated independent review and real-WPS fix waves approved; macOS WPS 12.1.26055 professional BuildUp no-op is detected and honestly degrades to the marked image/source ladder).
M4 Task 7: complete (commit 60d1ad8; offline acceptance and non-skipped real macOS create/save/reopen/refresh/move/recovery gates approved; final evidence `build/longform-m4/macos-native-20260823-184331-936811`, including DOCX/PDF/screenshots and privacy-safe evidence JSON).
M4 Task 8: complete in this integration commit (independent review approved after one fix wave; 18 Task8 gates, 90 focused integration tests, 1981 expanded tests + 7 gated skips, and full suite 2403 passed + 7 gated skips).
M4 evidence: final macOS run `build/longform-m4/macos-native-20260823-184331-936811`; all three real gates passed with formula fallback/number/bookmark/reference continuity, citation/bibliography ordering, notice styling/placement, PDF non-overlap, reopen/move stability, and fatal engine behavior inspected.
M4 carry-forward only: M5 PDF geometry/bbox quality mapping, deterministic full re-layout and notice-only patching, performance gates, public `generate()` default migration, final real Windows/cross-platform verification, and final release/version/publication work.
M5 plan: commit 7c913d0; Tasks 1-7 complete in commits bd1a326..30596af (typed quality model, normalized PDF analysis, closed relayout, notice patching, bounded lifecycle, production adapters, default DOCX/PDF route).
M5 Task 8: complete (commit 0d61345; 1228 M2-M5 tests passed + 8 gated skips; macOS WPS 12.1.26055 arm64 evidence under build/longform-m5/pytest-real-5, -6, and -7).
M5 Task 8 real gate: three consecutive post-fix runs passed; each produced six fixtures plus a 63-page performance PDF with one generation/export/analysis, zero performance issue codes, zero patch, and approximately 139-140 seconds total.
M5 Task 9: complete (commit 8574b6e; public docs, Markdown guide, macOS verification truth, selected operator-doc installation, wheel/import and fresh-plugin installation checks).
M5 Task 10 local audit: Round A found and fixed stale heading-numbering expectations, disconnected Presentation conversion WebView reuse, and duplicate quality-notice codes (6907311); Round B found test-double interface drift (165b3da); Round C full suite passed 2537 + 8 gated skips.
M5 Windows handoff: runnable cross-platform evidence contract and `test_windows_real_wps_m5.py` added locally; final three-round Windows COM/UI/performance evidence remains pending before 0.8.0.
M5 Task 10 local closure: Windows gate commit a68f93b; Round D full suite 2541 passed + 9 explicit platform skips; generalized runner macOS real gate `build/longform-m5/pytest-real-8` passed with 63 performance pages, one generation/export/analysis, zero patch/issues, 139.7457 seconds.
M5 LOCAL COMPLETE: worktree clean, JS syntax/asset manifest exact, fresh wheel import and fresh full-plugin install/public generation/conversion verified. Only the documented three-round real Windows gate and any resulting cross-platform fix cycle remain before release work.
M5 Windows gate: completed on 2026-08-24 in commits e9cac73 and 5148b1a; 2512 platform-independent tests passed on Windows, the real M5 gate and pooled-code rerun each passed three consecutive times, and same-process multi-format generation passed 15/15.
M5 post-Windows macOS acceptance: found and fixed M4 env-gate failure, stale documentation truth, reusable-executor front-matter leakage, failed pooled-COM apartment leakage, and silently omitted fenced code blocks. Fresh suite passed 2542 + 12 explicit platform skips. Final real evidence `build/longform-m5/final-postfix-{1,2,3}` passed three consecutive runs: Unicode code visible with one generation/no patch; performance 63 pages, one generation/export/analysis, zero patch/issues, 144.5114-170.3104 seconds. One final Windows clean-checkout rerun of the new shared-plan/Windows-executor fixes remains before 0.8.0.
