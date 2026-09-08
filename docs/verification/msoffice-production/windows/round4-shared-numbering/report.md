# Windows shared numbering and auto selection — stage 4

Native scheme gates: **PASS** on Microsoft Word 16.0.17932 and WPS 12.1.0.28505. Six cases cover two chapters, each with H1-H4, plus insertion, deletion, save/reopen and PDF numbering (30 checks). Word binds heading styles through ListLevels.LinkedStyle; WPS retains Style.LinkToListTemplate. Chinese-formal follows chapter/section/Chinese-list/parenthesized-Chinese; hybrid uses styles 37/253/253/22 and displays 第一章 / 1.1 / 1.1.1 / 关键工法01：, with second-chapter descendants 2.1 / 2.1.1. Actual Word Chinese/hybrid and WPS hybrid page images were inspected.

Auto selection: **PASS**, read-only detection finds the real 32-bit WPS LocalServer32 registration from 64-bit Python, selects WPS, and public generate plus convert produce WPS Office metadata without changing source DOCX bytes. auto-01 records a helper dependency failure (psutil unavailable); auto-02 uses existing WMI and passes. Detection starts no Office process. No registry/settings writes were performed.

Unit RED/GREEN: scheme tests 4 failed / 4 passed before fix; affected suite 902 passed / 10 skipped after fix (20 existing Pillow deprecation warnings). Registry tests 3 failed / 2 passed before fix; detection/routing 27 passed after fix.

Raw native-edit-01 shows the original Word split-list defect and a separate localized-style helper error. LinkedStyle Word capability probe passes all four levels; WPS rejects that primitive, so its existing binding path remains. These failures are retained alongside successful production-composer scheme cases.

Production admission remains **NO_GO** pending final representative structural pagination/TOC plus actual UI reopen, final PID observations and complete Windows pytest. The new public representative DOCX has passed the shared numId gate; its remaining PDF paths are still running at this stage. Prior media/sentinel/UI evidence remains under round3/round1.
