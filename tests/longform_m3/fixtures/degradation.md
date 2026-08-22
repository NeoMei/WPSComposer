---
title: M3 deterministic degradation
toc: true
figure_index: true
table_index: true
heading_numbering: decimal
caption_numbering: auto
---

# 1 Degradation chapter

:::figure {#fig:missing-caption caption=""}
![kept image](media/png-wrong.jpg)
:::

:::table {#tab:missing-caption caption="" style="grid"}
| A | B |
|---|---|
| 1 | 2 |
:::

:::table {#tab:invalid-merge caption="Invalid merge falls back" style="three-line" merges="A2:B3"}
| Group | Value |
|:------|------:|
| Anchor | occupied |
| also occupied | 3 |
:::

:::table {#tab:tall-group caption="Tall vertical group" style="grid" merges="A2:A17"}
| Group | Detail |
|:------|:-------|
| Tall | row 01 |
| | row 02 |
| | row 03 |
| | row 04 |
| | row 05 |
| | row 06 |
| | row 07 |
| | row 08 |
| | row 09 |
| | row 10 |
| | row 11 |
| | row 12 |
| | row 13 |
| | row 14 |
| | row 15 |
| | row 16 |
:::

Missing caption target: {{ref:fig:missing-caption}}.
