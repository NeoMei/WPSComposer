---
title: M3 auto boundary
toc: true
figure_index: true
table_index: true
heading_numbering: decimal
caption_numbering: auto
---

:::figure {#fig:before-h1 caption="Before numbered chapter"}
![before](media/gif-wrong.bmp)
:::

:::table {#tab:before-h1 caption="Table before chapter" style="three-line"}
| A | B |
|---|---|
| 1 | 2 |
:::

# 1 Numbered chapter

:::figure {#fig:after-h1 caption="After numbered chapter"}
![after](media/tiff-wrong.dat)
:::

:::table {#tab:after-h1 caption="Table after chapter" style="grid"}
| A | B |
|---|---|
| 3 | 4 |
:::

Boundary references: {{ref:fig:before-h1}} then {{ref:fig:after-h1}}.

# 第二章 Unnumbered boundary

:::figure {#fig:after-unnumbered-h1 caption="After unnumbered H1"}
![after unnumbered](media/png-wrong.jpg)
:::

:::table {#tab:after-unnumbered-h1 caption="Table after unnumbered H1" style="grid"}
| A | B |
|---|---|
| 5 | 6 |
:::
