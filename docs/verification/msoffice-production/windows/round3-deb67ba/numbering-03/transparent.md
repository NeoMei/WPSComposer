---
title: Native transparent heading
title_page: false
toc: true
heading_numbering: decimal
caption_numbering: auto
---
# 1 Numbered chapter

:::figure {#fig:first caption="First numbered figure"}
![first](figure.png)
:::

# 第二章 Unnumbered boundary

:::figure {#fig:second caption="Second numbered figure"}
![second](figure.png)
:::

:::table {#tab:values caption="Measured values"}
| Signal | Value |
| --- | --- |
| Alpha | 10 |
| Beta | 20 |
:::

Reference pair: {{ref:fig:first}} and {{ref:fig:second}}. Table reference: {{ref:tab:values}}.
