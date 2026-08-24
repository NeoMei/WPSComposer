---
title: M3 chapter evidence
author: WPSComposer
title_page: true
toc: true
figure_index: true
table_index: true
heading_numbering: decimal
caption_numbering: auto
design: academic
---

# 1 Native chapter

:::figure {#fig:chapter-one caption="Chapter diagram" width="full" kind="diagram"}
![native diagram](media/png-wrong.jpg)
:::

:::figure {#fig:chapter-two caption="Two-column comparison" layout="columns" columns="2" width="full"}
![left comparison](media/jpeg-wrong.png)
![right comparison](media/oriented.jpg)
:::

:::table {#tab:chapter-one caption="Three-line grouped data" style="three-line" merges="A2:A3"}
| Group | Metric | Value |
|:------|:------:|------:|
| A | Count | 10 |
|   | Rate | 20 |
:::

:::table {#tab:chapter-two caption="Grid data" style="grid"}
| Name | Status |
|:-----|:------:|
| Alpha | Ready |
| Beta | Done |
:::

:::table {#tab:chapter-delete caption="Disposable chapter grid" style="grid"}
| Name | Status |
|:-----|:------:|
| Temporary | Delete |
:::

:::formula {#eq:chapter-one}
E=mc^2
:::

See {{ref:fig:chapter-one}}, {{ref:tab:chapter-one}}, and {{ref:eq:chapter-one}}. Missing stays inline: {{ref:fig:not-found}}.
