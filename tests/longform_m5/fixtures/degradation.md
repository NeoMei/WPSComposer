---
title: M5 受控降级验收
author: WPSComposer
header: 受控降级验收
toc: true
title_page: true
heading_numbering: decimal
design: academic
---
# 降级边界

无法解析的引用必须保留可见占位：{{ref:fig:missing}}。

:::equation {#eq:malformed fallback_image="media/missing.svg"}
x_{
:::

# 后续正文

即使局部对象插入失败，文档仍应生成，并在相应位置显示受控提示。
