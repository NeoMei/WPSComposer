---
title: M4 native acceptance
author: WPSComposer
title_page: false
toc: false
heading_numbering: decimal
caption_numbering: chapter
bibliography_include_uncited: true
design: academic
---

# Native formula families

Inline continuity begins {{cite:b}}, repeats {{cite:b}}, cites {{cite:a}}, and keeps missing {{cite:missing}} in this paragraph.

:::equation {#eq:power fallback_image="media/formula.svg"}
E = mc^2
:::

:::formula {#eq:scripts}
a_i^2+b^{n}
:::

:::equation {#eq:fraction}
\frac{a+b}{c}
:::

:::formula {#eq:radical}
\sqrt{x}+\sqrt[3]{y}
:::

:::equation {#eq:large-operators}
\sum_{i=1}^{n}i+\prod_{j=1}^{m}j
:::

:::formula {#eq:integral}
\int_{0}^{1}xdx
:::

:::equation {#eq:delimiters}
\left(\frac{x}{y}\right)
:::

:::formula {#eq:symbols}
\alpha+\Gamma\leq\infty
:::

:::equation {#eq:matrix}
\begin{pmatrix}a&b\\c&d\end{pmatrix}
:::

:::formula {#eq:cases}
\begin{cases}x&x>0\\-x&x\leq0\end{cases}
:::

:::equation {#eq:nested}
\frac{1}{1+\sqrt{x_i^2}}
:::

:::formula {#eq:unicode}
α + 中文变量 = é
:::

Formula references remain native: {{ref:eq:power}} then {{ref:eq:matrix}}.

:::bibliography
[a] Alpha uncited-order declaration.
[b] Beta cited first.
[c] Gamma remains uncited.
malformed bibliography declaration
:::
