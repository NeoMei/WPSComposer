# 排版回归守则（Regression Guardrails）

> **更新/重构 M5 排版前必读。** 本文固定 2026-09-01 修复（commit `66d4b9b`）调试好的文档行为。
> 任何 M5 路由的改动必须保持下述功能不回退；改前跑基线测试，改后跑全套 + 快照 diff + 实机 PDF 验证。

## 背景

0.8.0 将 DOCX/PDF 默认路由切到 M5 长文引擎时，legacy 路由已有的四项排版能力出现回退，用户实际生成商业计划书时发现（字号坍缩、无首行缩进、编号关联消失、封面标题重复）。当日修复并推送。本文档防止下次升级再犯。

## 固定的核心功能（升级不可抹掉清单）

### 1. 封面页
- markdown front matter `author` / `date` / `title_page: true` → 标准封面（Title 样式 + 署名 + 日期）
- **标题全文只出现一次**（仅封面）；正文从第一章直接开始
- Title 显式 `outlineLevel=10`（正文大纲级别），避免 WPS 内置 Title 被目录再次收录；验收必须计入目录缓存中的重复标题
- 教训：H1 下手写署名行且无 front matter 时，引擎会把标题块渲染两遍——封面信息一律走 front matter，不要在 H1 下重复手写

### 2. 标题分级排版（正式公文层级）
- **H1 = 16pt 居中加粗；H2/H3 = 15pt；H4/H5 = 14pt；H6 = 12pt**，全部加粗 + `keep_with_next` + 段前/段后距
- 封面 Title = 22pt
- 常量在 `skills/WPSComposer/scripts/longform/plan.py`：`_HEADING_LEVEL_SIZE_PT` / `_HEADING_SPACE_BEFORE_PT` / `_HEADING_SPACE_AFTER_PT` / `_TITLE_SIZE_PT`（源自 `reference_styles.HEADING_STYLE_MAP`）
- **禁止**退回"单一基准值递减"（0.8.0 回退形态：H1=14→H2=13→H3=12）

### 3. 正文排版
- **仿宋**（GB/T 9704 公文规范）12pt，两端对齐（align=3），1.5 倍行距——`policy.py` `body_font={"cjk": "仿宋"}`
- **首行缩进 2 字符**：Body Text 样式 `indentFirst = 2 × body_size_pt`（24pt）——在 plan.py 的 ensure_styles 发射中
- 表格/目录小字（10pt 书宋族）是 m0 层 macOS 字体兼容设计（`writer-longform-m0.js` chooseFont），**不是 bug，保持不动**

### 4. 原生标题编号关联（WPS 内增删章节自动重编号）
- `# 标题 / ## 01 章 / ### 2.1 节` 惯用结构：标题被封面消费后，正文无 H1 且从 L2 开始 → **semantic 层整体升一级**（层级降级），使 "01" 对上 L1 前缀模式、"2.1" 对上 L2——`semantic.py` `_apply_heading_numbering`
- **方案检测前缀证据优先**于汉字占比回退："01/2.1"→decimal，"第一章+1.1"→hybrid-bid——`_detect_heading_scheme`
- preface 门控不得吞掉全部 L2+ 章节（降级后章节即 L1，门控自然通过）
- 编号通过样式级 `LinkToListTemplate` 绑定（段落 XML 中无 numPr 属正常）

### 5. Windows COM 样式键兼容
- plan 发射 **camelCase** 样式键（对齐 macOS addin），`writer.py` `_STYLE_CAMEL_KEYS` 提供 snake_case 别名
- 删除别名 = Windows 路径 M5 样式**整体静默丢弃**（0.8.0 隐藏 bug）

## 保护机制

| 回归测试 | 锁定内容 |
|---|---|
| `tests/longform/test_plan.py::test_heading_styles_restore_formal_document_typography` | 分级字号/居中/加粗/keep_with_next/Title 22pt/仿宋/两端对齐 |
| `tests/longform/test_plan.py::test_build_policy_returns_stable_defaults` | 正文仿宋默认值 |
| `tests/longform/test_semantic.py::test_title_anchored_level2_headings_get_native_numbering` | 层级降级 + 前缀剥离 + decimal 编号 |
| `tests/test_writer_renderer.py::test_configure_style_accepts_camelcase_longform_keys` | Windows 键别名 |

快照（改排版必然变化，diff 必须只含预期块）：
- `tests/longform_m2/snapshots/*_ops.json`（academic、degradation）
- `tests/longform_m3/snapshots/*.json`（auto_boundary、chapter_native、degradation、global_media）

## 更新前检查单

1. **基线**：`python3 -m pytest tests/longform tests/longform_m2 tests/longform_m3 -q` 全绿再动手
2. **改动后**：全套 `tests/longform*` + 快照 diff 审查（非预期块变化 = 行为漂移，先弄清楚再接受）
3. **实机验证**：`generate()` 出 docx + pdf，用 pdfplumber 程序化检查——字体分布（仿宋/黑体）、标题字号分级、段落首行缩进（段首 x0 与续行差 ≈24pt）、标题编号渲染；docx 解包查 styles.xml（heading 样式 numPr、Body Text firstLine=480）
4. **已知预存失败（勿误判为回归）**：`tests/longform_m4/test_acceptance_m4.py` 的两个 PDF gate 测试在干净树上亦失败（环境性）

## 历史索引

- `66d4b9b` Restore M5 document typography and native heading numbering（2026-09-01，13 文件 +224/-44）
