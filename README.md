# WPSComposer

> 一个 Codex 插件：通过 WPS Office COM 自动化，让 AI agent 能够生成和编辑高质量排版的 DOCX / PPTX / XLSX / PDF 文档。

## 为什么需要这个插件？

python-docx、openpyxl、python-pptx 等库通过直接拼接 OOXML 来生成文档，但它们有一个根本性缺陷：**无法计算排版**。分栏后的文本流如何分布、浮动图片周围怎样环绕、表格跨页后表头是否重复——这些问题只有真正的排版引擎才能回答。

WPSComposer 的做法是直接驱动 WPS Office（或 MS Office）的 COM 接口，把排版工作交给真实的排版引擎完成。agent 只需描述"我要什么"，引擎负责计算"怎么排"，最终输出和你在 WPS 界面里看到的一模一样。

## 它能做什么？

**生成新文档**（Markdown → 排版文档）：
- 分栏排版、浮动文本框（带文字环绕）、艺术字
- 合并/底纹/边框表格、自动生成目录和页码字段
- 公式、冻结窗格、条件格式、图表
- 16:9 幻灯片、标题版式、演讲者备注
- 五种设计预设（`academic`、`consultant`、`business`、`tech`、`proposal`）快速切换风格

**编辑已有文档**（对话式精排）：
- 打开已有文件或直接接管 WPS 中正在编辑的文档
- 读取任意段落的当前格式（字体、字号、行距、缩进……）并以 JSON 返回
- 对单个元素做精确的格式补丁，不影响无关属性
- 通过对话微调每一个段落的字号、颜色、对齐方式

**编辑 PDF**（合并/拆分/水印/旋转/提取）。

**Office 转 PDF**（单文件、跨平台）：
- Word：`.doc`、`.docx`
- Excel：`.xls`、`.xlsx`，导出全部可见工作表
- PowerPoint：`.ppt`、`.pptx`
- Windows 使用 WPS/MS Office COM；macOS 使用 WPS JSAPI 容器暂存后原子发布

## 原理架构

```
用户对话
  │
  ▼
Markdown 源文本
  │  md_parser.py 解析
  ▼
StructuredDocument（结构化文档模型）
  │  格式渲染器 (writer_renderer / sheet_renderer / slide_renderer)
  │  + 设计预设 (design_presets) + 排版规范 (reference_styles)
  ▼
Composer 引擎（WriterComposer / SheetComposer / SlideComposer）
  │  通过 COM 接口驱动 WPS Office
  ▼
用户请求的单一产物：DOCX / PPTX / XLSX / PDF
```

排版规范（`reference_styles.py`）默认遵循中文公文/政企方案标准：
- 正文：仿宋 12pt，两端对齐，首行缩进 2 字符，1.5 倍行距
- 标题：黑体加粗（H1 16pt 居中，H2/H3 15pt 左对齐）
- 西文：Times New Roman

## 快速开始

### 安装

```bash
# 克隆仓库
git clone https://github.com/NeoMei/WPSComposer.git

# 安装到 Codex 个人插件 marketplace
cd WPSComposer
python3 install.py            # 首次安装
python3 install.py --force    # 覆盖更新
```

Windows 也可使用 `pwsh ./install.ps1`，macOS/Linux 可使用
`./install.sh`。安装后重启 ChatGPT/Codex Desktop，在 Plugins 目录中安装或启用
`wps-composer`，然后在 Codex 中引用 `$WPSComposer`。

### 运行时要求

- Python 3.9 或更高版本。
- Windows 上生成 DOCX/PPTX/XLSX 需要 WPS Office 或 MS Office，以及 `pywin32`。
- macOS：公开 `generate()` 仍为 **NO-GO**，因为 WPS 12.1.26035 的
  Writer `SaveAs2` 会打开原生保存面板；已有 Office 文件转 PDF 已通过
  六格式双轮真实门禁并启用。详见 [macOS Phase 0](docs/macos-phase0.md)。
- Markdown 解析、文档模型和 PDF 编辑模块可在非 Windows 系统导入。
- PDF 编辑需要 `pypdf` + `pdfplumber`，文本水印额外需要 `reportlab`。

## 用法

### 一行生成文档

```python
from skills.WPSComposer import generate

generate("报告.md", format="docx", preset="academic")    # 学术风格 DOCX
generate("演示.md", format="pptx", preset="business")     # 商务风格 PPTX
generate("数据.md", format="xlsx")                        # 带公式的 XLSX
generate("报告.md", format="pdf",  preset="consultant")   # 咨询风格 PDF
```

### 编辑已有文档（对话式精排）

```python
from skills.WPSComposer import inspect, edit, attach_active

# 打开已有文件，查看结构化内容和格式
tree = inspect("报告.docx")

# 接管 WPS 中正在编辑的文档，读取光标位置的格式
w = attach_active("writer")
print(w.inspect_selection())

# 对第 3 段做格式补丁：字号 11、颜色 #333、两端对齐、行距 20pt
w.apply_format_patch(
    "paragraph:3",
    font={"size": 11, "color": "#333333"},
    paragraph={"alignment": 3, "line_spacing": 20},
)
w.save_current()

# 批量补丁 + 另存 + 导出 PDF 用于核对
edit(
    "报告.docx",
    output="报告-修订.docx",
    export_pdf="报告-修订.pdf",
    patches=[
        {"target": "paragraph:1", "font": {"size": 20, "bold": True}},
        {"target": "table:1/cell:2,1", "fill": {"color": "#FFF2CC"}},
    ],
)
```

可寻址的元素包括：段落、表格、单元格、形状、节、页眉页脚等。目标语法和补丁字段详见 [API 参考](skills/WPSComposer/references/api.md)。

### Word / Excel / PowerPoint 转 PDF

```python
from skills.WPSComposer import convert_to_pdf

# 默认输出同目录、同文件名的 .pdf
pdf = convert_to_pdf("季度报告.xlsx")

# 指定输出；只有显式 overwrite=True 才覆盖已有文件
pdf = convert_to_pdf(
    "演示稿.pptx",
    "exports/演示稿.pdf",
    overwrite=True,
)
```

`source` 只接受 `.doc/.docx/.xls/.xlsx/.ppt/.pptx`，大小写不敏感；一次只
转换一个文件。缺失源文件抛出 `FileNotFoundError`，已有目标默认抛出
`FileExistsError`，不支持的扩展名抛出 `ValueError`。WPS/Office 运行时错误
抛出 `ConversionError`，并提供 `code`、`source`、`component`、`backend`
和 `message` 字段。macOS 上输入和输出都先进入 WPS 私有容器会话，校验后
再通过目标目录内临时文件、`fsync` 和原子替换发布；不需要完全磁盘访问、
辅助功能、AppleScript 或屏幕控制权限。

### PDF 编辑

```python
from skills.WPSComposer import PdfComposer

PdfComposer.merge(["a.pdf", "b.pdf"], "合并.pdf")
PdfComposer.split("输入.pdf", "输出目录/")
PdfComposer.extract_pages("输入.pdf", [1, 3], "提取.pdf")
PdfComposer.rotate("输入.pdf", 90, "旋转.pdf")
PdfComposer.add_text_watermark("输入.pdf", "机密", "加水印.pdf")
print(PdfComposer.extract_text("输入.pdf"))
```

## 能力矩阵

| 能力 | Writer (docx) | Sheet (xlsx) | Slide (pptx) |
|---|:---:|:---:|:---:|
| 分栏 / 页面设置 | ✅ | ✅ | ✅ |
| 表格（合并/底纹/边框） | ✅ | ✅ | ✅ |
| 浮动形状 + 文字环绕 | ✅ | — | ✅ |
| 艺术字 | ✅ | — | — |
| 图表（柱/折/饼） | — | ✅ | — |
| 公式 + 数字格式 | — | ✅ | — |
| 冻结窗格 / 条件格式 | — | ✅ | — |
| 图片 | ✅ | — | ✅ |
| 项目符号 / 编号列表 | ✅ | — | — |
| 页眉 / 页脚 | ✅ | ✅ | — |
| 自动目录 / 页码字段 | ✅ | — | — |
| 检查已有文档格式 | ✅ | ✅ | ✅ |
| 元素级格式补丁 | ✅ | ✅ | ✅ |
| 转为 PDF（含旧格式） | ✅ DOC/DOCX | ✅ XLS/XLSX（全部可见表） | ✅ PPT/PPTX |

## 模块结构

```
WPSComposer/
├── .codex-plugin/plugin.json        # 插件清单
├── skills/WPSComposer/
│   ├── SKILL.md                     # Codex 加载的 skill 入口
│   ├── references/api.md            # API 完整参考
│   └── scripts/
│       ├── wps_engine.py            # 统一入口（re-export facade）
│       ├── orchestrator.py          # generate() 一行生成
│       ├── conversion.py            # convert_to_pdf() 跨平台入口
│       ├── artifact_transport.py     # 校验、fsync、原子发布
│       ├── windows_conversion.py     # Windows COM 转换后端
│       ├── macos_probe/              # macOS WPS JSAPI、容器暂存与转换
│       ├── _dispatch.py             # COM 连接、ProgID 回退链
│       ├── _colors.py               # 统一颜色模型（hex ↔ BGR）
│       ├── _base.py                 # BaseComposer 通用生命周期
│       ├── writer.py                # WriterComposer → DOCX
│       ├── sheet.py                 # SheetComposer → XLSX
│       ├── slide.py                 # SlideComposer → PPTX
│       ├── pdf.py                   # PdfComposer → PDF 编辑
│       ├── md_parser.py             # Markdown 解析器
│       ├── document_model.py        # 结构化文档模型
│       ├── reference_styles.py      # 排版规范定义
│       ├── design_presets.py        # 设计预设（5 套配色+字体）
│       ├── layout_templates.py      # 幻灯片版式（14 种）
│       ├── quality_checks.py        # 质量校验
│       ├── formatting.py            # 格式工具函数
│       ├── heading_numbering.py     # 标题自动编号
│       └── renderers/               # 格式渲染器
│           ├── writer_renderer.py
│           ├── sheet_renderer.py
│           └── slide_renderer.py
├── install.py                       # 跨平台安装器
├── install.ps1                     # PowerShell 包装脚本
├── install.sh                      # macOS/Linux 包装脚本
└── README.md
```

## 开发调试

在本项目目录里直接修改 `skills/WPSComposer/scripts/` 下的文件，用 pytest 和临时脚本验证。原生 WPS 排版改动可另外导出 PDF 作为开发验证证据，但公开生成接口只返回用户请求的格式。调试满意后运行 `python3 install.py --force` 更新本地插件。

## License

MIT
