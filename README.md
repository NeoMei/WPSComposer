# WPSComposer

> 一个 Codex 插件：通过 WPS Office 自动化，让 AI agent 能够**生成**和**编辑**高质量排版的 DOCX / PPTX / XLSX / PDF 文档。
>
> 当前版本：**0.5.0**（Windows 10/11 + WPS Office 12.1 实测验证；macOS 生成与转 PDF 经 JSAPI 门禁验证）

## 为什么需要这个插件？

python-docx、openpyxl、python-pptx 等库通过直接拼接 OOXML 来生成文档，但它们有一个根本性缺陷：**无法计算排版**。分栏后的文本流如何分布、浮动图片周围怎样环绕、表格跨页后表头是否重复——这些问题只有真正的排版引擎才能回答。

WPSComposer 的做法是直接驱动 WPS Office（或 MS Office）的排版引擎：Windows 走 COM 接口，macOS 走 WPS JSAPI 回环桥。agent 只需描述"我要什么"，引擎负责计算"怎么排"，最终输出和你在 WPS 界面里看到的一模一样。

## 核心能力一览

| 能力域 | 说明 |
|---|---|
| **Markdown → 排版文档** | 一行 `generate()` 把 Markdown 渲染成带设计预设的 DOCX / PPTX / XLSX / PDF |
| **对话式精排** | `inspect` / `edit` 读取并补丁已有文档的任意元素格式，可接管 WPS 中正在编辑的文档 |
| **结构化编辑** | insert / remove / move / clone 四种结构动词，覆盖段落、表格、形状、幻灯片、行列、工作表 |
| **稳定 ID 寻址** | `@paraId` / `@id` / `@name`，结构变动后仍能命中原元素 |
| **原子批处理** | 任一丁补丁失败则整批不落盘，源文件零改动 |
| **dump → replay** | `snapshot_to_patches()` 把一份文档的格式快照重放到另一份文档 |
| **Office → PDF** | 六种格式（doc/docx/xls/xlsx/ppt/pptx）跨平台转换，macOS 容器暂存 + 原子发布 |
| **PDF 编辑** | 合并 / 拆分 / 提取 / 旋转 / 水印 / 文本抽取 |

## 功能特点详解

### 1. 真实排版引擎渲染

- **Writer**：分栏排版、浮动文本框（带文字环绕）、艺术字、页眉页脚、自动目录与页码字段、项目符号/编号列表、标题自动编号
- **Sheet**：公式、数字格式、冻结窗格、条件格式、图表（柱/折/饼）、合并单元格
- **Slide**：16:9 版式、14 种布局模板、演讲者备注、形状与图片
- **表格**：合并 / 底纹 / 边框，跨页表头重复
- 排版规范默认遵循中文公文/政企标准：正文仿宋 12pt 两端对齐首行缩进 2 字符 1.5 倍行距，标题黑体加粗，西文 Times New Roman

### 2. 五种设计预设

`academic`、`consultant`、`business`、`tech`、`proposal`——一键切换整套配色、字体、表格样式与标题层级风格：

```python
generate("报告.md", format="docx", preset="academic")
generate("演示.md", format="pptx", preset="business")
```

### 3. 对话式精排（inspect / edit）

打开已有文件，或直接接管 WPS 窗口里正在编辑的文档（`attach_active`）。`inspect()` 以 JSON 返回文档树：每个段落/单元格/形状的文本与完整格式（字体、字号、颜色、行距、缩进、底纹……）。`edit()` 对指定目标做精确补丁，**不影响无关属性**：

```python
from skills.WPSComposer import inspect, edit

tree = inspect("报告.docx")
edit("报告.docx", output="修订.docx", patches=[
    {"target": "paragraph:1", "font": {"size": 20, "bold": True}},
    {"target": "table:1/cell:2,1", "fill": {"color": "#FFF2CC"}},
])
```

### 4. 稳定 ID 寻址（0.5.0 新）

位置寻址（`paragraph:3`）在结构变动后会漂移。0.5.0 引入宿主分配的稳定 ID：

- Writer：`paragraph:@paraId=3EF26A1B`（取自 docx 的 `w14:paraId`，WPS 保存时保留）
- Slide / Sheet：`shape:@id=7`、`shape:@name=Rectangle 3`

`inspect()` 自动回报稳定 ID；先插入段落再用 `@paraId` 打补丁，仍命中**原来的**那一段。无稳定 ID 的旧文档自动回退位置寻址。

### 5. 结构化编辑动词（0.5.0 新）

`edit(ops=[...])` 支持四种结构动词，三端全覆盖：

| 动词 | Writer | Slide | Sheet |
|---|---|---|---|
| insert | 段落/标题/分页/表格/图片/文本框 | 幻灯片/文本框/图片 | 行/列/工作表 |
| remove | 段落/形状/表格 | 幻灯片/形状 | 行/列/形状/图表/工作表 |
| move | 段落/形状/表格 | 幻灯片/形状（跨页） | 行/列/工作表 |
| clone | 段落/形状/表格 | 幻灯片/形状 | 行/工作表 |

位置语义完整：`"end"` / `"start"` / `{"after": ...}` / `{"before": ...}` / `{"index": N}`，含 WPS `MoveTo` 真实语义（插到移动前第 N 张之前）的完整象限处理。整表删除有最后一张可见工作表保护。

### 6. 原子批处理与结构化错误

- `edit(atomic=True)`：任一补丁失败 → 不写盘、源文件不变、`saved=False`
- 错误分类机器可读：`invalid_target`（附 `valid_forms` 自愈建议）/ `invalid_value` / `apply_failed`，`PatchError.reports` 携带逐条报告
- 内置帮助：`patch_grammar()` 返回全部目标语法与补丁字段；`validate_target()` 预检目标并给出最接近的合法形式

### 7. dump → replay 格式迁移

```python
snap = inspect("模板.docx")
patches = snapshot_to_patches(snap, dimensions=("font",))
edit("新文档.docx", output="套用模板格式.docx", patches=patches)
```

快照自动重写为位置寻址（稳定 ID 不可跨文档解析），`None` 值字段自动丢弃，不兼容字段落入 `rejected` 列表而非报错。

### 8. 跨平台转换与 PDF

- Windows：COM 直连；macOS：WPS JSAPI 桥 + 私有容器暂存 + fsync + 原子替换发布，无需完全磁盘访问/辅助功能权限
- PDF 编辑纯 Python（pypdf/pdfplumber/reportlab），全平台可用

## 原理架构

```
用户对话
  │
  ▼
Markdown 源文本
  │  md_parser.py 解析
  ▼
StructuredDocument（结构化文档模型）
  │  格式渲染器 (writer/sheet/slide_renderer)
  │  + 设计预设 (design_presets) + 排版规范 (reference_styles)
  ▼
Composer 引擎（WriterComposer / SheetComposer / SlideComposer）
  │  Windows: COM 接口 ／ macOS: WPS JSAPI 回环桥
  ▼
用户请求的单一产物：DOCX / PPTX / XLSX / PDF
```

## 快速开始

### 安装

```bash
git clone https://github.com/NeoMei/WPSComposer.git
cd WPSComposer
python3 install.py            # 首次安装（macOS 会安装锁定的 JSAPI 运行时）
python3 install.py --force    # 覆盖更新
```

Windows 也可 `pwsh ./install.ps1`，macOS/Linux 可 `./install.sh`。安装后重启 Codex Desktop，在 Plugins 目录启用 `wps-composer`，然后在 Codex 中引用 `$WPSComposer`。

### 运行时要求

- Python 3.9+
- Windows 生成/编辑 DOCX/PPTX/XLSX：WPS Office 或 MS Office + `pywin32`
- macOS：公开 `generate()` 与 `convert_to_pdf()` 已启用（Node.js 20+）；对话式编辑无 macOS 后端
- PDF 编辑：`pypdf` + `pdfplumber`，文本水印另需 `reportlab`

## 用法

### 一行生成文档

```python
from skills.WPSComposer import generate

generate("报告.md", format="docx", preset="academic")
generate("演示.md", format="pptx", preset="business")
generate("数据.md", format="xlsx")
generate("报告.md", format="pdf",  preset="consultant")
```

### 编辑已有文档

```python
from skills.WPSComposer import inspect, edit, attach_active

# 接管 WPS 中正在编辑的文档，读取光标处格式
w = attach_active("writer")
print(w.inspect_selection())

# 批量补丁 + 另存 + 导出 PDF 核对；原子模式任一失败不落盘
edit(
    "报告.docx",
    output="报告-修订.docx",
    export_pdf="报告-修订.pdf",
    atomic=True,
    patches=[
        {"target": "paragraph:@paraId=3EF26A1B", "font": {"size": 20, "bold": True}},
        {"target": "table:1/cell:2,1", "fill": {"color": "#FFF2CC"}},
    ],
)
```

### 结构化编辑

```python
edit("f.docx", output="o.docx", ops=[
    {"op": "insert", "type": "table", "props": {"rows": 2, "cols": 2,
         "data": [["a", "b"], ["c", "d"]]}},
    {"op": "clone", "target": "paragraph:1", "to": "end"},
    {"op": "move", "target": "paragraph:2", "to": {"after": "paragraph:4"}},
])

edit("f.pptx", output="o.pptx", ops=[
    {"op": "clone", "target": "slide:1/shape:@id=7", "to": {"slide": 2}},
    {"op": "move", "target": "slide:1", "to": {"after": "slide:3"}},
])

edit("f.xlsx", output="o.xlsx", ops=[
    {"op": "insert", "type": "sheet", "props": {"name": "Summary"}},
    {"op": "remove", "target": "sheet:1/cell:C1", "axis": "column"},
])
```

目标语法与补丁字段完整清单见 [API 参考](skills/WPSComposer/references/api.md)。

### Office 转 PDF

```python
from skills.WPSComposer import convert_to_pdf

pdf = convert_to_pdf("季度报告.xlsx")                          # 同目录同名 .pdf
pdf = convert_to_pdf("演示稿.pptx", "exports/演示稿.pdf", overwrite=True)
```

缺失源文件抛 `FileNotFoundError`，已有目标默认抛 `FileExistsError`，运行时错误抛 `ConversionError`（含 `code`/`source`/`component`/`backend` 字段）。

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
| 稳定 ID 寻址 | ✅ `@paraId` | ✅ `@id`/`@name` | ✅ `@id`/`@name` |
| 结构动词 insert/remove/move/clone | ✅ | ✅ | ✅ |
| 原子批处理 | ✅ | ✅ | ✅ |
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
│       ├── document_api.py          # inspect/edit/apply_ops/validate_op 编排
│       ├── conversion.py            # convert_to_pdf() 跨平台入口
│       ├── artifact_transport.py    # 校验、fsync、原子发布
│       ├── windows_conversion.py    # Windows COM 转换后端
│       ├── macos_probe/             # macOS WPS JSAPI、容器暂存与转换
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
├── fixtures/                        # COM 实机验证脚本（verify_*.py）
├── install.py / install.ps1 / install.sh
└── README.md
```

## 质量保障

- 平台无关 pytest 套件：**600 passed, 11 skipped**（Windows 与 macOS 均绿；跳过项为 POSIX 专属与可选依赖）
- 活 WPS COM 验证：`fixtures/verify_*.py` 8 个脚本全过（错误分类、原子不落盘、稳定 ID、边缘目标、dump→replay、结构动词全矩阵 + 精确落点）
- 完整验证记录与已修复 bug 清单见 [docs/windows-verification.md](docs/windows-verification.md)

## 开发调试

在本项目目录里直接修改 `skills/WPSComposer/scripts/` 下的文件。运行测试：

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'   # Windows: .venv\Scripts\python
.venv/bin/python -m pytest -v
```

macOS JSAPI 门禁前先在 `macos/wps-jsapi-probe/` 执行 `npm ci`。调试满意后运行 `python3 install.py --force` 更新本地插件。

COM 实机验证（Windows + WPS）：二进制 fixtures 已被 gitignore，新克隆需先生成：

```powershell
.venv\Scripts\python fixtures\make_fixtures.py
.venv\Scripts\python fixtures\add_extras.py
.venv\Scripts\python fixtures\verify_g.py   # 及其他 verify_*.py
```

## License

MIT
