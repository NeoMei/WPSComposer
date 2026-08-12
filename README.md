# WPSComposer

**WPS-backed rich document composition for AI agents**

WPSComposer 是一个强大的文档生成工具，让 AI agent 能够通过 WPS Office 生成高质量排版的 DOCX、PDF、XLSX、PPTX 文档。

## ✨ 核心特性

### 🎨 专业排版
- **样式驱动**：所有段落/字符样式一次性定义，通过样式名称引用
- **中文优化**：遵循中文公文/政企方案标准（仿宋 12pt、黑体标题等）
- **多语言支持**：支持中文、英文、日文、法文等
- **设计预设**：5 种快速切换风格（academic、consultant、business、tech、proposal）

### 📊 丰富内容
- **分栏排版**：支持多栏布局
- **浮动元素**：文本框（带文字环绕）、艺术字
- **表格**：合并/底纹/边框表格、自动生成目录和页码字段
- **公式**：数学公式支持
- **图表**：Excel 冻结窗格、条件格式、图表

### 🖼️ 插件系统
- **Excalidraw 插件**：自动将 `.excalidraw.md` 文件渲染为 PNG 图片并嵌入文档
- **可扩展**：支持自定义插件，在 Markdown 解析前预处理内容

### 🌍 跨平台
- **Windows**：通过 COM 接口驱动 WPS Office / MS Office
- **macOS**：通过 JSAPI 容器驱动 WPS Office
- **统一 API**：相同的 Python 接口，跨平台一致体验

## 🚀 快速开始

### 安装

```bash
# 克隆仓库
git clone https://github.com/NeoMei/WPSComposer.git
cd WPSComposer

# 安装到 Codex 个人插件 marketplace
python3 install.py            # 首次安装（macOS 会安装锁定的 JSAPI 运行时）
python3 install.py --force    # 覆盖更新
```

### 基本用法

```python
from skills.WPSComposer import generate

# 生成 DOCX 文档
generate("report.md", format="docx", preset="academic", output="report.docx")

# 生成 PDF 文档
generate("report.md", format="pdf", preset="consultant", output="report.pdf")

# 生成 PPTX 演示文稿
generate("slides.md", format="pptx", preset="business", output="slides.pptx")

# 生成 XLSX 电子表格
generate("data.md", format="xlsx", output="data.xlsx")
```

### 使用插件

```python
from skills.WPSComposer import generate

# 启用 Excalidraw 插件（自动渲染 .excalidraw.md 为 PNG）
generate(
    "notes.md",
    format="pdf",
    preset="consultant",
    plugins=["excalidraw"],
    overwrite=True,
)
```

## 📖 详细文档

### Markdown 语法支持

WPSComposer 支持标准 Markdown 语法，并扩展了以下功能：

#### 基本语法
- **标题**：`# H1`、`## H2`、`### H3` ...
- **粗体**：`**bold**`
- **斜体**：`*italic*`
- **删除线**：`~~strikethrough~~`
- **代码**：`` `inline code` ``
- **链接**：`[text](url)`
- **图片**：`![alt](url)`

#### 扩展语法
- **任务列表**：`- [ ] task`、`- [x] done`
- **表格**：标准 Markdown 表格语法
- **代码块**：带语法高亮
- **引用块**：`> blockquote`
- **水平线**：`---`、`***`、`___`

#### Obsidian 支持
- **Wikilink 图片**：`![[path]]`、`![[path|width]]`、`![[path|widthxheight]]`
- **Vault 根目录自动检测**：通过 `.obsidian` 文件夹定位

### 设计预设

| 预设名称 | 风格描述 | 适用场景 |
|---------|---------|---------|
| `academic` | 学术论文风格 | 研究报告、学位论文 |
| `consultant` | 咨询报告风格 | 商业分析、方案文档 |
| `business` | 商务风格 | 企业报告、提案 |
| `tech` | 技术文档风格 | API 文档、技术手册 |
| `proposal` | 提案风格 | 项目提案、投标书 |

### 插件系统

#### 内置插件

**Excalidraw 插件** (`excalidraw`)

自动将 Obsidian Excalidraw 图形渲染为 PNG 图片并嵌入文档。

**输入 Markdown：**
```markdown
## 系统架构

![[系统架构图.excalidraw]]

![[合作分工流程图.excalidraw|800]]
```

**插件自动处理：**
1. 扫描所有 `![[*.excalidraw]]` 引用
2. 用 Playwright + Excalidraw 官方导出 API 渲染为 PNG（无工具栏）
3. 替换为标准 Markdown 图片语法：`![name](/absolute/path/name.png)`
4. 传给 WpsComposer 正常生成 PDF/DOCX

**依赖安装：**
```bash
pip install lzstring playwright
playwright install chromium
```

#### 自定义插件

```python
from skills.WPSComposer import register_plugin

def my_plugin(content: str, base_dir: str) -> str:
    """修改 Markdown 内容，返回修改后的内容。"""
    return content.replace("foo", "bar")

register_plugin("my_plugin", my_plugin)

# 使用
generate("input.md", format="pdf", plugins=["my_plugin"])
```

**插件接口：**
```python
PluginFunc = Callable[[content: str, base_dir: str], str]
```

- `content`：原始 Markdown 文本
- `base_dir`：Markdown 文件所在目录（用于解析相对路径）
- 返回值：修改后的 Markdown 文本

### API 参考

#### `generate()` 函数

```python
def generate(
    source: str,
    format: str = "docx",
    preset: Optional[str] = None,
    output: Optional[str] = None,
    source_is_text: bool = False,
    plugins: Optional[List[str]] = None,
    timeout: float = 600,
    overwrite: bool = False,
) -> str:
```

**参数说明：**

| 参数 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| `source` | `str` | - | Markdown 文件路径或文本（`source_is_text=True` 时） |
| `format` | `str` | `"docx"` | 输出格式：`"docx"`、`"pptx"`、`"xlsx"`、`"pdf"` |
| `preset` | `Optional[str]` | `None` | 设计预设名称 |
| `output` | `Optional[str]` | `None` | 输出文件路径（自动生成） |
| `source_is_text` | `bool` | `False` | 将 `source` 视为文本而非文件路径 |
| `plugins` | `Optional[List[str]]` | `None` | 插件列表 |
| `timeout` | `float` | `600` | WPS 生成超时（秒） |
| `overwrite` | `bool` | `False` | 是否覆盖已存在的输出文件 |

**返回值：** 生成的文件绝对路径

**异常：**
- `ValueError`：未知格式或预设名称
- `FileNotFoundError`：源文件未找到
- `FileExistsError`：输出文件已存在（除非 `overwrite=True`）

#### 其他函数

```python
# 查看可用格式
list_formats() -> list  # ["docx", "pptx", "xlsx", "pdf"]

# 查看可用预设
list_available_presets() -> list  # ["academic", "consultant", ...]

# 查看可用插件
list_plugins() -> list  # ["excalidraw", ...]

# 注册自定义插件
register_plugin(name: str, func: PluginFunc) -> None
```

## 🔧 技术原理

### 架构设计

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
  │  通过 COM 接口驱动 WPS Office（Windows）
  │  或通过 JSAPI 容器驱动 WPS Office（macOS）
  ▼
用户请求的单一产物：DOCX / PPTX / XLSX / PDF
```

### 核心组件

#### 1. Markdown 解析器 (`md_parser.py`)
- 零外部依赖的轻量级解析器
- 支持 CommonMark 风格 Markdown
- 支持中文内容和 YAML frontmatter
- 扩展支持 Obsidian wikilink 语法

#### 2. 文档模型 (`document_model.py`)
- `StructuredDocument`：顶层容器（元数据 + 有序章节）
- `Section`：章节（标题级别 + 标题 + 块级元素）
- 块级元素：`Paragraph`、`ListBlock`、`TableBlock`、`CodeBlock`、`ImageBlock` 等
- 行级元素：`Span`（文本 + 格式属性）

#### 3. 渲染器 (`renderers/`)
- `writer_renderer.py`：DOCX/PDF 渲染
- `sheet_renderer.py`：XLSX 渲染
- `slide_renderer.py`：PPTX 渲染
- 样式驱动：所有样式通过 `reference_styles.py` 统一定义

#### 4. Composer 引擎
- **Windows**：`wps_engine.py` 通过 COM 接口驱动 WPS Office
- **macOS**：`macos_probe/` 通过 JSAPI 容器驱动 WPS Office
- 统一抽象：相同的操作接口，不同平台实现

#### 5. 插件系统 (`plugins/`)
- 在 Markdown 解析前运行
- 接收原始文本和基础目录
- 返回修改后的文本
- 支持懒加载和自定义注册

### 排版规范

默认遵循中文公文/政企方案标准：

| 元素 | 字体 | 字号 | 其他 |
|------|------|------|------|
| 正文 | 仿宋 | 12pt | 两端对齐，首行缩进 2 字符，1.5 倍行距 |
| H1 标题 | 黑体加粗 | 16pt | 居中 |
| H2/H3 标题 | 黑体加粗 | 15pt | 左对齐 |
| 西文 | Times New Roman | - | - |
| 代码 | Consolas | 10pt | 灰色背景 |
| 表格 | 仿宋 | 10pt | 灰色表头，斑马纹 |

##  依赖要求

### 运行时要求

- **Python**：3.9 或更高版本
- **Windows**：WPS Office 或 MS Office，`pywin32`
- **macOS**：WPS Office 12.1.26035 或更高版本，Node.js 20+（JSAPI 运行时）
- **PDF 编辑**：`pypdf` + `pdfplumber`，文本水印额外需要 `reportlab`

### 可选依赖

```bash
# Excalidraw 插件
pip install lzstring playwright
playwright install chromium

# PDF 编辑功能
pip install pypdf pdfplumber reportlab

# Windows COM 支持
pip install pywin32
```

## 🧪 测试

```bash
# 运行测试
pytest

# 运行特定测试
pytest tests/test_md_parser.py
pytest tests/test_writer_renderer.py
```

## 📝 更新日志

### v0.7.1 (2026-08-12)
- 🐛 文章标题（md 第一个 `#`）不再进入正文、目录和编号：仅渲染在封面，正文从第一节开始
- 🐛 目录标题（"目  录"）不再被目录自身收录（改用无大纲级别的 Body Text 样式）

### v0.7.0 (2026-08-12)
- ✨ 原生标题编号：生成的 DOCX 使用 Word/WPS 原生多级编号（第一章→第一节→一、→（一）），增删或移动章节后编号自动重排
- ✨ `generate()` 生成后自动应用，无需手工处理；标题样式新建段落自动接续编号
-  幂等处理，失败不阻断生成

### v0.5.0 (2026-07-31)
- ✨ 新增插件系统
- ✨ 新增 Excalidraw 插件（自动渲染 .excalidraw.md 为 PNG）
-  支持 Obsidian wikilink 图片语法
- ✨ 支持自定义插件注册
-  增加默认超时时间到 600 秒
- 🔧 添加 `overwrite` 参数支持覆盖输出文件
-  修复图片路径解析（支持空格路径）
- 🐛 修复 Excalidraw 渲染空白图片问题
-  修复 HTML 模板中 JSON 数据转义问题

### v0.4.0 (2026-07-30)
- 初始版本发布
- 支持 DOCX、PDF、XLSX、PPTX 生成
- 支持 5 种设计预设
- 支持 Windows COM 和 macOS JSAPI

##  贡献

欢迎提交 Issue 和 Pull Request！

### 开发环境设置

```bash
# 克隆仓库
git clone https://github.com/NeoMei/WPSComposer.git
cd WPSComposer

# 创建虚拟环境
python3 -m venv .venv
source .venv/bin/activate

# 安装开发依赖
pip install -e ".[dev]"

# 运行测试
pytest
```

## 📄 许可证

MIT License

## 🙏 致谢

- [WPS Office](https://www.wps.com/) - 提供强大的文档处理能力
- [Excalidraw](https://excalidraw.com/) - 提供手绘风格图表工具
- [Playwright](https://playwright.dev/) - 提供浏览器自动化能力

## 📞 联系方式

- GitHub: https://github.com/NeoMei/WPSComposer
- Issues: https://github.com/NeoMei/WPSComposer/issues
