# WPSComposer

Rich-layout document composer driven by WPS Office COM automation. Generates
DOCX / PPTX / XLSX / PDF with full layout control (multi-column, floating shapes,
WordArt, shaded/merged tables, charts, auto-updated TOC and fields).

It can also open existing files or attach to the document currently visible in
WPS, inspect element-level formatting as JSON-compatible data, and apply precise
partial patches without resetting unrelated formatting.

```python
from skills.WPSComposer.scripts.wps_engine import attach_active

w = attach_active("writer")
print(w.inspect_selection())
w.apply_format_patch(
    "paragraph:3",
    font={"size": 11, "color": "#333333"},
    paragraph={"alignment": 3, "line_spacing": 20},
)
w.save_current()
```

## 目录结构

`
WPSComposer/
├── .codex-plugin/
│   └── plugin.json              # 插件清单（name: WPSComposer）
├── skills/
│   └── WPSComposer/
│       ├── SKILL.md             # skill 入口（frontmatter name: WPSComposer）
│       ├── references/
│       │   └── api.md           # 四个 Composer 的 API 参考
│       └── scripts/
│           └── wps_engine.py    # 核心引擎（WriterComposer/SlideComposer/SheetComposer/PdfComposer）
├── install.ps1                  # 一键发布到 ~/.codex/skills/
└── .gitignore
`

## 工作流：调试 → 发布

1. **调试**：在本项目目录里直接修改 \skills/WPSComposer/scripts/wps_engine.py\
   等文件。用项目里的测试脚本（放在任意子目录）验证排版效果。
2. **发布**：调试满意后运行
   \\\powershell
   pwsh ./install.ps1            # 首次安装
   pwsh ./install.ps1 -Force     # 覆盖更新已发布版本
   \\\
   脚本会把插件拷贝到 \~/.codex/skills/WPSComposer\，Codex 即可加载。

## 运行时要求

- Windows + WPS Office（Writer/Spreadsheets/Presentation），或 MS Office 作为回退
- Python 运行时带 \pywin32\、\pypdf\、\pdfplumber\

## 引擎概览

\\\python
from wps_engine import WriterComposer, SlideComposer, SheetComposer, PdfComposer

with WriterComposer() as w:
    w.set_columns(2)
    w.add_floating_textbox("note", 80, 280, 150, 50)
    w.save_docx("out.docx")
    w.export_pdf("out.pdf")
\\\

详见 \skills/WPSComposer/references/api.md\。
