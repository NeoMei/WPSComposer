# Microsoft Word 原生排版

公共 `generate()` 和 `convert_to_pdf()` 使用 `engine="msoffice"` 调用本机桌面版 Microsoft Word。文档仍经过 WPSComposer 的 Markdown 解析、M5 文档计划、分页质量检查和原子发布流程。

```python
from skills.WPSComposer import generate, convert_to_pdf

generate("report.md", output="report.docx", engine="msoffice", open_result=True)
generate("report.md", format="pdf", output="report.pdf", engine="msoffice")
convert_to_pdf("existing.docx", "existing.pdf", engine="msoffice", timeout=600)
```

## 引擎选择

- `wps`：默认值，使用 WPS。Windows 公共生成/转换不再通过 Office ProgID 自动退到 Word。
- `msoffice`：使用 Microsoft Word，支持 DOCX/PDF 生成及 DOC/DOCX 转 PDF。
- `auto`：在开始任务前检测安装，优先 WPS，再选择支持该格式的 Word。任务开始后不会因排版、保存或质量检查失败而切换引擎。

Excel、PowerPoint 生成继续使用 WPS。MS Word 的活动文档编辑、`layout_engine: legacy` 和 Excel/PowerPoint 后端不属于本接口。现有直接 Composer API 保持原有签名。

## 安装要求

Windows 使用与 Microsoft Word 注册一致的 Python 环境及 `pywin32`，例如安装项目的 `windows` 可选依赖。仅检查到 `Word.Application` 不足以证明可用：运行时还核对真实 WINWORD 进程和窗口身份。WPS 占用 Word ProgID 时会明确失败，不能当作 Word 验收通过。

macOS 使用 `/Applications/Microsoft Word.app`。请先正常启动 Word 完成首次初始化，并允许运行 Python 的终端或应用通过 macOS 自动化控制 Word。原生 Word 执行本身不使用 WPS 加载项；完整插件安装器同时安装 WPS 运行时，因此仍需要 Node.js 20+ 和 npm。自动 WPS 检测对应 `/Applications/wpsoffice.app`。

PDF 质量检查沿用项目的 PDF 依赖。必须能读取并验证原生导出的 PDF，不能跳过质量检查直接返回 DOCX。

## 支持范围

两端支持常规标题层级、正文、封面、目录、原生标题编号、原生表格和图片、页眉页脚及分页信息。Word 与 WPS 分别执行原生排版，页数、字体替代及未指定颜色的主题表现可以不同。

macOS Word 会在启动原生任务前明确拒绝尚未实现的高级操作，例如原生公式、合并或带单元格引用的语义表格、横向媒体节、多列组合图、混合编号方案，以及未实现的自定义页面/样式参数。不能用自定义 OOXML 渲染器绕过此错误；可调整输入或明确选择 WPS。可见的资源缺失等计划降级会保留到质量报告。

## 文件与失败恢复

默认 `overwrite=False`，输出已存在时拒绝覆盖。转换操作处理源文档的私有副本。`open_result=True` 在成功发布及清理后打开结果：DOCX 使用本次选择的应用，PDF 使用系统默认阅读器。

Word 操作遵守一个总超时预算。Windows 超时只终止任务的 Python 工作进程，不终止 Word；未确认关闭的文档与诊断保留在错误信息指定的暂存目录。

macOS 不会退出共享 Word。超时或不能确认清理完成后，会隔离本次任务并拒绝新的 Word 任务。关闭错误所指向的任务文档后，显式执行：

```bash
python -m skills.WPSComposer.scripts.msoffice.macos_runtime --recover --timeout 30
```

恢复命令核查旧脚本进程和本次暂存文档；仍有不确定的任务时拒绝解除隔离。保留的诊断用于排查，恢复命令不会关闭用户文档。
