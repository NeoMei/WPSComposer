# Long-form M0 原生能力闸门

本闸门只验证 WPSComposer 0.8.0 长文档引擎所依赖的原生能力，不修改公共 `generate()` 路由，也不生成 M1 产品代码。当前 macOS 原生证据已经取得；**Windows native gate is pending**，因此当前总体决策仍是 no-go，不能进入 M1。

## 判定规则

能力 1–14 是强制项：Windows 与 macOS 必须全部为 `passed`，合并结果才是 `go`；任一平台、任一强制项失败或未运行均为 `no-go`。能力 15（SVG）是可选项：两端都通过才纳入 0.8.0，否则从本版本排除，但不单独阻塞 1–14。

| ID | 原生能力 | M0 验收重点 |
|---:|---|---|
| 1 | 封闭协议与资源清单 | 版本、固定资源摘要和未知字段均严格校验 |
| 2 | Windows 进程归属 | `DispatchEx`、不可变进程身份、真实协作超时且用户进程存活；macOS 记为不适用 |
| 3 | 字体发现与替代 | CJK、Latin、等宽字体均来自 WPS 字体集合，PDF 中存在字体资源 |
| 4 | Unicode 与 UTF-16 范围 | 组合字符、emoji、扩展汉字和两种规范化形态 |
| 5 | 坐标语义 | WPS 点坐标与 PDF 位置误差不超过 1 pt |
| 6 | 原生多级编号 | 中文、十进制、混合三套四级编号，插入/移动/删除后重新编号 |
| 7 | 分节与页码 | 封面/正文、横竖版往返、显式分页、页码重启 |
| 8 | 原生目录 | 标题排除、三级 TOC 样式及紧凑段后距，避免样例目录间距过大 |
| 9 | 题注与图表目录 | `SEQ` 字段、图目录/表目录和重置语义 |
| 10 | 交叉引用 | 稳定书签、冲突重试和 `REF` 字段 |
| 11 | 原生公式 | 可编辑 OMath 置于无边框容器 |
| 12 | 跨页布局 | 跨页表格逐行定位、物理页片段和 UTF-16 范围 |
| 13 | 局部降级 | 清理失败子对象、在原位置插入可见标注并继续后续排版 |
| 14 | 收敛与导出 | 有界字段刷新、保存、关闭、重开、再刷新、再保存和一次 PDF 导出 |
| 15 | SVG | 固定清单中的 SVG 以 WPS 原生图片对象插入；不支持时排除 |

文档中本来没有图像时按正常版式输出；图片插入或局部对象排版失败时，只在该对象的检查点内降级，并在相应位置加入可见标注。WPS 引擎不存在、协议/资源不匹配、保存或导出失败时直接终止平台探针，不继续伪造结果。

## 运行命令

输出目录必须是新目录或空目录。原生运行只公开用户请求的 DOCX/PDF 探针产物和脱敏 JSON；临时请求、PID 和绝对路径不会进入证据。

macOS 必须在允许访问 WPS 容器的 unrestricted local run 中执行：

```bash
.venv/bin/python -m skills.WPSComposer.scripts.longform_m0 \
  --platform macos \
  --output-dir build/longform-m0/macos-<run-id> \
  --timeout 600
```

Windows 需要已安装 WPS Writer、Python 3.9+、`pywin32`、`pypdf`、`pdfplumber` 和 Pillow。运行前请在独立的 WPS 进程中打开一份可识别的测试文档（separate user WPS document），用于证明超时清理不会关闭用户文档：

```powershell
.venv\Scripts\python.exe -m skills.WPSComposer.scripts.longform_m0 `
  --platform windows `
  --output-dir build\longform-m0\windows-<run-id> `
  --timeout 600
```

两个平台都完成后合并：

```bash
.venv/bin/python -m skills.WPSComposer.scripts.longform_m0 \
  --platform verify \
  --windows-evidence build/longform-m0/windows-<run-id>/platform-evidence.json \
  --macos-evidence build/longform-m0/macos-<run-id>/platform-evidence.json \
  --output-dir build/longform-m0/matrix-<run-id>
```

退出码：`0` 表示平台强制项全部通过或矩阵为 `go`；`1` 表示原生失败或矩阵为 `no-go`；`2` 表示参数或证据无效。

## 恢复与安全检查

macOS 重跑前先确认探针 registration 已恢复到运行前内容；若恢复失败，不要覆盖容器中的注册文件，应保留失败证据并人工比较备份。Windows 重跑前确认原先打开的用户文档仍在、内容未变化，且其进程身份未变化。探针只能终止再次核对 PID、可执行文件、创建时间和父 PID 后仍完全匹配的自有进程树；不得终止任何无法证明归属的 WPS 进程。

如果 WPS 引擎未安装或无法启动，先安装/修复 WPS 后使用全新的输出目录重跑。若对象级能力失败，打开 `probe.docx` 查看红色失败标注或降级标注；不得用 OOXML 后处理掩盖原生失败。

## 视觉证据与隐私

在 WPS 中重新打开 DOCX、刷新字段并另存后，再检查 PDF。代表性截图应覆盖封面/目录、多级编号、横版页、图表目录、公式、跨页表格、局部降级页和 SVG 页。截图只以 relative filename 记录，不能写绝对路径。

证据 JSON must not contain 文档正文、源文件路径、临时目录、用户名、环境变量、书签映射、字段哈希、资源载荷哈希或进程命令行。允许记录的内容限于能力状态、脱敏错误码、WPS 版本、相对文件名、交付产物 SHA-256、页数/页面框/旋转、字体名、坐标误差和计数指标。
