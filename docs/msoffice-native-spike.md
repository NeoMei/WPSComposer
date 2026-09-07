# Microsoft Office 原生 Word/PDF 原型验证（2026-09-08）

结论：macOS Word 原生 AppleScript 路线已通过本次有界能力原型，两次独立运行完成原生创建、编号、目录更新、82 行表格、保存重开与 PDF 导出。建议进入 Word 执行器的正式设计；不能据此宣称完整 M5 或双平台 Office 已支持。Windows 已完成原生实测、未保存哨兵保护及最终异常清理修复复验：同一候选 f0dc91f 的 run09/run10 均通过，证据提交 f317b16 已由控制端独立核对。

## 范围与版本

- 基于 WPSComposer 0.8.1 / master `54e277b`；独立分支 `codex/msoffice-native-spike`。
- Mac 实机：macOS 26.6.2 arm64；Microsoft Word 16.112.3。
- Windows 实机：系统 build 26100；Microsoft Word LTSC 16.0.17932.20910；64 位 Python 3.11.9、pywin32 312。
- 只有独立调查脚本与文档；没有修改 `generate()`、`convert_to_pdf()`、生产源代码或本机已安装插件。
- Mac 通过 AppleScript、Windows 通过已核对 WINWORD.EXE 进程身份的 COM 原生对象模型完成产物创建、保存及 PDF 导出；没有让 WPS 或文件生成库代替 Word 排版，也没有使用宏。

## 能力矩阵

| 验收项 | Mac Word 实测 | Windows Word |
|---|---|---|
| 明确指定 Microsoft Word | 通过，Word 16.112.3 | 通过；实际 WINWORD.EXE，窗口 PID 与独立进程一致 |
| 中文段落字体、字号、缩进 | 指定样例段落实际 FangSong 12pt，firstLine=480 twips、firstLineChars=200 | 通过；仿宋 12pt、firstLine=480、firstLineChars=200 |
| 分级标题字号 | H1=16pt、H2/H3=15pt 保存后持久化 | 通过；H1=16pt、H2/H3=15pt 持久化 |
| 原生标题编号 | Heading1/2/3 样式关联原生 numId，PDF显示1/1.1/1.1.1/2 | 通过；原生 numId/abstractNum 关联及 PDF 编号 |
| 目录与字段更新 | 原生TOC和4个PAGEREF，目录包含层级及结果页码 | 通过；原生 TOC 和 4 个 PAGEREF |
| 长表格 | 原生表格82行3列，80条数据和末尾标记全部保留 | 通过；82 行 3 列，80 条记录及末尾标记 |
| 重分页和PDF | 两次均9页；每页显示重复表头，允许行跨页 | 通过；9 页，逐页重复表头，允许行跨页 |
| 保存、关闭、重开、导出 | 通过；重新打开精确目标路径并检查末尾标记 | 通过；原生保存、重开精确路径、PDF 导出 |
| 不干扰已有未保存文档 | 一份44字符专用未保存哨兵，前后名称/路径/saved状态/全文均相同 | 通过；62 字符专用未保存哨兵前后名称/路径/状态/哈希一致，明确不保存关闭；范围限已注册实例 |
| 无补操作的文件发布 | 两次在Word沙箱内新建临时目录执行后由Python原样复制，无UI补操作 | 通过；新目录直接保存和导出，无 UI 补操作 |

## 实验过程与保留的失败

1. `mac-run-01`：直接遍历 Word documents 集合报 -1708；在创建任务文档之前退出。改为先读数量、逐个按序号访问后通过。
2. `mac-run-02`：DOCX和4页PDF生成通过；`convert to table` 报 -1708，实际DOCX没有表格，明确标为部分能力失败。保存到工作区新目录触发Mac Word沙箱授权；仅在UI中授权这一轮专用输出目录，故这轮不算全自动文件流程。
3. `mac-run-03`：改为Word原生建表后逐个填写单元格；在Word自己的 `Data/tmp/msoffice-spike-*` 新目录保存、重开和导出，再由Python复制产物。所有原型操作通过，9页PDF，未发生UI补操作。
4. `mac-run-04`：同一脚本、全新工作区输出目录及全新Word沙箱目录再次通过；9页、82行表格、原生目录字段和字体检查一致，哨兵仍未保存且全文不变。
5. 验收结束后核对哨兵标记和未保存状态，明确不保存关闭这份测试文档。Word恢复到原先0份打开文档的状态，应用保持运行。

## 证据与复跑

原始证据位于本分支工作区 `build/msoffice-spike/`：

- `baseline-pytest.log`：2624 passed，12 skipped，170.48秒。更早的61项排版契约基线通过。
- `mac-run-01/result.json`、`mac-run-02/result.json`：原始失败与部分成功记录。
- `mac-run-03/result.json`、`mac-run-04/result.json`：真实原生执行、每项状态、原生保存路径、SHA256、已有文档保护记录。
- 两轮的 `probe.docx`、`probe.pdf`、`artifact-verification.json`、`pdf-text.txt`；run03另有 `acceptance.json`。
- `page-1.png`、`page-2.png`、`page-9.png`：由原生Word PDF渲染；已目视检查run03首页、中间表页和末页。
- `controller-baseline.json`：初始0份文档与专用哨兵信息。原生沙箱临时目录保留作证据，不作为正式安装方案。

```sh
python3 fixtures/msoffice_spike/mac_word.py --output-dir build/msoffice-spike/mac-next
```

Windows在已登录桌面、已安装Microsoft Word和pywin32的机器执行（输出目录必须不存在）：

```powershell
python fixtures/msoffice_spike/windows_word.py --output-dir build/msoffice-spike/windows-run-01
```

Windows脚本的退出码0只表示原生操作完成，仍需检查DOCX结构、PDF内容与版面。实测原始证据、失败记录、产物校验和及哨兵保护见 [Windows 报告](verification/windows-word-spike/report.md)，控制端复核与异常清理补丁见 [独立复核](verification/windows-word-spike-controller-review.md)。上表记录的能力在最终清理补丁 f0dc91f 的 run09/run10 中再次通过；[最终 Windows 复验](verification/windows-word-spike/final-cleanup-review/report.md) 保留精确脚本来源、命令、产物和非空哨兵检查。

## 对接入方案的影响

- Mac Word基础执行器优先采用AppleScript是可行候选，当前这组能力不需要VBA宏或Office.js加载项。
- 采用任务拥有的Word文档及Word沙箱内暂存区，避免依赖活动文档或每次请求访问用户输出目录。正式实现还需原子发布、异常清理、取消与超时回收；本次脚本仅保留诊断证据，不具备生产事务保证。
- 正式执行器应消费现有GenerationPlan并返回ExecutionOutcome/分页映射，再接入已有质量检查、有限重排与发布生命周期；本次没有接入这些生产接口。
- Windows应显式选择软件，贯穿生成、重排和导出；现有COM方法可复用，但需要Word独立验收，不能从WPS的通过记录推断Word也通过。
- Word进程隔离、会话串行化和文档归属需要按平台设计。当前Mac实测共享Word进程但仅操作自己创建的文档，未证明独立Word进程方案。

## 尚未证明的能力

本次是执行能力原型，不是正式文档模板验收：标题仍沿用Word主题颜色/字体，只有指定正文样例段落应用仿宋，未覆盖完整封面、页眉页脚、正文全局样式及图表公式；未证明编辑后自动重编号、跨页目录变化收敛、完整节点分页/坐标映射、M5质量提示补丁、失败取消回收、冷启动无额外授权、并发、不同Office版本或Excel/PowerPoint。重复运行证明当前已启动且可自动化的Word环境下流程可重复，不代表所有首次安装环境免授权。

Windows 异常清理补丁的复验门已关闭。进入正式实现前，还需对 Mac 补充复杂对象及分页映射探测；Office.js仍作为原生接口缺口明显时的备选，需要独立验证而非直接替换。

## 审查与当前状态

任务审查与整体审查均通过，无Critical/Important问题；两轮原生脚本与HEAD一致、原生文件与工作区产物SHA256一致均获独立核对。两项非阻塞诊断改进留待正式适配器阶段：Mac超时后JSON未解析已执行阶段，但保留原始stderr且总结果失败；Windows早期快照失败仅归入core错误。Windows非Windows平台保护已实测拒绝运行且未生成Office产物。

原型与 Windows 证据、审查修复已推送独立 codex 分支，未合并主分支或安装。Windows 的成功原生运行和合成未保存哨兵门已获产物证明；清理修复 f0dc91f 通过 2630 项平台无关测试（12 项跳过）、独立复审以及 Windows run09/run10 最终原生复验。控制端核对最终原始脚本、全部 65 项证据哈希及 25 项最终复验子清单、DOCX 语义和 PDF 全部 9 页像素，与已目视检查版面一致。首轮遗留 Word 进程启动归属仍未完全证明，没有强制关闭；不宣称完整会话清理或正式双平台产品支持完成。
