# MD → Word 章节编号关联：复发调查与防回归方案

调查日期：2026-10-14。代码基线：`a639382ffa7064de3765ce28875d8125a65a6c05`（0.9.3）。

本次完成代码、历史验收、安装入口与离线复现调查；尚未收到本次失效的 MD / DOCX，因此没有认定该文档的根因，也没有宣称完成新的原生验收。本文记录已落地的结构性防线和仍需环境门禁的部分。

## 结论

“初次显示有编号”与“编辑后章节及子孙编号仍联动”是不同的验收条件。现有修复分别覆盖了解析、引擎绑定和后处理，但最终交付链没有对这项用户承诺实行统一的硬检查；历史动态验收也未成为持续门禁。多个安装入口又使运行的修复版本难以确认。

应把编号关联作为交付条件：来源可追溯、语义明确、最终 DOCX 结构正确、原生编辑回放通过。编号有歧义或检查失败时，不能继续以成功产物交付。当前已落地的是前缀解析统一和 DOCX 结构门禁；原生编辑回放仍需环境门禁。

## 已验证的证据

### 1. 已修复的解析缺口（保留复现证据）

`longform/semantic.py` 的 `_scheme_level_match` 接受 `1.1：范围`，但 `_strip_prefix` 的 decimal 二至四级、hybrid-bid 二至三级规则要求编号后有空白。hybrid-bid 四级的 `关键工法01：` 已支持冒号，不属于此缺口。当前离线生成计划的实际结果：

| 输入（位于 `# 报告 / ## 1 总体` 下） | 规范化结果 | 诊断 |
|---|---|---|
| `### 1.1 范围` | 正文二级标题，`numbering=decimal`，标题文本“范围” | 无 |
| `### 1.1：范围` | 修复前：`numbering=none`；修复后：`numbering=decimal`，标题文本“范围” | 修复前为 `HEADING_PREFIX_AMBIGUOUS` |

第二章的 `2.1：步骤` 在修复前同样取消原生编号；现在识别和剥离共用边界规则，回归测试覆盖中英文冒号。这是已确认并已修复的代码缺口，但还不能认定是本次用户文档的触发条件。

复现（仓库根目录，无需启动 Office）：

```python
from skills.WPSComposer.scripts.longform.pipeline import build_longform_generation

for heading in ("1.1 范围", "1.1：范围"):
    build = build_longform_generation(
        f"# 报告\n\n## 1 总体\n\n### {heading}\n\n"
        "## 2 实施\n\n### 2.1 步骤\n"
    )
    print([(s.level, s.heading, s.numbering)
           for s in build.semantic.document.sections if s.heading])
    print([issue.code for issue in build.issues])
```

### 2. 历史修复曾在不同层互相覆盖

- `66d4b9bf` 修复 M5 路由切换后，题名消费与标题层级导致正文被当作前言的问题。
- 原生 Word 曾为不同标题级别生成不同列表实例：插入一级章后，章号变成 2，但子节仍为 1.1 / 1.1.1。证据见 `docs/verification/msoffice-production/windows/round3-deb67ba/report.md`。
- `116e3e49` 在 M5 生成完成后添加 legacy `apply_native_numbering()`；`ed2f0327` 又移除这一后处理，因为它覆盖原生关联并使目录及质量证据过期。当前 `orchestrator.py:205` 已解释这个约束。

需要一个负责从语义到原生列表的编号契约；原生生成、字段刷新与校验完成之后不得再用另一套规则改编号。

### 3. 有过真实动态验收，但常规门禁没有承接

历史 Windows Word / WPS 验收实际覆盖过插入、删除、保存、重开和 PDF，共三套方案、两个章节、四级标题；macOS 也检查过子孙编号联动。不能把历史工作概括成“只测字符串”。

可复用的真实编辑脚本：

- `docs/verification/msoffice-production/windows/round4-shared-numbering/helpers/numbering_schemes_native.py`
- `docs/verification/msoffice-production/macos-word/structural-shared-outline/run.py`
- `docs/verification/msoffice-production/macos-word/structural-shared-outline/validate.py`

当前 M5 native runner 的交付目标是 PDF（`longform_m5_evidence.py:162`），未把 DOCX 编辑关联列为门禁；原生测试默认也会跳过。新加入的 `chinese-outline` 不在上述历史三方案动态验收范围。

`fixtures/verify_msoffice_production.py:198` 已有 effective `numId` 一致性检查及分裂列表负例，可作为实现基础；它还没有进入每次 public `generate()` 的交付检查。`longform/platform_runtime.py:234` 的发布验证主要检查 Office 包有效性，不能据此证明关联正确。

### 4. 运行入口身份不一致，目录名不可信

本机实查：

| 入口 | 实际状态 |
|---|---|
| 当前工作仓库 | `a639382f` / 0.9.3 |
| `~/.codex/skills/WPSComposer` | 链接目录名为 `702f709...`，但该目录的实际 Git HEAD 已是 `a639382f`；关键编号文件哈希与工作仓库一致 |
| `~/.codex/plugins/wps-composer` | 0.9.1，缺少后续修复 |
| `~/.superwriter/dependencies/WPSComposer/current` | 指向 `ac839952...` / 0.9.2，仍有后来被撤销的 M5 后处理 |

因此“本地仓库最新”不能代表所有入口最新；也不能把旧目录名等同于旧代码。本会话列出的技能入口实际含最新修复。本次生成到底使用哪条入口仍需原始调用或产物证据。

关联的 SuperWriter 解析器存在额外风险：`current` 对应缓存不是 Git checkout，读取其 revision 出错会进入旧缓存回退；安装再用该缓存重设技能链接可能造成倒退。该行为经只读解析及内存 remote-read stub 验证，未实际安装或改写链接。修复属于 SuperWriter 的依赖管理边界。

### 5. 初始基线与最终验证

执行：

```sh
.venv/bin/python -m pytest tests/longform tests/longform_m2 tests/longform_m3 -q
```

修改前的聚焦基线是 **888 passed、4 failed、7 skipped**。四个失败均为 M3 快照；逐项比较确认仅为标题操作新增 `keepWithNext: true`，未有其他参数或操作数量变化。随后同步了这些已提交行为对应的快照和测试断言，并修正了一个低于当前 64 单位阈值的测试数据。

最终全套验证为 **2883 passed、13 skipped**；跳过项是 PyMuPDF、WPS 组件注册和真实 macOS/Windows WPS 环境门控。合成 DOCX 结构门禁和语义编号回归均在全套结果内通过；这仍不替代真实 Office 插章、删章、保存重开验收。

## 按顺序落地的防回归方案

### A. 固定输入编号契约并在生成前阻断歧义

1. 统一前缀识别与剥离，让两者共享解析结果，避免各写一套正则。
2. 为每个预期编号的标题保留来源、有效级别、所属序列、编号方案以及前缀处理结果。前言、附录、明确 `heading_numbering: none` 分别遵循既有契约。
3. 不能仅从最终 `numbering=true` 的标题反推检查对象，否则被误降级成 `none` 的标题会逃过检查。源结构中要求编号的章节出现歧义或缺失父级时，应带位置报错，不静默降级后继续交付。
4. 保留自动连续编号的既有语义；不在未确认需求时把手写源数字（如 1 后接 3）解释为必须保留的跳号要求。

验收：空格 / 中英文冒号、中文章节、题名消费、前言、跳级、四套方案均有正反例；特别防止标题“外观有数字、实际无关联”。

### B. 已落地：将最终 DOCX 的编号结构检查接入交付流程

1. `longform/numbering_validation.py` 以源计划为预期，通过稳定书签定位真正的正文标题，排除目录缓存。
2. 解析段落及样式继承后的有效 `numId`、`ilvl` 和 abstract 定义。同一编号层级树必须属于同一有效序列；检查起始值和方案契约。
3. 检查计划要求编号却未绑定、透明标题仍绑定列表、手写与原生前缀重复等情况。
4. 生命周期在首次生成、重排版和质量提示补丁后、PDF 导出前执行；失败会阻断发布。PDF 输出也先验证其来源 DOCX；PDF 文本检查只证明显示结果。
5. 故障注入已覆盖分裂列表、`numId=0`、缺失 abstract、错误级别、错误起始值和残留字面前缀。

静态结构检查无法独自证明 Word / WPS 编辑行为，仍须 C。

### C. 把历史真实编辑脚本提升为持续原生验收

从 public `generate()` 开始：至少两章、每章多节、多级标题，并穿插正文、表格、普通列表和分页。依次检查：

1. 初次生成的全部标题编号与目录。
2. 插入一级章：后续章及全部子孙同步变化。
3. 在章内插入二级节：同级及子孙变化，其他章按方案正确重启。
4. 删除、撤销、保存、关闭重开后再次读回全部编号和目录。
5. 导出 PDF，核对显示与原生读回一致。

分别覆盖 `decimal`、`chinese-formal`、`hybrid-bid`、`chinese-outline`，以及 macOS / Windows 的 Word / WPS 支持路径。普通测试缺乏原生环境可跳过，但发布检查必须列为待验，不计通过。复用已有脚本，避免每次修复重新编写一次性演示。

### D. 固定运行来源与修复验收标准

每次诊断记录实际模块路径、源码提交或内容哈希、版本、引擎、应用版本、编号方案、源与产物哈希；不增加未请求的公开伴随文件，也不在日志保存客户正文。

缓存必须以内容身份核验，不允许把一个以旧提交命名的缓存目录直接更新成新提交。安装应明确哪些入口已更新，遇到缓存损坏与网络异常分别处理。发布和本机安装验收分开记录，避免“仓库修好”被当成“用户已在用”。

本次问题关闭条件：取得真实 MD/DOCX 与实际调用入口 → 最小复现成为永久 fixture → 新交付门禁能拦住旧错误 → 修复后真实插入/删除/重开回放通过。未完成前保持待验。
