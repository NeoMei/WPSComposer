# Microsoft / WPS 能力对齐设计

状态：用户已于 2026-09-08 确认；按原生适配优先路线进入实现。
日期：2026-09-08。基线：v0.9.0 / 6dd3a00dff226096ad963cc68a65088865541c25。

## 目标与完成定义

让使用者选择 WPS Office 或 Microsoft Office 后，能以相同的公共调用完成当前 WPSComposer 提供的同类工作。目标覆盖 Windows 和 macOS 的 Word、Excel、PowerPoint：生成、原生排版、转换、检查、格式修改、结构编辑、活动文档操作及保存/恢复。

“对齐”比较 WPSComposer 暴露且实际实现的业务能力，不是 Office 软件全部菜单。两种引擎应有一致的输入、输出、错误、覆盖与文档保护语义；允许宿主字体度量和分页不同，不允许丢内容、把可编辑对象静默栅格化或把失败写成成功。

以 WPS 已实现能力的并集建立目标清单，并记录其当前平台限制。Microsoft 两端以该清单为目标；平台差异必须显式列项、调查和验收，不能通过缩小 WPS 基线或静默 fallback 宣称对齐。确实无法实现的项目需要单独说明并重新确认范围，否则保持未完成。内部 COM 对象指针并非跨平台结果契约；现有业务方法与可检查的文档行为才是对齐对象。

## 现场审查结论

- public generate/convert_to_pdf 已接受 engine=wps/msoffice/auto；office_engines.py 显式拒绝 Microsoft spreadsheet/presentation。
- inspect/edit/open_document/attach_active 当前没有明确的 engine 参数，不能靠构造文件扩展名获得可验证的 Microsoft 路由。
- Windows SheetComposer 已有公式、合并、边框、冻结窗格、条件格式、图表、页眉页脚，以及行列/工作表结构操作；SlideComposer 已有文本、图形、图片、表格、备注、版式以及页/对象结构操作。可复用业务语义，但须隔离 WPS quirks 与 Microsoft 原生绑定。
- macOS WPS 的 GenerationPlan 白名单比直接 COM Composer 更窄；公共文件检查支持文字、表格、演示，公共文件编辑入口当前只将 PPTX set patches 路由给 JSAPI。结构操作和 active attachment 仍走 COM。不能把 README 的概括性支持描述当成两端全部相同。
- macOS Microsoft Word 仍明确拒绝原生公式、合并/带单元格语义表格、横向媒体节、多列图及部分自定义样式参数；这些也是对齐工作，不能仅补 Excel/PPT 后结束。
- 本机 Word / Excel / PowerPoint 都安装了 16.112.3。只读检查其应用自带 sdef：Excel 有 workbook/range/chart 等对象，PowerPoint 有 slide/shape/table/selection 等对象。这证明有可调查的自动化入口，不代表所有操作已经原生验证。
- 独立 PDF 编辑器继续共享；不要求为 merge/split/rotate/watermark 再造一个 Office 后端。

代码依据：skills/WPSComposer/__init__.py、scripts/office_engines.py、document_api.py、generation_plan.py、writer.py、sheet.py、slide.py、macos_probe/inspection.py、msoffice/macos_script.py。

## 业务能力清单

| 领域 | 对齐目标 | 当前主要差距 |
|---|---|---|
| Word | 六级标题、目录、分节/页码、图片、表格、公式、引用/索引、列布局及可编辑对象 | 补齐 Mac Word 拒绝项，增加明确的 Microsoft 检查/编辑入口 |
| Excel | 多工作表、单元格和区域、公式与计算值、格式/合并/边框、行列尺寸、冻结、条件格式、图表、打印排版 | Microsoft 公共生成/转换未接入；Mac 需原生适配；检查与编辑需路由 |
| PowerPoint | 页面尺寸、文本/图形/图片/表格、布局与主题、备注、对象和页面顺序 | Microsoft 公共生成/转换未接入；Mac 需原生适配；检查与编辑需路由 |
| 共同编辑 | 文件与选区检查、稳定定位、set/insert/remove/move/clone、保存副本/原子发布、编辑后检查 | 三应用明确的 Microsoft 会话；宿主支持的撤销与事务边界 |
| 转换 | 对照现有 WPS 转换白名单逐格式处理；保持原文件、拒绝意外覆盖 | Excel/PowerPoint 原生 PDF 导出及旧格式读取 |
| 生命周期 | 只控制已绑定文档，保留未保存用户文档，截止时间、诊断、清理与恢复 | 不同应用进程/窗口模型；不能复制 Word 专属身份逻辑 |

每个操作还要拆分为 target 类型与属性维度，例如 cell/region 的 number format、font、fill、border，slide/shape 的位置、尺寸、文本、填充。清单绑定源码符号、对应平台、输入样例和验收结果，不以方法数量冒充覆盖率。

## 技术路线比较

### A. 原生适配优先（推荐）

Windows 使用 Word/Excel/PowerPoint 的 COM 对象模型，复用现有 Composer 的业务操作；将宿主发现、身份绑定、创建/打开/附着、保存/导出、退出规则分应用实现。

macOS 保留 Word 的 AppleScript 路径，为 Excel 和 PowerPoint 分别实现原生对象适配。先用应用自带 scripting dictionary 建立能力表，再以真实保存、重开、PDF 和编辑结果确认。公共 API、计划/操作校验、资源管理、原子输出、诊断与报告共享。

优点：沿用现有架构和本地原生排版，安装负担较小，Windows 可复用较多代码。风险：Mac 各应用的脚本覆盖范围和保存/选择语义不同；PowerPoint 的实例与窗口模型需要单独验证，不能假设 DispatchEx 就是隔离进程。

### B. 三应用统一 Office.js 加载项

Office Add-ins 支持跨平台且具有应用专属 API；可评估作为 Mac 原生 API 缺口的补充。它需要额外 manifest、Web 应用和部署/权限/会话绑定机制，不能假设不同客户端 API 集合相同，也不能未经验证就承诺全部 WPS 操作覆盖。

优点：将来可共享一部分跨平台实现。代价：增加安装及宿主会话基础设施，且无法直接继承当前原生适配器的验收。建议仅在能力探针证明具体操作无法通过路线 A 达到目标时，提出针对性补充设计。

来源：[Excel object model](https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model)、[PowerPoint object model](https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint/object-model)、[Office Add-ins overview](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)。

## 公共接口方向

保留包名 WPSComposer、现有导入与默认 engine=wps；扩展现有 msoffice/auto 到三应用。在 inspect、edit、open_document、attach_active 中增加显式 keyword-only engine；获得的会话绑定引擎，apply_patches/apply_ops 不重新猜测。没有匹配活动文档时明确报错，不能启动另一个应用并假装已附着。

能力发现独立于实际执行，按 engine/platform/component/operation/target/属性返回支持状态与限制。auto 在变更前按请求所需能力选择，默认 WPS 优先；选定后不再切换。发现同一应用中的多个活动候选时要求明确绑定，不以窗口标题碰运气。

复用 M5 Word 管线；为 spreadsheet/presentation 使用现有录制计划与原生适配契约，补齐所需业务操作，禁止将所有工作塞入 Word adapter 或只移除 engine 守卫。跨平台公共业务能力稳定，COM 原始对象仅作为 Windows 实现细节保留。

## 原生验收规则

- 每个应用、每个平台分别生成原生可编辑文件、关闭、重开、检查语义，再导出 PDF 进行文字/图像/版面检查。基线 WPS 与 Microsoft 使用同一输入和断言。
- Excel：真实公式及重开后的计算值，多表引用、合并/格式、图表数据绑定、打印区域/分页；不能只检查 XLSX ZIP 能打开。
- PowerPoint：页序、对象数与类型、文字/图片/表格、备注、几何与布局；实际编辑对象、撤销、明确保存并重开，不能仅查看 PDF。
- Word：新增高级能力与已锁定 M5 标题/编号/目录/分节/页码/长表格同时回归。
- 编辑：按 target/verb 覆盖增删移克隆与格式修改；结构变更后重新检查稳定定位；活动文档的复合原子修改在无可靠回滚能力时仍须拒绝，不能破坏现有保护契约。
- 所有宿主保留一个不属于测试的状态哨兵；只关闭测试拥有的文档与已核实可退出实例。无法确认退出时保留恢复文件，不全局 kill、修改 Normal 模板、降低宏安全或运行文档携带的宏。
- 原生失败/未测明确记入矩阵；平台跳过、模拟对象通过、能创建空白文件，都不计为功能验收通过。

## 交付顺序与放行

1. 冻结 WPS 行为矩阵，审查各公开能力与平台差异；建立 Microsoft 三应用两平台的原生探针与失败证据。
2. 统一显式引擎路由及会话契约；实现 Windows Excel/PowerPoint 生成与转换，并补齐 Word 业务缺口。
3. 实现 Mac Excel/PowerPoint 原生适配与 Word 高级能力。发现 API 缺口即提出具体补充路线，不将它静默从矩阵移除。
4. 实现三应用文件检查/格式编辑，再实现结构编辑和活动文档操作；按已有保护契约逐项验证。
5. 运行同输入跨引擎对比、Windows/Mac 原生 UI 与故障回归、安装验证、多轮代码/任务审查。
6. 全部接受项有证据后更新 capability metadata、description、README 和版本，才宣布 Microsoft 与 WPS 能力对齐。可以分批提交实现，但不将首批可生成文件等同于完成目标。

现有文档分支中已经修正当前能力描述；这些修正不会把尚未实现的 Microsoft 能力提前宣传为支持。已发布 v0.9.0 标签保持不变。
