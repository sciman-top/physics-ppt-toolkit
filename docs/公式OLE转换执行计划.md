# 公式 OLE→原生公式转换执行计划（MathType/OLE 专项）

版本：v1.3
状态：执行中（B0 取消；B1–B5 全部完成：13.3 达成 16/16、14.2 达成 3/3，judge 两轮验收 PASS；剩余可执行切片仅可选 B6）
更新时间：2026-09-16
上游文档：[`产品需求与工程路线图.md`](产品需求与工程路线图.md)（产品边界）· [`公式处理说明.md`](公式处理说明.md)（行为说明）· [`公式识别转换实施计划.md`](公式识别转换实施计划.md)（程序级设计基线）

---

## 1. 目标与非目标

### 1.1 目标终态

真实课件中所有**可确证**的 MathType/OLE 公式（`Equation.DSMT4`）转换为微软原生公式（OfficeMath/OMML），字体/字号/颜色符合既有统一体系；无法确证的每一项都有明确归宿（保留原状并说明原因）。策略为 **OLE 先行、图片公式暂缓**：

- 内容真值永远来自 config `formulaWhitelist`（`targetUnicodeMath`/`targetTex`）；AI 视觉只产生**候选与证据**，不直接写回。
- 每批转换都是新副本，全套门禁通过才算数；失败删本批输出、从原始 PPTX 重跑。

### 1.2 量化指标

| 指标 | 定义 | 当前基线 | 目标 |
|---|---|---|---|
| FormulaConversionRate | 真公式 OLE 中 `Converted` 占比 | 13.3：11/16；14.1：5/8 | 13.3 达 16/16；14.1 除动画绑定组外全部转换或有不可转证据 |
| DispositionRate | 全部 OLE 有终态归宿（Converted/OriginalKept+原因/Skipped+原因） | 未统计 | 100% |
| 错误接受 | 写回后内容与审定 canonical 不一致 | 0 | 保持 0（任何一起即停机回滚） |

### 1.3 非目标（明确不做）

- **图片公式自动转换**（FormulaImage/MixedImage 的识别、归一化与写回）。重启触发条件：活跃课件 OLE 侧只剩保护性跳过项，且某批样本独立公式图数量 ≥10 张，或用户显式授权启动 F2 评测。
- 自研 MTEF 二进制解析（`公式处理说明.md` §7 既有边界）。
- 动画绑定组合并 OLE 的强转（14.1 slide6 三段式类）：timing spid 悬空 = PowerPoint 拒开，`Apply-FormulaOmmlForOle.ps1` 的 `OleTimingDangling` 拒绝逻辑是保护，不改。
- 默认一键流程行为变化：本计划所有新能力走独立显式入口，`Invoke-PhysicsPptWorkflow.ps1` 默认链不静默改变。
- 小结页之外的新页型治理。

---

## 2. 现状基线（2026-09-14 工作区，行号以此为参照）

### 2.1 已验证能力（可直接复用）

| 环节 | 工具 | 状态 |
|---|---|---|
| 载体盘点 | `tools/Export-FormulaCarrierInventory.ps1` | 可用；records 含 `source.carrier/slide/shapeId/shapeName/bbox`、`recordId`、input sha256 绑定 |
| GoldSet→映射 | `tools/Export-FormulaOleMapping.ps1` | 可用；校验 ReviewStatus=Approved、SourceFormulaText 精确命中白名单、EvidencePath 存在、哈希绑定 |
| 候选片段 | `tools/Export-FormulaOmmlCandidates.ps1` | 可用；UnicodeMath→AST→FormulaIR/OMML/MathML/TeX |
| OLE 写回 | `tools/Apply-FormulaOmmlForOle.ps1` | 可用；mc:AlternateContent 等界盒替换、Fallback 保全、cNvPr id 沿用、m:nor 角色化样式 |
| 结构校验 | `tools/FormulaOfficeMathValidator`（dotnet） | 可用；门禁要求 0 错误 |
| 参考对照 | `tools/Compare-FormulaConverters.ps1`（Pandoc，只读） | 可用 |
| 官方转换表 | `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL` / `OMML2MML.XSL` | 本机已验证存在；可作结构 diff oracle（任务 T4） |
| 页型判定 | `tools/Normalize-PhysicsPpt.ps1` `Get-SlideKind`（:592） | 可用；已有 Cover/Ending/Resource/ContentSection/ExtensionSection/Exercise/AppendixText/Normal |

真实样本基线：13.3 v17（18 OLE → 11 转；2 非公式；5 条数值代入链待转）；14.1 v38（8 OLE → 5 转；slide6 三段式动画绑定保留；slide21/22 图片公式遗留）。

### 2.2 代码级缺口清单（本计划要闭的洞）

| ID | 缺口 | 位置 | 现状行为 |
|---|---|---|---|
| G1a | Unicode 上标字符（²³⁴⁵⁶⁷⁸⁹⁰ⁿ）无归一化 | `Export-FormulaOmmlCandidates.ps1` `Get-FormulaTokens`（:104） | 变成孤立 Char token，role 落 `Text`，**静默产出错误结构**（不抛错） |
| G1b | `/` 分母贪婪绑定 | 同文件 `Parse-FormulaSequence`（:179-188） | `J/(kg·℃)×50kg×20℃` 的 `×50kg×20℃` 全部进分母，**错误 AST 且无报错** |
| G1c | 负指数无规范路径 | 同文件 `Parse-FormulaAtom`（:150-167） | `10^-3` 把 `-` 当 script、剩余 `3` 触发 trailing 抛错（安全失败但不可用；`10^(-3)` 括号形式已可用） |
| G2 | OLE 裁剪图导出无工具 | （不存在；13.3 的 `ole-crops/` 是临时脚本产物） | 每批手工重做，证据不可再生成 |
| G3 | GoldSet 编制全靠人工填 CSV | 流程缺口 | 无 AI 视觉草拟 SOP，编制成本高、不可审计 |
| G4 | 小结页无页型 | ~~`Get-SlideKind`~~ | **已取消**（2026-09-14 用户决策：小结页不做例外，不排任何页型） |
| G4b | 页型排除未接入 OLE 链 | ~~盘点/映射~~ | **已取消**（同上） |
| G5 | 字号标定纯人工 | 流程缺口 | 无建议值辅助（注意：字号按字形实测，**不可按界盒估算**，工具只能给建议） |
| G6 | 批处理无编排 | （不存在） | 7 步手工串，断点不可续跑 |
| G7 | 尾等号白名单片段无生成路径 | `tools/Export-FormulaOmmlCandidates.ps1` 解析器 | `Q_吸=`、`cmΔt=`、`Q_放=` 等尾等号规则只能作匹配键；若 GoldSet 引用其作为生成目标，候选导出按设计抛错（fail-closed）。仅当真实 GoldSet 命中时才需要扩解析器，当前不修 |
| ~~G1a/G1b 遗留视觉缺陷~~ | ☑ 已修复（2026-09-16，13.3 v18 验收发现） | 上标脚本组括号泄漏（`10^(-3)` 渲染出字面括号）→ 解析时解包脚本 Group，canonical 层保持 `^(...)`；线性单位斜杠缺失 → 新增 `\/` 转义（`J\/(kg·℃)` 保持线性除号）；两者均有 fixture 与门禁 OMML 结构断言覆盖 |

---

## 3. 固定架构与数据流（所有任务不得违反）

```text
原始 PPTX（只读）
  → [B0] Normalize 报告（SlideKind，COM）
  → [B2] Export-FormulaCarrierInventory（Open XML，sha256 绑定源文件）
  → [T5] Export-FormulaOleCrops（按 bbox 从页面渲染图裁剪，只读）
  → [T6] AI 视觉草拟 GoldSet（Draft）→ 独立 agent/人工定稿（Approved）
  → Export-FormulaOleMapping（白名单精确匹配 + 证据 + 页型排除校验）
  → Export-FormulaOmmlCandidates（UnicodeMath→OMML/MathML/FormulaIR 片段）
  → Apply-FormulaOmmlForOle（副本写回，Fallback 保全）
  → FormulaOfficeMathValidator（0 错误）
  → PowerPoint 真实打开 + Export-PptxVisualAudit + judge 视觉验收
  → Compare-PptxInvariantSnapshot -AllowReviewedFormulaConversion
  → 批次 manifest/summary 收口
```

不变式（来自 AGENTS.md，逐条适用）：

1. 源 PPTX 永不覆盖；一切写回产生新文件。
2. AI/OCR 输出只是候选；草拟与定稿必须分离（同一会话不得自我验收）。
3. 文本内容、对象位置/尺寸、动画顺序、图片裁剪不变；组内对象不碰。
4. `tools/*.ps1` UTF-8 with BOM；CSV UTF-8 BOM；新工具必须纳入 `tools/Test-ToolkitFiles.ps1`。
5. `reports/` 交付走版本化目录 `reports/<课件名>_v<N>`（号最大最新），复用 `Invoke-PhysicsPptWorkflow.ps1` 现行版本解析行为，不在 `reports/` 内手工指定 OutputRoot。

---

## 4. 批次与任务清单

依赖关系：`B1 ∥ B2` → `B3(编排) 依赖两者` → `B4(GoldSet SOP) 依赖 B2` → `B5(样本批次) 依赖全部` → `B6(可选)`。（B0 已取消。）

### B0 页型排除：~~新增 Summary 页型~~（已取消）

**2026-09-14 用户决策：小结页不再做例外，OLE 程序不排除任何页型。** 本批次全部任务（T0-1 Summary 页型、T0-2 盘点注记、T0-3 映射排除与 `-IncludeSpecialSlides`）未合入即回退；编排器与 SOP 不含页型逻辑。若将来需要页型信息做证据展示（非排除），按当时的真实需求重立任务。

### B1 代入链 tokenizer 升级（`tools/Export-FormulaOmmlCandidates.ps1`）

#### T1-1 Unicode 上标归一化（闭 G1a）

- 新函数 `Convert-FromSuperscriptChars`，在 `Parse-FormulaAst`（:221）取 token 前调用：连续上标串（含 `ⁿ`、`⁻`）→ `^(合并内容)`，如 `10³` → `10^(3)`、`10⁻³` → `10^(-3)`；下标字符（₀₁₂…）同理 → `_(…)`。
- 验收：`Q吸=cmΔt=4.2×10³J/(kg·℃)×50kg×20℃=4.2×10⁶J` 归一化后可解析且 AST 含 Superscript 节点。

#### T1-2 分母绑定修复（闭 G1b）

- `Parse-FormulaSequence`（:179-188）`/` 分支：分母从 `Parse-FormulaSequence`（贪婪）改为**单原子** `Parse-FormulaAtom`（括号 Group、单 token、带 `_`/`^` 后缀的原子均可）；分母后返回外层 sequence 继续处理 `×…`。
- 同步约束 canonical 书写：`Convert-AstToUnicodeMath` 的 Fraction 输出改为 `(num)/(den)` 括号形式，保证 round-trip（parse→render→parse）AST 幂等；`Convert-AstToCanonicalTex` 不变（`\frac` 已正确）。
- **回归红线**：config 现有 33 条白名单 `targetUnicodeMath` 全部重新生成 fragments，OMML/MathML/TeX 与升级前逐字节一致（分母单原子的条目 AST 不受影响，必须证明）；`Compare-FormulaConverters.ps1` Pandoc 对照抽 3 条复跑。
- 验收 fixtures：新增 `examples/fixtures/formula-unicodemath.valid.json` 与 `.invalid.json`（模式仿 `formula-ir.*.json` + `Test-ToolkitFiles.ps1` 的 `Test-FormulaIrFixtureContract` :218 断言）。valid 至少覆盖：`η=Q_吸/Q_放`、`Q_吸=cmΔt`、`4.2×10^(3)J/(kg·℃)×50kg`、`10^(-3)`、多重等号链 `A=B=C`；invalid 至少覆盖：裸 `10^-3`（仍抛错）、`/3/4` 连除、空括号。

#### T1-3 书写规范落文档

- `docs/公式处理说明.md` §3/§4 增补：canonical UnicodeMath 禁裸 `^-`、分母含乘法链必须括号、上标一律 `^()`。仅文档，无代码。

### B2 OLE 裁剪图导出工具（闭 G2）

- 新文件：`tools/Export-FormulaOleCrops.ps1`（UTF-8 BOM，`Set-StrictMode Latest`，dot-source `PhysicsPpt.Common.ps1`）。
- 参数：`-CarrierInventoryJson`（必填，校验 input sha256 与源 PPTX 一致）、`-PagesDir`（`Export-PptxVisualAudit` 的 `pages/slide-NNN.png`）、`-OutputDir`、`-PageImageWidth`（默认 1600）、`-PaddingPx`（默认 4）、`-Carrier`（默认 MathTypeOle）。
- 实现：
  - 页宽换算不写死 12192000：解析 `ppt/presentation.xml` 的 `p:sldSz@cx`，`px = emu / sldSzCx * PageImageWidth`。
  - `Add-Type -AssemblyName System.Drawing` 裁剪（Windows PowerShell 5.1 兼容），越界 clamp 到页内。
  - 输出：`ole-crops/<recordId>_slide-NNN.png` + `ole-crops.csv`（recordId/Slide/ShapeId/ShapeName/bbox Emu/crop path/crop sha256/page png sha256）+ manifest JSON（`writeBackAllowed=false`）。
- 注册：`tools/Test-ToolkitFiles.ps1` 工具清单（:54-63 区域）加入；语法/编码断言自动覆盖。
- 验收：对 13.3 v17 链头跑通，crop 数 = MathTypeOle 数（18），图片可读、公式居中完整。
- 回滚：删输出目录 + revert 工具清单行。

### B3 批处理编排（闭 G6）

- 新文件：`tools/Invoke-FormulaOleBatch.ps1`（薄编排，不改默认一键链；UTF-8 BOM）。
- 参数：`-InputPath`、`-GoldSetCsv`、`-SlideKindCsv`（可选）、`-MaxItems`、`-SkipVisualAudit`、`-IncludeSpecialSlides`、`-Resume`。
- 步骤（每步产物落盘、记录 sha256，任一步失败即停并汇总已成功步骤）：
  1. `Export-FormulaCarrierInventory`（带 -SlideKindCsv）；
  2. `Export-FormulaOleCrops`；
  3. `Export-FormulaOleMapping`（页型排除生效）；
  4. `Export-FormulaOmmlCandidates`；
  5. `Apply-FormulaOmmlForOle`（输出=新副本）；
  6. `FormulaOfficeMathValidator`（dotnet run -c Release，0 错误门禁）；
  7. 未 `-SkipVisualAudit`：`Export-PptxVisualAudit` + `Export-PptxVisualConfirmation`；
  8. 批次 manifest：每步命令行、产物路径+sha256、`automationGateStatus`。
- `-Resume`：按 manifest 中既有产物 sha256 跳过已成功步骤；发现 sha 不符即失败（不允许静默混用旧产物）。
- 输出根：按 `reports/<课件名>_v<N>` 版本化（版本解析行为与 `Invoke-PhysicsPptWorkflow.ps1` 保持一致，必要时提取共享函数；不得在 reports/ 内手工指定 OutputRoot）。
- 注册：`tools/Test-ToolkitFiles.ps1` 清单。
- 验收：14.1 链头 dry-run（GoldSet 用已定稿的 5 行）产出与 v38 等价的产物集；中途杀进程后 `-Resume` 能续跑且最终 manifest 完整。

### B4 GoldSet 编制 SOP + AI 视觉草拟（闭 G3；流程文档，不新增写回能力）

- 新文件：`docs/公式GoldSet编制SOP.md` + `examples/fixtures/formula-goldset.sample.csv`（表头与 `Export-FormulaOleMapping.ps1` 注释一致：ReviewStatus,Slide,ShapeIds,WhitelistName,SourceFormulaText,SizePt,MainColorHex,SubColorHex,Note,EvidencePath）。
- SOP 要点（写进文档）：
  1. 草拟：执行 agent 读 `ole-crops/` 逐图提出 WhitelistName/SourceFormulaText/颜色/备注，SizePt 仅给建议值；产出 `ReviewStatus=Draft` 草稿。
  2. 定稿：**另一个** agent 或人工对照 crop + 整页图逐条核验后改 `Approved`；同一会话不得自我验收。
  3. `SourceFormulaText` 无法精确命中现有白名单 → 不批准，走 `Export-FormulaWhitelistSuggestions` + 人工确认进 config 的白名单晋升流程。
  4. 证据绑定：EvidencePath 指向 crop 与整页图双证据。
- 无代码改动（sample.csv 进 fixtures 需在 Test-ToolkitFiles 有断言则一并加，无则仅登记路径卫生）。

### B5 样本批次执行（收口验证）

| # | 内容 | 通过标准 |
|---|---|---|
| 1 | 13.3：5 条代入链白名单晋升 + GoldSet + 全链转换副本 | Validator 0 错；PowerPoint 打开 40 页；5 处 judge 前后对比 pass；FormulaConversionRate 16/16 |
| 2 | 14.1：slide6 三段式复评（确认 OleTimingDangling 仍拒绝，写入明确 OriginalKept 原因） | DispositionRate 100%，无强转 |
| 3 | 小结页排除 dry-run：GoldSet 混入小结页行 | 默认 Failed(`SlideKindExcluded`)，开关放行后 judge 正常 |
| 4 | 全局：`Test-ToolkitFiles.ps1` exit=0；`Test-PhysicsPptPolicy.ps1`（config/白名单有变更时）通过 | exit=0 |

每批交付走版本化目录；失败删本批输出目录，从原始 PPTX 重跑。

### B6（可选）字号建议辅助

- 新函数放 `Export-FormulaOleCrops.ps1` 内：按 crop 中字形像素高度（去白边后主体高度）→ 建议 SizePt（换算含 96dpi 假设，标注置信度）。
- **只进 GoldSet 的建议列，不进 SizePt 权威值**；权威值仍以 judge 视觉验收定版（已知教训：24pt 过小、字号按字形实测不可按界盒估）。用户未要求前可不做。

---

## 5. 全局门禁与收口

- 每个代码任务收口即跑最低门禁：`pwsh -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File tools/Test-ToolkitFiles.ps1`。
- 涉及 OMML 生成逻辑的任务（B1）另跑：构建 + `FormulaOfficeMathValidator`（对含新片段的样本 PPTX）。
- config/白名单变更（B5-1）另跑 `tools/Test-PhysicsPptPolicy.ps1`。
- 无新外部依赖（MML2OMML.XSL 为本机已有 Office 资产，Pandoc 链路已存在），`Assert-Toolchain -Deep` 不触发；若 B6 引入图像库之外的新依赖则必须补跑。
- 文档任务（T1-3、B4）至少 `git diff --check`。

## 6. 风险与对策

| 风险 | 对策 | 触发条件→动作 |
|---|---|---|
| T1-2 分母语义变更破坏既有 33 条白名单 | 逐字节回归 + fixtures | 任一条目 fragments 变化 → 停，改实现而非改白名单 |
| Summary 误判正文页 | 标题级精确匹配 + 判定置后 + 全页型回归对比 | 一起误判 → 收紧规则 |
| AI 草拟错读（放/吸颠倒等） | Draft/Approved 分离 + 白名单精确匹配 + judge | 一起错误接受 → 停机回滚本批，复盘 SOP |
| 代入链写回过宽/过窄溢出界盒 | 等界盒 + normAutofit + judge 逐处验收 | judge fail → 单处调 SizePt 重跑该页 |
| COM 瞬态失败 | 沿用 `Invoke-WithComRetry` 边界模式 | — |
| PowerPoint 保存包扰动误报门禁 | 不变量比较已是规范化比较；先比媒体 sha256 多重集 | — |

## 7. AI 编码执行提示（执行 agent 必读）

- 仓库 gotchas（来自既有实践，违反即返工）：
  - pwsh 7 binder：函数返回可枚举对象必须 `return ,$x`；`@()` 包 `List[object]` 不枚举、会把本体传给 .NET 方法。
  - COM 路径必须 `[System.IO.Path]::GetFullPath()`（相对路径按 PowerPoint 进程 CWD 解析）。
  - PowerPoint 保存会重新本地化默认形状名：跨 SaveAs 按名匹配需唯一文本兜底；本计划写回走 Open XML（cNvPr id 沿用），不依赖形状名。
  - CSV 一律 UTF-8 BOM（`Write-Utf8BomCsv`）；新 `tools/*.ps1` 一律 UTF-8 BOM；docs 为 UTF-8 无 BOM。
  - 版本化交付目录号最大最新；门禁查目录卫生（`_v<N>` 前邻 ASCII 字母即违规）。
- 改 `Export-FormulaOmmlCandidates.ps1` 时：先加 fixtures 与回归断言（T1-2 的 33 条逐字节对比），再改 parser；一次只动一个语法点。
- 改 `Normalize-PhysicsPpt.ps1` 时：`Get-SlideKind` 是 COM 热路径，保持幂等只读；新增判定不得抛异常冒泡（try/catch 归 Normal 并记报告行）。
- 禁止事项：不改 `Apply-FormulaOmmlForOle.ps1` 的拒绝语义（白名单/CurrentConfig/动画悬空）；不给默认一键链加 OLE 自动转换；不把 AI 读图结果写进任何 `ReviewStatus=Approved` 的文件。

## 8. 任务状态跟踪

| 任务 | 状态 | 完成证据（产物/门禁输出） |
|---|---|---|
| T0-1 Summary 页型 | ⊘ 已取消（随 B0） | 2026-09-14 用户决策 |
| T0-2 盘点页型注记 | ⊘ 已取消（随 B0） | 同上 |
| T0-3 映射页型排除 | ⊘ 已取消（随 B0） | 同上 |
| T1-1 上标归一化 | ☑ 完成 | `Convert-SuperscriptSubscriptCharsToAscii`；fixtures 含 `10^(3)`/`10^(-3)` 用例；门禁 parser 函数断言 |
| T1-2 分母绑定修复 | ☑ 完成 | 分母单原子 `Parse-FormulaAtom`；canonical Fraction 括号形式；valid/invalid fixtures + 门禁 round-trip 断言 |
| T1-3 书写规范文档 | ☑ 完成 | `docs/公式处理说明.md` §3/§4 规则行（禁裸 `10^-3`、显式乘法结束分母等） |
| B2 裁剪图工具 | ☑ 完成 | `tools/Export-FormulaOleCrops.ps1` 已入门禁；`reports/13.3比热容（王耀强）_v17/00_检查报告/ole-crops` 为真实产物 |
| B3 批处理编排 | ☑ 代码完成 | `tools/Invoke-FormulaOleBatch.ps1` 已入门禁；14.1 链头 dry-run 因源样本已被移出 `PPTX/` 暂无法复跑（样本恢复后即可） |
| B4 GoldSet SOP | ☑ 完成 | `docs/公式GoldSet编制SOP.md` + `examples/fixtures/formula-goldset.sample.csv`（表头门禁断言） |
| B5-1 13.3 代入链批次 | ☑ 完成（2026-09-16） | `reports/13.3比热容（王耀强）_v18`：链头=v17 交付副本；5 条代入链转换 + 5 条白名单晋升；16/16 达成；validator 0 错、不变量 0 阻断、judge 两轮验收（第 1 轮 FAIL 挖出并修复解析器上标括号泄漏/线性除号/长链溢出三缺陷，第 2 轮 PASS）。执行期间临时经 LFS 恢复 13.3 样本，收口后已删除还原 |
| B5-2 14.1 slide6 复评 | ☑ 已被 v5 覆盖 | v5 实际交付优于原定 OriginalKept：动画 timing id 保全后 slide6 三段式全部转换成功；29/29 页导出、validator 0 错、不变量 passed（40 项均为授权的 OLE→OMML 载体差异） |
| B5-3 小结页排除 dry-run | ⊘ 已取消（随 B0） | OLE 程序不排除任何页型 |
| B6 字号建议（可选） | ☐ 未开始 | 用户未要求 |
| 14.2 热机效率迁移批次 | ☑ 完成（2026-09-16，计划外新增） | `reports/14.2热机、效率（王耀强）_v32`：用户放入的新样本，3/3 OLE 转换（热机效率定义式/燃料放热量公式/气体燃料放热量公式）；validator 0 错、不变量 0 阻断、judge PASS；B3 批处理编排首次真实链头实跑并修复三缺陷 |
