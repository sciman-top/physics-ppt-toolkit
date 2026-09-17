# 公式 GoldSet 编制 SOP（MathType/OLE 迁移批次）

版本：v1.0
更新时间：2026-09-14
上游文档：[`公式OLE转换执行计划.md`](公式OLE转换执行计划.md) · [`公式处理说明.md`](公式处理说明.md)

## 1. 目的与边界

GoldSet 是 OLE→原生公式迁移的**人工审定 canonical 源**：每一行把一个（或一组）`Equation.DSMT4` OLE 对象绑定到白名单公式名、证据图和目标样式。`Export-FormulaOleMapping.ps1` 只接受 `ReviewStatus=Approved` 的行，且 `SourceFormulaText` 必须精确命中当前 config 白名单。

- AI 视觉只能产出 `Draft` 草稿；草稿与定稿必须分离（不得由同一会话自我验收）。
- 全程不写 PPTX；写回只发生在后续 `Apply-FormulaOmmlForOle.ps1` 的独立副本步骤。

## 2. 文件与列契约

CSV 文件（UTF-8 BOM），列与 `Export-FormulaOleMapping.ps1` 校验一致：

| 列 | 必填 | 说明 |
|---|---|---|
| ReviewStatus | 是 | `Draft` 或 `Approved`；只有 `Approved` 进入映射 |
| Slide | 是 | 1 起始的页码（演示顺序） |
| ShapeIds | 是 | 该页内 OLE 形状 id，分号分隔（多 OLE 合并一条公式时用；动画绑定组合并会被 `OleTimingDangling` 拒绝） |
| WhitelistName | 是 | 必须是 config `formulaWhitelist` 中已存在的 `name` |
| SourceFormulaText | 是 | 从裁剪图转录的线性文本；`Test-FormulaWhitelistMatch` 必须以该文本命中 `WhitelistName` 对应规则 |
| SizePt | 是 | 8–96 整数；可用裁剪清单的 `SuggestedSizePt` 作起点，由 judge 视觉验收定版 |
| MainColorHex | 是 | 6 位十六进制，无 # |
| SubColorHex | 否 | 中文下标色；空则同主色 |
| Note | 否 | 复核提示（如"黄框主式"） |
| EvidencePath | 是 | 证据图路径，分号分隔；至少包含裁剪图 + 整页图，文件必须存在 |

## 3. 编制流程

1. **盘点**：`Export-FormulaCarrierInventory.ps1 -InputPath <链头PPTX> -OutputDir <版本目录>/00_检查报告/inventory`，得到 `formula-carrier-inventory.json`。
2. **渲染**：`Export-PptxVisualAudit.ps1 -InputPath <链头PPTX> -OutputDir <版本目录>/00_检查报告/before-audit`，得到 `pages/slide-NNN.png`。
3. **裁剪**：`Export-FormulaOleCrops.ps1 -CarrierInventoryJson <inventory.json> -PagesDir <before-audit/pages> -OutputDir <版本目录>/00_检查报告/crops`，得到 `ole-crops/*.png` 与 `ole-crops.csv`（含 `SuggestedSizePt` 建议列）。
4. **AI 草拟（Draft）**：执行 agent 逐张读取裁剪图，填 `WhitelistName/SourceFormulaText/颜色/Note`，`SizePt` 取 `SuggestedSizePt`，`ReviewStatus=Draft`。
5. **独立定稿（Approved）**：**另一个** agent 或人工对照裁剪图与整页图逐条核验转录与公式名，确认无误后改 `ReviewStatus=Approved`。同一会话不得自我验收。
6. **白名单不命中时**：不批准。走 `Export-FormulaWhitelistSuggestions.ps1` 证据整理 → 人工确认后在 config `formulaWhitelist` 晋升（含 `targetUnicodeMath` canonical，书写遵守《公式处理说明》§4.1）→ 重跑本流程。
7. **映射**：`Export-FormulaOleMapping.ps1 -CarrierInventoryJson <inventory.json> -GoldSetCsv <goldset.csv> -OutputDir <版本目录>/00_检查报告`。2026-09-14 起小结页不再例外，不排任何页型。

## 4. 拒绝与降级语义（自动，不可绕过）

| 现象 | 结果 |
|---|---|
| ReviewStatus 非 Approved | 行被拒（`ReviewStatus must be Approved`） |
| SourceFormulaText 未命中指定白名单规则 | 行被拒 |
| 页码/形状 id 不在盘点或非 MathTypeOle | 行被拒 |
| 证据文件缺失 | 行被拒 |
| OLE 内容指纹缺少 SourceFormulaText 声明的中文字符 | 行被拒（`OLE content fingerprint mismatch`；映射读取 MathType embedding 的 UTF-16LE 字符做交叉验证，防裁剪图-形状错标——2026-09-17 slide16 sh13/sh16 对调事故后加入） |
| 多 OLE 合并且次形状有动画绑定 | 写回阶段 `OleTimingDangling` 拒绝，保留 MathType |
| 任何一行被拒 | `Export-FormulaOleMapping.ps1` 整体 Failed（有 error 即 throw） |

## 5. 样例

见 [`examples/fixtures/formula-goldset.sample.csv`](../../examples/fixtures/formula-goldset.sample.csv)。真实批次样例参考 14.1 v38 的 `00_检查报告/formula-ole-goldset-manifest.json` 与 `formula-ole-mapping.csv`。
