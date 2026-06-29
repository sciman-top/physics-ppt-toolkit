# AGENTS.md - physics-ppt-toolkit 项目规则（Codex）

**项目**: physics-ppt-toolkit  
**类型**: Windows PowerPoint / PowerShell COM 自动化工具包  
**适用范围**: 仓库根目录  
**最后更新**: 2026-05-29

## A. 共性基线

### A.1 输出与协作
- 默认输出结构：`理解 / 改动范围 / 最小方案 / 验证方法 / 风险与回滚`。
- 修改前先确认需求边界、技术栈依据、目标归宿、迁移批次。
- 优先保持现有一键入口、报告格式和低风险默认行为稳定。

### A.2 项目不变约束
- 只统一样式和可验证的低风险结构，不静默改写教学内容。
- 默认不覆盖输入 PPTX；所有写回都必须输出新文件或新目录。
- PowerPoint COM、Open XML、外部工具失败必须进入报告或降级，不得让批量流程无提示地产生半成品。
- `PPTX/` 样本、`config/physics-ppt-style.config.json`、`vba/` 离线宏兼容性不得被破坏。

### A.3 终态闸门
- 最小门禁：`tools/Test-ToolkitFiles.ps1`。
- 依赖门禁：`tools/Assert-Toolchain.ps1 -Deep`；需要证明 PowerPoint 可启动时才额外使用 `-LaunchPowerPoint`。
- 涉及 OfficeMath/OMML 的变更必须构建并运行 `tools/FormulaOfficeMathValidator`。
- 涉及输出 PPTX 的变更必须至少用一个真实课件样本生成 `summary.md` 和 `review-manifest.json`。

### A.4 反过度优化
- 未经明确要求，不扩大默认产物体积，不默认导出逐页图片，不默认启用 OCR、AI 清晰化或联网视觉 API。
- 不为单次样本问题做全局重排版、删页、移动对象、裁图或内容重写。
- 大文件、报告目录、便携工具和一次性验证产物优先作为可再生成产物治理，不混入普通源码改动。

## B. Codex 项目内差异

### B.1 入口优先级
- 主入口：`tools/Invoke-PhysicsPptWorkflow.ps1`。
- 核心规范化：`tools/Normalize-PhysicsPpt.ps1`。
- 快速入口：`一键规范化并导出PDF.cmd`；增强公式入口：`一键规范化导出并转换可编辑公式.cmd`。
- VBA 为离线备用方案，不读取 JSON；修改 JSON 字体字号时必须检查 `vba/PhysicsPptCommon.bas`。

### B.2 文件与编码
- `tools/*.ps1`、`examples/*.ps1` 必须保持 UTF-8 with BOM，兼容 Windows PowerShell 5.1。
- 读取中文配置、CSV、Markdown、VBA 时显式指定 `-Encoding UTF8`。
- 给 Excel 使用的 CSV 必须写入 UTF-8 BOM。

## C. 项目差异

### C.1 目录与模块归宿
- `config/`：样式、公式白名单和安全开关。
- `tools/`：PowerShell/Node/Python/.NET 工具；新增脚本必须进入 `Test-ToolkitFiles.ps1` 自检。
- `tools/FormulaOfficeMathValidator/`：OfficeMath/Open XML 结构验证。
- `vba/`：PowerPoint 内离线宏方案。
- `docs/`：用户说明、边界、路线图和验证证据。
- `examples/`：示例命令和最小 fixture。
- `reports/`：可再生成验证产物；默认不应新增跟踪文件。
- `PPTX/`：真实课件样本；新增或替换大样本需说明用途。

### C.2 公式链路自动化边界
- 文本公式自动转换只允许走白名单精确匹配：`formula-review.csv` 中 `SuggestedAction=ReviewWhitelistConversion`。
- 视觉审查可以替代人工确认，但必须作为门禁使用：原始/规范化基线、OMML 副本 Open XML 校验、PowerPoint PDF/PNG 导出、视觉审查错误不增加。
- 图片公式 OCR 不能直接写回 PPTX；必须先归一化到白名单或明确公式源码，再复用同一视觉审查门禁。
- 任何公式写回都只保存副本，不覆盖原始、规范化或视觉修复 PPTX。

### C.3 媒体与视觉链路
- 媒体优化默认只做保守、可回滚处理；视频、GIF、SVG、EMF、WMF 默认只报告。
- AI 清晰化默认关闭；必须先候选扫描、抽样验证、体积不增或收益明确，再写回副本。
- 视觉修复只允许确定性低风险动作；删除空白页、改版、移动教学对象必须单独批次。

## D. 维护清单

- 保持 README、docs、示例命令、cmd 入口与实际参数一致。
- 每次新增工具后更新 `tools/Test-ToolkitFiles.ps1`。
- 每次新增输出目录后更新 README/docs 和 `.gitignore`。
- 交付前至少运行：`powershell -NoProfile -ExecutionPolicy Bypass -File tools\Test-ToolkitFiles.ps1`。
- 批量或跨模块改动必须说明影响模块、验证证据和回滚动作。
