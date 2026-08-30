# AGENTS.md - physics-ppt-toolkit
**项目契约**: 2.0
**全局规则复核**: 9.77
**最后更新**: 2026-08-30

## 1. 产品边界与入口
- `docs/产品需求与工程路线图.md` 是产品边界与默认主线的 source of truth；`config/physics-ppt-style.config.json` 是 PowerShell 样式运行配置，VBA 常量必须按文件注释显式同步。
- 主 entrypoint 是 `tools/Invoke-PhysicsPptWorkflow.ps1`；快速入口为根目录两个一键 `.cmd`，核心规范化实现是 `tools/Normalize-PhysicsPpt.ps1`。
- 本仓只对既有初中物理 PPT 做低风险样式规范化、公式审查和导出，不自动改写教学内容或重排课件。

## A. 仓库真值与领域不变量
- 默认不覆盖输入 PPTX；所有写回必须产生新文件/目录，并保留原始备份和可复核报告。
- 禁止修改文本内容、对象位置/尺寸、动画顺序和图片裁剪；分组对象默认跳过，公式转换仅允许白名单精确匹配。
- COM 调用必须重试并在边界释放对象；异常不得冒泡为无说明 UI/脚本崩溃。
- 视觉审查可作为交付门禁，但 AI/OCR 结果不能直接写回 PPTX；必须先形成受验证的结构化结果，再走同一副本与视觉复核链。
- `tools/*.ps1`、`examples/*.ps1` 保持 UTF-8 with BOM 以兼容 Windows PowerShell 5.1；Excel CSV 必须写 UTF-8 BOM。

## B. 执行边界
- `config/` 管样式、schema 与安全开关；`tools/` 管 PowerShell/Node/Python/.NET 自动化；`vba/` 是不读取 JSON 的离线备用实现。
- 新工具必须纳入 `tools/Test-ToolkitFiles.ps1`；`reports/` 是可再生成产物，默认不新增跟踪文件。
- OCR、AI 清晰化、逐页图片、媒体重编码和联网视觉 API 默认关闭；只有当前任务明确需要并有样本收益证据时才启用。
- `PPTX/` 是真实样本资产；新增、替换或删除大样本必须说明用途，普通代码回滚不得触碰无关样本变更。

## C. 最低门禁
- minimum gate：`powershell -NoProfile -ExecutionPolicy Bypass -File tools/Test-ToolkitFiles.ps1`。
- 依赖/工具链变更追加 `tools/Assert-Toolchain.ps1 -Deep`；只有需要证明 PowerPoint 可启动时才加 `-LaunchPowerPoint`。
- OfficeMath/OMML 变更必须构建并运行 `tools/FormulaOfficeMathValidator`；输出 PPTX 变更至少用一个真实样本生成 `summary.md` 与 `review-manifest.json`。
- High DPI、PowerPoint 实际导出和视觉质量属于人工/真实宿主验收，静态测试不能替代。

## D. 回滚与收口
- Git baseline=`main`，upstream=`origin/main`；只回滚本次切片，保留用户现有 PPTX、reports 和并发改动。
- 公式、媒体或视觉修复必须逐批输出副本；失败时删除本批输出并从原始 PPTX 重跑，不原地修补源文件。
- 证明受影响门禁后立即停止，不为单一样本做全局重排、删页、移动对象或内容重写。
