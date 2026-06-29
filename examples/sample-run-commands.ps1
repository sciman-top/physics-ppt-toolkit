# 示例命令：推荐使用一键工作流。

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 进入工具包根目录
cd D:\tools\physics-ppt-toolkit

# 1. 推荐：检查、规范化、导出 PDF，不导出页面图片，并打开生成的 PPTX
.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\原始PPT" `
  -Recurse `
  -Mode NormalizeAndPdf `
  -SkipPreflightReport `
  -OpenGeneratedPptx

# 2. 只生成检查报告，不修改 PPT
.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\原始PPT" `
  -Recurse `
  -Mode CheckOnly

# 3. 单个文件处理
.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\18.2电功率.pptx" `
  -Mode NormalizeAndPdf

# 4. 只输出规范化 PPTX，不导出 PDF/图片
.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\18.2电功率.pptx" `
  -Mode SafeNormalize

# 5. 需要详细复核产物、不自动生成视觉修复副本时
.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\原始PPT" `
  -Recurse `
  -Mode NormalizeAndPdf `
  -IncludeReviewArtifacts `
  -IncludeVisualAudit

# 6. 增强一键：白名单公式转换为可编辑 OfficeMath/OMML 副本，并做结构校验（中速默认）
.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\原始PPT" `
  -Recurse `
  -Mode NormalizeAndPdf `
  -SkipPreflightReport `
  -ApplyFormulaOmmlWhitelist `
  -OpenGeneratedPptx

# 7. 深度视觉审查：需要全量页面图、基线确认和低风险修复副本时再显式开启
.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\原始PPT" `
  -Recurse `
  -Mode NormalizeAndPdf `
  -IncludeReviewArtifacts `
  -IncludeVisualAudit `
  -ApplyVisualAuditFixes `
  -ApplyFormulaOmmlWhitelist `
  -FormulaOmmlVisualAudit `
  -OpenGeneratedPptx
