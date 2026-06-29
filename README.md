# Physics PPT Toolkit

面向初中物理课件的 PowerPoint 规范化、公式审查与导出工具包。

A Windows toolkit for normalizing, auditing, and exporting junior-high physics PowerPoint lessons.

本项目用于在 Windows + Microsoft PowerPoint 环境下，对既有物理课件进行低风险、可回滚、可验证的批量处理。核心目标是统一样式、保留教学内容、输出可复核结果，而不是静默改写课件内容。

核心原则：**只统一样式，不修改教学内容和对象位置。**

---

## 目录结构

```text
physics-ppt-toolkit/
├─ README.md
├─ docs/
│  ├─ 初中物理PPT统一排版规范.md
│  ├─ 使用方法.md
│  ├─ GitHub远端迁移说明.md
│  ├─ PPT母版制作说明.md
│  ├─ 自动化边界与风险控制.md
│  └─ 公式处理说明.md
├─ config/
│  └─ physics-ppt-style.config.json
├─ tools/
│  ├─ Normalize-PhysicsPpt.ps1
│  ├─ Report-PhysicsPptStyle.ps1
│  ├─ Invoke-PhysicsPptWorkflow.ps1
│  └─ Test-ToolkitFiles.ps1
├─ vba/
│  ├─ PhysicsPptCommon.bas
│  ├─ PhysicsPptNormalize.bas
│  ├─ PhysicsPptReportOnly.bas
│  └─ ApplyPhysicsPptMasterStyle.bas
└─ examples/
   └─ sample-run-commands.ps1
```

---

## 推荐工作流

```text
1. 把 PPTX 文件放进一个文件夹。
2. 双击 `一键规范化并导出PDF.cmd`，或把 PPTX/文件夹拖到该文件上（文件名保留兼容）。
3. 脚本自动检查、规范化 PPTX、导出 PDF，并生成简短报告；默认不导出页面图片。
4. 处理完成后直接打开生成的 PPTX 复核。
5. 如报告提示异常页，在生成的完整 PPTX 中定位查看并人工修正。
```

---

## 快速开始

在 PowerShell 中运行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

cd .\physics-ppt-toolkit

.\tools\Invoke-PhysicsPptWorkflow.ps1 `
  -InputPath "D:\课件\原始PPT" `
  -Recurse `
  -Mode NormalizeAndPdf `
  -SkipPreflightReport `
  -OpenGeneratedPptx
```

输出目录会自动创建，结构如下：

```text
_physics_ppt_output_yyyyMMdd_HHmmss/
├─ 00_检查报告/
├─ 01_规范化PPTX/
├─ 02_导出PDF/
├─ 03_原始备份/
├─ summary.md
└─ review-manifest.json
```

常用模式：

```powershell
# 只检查，不修改
.\tools\Invoke-PhysicsPptWorkflow.ps1 -InputPath "D:\课件\原始PPT" -Recurse -Mode CheckOnly

# 规范化 PPTX 并导出 PDF（默认，不导出页面图片）
.\tools\Invoke-PhysicsPptWorkflow.ps1 -InputPath "D:\课件\原始PPT" -Recurse -Mode NormalizeAndPdf -SkipPreflightReport

# 只输出规范化 PPTX，不导出 PDF/图片
.\tools\Invoke-PhysicsPptWorkflow.ps1 -InputPath "D:\课件\原始PPT" -Recurse -Mode SafeNormalize

# 需要详细视觉复核产物时，再额外生成页面图片、总览图和复核索引
.\tools\Invoke-PhysicsPptWorkflow.ps1 -InputPath "D:\课件\原始PPT" -Recurse -Mode NormalizeAndPdf -IncludeReviewArtifacts

# 增强一键：白名单公式转可编辑 OfficeMath/OMML，并做 Open XML 结构校验
.\tools\Invoke-PhysicsPptWorkflow.ps1 -InputPath "D:\课件\原始PPT" -Recurse -Mode NormalizeAndPdf -SkipPreflightReport -ApplyFormulaOmmlWhitelist

# 深度视觉审查：需要逐页 PNG、自动确认和低风险修复副本时再开启
.\tools\Invoke-PhysicsPptWorkflow.ps1 -InputPath "D:\课件\原始PPT" -Recurse -Mode NormalizeAndPdf -IncludeReviewArtifacts -IncludeVisualAudit -ApplyVisualAuditFixes -ApplyFormulaOmmlWhitelist -FormulaOmmlVisualAudit

# 忽略已有输出，重新处理
.\tools\Invoke-PhysicsPptWorkflow.ps1 -InputPath "D:\课件\原始PPT" -Recurse -Mode ForceRebuild
```

---

## 配置说明

样式参数集中在 `config/physics-ppt-style.config.json`，PowerShell 脚本启动时自动加载。
修改 JSON 配置即可调整字体、字号、颜色等参数，无需改动脚本代码。

VBA 宏是独立的离线方案，不读取 JSON 配置。如需调整，请手动修改 VBA 模块顶部的常量，并参考注释中的配置字段映射。

---

## 安全说明

- 脚本会默认复制输出新文件，不直接覆盖原文件。
- 脚本默认导出 PDF 作为固定版式复核产物，但不导出页面图片；需要图片复核产物时额外使用 `-IncludeReviewArtifacts`。
- 脚本不会修改文本内容。
- 脚本不会移动对象。
- 脚本不会裁剪图片。
- 脚本不会修改动画顺序。
- 公式转换默认不启用；低风险独立文本公式只做样式归一和候选报告。
- 启用 `-ApplyFormulaOmmlWhitelist` 时，只转换白名单精确匹配公式，并默认要求 Open XML 校验通过；逐页视觉审查和视觉自动确认需额外使用 `-FormulaOmmlVisualAudit`。

---

## 适用场景

- 批量统一标题、正文、表格、重点框样式。
- 批量设置普通页白底、视频页黑底。
- 批量检查小字号、疑似公式、非规范字体。
- 对既有 PPT 进行低风险视觉规范化。

---

## 不适用场景

- 自动重排版面。
- 自动拆页。
- 自动改写教学内容。
- 自动裁剪图片。
- 自动识别图片中的公式。
- 自动修改复杂公式含义。
