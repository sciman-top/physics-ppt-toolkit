# GitHub 远端迁移说明

本文说明 Physics PPT Toolkit 的 GitHub 远端位置，以及在新电脑上拉取与验证的方法。

> **状态（2026-10-09）**：迁移已完成。远端 `origin` 已配置为 `https://github.com/sciman-top/physics-ppt-toolkit.git` 并与本地 `main` 同步。

## 1. 仓库信息

- GitHub 仓库名：`physics-ppt-toolkit`
- 简介：`Windows toolkit for normalizing, auditing, and exporting junior-high physics PowerPoint lessons.`
- 历史：2026-10-09 首次发布时按瘦身策略重建了干净历史（排除 `reports/`、`node_modules/`、大体积 PPTX 样本与超限便携工具，方法见 git 历史中本文件的 v1 版本）；大样本与便携工具走 GitHub Release、网盘或内网共享。

## 2. 新电脑拉取与验证

新电脑准备：

- Windows 10 / Windows 11
- Microsoft PowerPoint 桌面版
- Node.js
- .NET SDK 10
- Git

拉取仓库：

```powershell
cd D:\tools
git clone https://github.com/sciman-top/physics-ppt-toolkit.git
cd D:\tools\physics-ppt-toolkit
npm install
```

如果仓库启用了 Git LFS：

```powershell
git lfs pull
```

验证：

```powershell
pwsh -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File tools\Test-ToolkitFiles.ps1
pwsh -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File tools\Assert-Toolchain.ps1 -Deep -LaunchPowerPoint
```

最后用一个真实课件样本跑完整流程，确认生成 `summary.md` 和 `review-manifest.json`。

## 3. 回滚

推荐先保留当前本地仓库不动，把 GitHub 发布仓库作为独立目录准备。若发布目录处理失败，直接删除发布目录即可，不影响当前工作仓库。

如果已经添加了错误远端，可以移除：

```powershell
git remote remove origin
```

如果已经推送了错误的大文件历史，不要继续强推；应新建干净仓库或重新创建远端仓库后再发布瘦身版本。
