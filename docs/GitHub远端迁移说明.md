# GitHub 远端迁移说明

本文说明如何把 Physics PPT Toolkit 迁移到 GitHub，并在另一台电脑上拉取使用。

## 1. 推荐仓库名

- GitHub 仓库名：`physics-ppt-toolkit`
- 显示名：`Physics PPT Toolkit`
- 简介：`Windows toolkit for normalizing, auditing, and exporting junior-high physics PowerPoint lessons.`

## 2. 当前仓库状态

截至 2026-06-29，本地仓库存在以下迁移风险：

- 当前没有配置 Git 远端。
- GitHub CLI `gh` 已安装，但需要先完成 `gh auth login`。
- Git LFS 已安装。
- Git 对象库约 `7.47 GiB`。
- 当前历史中包含超过 `100 MB` 的文件，例如 `PPTX/` 样本、`reports/` 产物和 `tools/vendor/pandoc/pandoc-3.9.0.2/pandoc.exe`。

这些文件会导致普通 GitHub 推送失败或仓库过大。即使从当前目录删除它们，只要历史中仍保留大文件，推送仍可能失败。

## 3. 推荐迁移策略

推荐建立一个面向 GitHub 的瘦身仓库，只保留源码、脚本、配置和说明文档。

应保留：

- `config/`
- `docs/`
- `examples/`
- `tools/`
- `vba/`
- `README.md`
- `AGENTS.md`
- `.gitattributes`
- `.gitignore`
- `package.json`
- `package-lock.json`
- 一键入口 `*.cmd`

应排除或单独分发：

- `reports/`
- `_physics_ppt_output_*/`
- `node_modules/`
- 大体积 `PPTX/` 样本
- 超过 GitHub 普通文件限制的大型便携工具，例如 `pandoc.exe`
- 一次性验证产物和可再生成产物

大文件有三种合适放法：

- Git LFS：适合必须随仓库版本管理的少量样本。
- GitHub Release：适合便携工具包、离线依赖包和发布附件。
- 网盘或内网文件夹：适合真实课件样本、批量验证报告和教学资料。

## 4. GitHub 创建远端

在 GitHub 页面创建新仓库：

- Owner：`sciman-top`
- Repository name：`physics-ppt-toolkit`
- Visibility：按需要选择 `Private` 或 `Public`
- 不要勾选自动生成 README、`.gitignore` 或 License，避免和本地文件冲突。

创建后，本地添加远端：

```powershell
git remote add origin https://github.com/sciman-top/physics-ppt-toolkit.git
git remote -v
```

如果使用 SSH：

```powershell
git remote add origin git@github.com:sciman-top/physics-ppt-toolkit.git
git remote -v
```

## 5. 不建议直接推送当前历史

当前仓库历史中已有大文件。直接运行下面命令大概率失败：

```powershell
git push -u origin master
```

原因是 GitHub 会检查整个推送历史，不只检查当前工作区。

## 6. 推荐发布流程

推荐新建一个干净的发布目录，再从当前仓库复制需要保留的文件，并重新初始化 Git。

示例：

```powershell
mkdir D:\CODE\physics-ppt-toolkit-publish
cd D:\CODE\physics-ppt-toolkit-publish
git init
```

复制保留清单中的源码、配置、文档和入口文件后，恢复 Node 依赖：

```powershell
npm install
```

运行门禁：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Test-ToolkitFiles.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Assert-Toolchain.ps1 -Deep
```

确认通过后提交并推送：

```powershell
git add .
git commit -m "Initial public toolkit release"
git branch -M main
git remote add origin https://github.com/sciman-top/physics-ppt-toolkit.git
git push -u origin main
```

如果必须保留部分 PPTX 样本，先启用 Git LFS：

```powershell
git lfs install
git lfs track "*.pptx"
git add .gitattributes
```

只把必要的小批量样本加入仓库。真实大课件优先放 Release、网盘或内网共享。

## 7. 新电脑拉取与验证

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
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Test-ToolkitFiles.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Assert-Toolchain.ps1 -Deep -LaunchPowerPoint
```

最后用一个真实课件样本跑完整流程，确认生成 `summary.md` 和 `review-manifest.json`。

## 8. 回滚

推荐先保留当前本地仓库不动，把 GitHub 发布仓库作为独立目录准备。若发布目录处理失败，直接删除发布目录即可，不影响当前工作仓库。

如果已经添加了错误远端，可以移除：

```powershell
git remote remove origin
```

如果已经推送了错误的大文件历史，不要继续强推；应新建干净仓库或重新创建远端仓库后再发布瘦身版本。
