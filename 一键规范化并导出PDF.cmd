@echo off
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "INPUT=%~1"

if "%INPUT%"=="" (
  set "INPUT=%SCRIPT_DIR%"
)

set "PS_HOST=pwsh.exe"
where pwsh.exe >nul 2>&1
if errorlevel 1 set "PS_HOST=powershell.exe"

"%PS_HOST%" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%SCRIPT_DIR%tools\Invoke-PhysicsPptWorkflow.ps1" -InputPath "%INPUT%" -Recurse -Mode NormalizeAndPdf -SkipPreflightReport -OpenGeneratedPptx

if errorlevel 1 (
  echo.
  echo 处理失败，请查看上方错误信息。
  pause
)
