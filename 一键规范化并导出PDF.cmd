@echo off
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "INPUT=%~1"

if "%INPUT%"=="" (
  set "INPUT=%SCRIPT_DIR%"
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%tools\Invoke-PhysicsPptWorkflow.ps1" -InputPath "%INPUT%" -Recurse -Mode NormalizeAndPdf -SkipPreflightReport -OpenGeneratedPptx

if errorlevel 1 (
  echo.
  echo 处理失败，请查看上方错误信息。
  pause
)
