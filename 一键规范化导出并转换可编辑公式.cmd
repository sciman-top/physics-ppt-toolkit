@echo off
setlocal
for /f "tokens=2 delims=:" %%i in ('chcp') do set /a "OLDCP=%%i"
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "INPUT=%~1"

if "%INPUT%"=="" (
  set "INPUT=%SCRIPT_DIR%"
)
rem %~dp0 ends with a backslash: a bare trailing \" would swallow the rest of
rem the argument line, so anchor the path with a trailing dot instead.
if "%INPUT:~-1%"=="\" set "INPUT=%INPUT%."

set "PS_HOST=pwsh.exe"
where pwsh.exe >nul 2>&1
if errorlevel 1 set "PS_HOST=powershell.exe"

"%PS_HOST%" -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%SCRIPT_DIR%tools\Invoke-PhysicsPptWorkflow.ps1" -InputPath "%INPUT%" -Recurse -Mode NormalizeAndPdf -SkipPreflightReport -ApplyFormulaOmmlWhitelist -OpenGeneratedPptx

if errorlevel 1 (
  echo.
  echo 处理失败，请查看上方错误信息。
  pause
  chcp %OLDCP% >nul
  exit /b 1
)
chcp %OLDCP% >nul
exit /b 0
