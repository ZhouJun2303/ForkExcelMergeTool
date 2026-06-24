@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "%~dp0ExcelMergeFork.exe" (
    "%~dp0ExcelMergeFork.exe" --uninstall-git-integration %*
    exit /b %ERRORLEVEL%
)

if exist "%~dp0ExcelMergeFork-lite.exe" (
    "%~dp0ExcelMergeFork-lite.exe" --uninstall-git-integration %*
    exit /b %ERRORLEVEL%
)

if exist "%~dp0ExcelMergeFork-python.cmd" (
    "%~dp0ExcelMergeFork-python.cmd" --uninstall-git-integration %*
    exit /b %ERRORLEVEL%
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\uninstall_git_integration.ps1"
exit /b %ERRORLEVEL%
