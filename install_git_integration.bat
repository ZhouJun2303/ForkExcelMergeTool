@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "%~dp0ExcelMergeFork.exe" (
    "%~dp0ExcelMergeFork.exe" --install-git-integration %*
    exit /b %ERRORLEVEL%
)

if exist "%~dp0ExcelMergeFork-lite.exe" (
    "%~dp0ExcelMergeFork-lite.exe" --install-git-integration %*
    exit /b %ERRORLEVEL%
)

if exist "%~dp0ExcelMergeFork-python.cmd" (
    "%~dp0ExcelMergeFork-python.cmd" --install-git-integration %*
    exit /b %ERRORLEVEL%
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\install_git_integration.ps1" %*
exit /b %ERRORLEVEL%
