@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "%~dp0ExcelMergeFork.exe" (
    echo ERROR: ExcelMergeFork.exe 不存在，请先运行 package.bat
    exit /b 1
)
"%~dp0ExcelMergeFork.exe" --install-git-integration %*
exit /b %ERRORLEVEL%
