@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~2"=="" (
    echo Usage: run_compare.bat ^<remote^> ^<local^>
    exit /b 1
)
"%~dp0ExcelMergeFork.exe" %1 %2
exit /b %ERRORLEVEL%
