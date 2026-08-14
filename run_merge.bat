@echo off
chcp 65001 >nul
cd /d "%~dp0"
if "%~4"=="" (
    echo Usage: run_merge.bat ^<local^> ^<base^> ^<remote^> ^<merged^>
    exit /b 1
)
"%~dp0ExcelMergeFork.exe" %1 %2 %3 %4
exit /b %ERRORLEVEL%
