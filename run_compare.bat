@echo off
chcp 65001 >nul
cd /d "%~dp0"
set A=%1
set B=%2
if "%A%"=="" set A=TestData\remote.xlsx
if "%B%"=="" set B=TestData\local.xlsx
if not exist "%A%" (
    echo Usage: run_compare.bat [remote] [local]
    echo Example: run_compare.bat TestData\remote.xlsx TestData\local.xlsx
    exit /b 1
)
python "%~dp0Scripts\MergeExcelFork.py" "%A%" "%B%"
exit /b %ERRORLEVEL%
