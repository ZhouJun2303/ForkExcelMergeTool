@echo off
chcp 65001 >nul
cd /d "%~dp0"
set A=%1
set B=%2
if "%A%"=="" set A=TestData\local.xlsx
if "%B%"=="" set B=TestData\remote.xlsx
if not exist "%A%" (
    echo Usage: run_compare.bat [fileA] [fileB]
    echo Example: run_compare.bat TestData\local.xlsx TestData\remote.xlsx
    exit /b 1
)
python "%~dp0Scripts\MergeExcelFork.py" "%A%" "%B%"
exit /b %ERRORLEVEL%
