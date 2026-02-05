@echo off
chcp 65001 >nul
cd /d "%~dp0"
set TD=TestData
set OUT=%TD%\_output
if not exist "%OUT%" mkdir "%OUT%"

if not exist "%TD%\base.xlsx" (
    echo Generating test data...
    python "%TD%\gen_test_data.py"
)
python "%~dp0MergeExcelFork.py" "%TD%\local.xlsx" "%TD%\base.xlsx" "%TD%\remote.xlsx" "%OUT%\merged.xlsx"
if %ERRORLEVEL% equ 0 (
    echo Opening output folder...
    start "" explorer "%OUT%"
)
exit /b %ERRORLEVEL%
