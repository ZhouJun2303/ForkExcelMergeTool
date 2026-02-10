@echo off
chcp 65001 >nul
cd /d "%~dp0"
set LOCAL=%1
set BASE=%2
set REMOTE=%3
set MERGED=%4
if "%LOCAL%"=="" (
    echo Usage: run_merge.bat ^<local^> ^<base^> ^<remote^> ^<merged^>
    echo Example: run_merge.bat TestData\local.xlsx TestData\base.xlsx TestData\remote.xlsx TestData\_output\merged.xlsx
    exit /b 1
)
python "%~dp0Scripts\MergeExcelFork.py" "%LOCAL%" "%BASE%" "%REMOTE%" "%MERGED%"
exit /b %ERRORLEVEL%
