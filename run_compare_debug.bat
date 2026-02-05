@echo off
chcp 65001 >nul
cd /d "%~dp0"
set A=%~1
set B=%~2
if "%A%"=="" set A=TestData\local.xlsx
if "%B%"=="" set B=TestData\remote.xlsx
echo 对比: %A% vs %B%
echo 日志: %CD%\MergeExcelFork.log
echo.
python MergeExcelFork.py "%A%" "%B%"
echo.
echo 退出码: %ERRORLEVEL%
pause
