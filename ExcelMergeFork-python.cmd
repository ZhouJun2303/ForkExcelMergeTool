@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo ERROR: Python was not found. Please install Python 3.7+ and run: pip install -r requirements.txt
    exit /b 1
)

python "%~dp0MergeExcelFork.py" %*
exit /b %ERRORLEVEL%
