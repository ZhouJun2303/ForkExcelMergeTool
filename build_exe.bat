@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Building ExcelMergeFork.exe with PyInstaller...
pyinstaller --onefile --console --name ExcelMergeFork Scripts\MergeExcelFork.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: PyInstaller failed. Install with: pip install pyinstaller
    exit /b 1
)
if exist dist\ExcelMergeFork.exe (
    copy /Y dist\ExcelMergeFork.exe ExcelMergeFork.exe
    echo Done: ExcelMergeFork.exe
) else (
    echo ERROR: dist\ExcelMergeFork.exe not found
    exit /b 1
)
exit /b 0
