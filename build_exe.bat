@echo off
chcp 65001 >nul
cd /d "%~dp0"

REM 参数：--gui 无控制台窗口，--test 打包后跑测试
set "WINMODE=--console"
set "RUN_TEST=0"
:parse
if "%~1"=="" goto done_parse
if /i "%~1"=="--gui" set "WINMODE=--windowed"
if /i "%~1"=="--test" set "RUN_TEST=1"
shift
goto parse
:done_parse

echo Bump version and build ExcelMergeFork.exe...
python Scripts\bump_version.py
pyinstaller --onefile %WINMODE% --name ExcelMergeFork --icon Assets\ExcelMergeFork.ico --add-data "Assets;Assets" Scripts\MergeExcelFork.py
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

if "%RUN_TEST%"=="1" (
    echo Running merge mode tests...
    if not exist TestData\mode_a_local.xlsx python TestData\gen_merge_mode_tests.py
    python TestData\run_merge_mode_tests.py
    if %ERRORLEVEL% neq 0 (
        echo WARNING: 模式合并测试失败
    ) else (
        echo 模式合并测试通过
    )
)
exit /b 0
