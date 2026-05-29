@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM 参数：--gui 无控制台窗口，--test 打包后跑测试，--dist 仅打包不跑测试（未传则默认跑测试），--release 打包成功后发布 GitHub Release
set "WINMODE=--console"
set "RUN_TEST=1"
set "PUBLISH_RELEASE=0"
set "TEST_FAILED=0"
:parse
if "%~1"=="" goto done_parse
if /i "%~1"=="--gui" set "WINMODE=--windowed"
if /i "%~1"=="--test" set "RUN_TEST=1"
if /i "%~1"=="--dist" set "RUN_TEST=0"
if /i "%~1"=="--release" set "PUBLISH_RELEASE=1"
if /i "%~1"=="no-test" set "RUN_TEST=0"
if /i "%~1"=="0" set "RUN_TEST=0"
if /i "%~1"=="skip" set "RUN_TEST=0"
shift
goto parse
:done_parse

echo ========== Fork Excel Merge Tool 一键打包 ==========
echo.

echo [1/4] 检查依赖...
pip install openpyxl pyinstaller -q
if %ERRORLEVEL% neq 0 (
    echo ERROR: pip 安装依赖失败
    pause
    exit /b 1
)
echo 依赖 OK
echo.

echo [2/4] 小版本号递增并打包 ( %WINMODE% )...
python Scripts\bump_version.py
pyinstaller --onefile %WINMODE% --name ExcelMergeFork --clean Scripts\MergeExcelFork.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: PyInstaller 打包失败
    pause
    exit /b 1
)
echo.

echo [3/4] 复制 exe...
if not exist dist\ExcelMergeFork.exe (
    echo ERROR: dist\ExcelMergeFork.exe 未生成
    pause
    exit /b 1
)
copy /Y dist\ExcelMergeFork.exe ExcelMergeFork.exe >nul
echo 已生成 ExcelMergeFork.exe
echo.

if "%RUN_TEST%"=="1" (
    echo [4/4] 快速验证（模式合并测试）...
    if not exist TestData\mode_a_local.xlsx (
        python TestData\gen_merge_mode_tests.py
    )
    python TestData\run_merge_mode_tests.py
    if !ERRORLEVEL! equ 0 (
        echo 模式合并测试通过
    ) else (
        echo WARNING: 模式合并测试失败，请检查
        set "TEST_FAILED=1"
    )
) else (
    echo [4/4] 已跳过测试（传参 --dist / no-test / 0 / skip 时跳过）
)

echo.
echo ========== 打包完成 ==========
echo 产物: %CD%\ExcelMergeFork.exe
echo 参数说明: --gui 无控制台窗口  --test 打包后跑测试  --dist 仅打包不跑测试  --release 发布 GitHub Release
echo.

if "%PUBLISH_RELEASE%"=="1" (
    if "%TEST_FAILED%"=="1" (
        echo ERROR: 测试失败，已停止发布 GitHub Release
        pause
        exit /b 1
    )
    echo [Release] 正在发布 GitHub Release...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\Scripts\publish-release.ps1"
    if !ERRORLEVEL! neq 0 (
        echo ERROR: GitHub Release 发布失败
        pause
        exit /b 1
    )
)

pause
