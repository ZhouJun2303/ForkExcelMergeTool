@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

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

echo [2/4] 使用 PyInstaller 打包...
pyinstaller --onefile --console --name ExcelMergeFork --clean MergeExcelFork.py
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

REM 传参 no-test / 0 / skip 时跳过快速验证；无参或其它参数则执行验证
set "RUN_TEST=1"
if /i "%1"=="no-test" set "RUN_TEST=0"
if /i "%1"=="0" set "RUN_TEST=0"
if /i "%1"=="skip" set "RUN_TEST=0"

if "%RUN_TEST%"=="1" (
    echo [4/4] 快速验证测试...
    if not exist TestData\base.xlsx (
        python TestData\gen_test_data.py
    )
    ExcelMergeFork.exe TestData\local.xlsx TestData\base.xlsx TestData\remote.xlsx TestData\_output\merged.xlsx
    if !ERRORLEVEL! equ 0 (
        echo 合并测试通过
        ExcelMergeFork.exe TestData\local.xlsx TestData\remote.xlsx
        if !ERRORLEVEL! equ 0 echo 对比测试通过
    ) else (
        echo WARNING: exe 测试失败，请检查
    )
) else (
    echo [4/4] 已跳过快速验证（传参 no-test/0/skip）
)

echo.
echo ========== 打包完成 ==========
echo 产物: %CD%\ExcelMergeFork.exe
echo.
pause
