@echo off
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

echo [4/4] 快速验证测试...
if not exist TestData\base.xlsx (
    python TestData\gen_test_data.py
)
ExcelMergeFork.exe TestData\local.xlsx TestData\base.xlsx TestData\remote.xlsx TestData\_output\merged.xlsx
if %ERRORLEVEL% equ 0 (
    echo 合并测试通过
    ExcelMergeFork.exe TestData\local.xlsx TestData\remote.xlsx
    if %ERRORLEVEL% equ 0 echo 对比测试通过
) else (
    echo WARNING: exe 测试失败，请检查
)

echo.
echo ========== 打包完成 ==========
echo 产物: %CD%\ExcelMergeFork.exe
echo.
pause
