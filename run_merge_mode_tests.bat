@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist TestData\mode_a_local.xlsx (
    echo 生成测试数据...
    python TestData\gen_merge_mode_tests.py
)
python TestData\run_merge_mode_tests.py
if %ERRORLEVEL% neq 0 (
    echo 模式合并测试失败
    exit /b 1
)
echo 5 种模式合并测试通过
exit /b 0
