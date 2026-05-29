@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Options: --gui for windowed exe, --test to run tests, --dist/no-test/skip to skip tests.
REM --release also publishes a GitHub Release after a successful package.
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

echo ========== Fork Excel Merge Tool Package ==========
echo.

echo [1/4] Checking dependencies...
pip install openpyxl pyinstaller -q
if %ERRORLEVEL% neq 0 (
    echo ERROR: pip install failed
    pause
    exit /b 1
)
echo Dependencies OK
echo.

echo [2/4] Bump minor version and build ( %WINMODE% )...
python Scripts\bump_version.py
pyinstaller --onefile %WINMODE% --name ExcelMergeFork --clean Scripts\MergeExcelFork.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: PyInstaller build failed
    pause
    exit /b 1
)
echo.

echo [3/4] Copy exe...
if not exist dist\ExcelMergeFork.exe (
    echo ERROR: dist\ExcelMergeFork.exe was not generated
    pause
    exit /b 1
)
copy /Y dist\ExcelMergeFork.exe ExcelMergeFork.exe >nul
echo Generated ExcelMergeFork.exe
echo.

if "%RUN_TEST%"=="1" (
    echo [4/4] Running merge mode tests...
    if not exist TestData\mode_a_local.xlsx (
        python TestData\gen_merge_mode_tests.py
    )
    python TestData\run_merge_mode_tests.py
    if !ERRORLEVEL! equ 0 (
        echo Merge mode tests passed
    ) else (
        echo WARNING: merge mode tests failed
        set "TEST_FAILED=1"
    )
) else (
    echo [4/4] Tests skipped
)

echo.
echo ========== Package Complete ==========
echo Artifact: %CD%\ExcelMergeFork.exe
echo Options: --gui windowed exe  --test run tests  --dist skip tests  --release publish GitHub Release
echo.

if "%PUBLISH_RELEASE%"=="1" (
    if "%TEST_FAILED%"=="1" (
        echo ERROR: tests failed; GitHub Release was not published
        pause
        exit /b 1
    )
    echo [Release] Publishing GitHub Release...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\Scripts\publish-release.ps1"
    if !ERRORLEVEL! neq 0 (
        echo ERROR: GitHub Release publish failed
        pause
        exit /b 1
    )
)

pause
