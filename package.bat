@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Options: default builds both full and lite exe. Use --console for debug console, --test to run tests, --dist/no-test/skip to skip tests.
REM --full-only builds only the bundled-Python exe. --lite-only builds only the non-bundled Python launcher exe.
REM --python-zip also creates a lightweight source zip for machines with Python installed.
REM --package-zip also creates distribution zips with exe + README + install/uninstall helpers.
REM --release also publishes a GitHub Release after a successful package.
set "WINMODE=--windowed"
set "RUN_TEST=1"
set "PUBLISH_RELEASE=0"
set "BUILD_PYTHON_ZIP=0"
set "BUILD_FULL_EXE=1"
set "BUILD_LITE_EXE=1"
set "BUILD_PACKAGE_ZIP=0"
set "TEST_FAILED=0"
set "NO_PAUSE=0"

:parse
if "%~1"=="" goto done_parse
if /i "%~1"=="--gui" set "WINMODE=--windowed"
if /i "%~1"=="--console" set "WINMODE=--console"
if /i "%~1"=="--test" set "RUN_TEST=1"
if /i "%~1"=="--dist" set "RUN_TEST=0"
if /i "%~1"=="--python-zip" set "BUILD_PYTHON_ZIP=1"
if /i "%~1"=="--package-zip" set "BUILD_PACKAGE_ZIP=1"
if /i "%~1"=="--release" set "PUBLISH_RELEASE=1"
if /i "%~1"=="--full-only" set "BUILD_LITE_EXE=0"
if /i "%~1"=="--lite-only" set "BUILD_FULL_EXE=0"
if /i "%~1"=="--lite-only" set "RUN_TEST=0"
if /i "%~1"=="--no-pause" set "NO_PAUSE=1"
if /i "%~1"=="no-test" set "RUN_TEST=0"
if /i "%~1"=="0" set "RUN_TEST=0"
if /i "%~1"=="skip" set "RUN_TEST=0"
shift
goto parse
:done_parse

echo ========== Fork Excel Merge Tool Package ==========
echo.

if "%BUILD_PYTHON_ZIP%"=="1" if "%RUN_TEST%"=="0" if "%BUILD_PACKAGE_ZIP%"=="0" (
    call :build_python_zip
    exit /b !ERRORLEVEL!
)

echo [1/5] Checking dependencies...
if "%BUILD_FULL_EXE%"=="1" (
    pip install openpyxl pyinstaller -q
    if %ERRORLEVEL% neq 0 (
        echo ERROR: pip install failed
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
) else (
    pip install openpyxl -q
    if %ERRORLEVEL% neq 0 (
        echo ERROR: pip install failed
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
)
echo Dependencies OK
echo.

echo [2/5] Bump minor version...
python Scripts\bump_version.py
if %ERRORLEVEL% neq 0 (
    echo ERROR: version bump failed
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)
echo.

if "%BUILD_FULL_EXE%"=="1" (
    echo [3/5] Building bundled-Python exe %WINMODE%...
    pyinstaller --onefile %WINMODE% --name ExcelMergeFork --clean --icon Assets\ExcelMergeFork.ico --add-data "Assets;Assets" Scripts\MergeExcelFork.py
    if !ERRORLEVEL! neq 0 (
        echo ERROR: PyInstaller build failed
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
    if not exist dist\ExcelMergeFork.exe (
        echo ERROR: dist\ExcelMergeFork.exe was not generated
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
    copy /Y dist\ExcelMergeFork.exe ExcelMergeFork.exe >nul
    echo Generated ExcelMergeFork.exe
) else (
    echo [3/5] Bundled-Python exe skipped
)
echo.

if "%BUILD_LITE_EXE%"=="1" (
    echo [4/5] Building lite exe...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\Scripts\build_lite_exe.ps1"
    if !ERRORLEVEL! neq 0 (
        echo ERROR: lite exe build failed
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
) else (
    echo [4/5] Lite exe skipped
)
echo.

if "%RUN_TEST%"=="1" (
    echo [5/5] Running merge mode tests...
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
    echo [5/5] Tests skipped
)

echo.
echo ========== Package Complete ==========
if "%BUILD_FULL_EXE%"=="1" echo Artifact: %CD%\ExcelMergeFork.exe
if "%BUILD_LITE_EXE%"=="1" echo Artifact: %CD%\ExcelMergeFork-lite.exe
echo Options: --console debug console  --test run tests  --dist skip tests  --full-only  --lite-only  --python-zip source zip  --package-zip distribution zip  --release publish GitHub Release  --no-pause
echo.

if "%BUILD_PYTHON_ZIP%"=="1" (
    call :build_python_zip
    if !ERRORLEVEL! neq 0 exit /b !ERRORLEVEL!
)

if "%BUILD_PACKAGE_ZIP%"=="1" (
    call :build_package_zip
    if !ERRORLEVEL! neq 0 exit /b !ERRORLEVEL!
)

if "%PUBLISH_RELEASE%"=="1" (
    if "%TEST_FAILED%"=="1" (
        echo ERROR: tests failed; GitHub Release was not published
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
    echo [Release] Publishing GitHub Release...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\Scripts\publish-release.ps1"
    if !ERRORLEVEL! neq 0 (
        echo ERROR: GitHub Release publish failed
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
)

if "%NO_PAUSE%"=="0" pause

exit /b 0

:build_python_zip
echo [Python ZIP] Creating lightweight Python package...
if not exist dist mkdir dist
if exist dist\ExcelMergeFork-python.zip del /Q dist\ExcelMergeFork-python.zip
powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\Scripts\make_python_zip.ps1"
if %ERRORLEVEL% neq 0 (
    echo ERROR: lightweight Python package failed
    exit /b 1
)
echo Artifact: %CD%\dist\ExcelMergeFork-python.zip
exit /b 0

:build_package_zip
echo [Package ZIP] Creating distribution packages...
if not exist dist mkdir dist
if "%BUILD_FULL_EXE%"=="1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\Scripts\make_distribution_zip.ps1" -Kind full
    if %ERRORLEVEL% neq 0 (
        echo ERROR: full distribution package failed
        exit /b 1
    )
)
if "%BUILD_LITE_EXE%"=="1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\Scripts\make_distribution_zip.ps1" -Kind lite
    if %ERRORLEVEL% neq 0 (
        echo ERROR: lite distribution package failed
        exit /b 1
    )
)
exit /b 0
