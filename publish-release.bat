@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Default: check GitHub CLI auth, build GUI exe, run tests, publish GitHub Release.
REM Options: --no-test skips tests, --prerelease publishes a prerelease, --no-pause skips pause prompts.
set "RUN_TEST=1"
set "PRERELEASE=0"
set "NO_PAUSE=0"

:parse
if "%~1"=="" goto done_parse
if /i "%~1"=="--no-test" set "RUN_TEST=0"
if /i "%~1"=="--dist" set "RUN_TEST=0"
if /i "%~1"=="no-test" set "RUN_TEST=0"
if /i "%~1"=="skip" set "RUN_TEST=0"
if /i "%~1"=="--prerelease" set "PRERELEASE=1"
if /i "%~1"=="--no-pause" set "NO_PAUSE=1"
shift
goto parse
:done_parse

echo ========== Fork Excel Merge Tool Release ==========
echo.

where gh >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: GitHub CLI ^(gh^) was not found
    echo.
    echo Install it first:
    echo   winget install --id GitHub.cli -e
    echo.
    echo Then reopen the terminal and run:
    echo   gh auth login
    echo.
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

gh auth status >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: GitHub CLI is not logged in or the session is invalid
    echo.
    echo Run:
    echo   gh auth login
    echo.
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

echo [1/3] Building GUI exe...
call "%~dp0build_exe.bat" --gui
if !ERRORLEVEL! neq 0 (
    echo ERROR: build failed
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

if "%RUN_TEST%"=="1" (
    echo.
    echo [2/3] Running merge mode tests...
    call "%~dp0run_merge_mode_tests.bat"
    if !ERRORLEVEL! neq 0 (
        echo ERROR: tests failed; release stopped
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
) else (
    echo.
    echo [2/3] Tests skipped
)

echo.
echo [3/3] Publishing GitHub Release...
if "%PRERELEASE%"=="1" (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\publish-release.ps1" -Prerelease
) else (
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\publish-release.ps1"
)
if !ERRORLEVEL! neq 0 (
    echo ERROR: GitHub Release publish failed
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

echo.
echo ========== Release Complete ==========
if "%NO_PAUSE%"=="0" pause
exit /b 0
