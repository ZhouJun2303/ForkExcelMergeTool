@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

set "RUN_TEST=1"
set "PRERELEASE=0"
set "NO_PAUSE=0"

:parse
if "%~1"=="" goto done_parse
if /i "%~1"=="--no-test" set "RUN_TEST=0"
if /i "%~1"=="--prerelease" set "PRERELEASE=1"
if /i "%~1"=="--no-pause" set "NO_PAUSE=1"
shift
goto parse
:done_parse

where gh >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ERROR: 需要 GitHub CLI。先执行 winget install --id GitHub.cli -e 和 gh auth login
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

if "%RUN_TEST%"=="1" (
    call "%~dp0package.bat" --package-zip --no-pause
) else (
    call "%~dp0package.bat" --dist --package-zip --no-pause
)
if !ERRORLEVEL! neq 0 (
    echo ERROR: package failed
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

set "TAG=v3.0"
if "%PRERELEASE%"=="1" (
    gh release create %TAG% --title %TAG% --notes "ExcelMergeFork %TAG% C# Fluent" --prerelease ExcelMergeFork.exe ExcelMergeFork.exe.sha256 dist\ExcelMergeFork-package.zip
) else (
    gh release create %TAG% --title %TAG% --notes "ExcelMergeFork %TAG% C# Fluent" ExcelMergeFork.exe ExcelMergeFork.exe.sha256 dist\ExcelMergeFork-package.zip
)
if "%NO_PAUSE%"=="0" pause
exit /b %ERRORLEVEL%
