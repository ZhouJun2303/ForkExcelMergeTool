@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

set "RUN_TEST=1"
set "NO_PAUSE=0"
set "BUILD_PACKAGE_ZIP=0"

:parse
if "%~1"=="" goto done_parse
if /i "%~1"=="--test" set "RUN_TEST=1"
if /i "%~1"=="--dist" set "RUN_TEST=0"
if /i "%~1"=="--package-zip" set "BUILD_PACKAGE_ZIP=1"
if /i "%~1"=="--no-pause" set "NO_PAUSE=1"
if /i "%~1"=="no-test" set "RUN_TEST=0"
if /i "%~1"=="skip" set "RUN_TEST=0"
shift
goto parse
:done_parse

echo ========== ExcelMergeFork Package ==========
echo.

if "%RUN_TEST%"=="1" (
    echo [1/3] Running tests...
    dotnet test src\ExcelMergeFork.Tests\ExcelMergeFork.Tests.csproj -c Release --nologo
    if !ERRORLEVEL! neq 0 (
        echo ERROR: tests failed
        if "%NO_PAUSE%"=="0" pause
        exit /b 1
    )
) else (
    echo [1/3] Tests skipped
)

echo.
echo [2/3] Publishing self-contained ExcelMergeFork.exe...
dotnet publish src\ExcelMergeFork.App\ExcelMergeFork.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=none -o dist\csharp
if !ERRORLEVEL! neq 0 (
    echo ERROR: publish failed
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)

copy /Y "dist\csharp\ExcelMergeFork.exe" "ExcelMergeFork.exe" >nul
if !ERRORLEVEL! neq 0 (
    echo ERROR: cannot replace ExcelMergeFork.exe — close the running app and retry
    if "%NO_PAUSE%"=="0" pause
    exit /b 1
)
certutil -hashfile ExcelMergeFork.exe SHA256 | findstr /v ":" > ExcelMergeFork.exe.sha256

echo.
echo [3/3] Done. Output: %CD%\ExcelMergeFork.exe

if "%BUILD_PACKAGE_ZIP%"=="1" (
    if exist dist\ExcelMergeFork-package.zip del /Q dist\ExcelMergeFork-package.zip
    powershell -NoProfile -Command "Compress-Archive -Force -Path ExcelMergeFork.exe,README.md,install_fork_integration.bat,uninstall_fork_integration.bat,install_git_integration.bat,uninstall_git_integration.bat -DestinationPath dist\ExcelMergeFork-package.zip"
    echo Package zip: %CD%\dist\ExcelMergeFork-package.zip
)

if "%NO_PAUSE%"=="0" pause
exit /b 0
