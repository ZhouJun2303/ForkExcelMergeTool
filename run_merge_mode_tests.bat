@echo off
chcp 65001 >nul
cd /d "%~dp0"
dotnet test src\ExcelMergeFork.Tests\ExcelMergeFork.Tests.csproj -c Release --nologo
exit /b %ERRORLEVEL%
