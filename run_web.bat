@echo off
cd /d "%~dp0"
echo.
echo ExcelMergeFork 调试台
echo http://127.0.0.1:5088
echo 改 web\ 里的页面后刷新浏览器。改接口后重新运行本脚本。
echo.
dotnet run --project src\ExcelMergeFork.Api\ExcelMergeFork.Api.csproj
