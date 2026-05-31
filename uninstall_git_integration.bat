@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\uninstall_git_integration.ps1"
exit /b %ERRORLEVEL%
