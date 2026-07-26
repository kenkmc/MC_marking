@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build_exe.ps1" %*
if errorlevel 1 exit /b %errorlevel%
echo Build complete: %~dp0dist\CheckMate\CheckMate.exe
