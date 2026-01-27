@echo off
setlocal

set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
%PY% -m pip install pyinstaller

%PY% -m PyInstaller .\mc_marking.spec

echo Build complete. EXE is in .\dist\CheckMate.exe
pause
