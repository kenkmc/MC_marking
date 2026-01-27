# Build standalone Windows EXE using PyInstaller
# Usage: Right click -> Run with PowerShell, or run in terminal:
#   powershell -ExecutionPolicy Bypass -File .\build_exe.ps1

$ErrorActionPreference = "Stop"

# Ensure venv is used if present
$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (Test-Path $venvPython) {
    $python = $venvPython
} else {
    $python = "python"
}

& $python -m pip install --upgrade pip
& $python -m pip install -r requirements.txt
& $python -m pip install pyinstaller

# Build
& $python -m PyInstaller .\mc_marking.spec

Write-Host "Build complete. EXE is in .\dist\CheckMate.exe" -ForegroundColor Green
