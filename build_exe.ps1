param(
    [switch]$SkipOcr,
    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (Test-Path -LiteralPath $venvPython) {
    $pythonExe = $venvPython
} else {
    $pythonExe = "python"
}

& $pythonExe --version
if ($LASTEXITCODE -ne 0) {
    throw "Python could not be started (exit code $LASTEXITCODE)."
}

if (-not $SkipInstall) {
    & $pythonExe -m pip install --upgrade pip
    if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed ($LASTEXITCODE)." }
    if ($SkipOcr) {
        & $pythonExe -m pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) { throw "Runtime dependency install failed ($LASTEXITCODE)." }
        & $pythonExe -m pip install "pyinstaller>=6.10,<7"
        if ($LASTEXITCODE -ne 0) { throw "PyInstaller install failed ($LASTEXITCODE)." }
    } else {
        & $pythonExe -m pip install torch==2.9.1 torchvision==0.24.1 `
            --index-url https://download.pytorch.org/whl/cpu
        if ($LASTEXITCODE -ne 0) { throw "CPU Torch install failed ($LASTEXITCODE)." }
        & $pythonExe -m pip install -r requirements-build.txt
        if ($LASTEXITCODE -ne 0) { throw "Build dependency install failed ($LASTEXITCODE)." }
    }
}

& $pythonExe -m unittest discover -s tests -v
if ($LASTEXITCODE -ne 0) { throw "Automated tests failed ($LASTEXITCODE)." }
& $pythonExe -m PyInstaller --clean --noconfirm .\mc_marking.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed ($LASTEXITCODE)." }

$exePath = Join-Path $PSScriptRoot "dist\CheckMate\CheckMate.exe"
$libraryPath = Join-Path $PSScriptRoot "dist\CheckMate\_internal\base_library.zip"
if (-not (Test-Path -LiteralPath $exePath)) {
    throw "Build failed: CheckMate.exe was not created."
}
if (-not (Test-Path -LiteralPath $libraryPath)) {
    throw "Build failed: base_library.zip is missing."
}

Copy-Item -LiteralPath (Join-Path $PSScriptRoot "LICENSE") `
    -Destination (Join-Path $PSScriptRoot "dist\CheckMate\LICENSE") -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "README.md") `
    -Destination (Join-Path $PSScriptRoot "dist\CheckMate\README.md") -Force

$process = Start-Process -FilePath $exePath -ArgumentList "--smoke-test" -PassThru
if (-not $process.WaitForExit(30000)) {
    Stop-Process -Id $process.Id -Force
    throw "Build failed: CheckMate smoke test timed out."
}
$process.Refresh()
if ($process.ExitCode -ne 0) {
    throw "Build failed: CheckMate smoke test returned $($process.ExitCode)."
}

Write-Host "Build complete: $exePath" -ForegroundColor Green
