# EasyWord Windows sidecar freeze
# Output: dist/easyword/easyword.exe (for pure-flow-client extraResources)
#
# Uses project .venv (see requirements.txt). Do NOT run PyInstaller in qhub_v1 etc.

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

$VenvDir = Join-Path $Root ".venv"
$VenvPython = Join-Path $VenvDir "Scripts\python.exe"

function Ensure-Venv {
    if (-not (Test-Path $VenvPython)) {
        Write-Host "==> Creating venv: .venv"
        python -m venv $VenvDir
    }

    Write-Host "==> Installing deps from requirements.txt..."
    & $VenvPython -m pip install -q --upgrade pip
    & $VenvPython -m pip install -q -r (Join-Path $Root "requirements.txt")
    & $VenvPython -m pip install -q "pyinstaller>=6.0"
}

$ActiveEnv = $env:CONDA_DEFAULT_ENV
if ($ActiveEnv) {
    Write-Warning "Active conda env: $ActiveEnv. Build will use .venv instead."
}

Ensure-Venv

Write-Host "==> Python:" (& $VenvPython -c "import sys; print(sys.executable)")
Write-Host "==> Cleaning previous dist/easyword..."
if (Test-Path "dist\easyword") {
    Remove-Item -Recurse -Force "dist\easyword"
}

Write-Host "==> Running PyInstaller (onedir)..."
& $VenvPython -m PyInstaller --noconfirm --clean easyword.spec

if (-not (Test-Path "dist\easyword\easyword.exe")) {
    Write-Error "Build failed: dist\easyword\easyword.exe not found"
    exit 1
}

Write-Host "==> Done: dist\easyword\"
Write-Host "==> Running smoke test..."
& (Join-Path $PSScriptRoot "test_packed_sidecar.ps1")
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
