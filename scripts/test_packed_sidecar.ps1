# Smoke test for packaged EasyWord: start -> /health -> auto stop.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\scripts\test_packed_sidecar.ps1
#   powershell -ExecutionPolicy Bypass -File .\scripts\test_packed_sidecar.ps1 -KeepAlive

param(
    [string]$DistDir = "",
    [string]$HealthUrl = "http://127.0.0.1:18765/health",
    [int]$ExpectedApiVersion = 1,
    [int]$TimeoutSec = 20,
    [switch]$KeepAlive
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
if (-not $DistDir) {
    $DistDir = Join-Path $Root "dist\easyword"
}
$ExePath = Join-Path $DistDir "easyword.exe"

if (-not (Test-Path $ExePath)) {
    Write-Error "Not found: $ExePath`nRun scripts\build_sidecar.ps1 first."
    exit 1
}

function Stop-EasyWord {
    Get-Process -Name easyword -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

Write-Host "==> Stopping any existing easyword..."
Stop-EasyWord
Start-Sleep -Milliseconds 500

Write-Host "==> Starting: $ExePath"
$proc = Start-Process `
    -FilePath $ExePath `
    -WorkingDirectory $DistDir `
    -WindowStyle Hidden `
    -PassThru

$health = $null
$deadline = (Get-Date).AddSeconds($TimeoutSec)

try {
    Write-Host "==> Waiting for health (timeout ${TimeoutSec}s)..."
    while ((Get-Date) -lt $deadline) {
        if ($proc.HasExited) {
            Write-Error "easyword.exe exited early with code $($proc.ExitCode)"
            exit 1
        }
        try {
            $health = Invoke-RestMethod -Uri $HealthUrl -TimeoutSec 2
            if ($health.ok -eq $true) { break }
        } catch {
            Start-Sleep -Milliseconds 400
            continue
        }
    }

    if (-not $health -or $health.ok -ne $true) {
        Write-Error "Health check failed: $HealthUrl"
        exit 1
    }

    if ($health.apiVersion -ne $ExpectedApiVersion) {
        Write-Error "apiVersion mismatch: expected $ExpectedApiVersion, got $($health.apiVersion)"
        exit 1
    }

    Write-Host "==> OK"
    $health | ConvertTo-Json -Compress
    exit 0
} finally {
    if ($KeepAlive) {
        Write-Host "==> -KeepAlive: process left running (PID $($proc.Id))"
    } else {
        Write-Host "==> Stopping easyword (PID $($proc.Id))..."
        if (-not $proc.HasExited) {
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }
        Stop-EasyWord
        Write-Host "==> Stopped."
    }
}
