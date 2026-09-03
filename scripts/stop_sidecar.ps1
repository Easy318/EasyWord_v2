# Stop running EasyWord sidecar processes (packed exe or dev leftovers on port 18765).

$ErrorActionPreference = "SilentlyContinue"

$procs = Get-Process -Name easyword
if ($procs) {
    $procs | Stop-Process -Force
    Write-Host "Stopped $($procs.Count) easyword process(es)."
} else {
    Write-Host "No easyword process found."
}
