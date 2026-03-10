[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
$PidFile = Join-Path $RepoRoot "out\web-pages-runtime\pids.json"

if (-not (Test-Path -LiteralPath $PidFile)) {
    Write-Output "No runtime pid file: $PidFile"
    exit 0
}

$payload = Get-Content -Raw -LiteralPath $PidFile | ConvertFrom-Json

foreach ($propertyName in @("pagesPid", "utilityPid")) {
    $targetPid = [int]$payload.$propertyName
    if ($targetPid -le 0) {
        continue
    }

    try {
        $process = Get-Process -Id $targetPid -ErrorAction Stop
        Stop-Process -Id $process.Id -Force -ErrorAction Stop
        Write-Output "Stopped $propertyName=$targetPid"
    } catch {
        Write-Output "Skip $propertyName=$targetPid"
    }
}

Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
