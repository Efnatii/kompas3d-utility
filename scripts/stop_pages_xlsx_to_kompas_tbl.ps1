[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRootPath = Split-Path -Parent $PSScriptRoot
$PidPath = Join-Path $RepoRootPath "out\web-pages-runtime\pids.json"

if (-not (Test-Path -LiteralPath $PidPath)) {
    Write-Host "Runtime pid file was not found: $PidPath"
    exit 0
}

$payload = Get-Content -Raw -LiteralPath $PidPath | ConvertFrom-Json

foreach ($propertyName in @("pagesPid", "utilityPid")) {
    $targetPid = $payload.$propertyName
    if (-not $targetPid) {
        continue
    }

    $process = Get-Process -Id $targetPid -ErrorAction SilentlyContinue
    if ($null -ne $process) {
        Stop-Process -Id $targetPid -Force
        Write-Host "Stopped $propertyName = $targetPid"
    }
}

Remove-Item -LiteralPath $PidPath -Force
