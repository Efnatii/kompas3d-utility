[CmdletBinding()]
param(
    [string]$PagesHost = "127.0.0.1",
    [int]$PagesPort = 5511,
    [string]$UtilityUrl = "http://127.0.0.1:38741",
    [string]$PairingToken = "kompas-pages-local"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRootPath = Split-Path -Parent $PSScriptRoot
$RuntimeRootPath = Join-Path $RepoRootPath "out\web-pages-runtime"
$ConfigPath = Join-Path $RuntimeRootPath "config.runtime.json"
$PidPath = Join-Path $RuntimeRootPath "pids.json"
$StaticUrl = "http://$PagesHost`:$PagesPort/index.html"
$LaunchUrl = "{0}?utilityUrl={1}&pairingToken={2}&autoConnect=1" -f `
    $StaticUrl, `
    [uri]::EscapeDataString($UtilityUrl), `
    [uri]::EscapeDataString($PairingToken)

function Resolve-UtilityExePath {
    $candidates = @(
        "C:\_GIT_\web-bridge-utility\artifacts\publish\utility\win-x64\WebBridge.Utility.exe",
        "C:\_GIT_\web-bridge-utility\src\WebBridge.Utility\bin\Release\net8.0\win-x64\WebBridge.Utility.exe",
        "C:\_GIT_\web-bridge-utility\artifacts\release\web-bridge-utility-1.0.0-win-x64\WebBridge.Utility.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw "WebBridge.Utility.exe was not found under C:\_GIT_\web-bridge-utility"
}

New-Item -ItemType Directory -Force -Path $RuntimeRootPath | Out-Null

$builderArgs = @(
    "tools\runtime\build_runtime_config.mjs",
    "--output", $ConfigPath,
    "--listen-url", $UtilityUrl,
    "--ui-url", $StaticUrl,
    "--origin", "http://$PagesHost`:$PagesPort",
    "--origin", "http://localhost:$PagesPort",
    "--pairing-token", $PairingToken,
    "--log-file", (Join-Path $RuntimeRootPath "utility.log"),
    "--diagnostics-dir", (Join-Path $RuntimeRootPath "diagnostics"),
    "--profile-dir", (Join-Path $RuntimeRootPath "profiles"),
    "--cache-dir", (Join-Path $RuntimeRootPath "cache")
)

$builder = Start-Process -FilePath "node.exe" -ArgumentList $builderArgs -WorkingDirectory $RepoRootPath -PassThru -Wait -NoNewWindow
if ($builder.ExitCode -ne 0) {
    throw "Runtime config build failed with exit code $($builder.ExitCode)."
}

$utilityExePath = Resolve-UtilityExePath
$pagesArgs = @("-3", "-m", "http.server", "$PagesPort", "--bind", $PagesHost, "--directory", "docs")
$utilityArgs = @("--config", $ConfigPath)

$pagesProcess = Start-Process -FilePath "py.exe" -ArgumentList $pagesArgs -WorkingDirectory $RepoRootPath -PassThru
$utilityProcess = Start-Process -FilePath $utilityExePath -ArgumentList $utilityArgs -WorkingDirectory (Split-Path -Parent $utilityExePath) -PassThru

$payload = [ordered]@{
    pagesPid     = $pagesProcess.Id
    utilityPid   = $utilityProcess.Id
    pagesHost    = $PagesHost
    pagesPort    = $PagesPort
    utilityUrl   = $UtilityUrl
    pairingToken = $PairingToken
    configPath   = $ConfigPath
    launchUrl    = $LaunchUrl
}

$payload | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $PidPath -Encoding utf8

Write-Host "Pages URL   : $StaticUrl"
Write-Host "Launch URL  : $LaunchUrl"
Write-Host "Utility URL : $UtilityUrl"
Write-Host "PIDs saved  : $PidPath"
