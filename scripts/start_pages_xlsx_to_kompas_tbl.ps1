[CmdletBinding()]
param(
    [string]$PagesHost = "127.0.0.1",
    [int]$PagesPort = 5511,
    [int]$UtilityPort = 38741,
    [string]$PairingToken = "kompas-pages-local",
    [switch]$OpenBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
$DocsRoot = Join-Path $RepoRoot "docs"
$RuntimeRoot = Join-Path $RepoRoot "out\web-pages-runtime"
$BuilderScript = Join-Path $RepoRoot "tools\runtime\build_runtime_config.mjs"
$StaticServerScript = Join-Path $RepoRoot "tools\runtime\serve_docs.mjs"
$PidFile = Join-Path $RuntimeRoot "pids.json"

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

    throw "WebBridge.Utility.exe was not found."
}

New-Item -ItemType Directory -Force -Path $RuntimeRoot | Out-Null

$StaticUrl = "http://{0}:{1}/index.html" -f $PagesHost, $PagesPort
$UtilityUrl = "http://127.0.0.1:{0}" -f $UtilityPort
$ConfigPath = Join-Path $RuntimeRoot "config.bootstrap.json"
$UtilityExePath = Resolve-UtilityExePath

$builderArgs = @(
    $BuilderScript,
    "--output", $ConfigPath,
    "--listen-url", $UtilityUrl,
    "--ui-url", $StaticUrl,
    "--pairing-token", $PairingToken,
    "--origin", ("http://{0}:{1}" -f $PagesHost, $PagesPort),
    "--origin", ("http://localhost:{0}" -f $PagesPort),
    "--log-file", (Join-Path $RuntimeRoot "utility.log"),
    "--diagnostics-dir", (Join-Path $RuntimeRoot "diagnostics"),
    "--profile-dir", (Join-Path $RuntimeRoot "profiles"),
    "--cache-dir", (Join-Path $RuntimeRoot "cache")
)
& node @builderArgs | Out-Null

$pagesProcess = Start-Process -FilePath "node" `
    -ArgumentList @($StaticServerScript, $DocsRoot, $PagesHost, "$PagesPort") `
    -WorkingDirectory $RepoRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput (Join-Path $RuntimeRoot "pages.stdout.log") `
    -RedirectStandardError (Join-Path $RuntimeRoot "pages.stderr.log") `
    -PassThru

$utilityProcess = Start-Process -FilePath $UtilityExePath `
    -ArgumentList @("--config", $ConfigPath) `
    -WorkingDirectory (Split-Path -Parent $UtilityExePath) `
    -WindowStyle Hidden `
    -RedirectStandardOutput (Join-Path $RuntimeRoot "utility.stdout.log") `
    -RedirectStandardError (Join-Path $RuntimeRoot "utility.stderr.log") `
    -PassThru

$launchUrl = "{0}?utilityUrl={1}&pairingToken={2}&autoConnect=1&workspaceRoot={3}" -f `
    $StaticUrl, `
    [uri]::EscapeDataString($UtilityUrl), `
    [uri]::EscapeDataString($PairingToken), `
    [uri]::EscapeDataString($RepoRoot)

@{
    pagesPid       = $pagesProcess.Id
    utilityPid     = $utilityProcess.Id
    staticUrl      = $StaticUrl
    utilityUrl     = $UtilityUrl
    launchUrl      = $launchUrl
    configPath     = $ConfigPath
    runtimeRoot    = $RuntimeRoot
    startedAtUtc   = [DateTime]::UtcNow.ToString("o")
} | ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 -Path $PidFile

Write-Output $launchUrl

if ($OpenBrowser) {
    Start-Process $launchUrl | Out-Null
}
