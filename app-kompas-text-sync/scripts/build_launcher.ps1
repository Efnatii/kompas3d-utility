[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $scriptDir
$source = Join-Path $scriptDir "AppLauncher.cs"
$icon = Join-Path $appRoot "assets\app.ico"
$outDir = Join-Path $appRoot "bin"
$outExe = Join-Path $outDir "app-kompas-text-sync.exe"

if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
    Write-Host "ERROR: source file not found: $source" -ForegroundColor Red
    exit 2
}

if (-not (Test-Path -LiteralPath $icon -PathType Leaf)) {
    Write-Host "ERROR: icon file not found: $icon" -ForegroundColor Red
    exit 3
}

if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -Path $outDir -ItemType Directory -Force | Out-Null
}

$cscCandidates = @(
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)

$csc = $null
foreach ($candidate in $cscCandidates) {
    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        $csc = $candidate
        break
    }
}

if ($null -eq $csc) {
    Write-Host "ERROR: csc.exe not found in expected paths." -ForegroundColor Red
    exit 4
}

& $csc `
    /nologo `
    /target:winexe `
    /optimize+ `
    /win32icon:"$icon" `
    /out:"$outExe" `
    /r:System.dll `
    /r:System.Windows.Forms.dll `
    /r:System.Core.dll `
    "$source"

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: launcher compilation failed with code $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}

Write-Host "OK: launcher built: $outExe" -ForegroundColor Green
