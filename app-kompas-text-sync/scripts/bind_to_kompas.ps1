[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $scriptDir
$outDir = Join-Path $appRoot "out"
$launcherVbs = Join-Path $scriptDir "kompas_button_launcher.vbs"
$launcherExe = Join-Path $appRoot "bin\app-kompas-text-sync.exe"
$buildLauncher = Join-Path $scriptDir "build_launcher.ps1"
$iconPath = Join-Path $appRoot "assets\app.ico"

if (-not (Test-Path -LiteralPath $launcherVbs)) {
    Write-Host "ERROR: launcher not found: $launcherVbs" -ForegroundColor Red
    exit 2
}

if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -Path $outDir -ItemType Directory -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $launcherExe) -and (Test-Path -LiteralPath $buildLauncher)) {
    try {
        & powershell -NoProfile -ExecutionPolicy Bypass -File $buildLauncher *> $null
    }
    catch {
    }
}

$program = ""
$arguments = ""
if (Test-Path -LiteralPath $launcherExe) {
    $program = $launcherExe
}
else {
    $program = Join-Path $env:WINDIR "System32\wscript.exe"
    $arguments = '"' + $launcherVbs + '"'
}

$bindingFile = Join-Path $outDir "kompas_button_binding.txt"
$bindingText = @"
KOMPAS-3D button binding (external command)
-------------------------------------------
Program : $program
Arguments: $arguments
Name    : KOMPAS Text Sync
Icon    : $iconPath

Action:
1) Open UI customization in KOMPAS-3D.
2) Add external command.
3) Copy Program and Arguments from this file.
4) Optionally set icon path to: $iconPath
"@

$bindingText | Set-Content -LiteralPath $bindingFile -Encoding UTF8

$clipboardText = "Program: $program`r`nArguments: $arguments`r`nIcon: $iconPath"
try {
    Set-Clipboard -Value $clipboardText
    $clipboardStatus = "OK"
}
catch {
    $clipboardStatus = "FAIL: $($_.Exception.Message)"
}

$kompasStatus = "COM_NOT_AVAILABLE"
try {
    $null = [Runtime.InteropServices.Marshal]::GetActiveObject("KOMPAS.Application.7")
    $kompasStatus = "RUNNING"
}
catch {
    try {
        $null = New-Object -ComObject "KOMPAS.Application.7"
        $kompasStatus = "COM_AVAILABLE"
    }
    catch {
        $kompasStatus = "COM_NOT_AVAILABLE"
    }
}

Write-Host ""
Write-Host "=== BIND TO KOMPAS-3D ===" -ForegroundColor Cyan
Write-Host "Program    : $program"
Write-Host "Arguments  : $arguments"
Write-Host "Icon       : $iconPath"
Write-Host "Output file: $bindingFile"
Write-Host "Clipboard  : $clipboardStatus"
Write-Host "KOMPAS COM : $kompasStatus"
Write-Host ""
