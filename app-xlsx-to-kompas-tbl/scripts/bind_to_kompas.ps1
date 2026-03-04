[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appRoot = Split-Path -Parent $scriptDir
$outDir = Join-Path $appRoot "out"
$launcherVbs = Join-Path $scriptDir "kompas_button_launcher.vbs"
$launcherExe = Join-Path $appRoot "bin\app-xlsx-to-kompas-tbl.exe"
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

$program = $null
$arguments = ""

if (Test-Path -LiteralPath $launcherExe) {
    $program = $launcherExe
}
else {
    $program = Join-Path $env:WINDIR "System32\wscript.exe"
    $arguments = '"' + $launcherVbs + '"'
    if (-not (Test-Path -LiteralPath $program)) {
        Write-Host "ERROR: wscript.exe not found: $program" -ForegroundColor Red
        exit 3
    }
}

$bindingFile = Join-Path $outDir "kompas_button_binding.txt"
$bindingText = @"
KOMPAS-3D button binding (external command)
-------------------------------------------
Program : $program
Arguments: $arguments
Name    : Excel -> TBL
Icon    : $iconPath

Steps in KOMPAS:
1) Open KOMPAS-3D.
2) Open UI customization and add an external command.
3) Paste Program and Arguments from this file.
4) If KOMPAS allows custom icon, set icon path to: $iconPath
5) Save button on toolbar.
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

$kompasStatus = "NOT_FOUND"
try {
    $null = [Runtime.InteropServices.Marshal]::GetActiveObject("Kompas.Application.5")
    $kompasStatus = "RUNNING"
}
catch {
    try {
        $null = New-Object -ComObject "Kompas.Application.5"
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
Write-Host "Add external button in KOMPAS with these values." -ForegroundColor Yellow
Write-Host "After binding, button opens the Excel -> TBL GUI app."
