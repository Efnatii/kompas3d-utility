[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Write-Info([string]$Message) {
    Write-Host "INFO: $Message"
}

function Write-Warn([string]$Message) {
    Write-Host "WARN: $Message" -ForegroundColor Yellow
}

function Test-ProgId([string]$ProgId) {
    try {
        $type = [type]::GetTypeFromProgID($ProgId)
        return $null -ne $type
    }
    catch {
        return $false
    }
}

$appRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$runCmd = Join-Path $appRoot "scripts\run_gui.cmd"
$runVbs = Join-Path $appRoot "scripts\run_gui.vbs"
$guiScript = Join-Path $appRoot "scripts\gui_sync.ps1"
$engineScript = Join-Path $appRoot "scripts\kompas_excel_text_sync.py"
$syncLogic = Join-Path $appRoot "scripts\sync_logic.py"
$buildLauncher = Join-Path $appRoot "scripts\build_launcher.ps1"
$launcherExe = Join-Path $appRoot "bin\app-kompas-text-sync.exe"
$iconPath = Join-Path $appRoot "assets\app.ico"
$buttonVbs = Join-Path $appRoot "scripts\kompas_button_launcher.vbs"

Write-Info "App root: $appRoot"

foreach ($path in @($runCmd, $runVbs, $guiScript, $engineScript, $syncLogic, $buildLauncher, $launcherExe, $iconPath, $buttonVbs)) {
    if (Test-Path -LiteralPath $path) {
        Write-Info "Found: $path"
    }
    else {
        Write-Warn "Missing: $path"
    }
}

if (Get-Command python -ErrorAction SilentlyContinue) {
    Write-Info "python found."
    try {
        python -c "import win32com.client" *> $null
        if ($LASTEXITCODE -eq 0) {
            Write-Info "pywin32 available."
        }
        else {
            Write-Warn "pywin32 missing (pip install pywin32)."
        }
    }
    catch {
        Write-Warn "pywin32 probe failed: $($_.Exception.Message)"
    }
}
else {
    Write-Warn "python not found."
}

if (Test-ProgId "Excel.Application") {
    Write-Info "Excel COM available."
}
else {
    Write-Warn "Excel COM not available."
}

if (Test-ProgId "Kompas.Application.5") {
    Write-Info "KOMPAS COM available."
}
else {
    Write-Warn "KOMPAS COM not available."
}

Write-Info "Selfcheck completed."
