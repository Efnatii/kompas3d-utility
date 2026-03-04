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
$repoRoot = Split-Path -Parent $appRoot
$runCmd = Join-Path $appRoot "scripts\run_gui.cmd"
$runVbs = Join-Path $appRoot "scripts\run_gui.vbs"
$buildLauncher = Join-Path $appRoot "scripts\build_launcher.ps1"
$launcherExe = Join-Path $appRoot "bin\app-xlsx-to-kompas-tbl.exe"
$appIcon = Join-Path $appRoot "assets\app.ico"
$guiScript = Join-Path $appRoot "scripts\gui_import.ps1"
$buttonVbs = Join-Path $appRoot "scripts\kompas_button_launcher.vbs"
$resolveBridge = Join-Path $appRoot "scripts\resolve_kompas_doc_dir.py"
$exporter = Join-Path $repoRoot "xlsx-to-kompas-tbl\src\create_tbl.vbs"
$insertBridge = Join-Path $repoRoot "xlsx-to-kompas-tbl\src\insert_tbl_bridge.py"
$layoutConfig = Join-Path $repoRoot "xlsx-to-kompas-tbl\config\table_layout.ini"

Write-Info "App root: $appRoot"

foreach ($path in @($runCmd, $runVbs, $buildLauncher, $launcherExe, $appIcon, $guiScript, $buttonVbs, $resolveBridge, $exporter, $insertBridge, $layoutConfig)) {
    if (Test-Path -LiteralPath $path) {
        Write-Info "Found: $path"
    }
    else {
        Write-Warn "Missing: $path"
    }
}

if (Get-Command cscript.exe -ErrorAction SilentlyContinue) {
    Write-Info "cscript.exe found."
}
else {
    Write-Warn "cscript.exe not found in PATH."
}

if (Get-Command python -ErrorAction SilentlyContinue) {
    Write-Info "python found."
    try {
        python -c "import win32com.client" *> $null
        if ($LASTEXITCODE -eq 0) {
            Write-Info "pywin32 available."
        }
        else {
            Write-Warn "pywin32 is missing (pip install pywin32)."
        }
    }
    catch {
        Write-Warn "pywin32 probe failed: $($_.Exception.Message)"
    }
}
else {
    Write-Warn "python not found (needed for fallback bridge)."
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
