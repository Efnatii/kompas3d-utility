[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Write-Info([string]$Message) {
    Write-Host "INFO: $Message"
}

function Write-Warn([string]$Message) {
    Write-Host "WARN: $Message" -ForegroundColor Yellow
}

function Write-Fail([string]$Message) {
    Write-Host "FAIL: $Message" -ForegroundColor Red
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

function Test-VbsEncoding([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return @{
            Ok      = $false
            Message = "File not found: $Path"
        }
    }

    [byte[]]$bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return @{
            Ok      = $true
            Message = "Encoding is UTF-16 LE with BOM (recommended)."
        }
    }

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return @{
            Ok      = $false
            Message = "UTF-8 BOM detected. VBScript may fail with 'Nedopustimyy znak (1,1)'."
        }
    }

    return @{
        Ok      = $true
        Message = "No BOM detected (likely ANSI/cp1251). Acceptable, but UTF-16 LE is safer."
    }
}

function Test-LayoutConfig([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        return @{
            Ok      = $false
            Message = "Layout config not found: $Path"
        }
    }

    $pairs = @{}
    foreach ($line in Get-Content -LiteralPath $Path) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($trimmed.StartsWith("#") -or $trimmed.StartsWith(";")) { continue }
        $chunks = $trimmed.Split("=", 2)
        if ($chunks.Count -eq 2) {
            $pairs[$chunks[0].Trim().ToLowerInvariant()] = $chunks[1].Trim()
        }
    }

    if (-not $pairs.ContainsKey("mode")) {
        return @{
            Ok      = $false
            Message = "Missing required key 'mode' in $Path"
        }
    }

    $mode = $pairs["mode"].ToLowerInvariant()
    if ($mode -ne "cell" -and $mode -ne "table") {
        return @{
            Ok      = $false
            Message = "Invalid mode '$mode' (expected: cell or table)."
        }
    }

    return @{
        Ok      = $true
        Message = "Layout config looks valid (mode=$mode)."
    }
}

$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$vbsPath = Join-Path $root "src\create_tbl.vbs"
$layoutConfigPath = Join-Path $root "config\table_layout.ini"

Write-Info "Project root: $root"

$encodingCheck = Test-VbsEncoding -Path $vbsPath
if ($encodingCheck.Ok) {
    Write-Info "create_tbl.vbs: $($encodingCheck.Message)"
}
else {
    Write-Fail "create_tbl.vbs: $($encodingCheck.Message)"
}

$layoutCheck = Test-LayoutConfig -Path $layoutConfigPath
if ($layoutCheck.Ok) {
    Write-Info "table_layout.ini: $($layoutCheck.Message)"
}
else {
    Write-Warn "table_layout.ini: $($layoutCheck.Message)"
}

$hasExcel = Test-ProgId "Excel.Application"
$hasKompas = Test-ProgId "Kompas.Application.5"

if ($hasExcel) { Write-Info "Excel COM ProgID found (Excel.Application)." }
else { Write-Warn "Excel COM ProgID not found. XLSX->COM export will fail." }

if ($hasKompas) { Write-Info "KOMPAS COM ProgID found (Kompas.Application.5)." }
else { Write-Warn "KOMPAS COM ProgID not found. Integration test will be skipped." }

$pythonOk = $false
try {
    $null = & python --version
    if ($LASTEXITCODE -eq 0) {
        $pythonOk = $true
        Write-Info "Python is available."
    }
}
catch {
    $pythonOk = $false
}

if (-not $pythonOk) {
    Write-Fail "Python is not found in PATH."
    exit 1
}

foreach ($pkg in @(
    @{ Name = "pytest"; Import = "import pytest" },
    @{ Name = "openpyxl"; Import = "import openpyxl" },
    @{ Name = "pywin32"; Import = "import win32com.client" }
)) {
    & python -c $pkg.Import 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Info "Python package '$($pkg.Name)' is available."
    }
    else {
        Write-Warn "Python package '$($pkg.Name)' is missing. Install: pip install -r requirements.txt"
    }
}

if (-not $encodingCheck.Ok) {
    exit 2
}

Write-Info "Selfcheck completed."
