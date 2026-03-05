[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$OutputDir = "",
    [switch]$NoClipboard,
    [switch]$SkipBuild,
    [switch]$WritePerAppFiles,
    [string]$KompasKitConfigPath = "",
    [switch]$SkipConfiguratorUpdate,
    [switch]$WaitForKompasExit,
    [int]$WaitTimeoutSec = 600
)

$ErrorActionPreference = "Stop"

function Convert-AppFolderToDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )

    $baseName = $FolderName -replace "^app-", ""
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        return $FolderName
    }

    $tokens = $baseName -split "[-_ ]+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    if ($tokens.Count -eq 0) {
        return $FolderName
    }

    $words = foreach ($token in $tokens) {
        if ($token.Length -le 1) {
            $token.ToUpperInvariant()
        }
        else {
            $token.Substring(0, 1).ToUpperInvariant() + $token.Substring(1)
        }
    }

    return ($words -join " ")
}

function Get-WindowsDir {
    if (-not [string]::IsNullOrWhiteSpace($env:WINDIR) -and (Test-Path -LiteralPath $env:WINDIR)) {
        return $env:WINDIR
    }

    if (-not [string]::IsNullOrWhiteSpace($env:SystemRoot) -and (Test-Path -LiteralPath $env:SystemRoot)) {
        return $env:SystemRoot
    }

    if (Test-Path -LiteralPath "C:\Windows") {
        return "C:\Windows"
    }

    return ""
}

function Get-WscriptPath {
    $windowsDir = Get-WindowsDir
    $candidate = ""
    if (-not [string]::IsNullOrWhiteSpace($windowsDir)) {
        $candidate = Join-Path $windowsDir "System32\wscript.exe"
    }
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return "wscript.exe"
}

function Get-CmdPath {
    if (-not [string]::IsNullOrWhiteSpace($env:ComSpec) -and (Test-Path -LiteralPath $env:ComSpec)) {
        return $env:ComSpec
    }

    $windowsDir = Get-WindowsDir
    $candidate = ""
    if (-not [string]::IsNullOrWhiteSpace($windowsDir)) {
        $candidate = Join-Path $windowsDir "System32\cmd.exe"
    }
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return "cmd.exe"
}

function Get-PowerShellExecutable {
    $powershellCmd = Get-Command powershell.exe -ErrorAction SilentlyContinue
    if ($null -ne $powershellCmd) {
        return $powershellCmd.Source
    }

    $pwshCmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($null -ne $pwshCmd) {
        return $pwshCmd.Source
    }

    $pwshAlias = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($null -ne $pwshAlias) {
        return $pwshAlias.Source
    }

    return $null
}

function Resolve-KompasKitConfig {
    param(
        [string]$ExplicitPath
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        if (Test-Path -LiteralPath $ExplicitPath) {
            return (Resolve-Path -LiteralPath $ExplicitPath).Path
        }

        throw "KOMPAS kit config not found: $ExplicitPath"
    }

    if ([string]::IsNullOrWhiteSpace($env:USERPROFILE)) {
        return $null
    }

    $kompasRoot = Join-Path $env:USERPROFILE "AppData\Roaming\ASCON\KOMPAS-3D"
    if (-not (Test-Path -LiteralPath $kompasRoot)) {
        return $null
    }

    # Prefer config matching the currently running KOMPAS version, if detected.
    try {
        $runningKompas = Get-Process -Name KOMPAS -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $runningKompas -and -not [string]::IsNullOrWhiteSpace($runningKompas.Path)) {
            $installRoot = Split-Path -Parent (Split-Path -Parent $runningKompas.Path)
            if ($installRoot -match "KOMPAS-3D v(?<ver>\d+(?:\.\d+)?)$") {
                $versionName = $Matches["ver"]
                $runningCandidate = Join-Path $kompasRoot "$versionName\KOMPAS.kit.config"
                if (Test-Path -LiteralPath $runningCandidate) {
                    return $runningCandidate
                }
            }
        }
    }
    catch {
    }

    $versionDirs = @(
        Get-ChildItem -LiteralPath $kompasRoot -Directory -ErrorAction SilentlyContinue |
            Sort-Object -Property {
                try {
                    [version]$_.Name
                }
                catch {
                    [version]"0.0"
                }
            } -Descending
    )

    foreach ($dir in $versionDirs) {
        $candidate = Join-Path $dir.FullName "KOMPAS.kit.config"
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Wait-ForKompasProcessExit {
    param(
        [int]$TimeoutSec = 600
    )

    if ($TimeoutSec -lt 0) {
        $TimeoutSec = 0
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ($true) {
        $running = @(Get-Process -Name KOMPAS -ErrorAction SilentlyContinue)
        if ($running.Count -eq 0) {
            return $true
        }

        if ($TimeoutSec -eq 0) {
            return $false
        }

        if ((Get-Date) -ge $deadline) {
            return $false
        }

        Start-Sleep -Seconds 1
    }
}

function Set-XmlAttributeValue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlElement]$Element,
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [AllowNull()]
        [string]$Value
    )

    if ($null -eq $Value) {
        $Value = ""
    }

    $hasAttribute = $Element.HasAttribute($Name)
    $current = $Element.GetAttribute($Name)
    if ($hasAttribute -and $current -eq $Value) {
        return $false
    }

    $Element.SetAttribute($Name, $Value)
    return $true
}

function Update-KompasConfiguratorUtilities {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$Bindings,
        [Parameter(Mandatory = $true)]
        [string]$KitConfigPath
    )

    $ready = @($Bindings | Where-Object { $_.Status -eq "READY" })
    if ($ready.Count -eq 0) {
        return [pscustomobject]@{
            Status     = "SKIPPED_NO_READY"
            Path       = $KitConfigPath
            BackupPath = ""
            Added      = 0
            Updated    = 0
            Removed    = 0
            Changed    = $false
        }
    }

    [xml]$xml = Get-Content -LiteralPath $KitConfigPath -Raw
    $group = $xml.SelectSingleNode("/Kit_Config/Groups/Group")
    if ($null -eq $group) {
        throw "Group node not found in: $KitConfigPath"
    }

    $added = 0
    $updated = 0
    $removed = 0
    $changed = $false
    $managedPrefix = "Managed by kompas3d-utility"

    foreach ($item in $ready) {
        $description = "$managedPrefix ($($item.AppFolder))"
        $existingNode = $null
        $matchedNodes = New-Object System.Collections.Generic.List[System.Xml.XmlElement]

        foreach ($utility in @($group.SelectNodes("Utility"))) {
            $utilityPath = $utility.GetAttribute("path")
            $utilityDisplayName = $utility.GetAttribute("displayName")
            $utilityParams = $utility.GetAttribute("params")
            $utilityDescription = $utility.GetAttribute("description")
            $isManagedNode = (
                -not [string]::IsNullOrWhiteSpace($utilityDescription) -and
                $utilityDescription.StartsWith($managedPrefix)
            )

            if (
                $utilityPath -eq $item.Program -or
                (
                    $utilityDisplayName -eq $item.Name -and
                    $utilityParams -eq $item.Arguments
                ) -or
                (
                    $utilityDescription -eq $description
                ) -or
                (
                    $isManagedNode -and
                    $utilityDisplayName -eq $item.Name
                )
            ) {
                $matchedNodes.Add($utility) | Out-Null
            }
        }

        if ($matchedNodes.Count -gt 0) {
            $existingNode = $matchedNodes[0]
        }

        $isNewNode = $false
        if ($null -eq $existingNode) {
            $existingNode = $xml.CreateElement("Utility")
            [void]$group.AppendChild($existingNode)
            $isNewNode = $true
            $added++
            $changed = $true
        }

        $entryChanged = $false
        $entryChanged = (Set-XmlAttributeValue -Element $existingNode -Name "path" -Value $item.Program) -or $entryChanged
        $entryChanged = (Set-XmlAttributeValue -Element $existingNode -Name "displayName" -Value $item.Name) -or $entryChanged
        $entryChanged = (Set-XmlAttributeValue -Element $existingNode -Name "params" -Value $item.Arguments) -or $entryChanged
        $entryChanged = (Set-XmlAttributeValue -Element $existingNode -Name "description" -Value $description) -or $entryChanged

        if ((-not $isNewNode) -and $entryChanged) {
            $updated++
        }
        if ($entryChanged) {
            $changed = $true
        }

        if ($matchedNodes.Count -gt 1) {
            for ($idx = 1; $idx -lt $matchedNodes.Count; $idx++) {
                [void]$group.RemoveChild($matchedNodes[$idx])
                $removed++
                $changed = $true
            }
        }
    }

    if (-not $changed) {
        return [pscustomobject]@{
            Status     = "NO_CHANGES"
            Path       = $KitConfigPath
            BackupPath = ""
            Added      = 0
            Updated    = 0
            Removed    = 0
            Changed    = $false
        }
    }

    $backupPath = "$KitConfigPath.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Copy-Item -LiteralPath $KitConfigPath -Destination $backupPath -Force
    $xml.Save($KitConfigPath)

    return [pscustomobject]@{
        Status     = "UPDATED"
        Path       = $KitConfigPath
        BackupPath = $backupPath
        Added      = $added
        Updated    = $updated
        Removed    = $removed
        Changed    = $true
    }
}

function Resolve-AppLauncher {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$AppDir,
        [switch]$SkipBuildLauncher
    )

    $appRoot = $AppDir.FullName
    $appFolderName = $AppDir.Name
    $scriptsDir = Join-Path $appRoot "scripts"
    $binDir = Join-Path $appRoot "bin"
    $expectedExe = Join-Path $binDir ($appFolderName + ".exe")
    $buildScript = Join-Path $scriptsDir "build_launcher.ps1"
    $note = ""

    if (-not $SkipBuildLauncher -and -not (Test-Path -LiteralPath $expectedExe) -and (Test-Path -LiteralPath $buildScript)) {
        $powerShellExe = Get-PowerShellExecutable
        if ($null -eq $powerShellExe) {
            $note = "cannot run build_launcher.ps1: PowerShell host not found"
        }
        else {
            try {
                & $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $buildScript *> $null
            }
            catch {
                $note = "build_launcher.ps1 failed: $($_.Exception.Message)"
            }
        }
    }

    if (Test-Path -LiteralPath $expectedExe) {
        return [pscustomobject]@{
            Program        = $expectedExe
            Arguments      = ""
            LauncherSource = "bin/<app>.exe"
            Status         = "READY"
            Note           = $note
        }
    }

    if (Test-Path -LiteralPath $binDir) {
        $exeCandidates = Get-ChildItem -LiteralPath $binDir -Filter *.exe -File -ErrorAction SilentlyContinue | Sort-Object Name
        if ($exeCandidates.Count -gt 0) {
            $selectedExe = $exeCandidates[0].FullName
            if ($exeCandidates.Count -gt 1) {
                $allExeNames = ($exeCandidates | Select-Object -ExpandProperty Name) -join ", "
                if ([string]::IsNullOrWhiteSpace($note)) {
                    $note = "multiple EXE found, selected first alphabetically: $allExeNames"
                }
                else {
                    $note = "$note; multiple EXE found, selected first alphabetically: $allExeNames"
                }
            }

            return [pscustomobject]@{
                Program        = $selectedExe
                Arguments      = ""
                LauncherSource = "bin/*.exe"
                Status         = "READY"
                Note           = $note
            }
        }
    }

    $buttonLauncherVbs = Join-Path $scriptsDir "kompas_button_launcher.vbs"
    if (Test-Path -LiteralPath $buttonLauncherVbs) {
        return [pscustomobject]@{
            Program        = (Get-WscriptPath)
            Arguments      = ('"' + $buttonLauncherVbs + '"')
            LauncherSource = "scripts/kompas_button_launcher.vbs"
            Status         = "READY"
            Note           = $note
        }
    }

    $runGuiVbs = Join-Path $scriptsDir "run_gui.vbs"
    if (Test-Path -LiteralPath $runGuiVbs) {
        return [pscustomobject]@{
            Program        = (Get-WscriptPath)
            Arguments      = ('"' + $runGuiVbs + '"')
            LauncherSource = "scripts/run_gui.vbs"
            Status         = "READY"
            Note           = $note
        }
    }

    $runCmd = Join-Path $scriptsDir "run.cmd"
    if (Test-Path -LiteralPath $runCmd) {
        return [pscustomobject]@{
            Program        = (Get-CmdPath)
            Arguments      = ('/c "' + $runCmd + '"')
            LauncherSource = "scripts/run.cmd"
            Status         = "READY"
            Note           = $note
        }
    }

    if ([string]::IsNullOrWhiteSpace($note)) {
        $note = "launcher not found"
    }
    else {
        $note = "$note; launcher not found"
    }

    return [pscustomobject]@{
        Program        = ""
        Arguments      = ""
        LauncherSource = ""
        Status         = "NO_LAUNCHER"
        Note           = $note
    }
}

function Write-PerAppBindingFile {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Binding
    )

    $appOutDir = Join-Path $Binding.AppPath "out"
    if (-not (Test-Path -LiteralPath $appOutDir)) {
        New-Item -Path $appOutDir -ItemType Directory -Force | Out-Null
    }

    $bindingFile = Join-Path $appOutDir "kompas_button_binding.txt"
    $bindingText = @"
KOMPAS-3D button binding (external command)
-------------------------------------------
Program : $($Binding.Program)
Arguments: $($Binding.Arguments)
Name    : $($Binding.Name)
Icon    : $($Binding.Icon)
Status  : $($Binding.Status)
Source  : $($Binding.LauncherSource)
"@

    if (-not [string]::IsNullOrWhiteSpace($Binding.Note)) {
        $bindingText = $bindingText + "`r`nNote    : $($Binding.Note)"
    }

    $bindingText | Set-Content -LiteralPath $bindingFile -Encoding UTF8
    return $bindingFile
}

if (-not (Test-Path -LiteralPath $RepoRoot)) {
    Write-Host "ERROR: Repo root does not exist: $RepoRoot" -ForegroundColor Red
    exit 2
}

$RepoRoot = (Resolve-Path -LiteralPath $RepoRoot).Path

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $RepoRoot "out"
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
}

$appDirs = @(Get-ChildItem -LiteralPath $RepoRoot -Directory -Filter "app-*" | Sort-Object Name)
if ($appDirs.Count -eq 0) {
    Write-Host "WARN: no app-* directories found under $RepoRoot" -ForegroundColor Yellow
    exit 1
}

$bindings = New-Object System.Collections.Generic.List[object]
foreach ($appDir in $appDirs) {
    $launcher = Resolve-AppLauncher -AppDir $appDir -SkipBuildLauncher:$SkipBuild
    $iconPath = Join-Path $appDir.FullName "assets\app.ico"
    if (-not (Test-Path -LiteralPath $iconPath)) {
        $iconPath = ""
    }

    $binding = [pscustomobject]@{
        AppFolder      = $appDir.Name
        AppPath        = $appDir.FullName
        Name           = Convert-AppFolderToDisplayName -FolderName $appDir.Name
        Program        = $launcher.Program
        Arguments      = $launcher.Arguments
        Icon           = $iconPath
        Status         = $launcher.Status
        LauncherSource = $launcher.LauncherSource
        Note           = $launcher.Note
        BindingFile    = ""
    }

    $bindingFile = Join-Path $binding.AppPath "out\kompas_button_binding.txt"
    if ($WritePerAppFiles) {
        $bindingFile = Write-PerAppBindingFile -Binding $binding
    }
    elseif (-not (Test-Path -LiteralPath $bindingFile)) {
        $bindingFile = ""
    }

    $binding.BindingFile = $bindingFile
    $bindings.Add($binding)
}

$readyBindings = @($bindings | Where-Object { $_.Status -eq "READY" })
$readyCount = $readyBindings.Count

$configUpdateStatus = "SKIPPED (-SkipConfiguratorUpdate)"
$configUpdatePath = ""
$configBackupPath = ""
$configAdded = 0
$configUpdated = 0
$configRemoved = 0

if (-not $SkipConfiguratorUpdate) {
    try {
        $resolvedKitConfig = Resolve-KompasKitConfig -ExplicitPath $KompasKitConfigPath
        if ([string]::IsNullOrWhiteSpace($resolvedKitConfig)) {
            $configUpdateStatus = "NOT_FOUND"
        }
        else {
            $runningKompas = @(Get-Process -Name KOMPAS -ErrorAction SilentlyContinue)
            if ($runningKompas.Count -gt 0) {
                if ($WaitForKompasExit) {
                    $closed = Wait-ForKompasProcessExit -TimeoutSec $WaitTimeoutSec
                    if (-not $closed) {
                        $configUpdateStatus = "SKIPPED_KOMPAS_RUNNING_TIMEOUT"
                    }
                }
                else {
                    $configUpdateStatus = "SKIPPED_KOMPAS_RUNNING"
                }
            }

            if ($configUpdateStatus -ne "SKIPPED_KOMPAS_RUNNING" -and $configUpdateStatus -ne "SKIPPED_KOMPAS_RUNNING_TIMEOUT") {
                $configResult = Update-KompasConfiguratorUtilities -Bindings $bindings -KitConfigPath $resolvedKitConfig
                $configUpdateStatus = $configResult.Status
                $configUpdatePath = $configResult.Path
                $configBackupPath = $configResult.BackupPath
                $configAdded = $configResult.Added
                $configUpdated = $configResult.Updated
                $configRemoved = $configResult.Removed
            }
        }
    }
    catch {
        $configUpdateStatus = "FAIL: $($_.Exception.Message)"
    }
}

$bulkTextPath = Join-Path $OutputDir "kompas_bulk_binding.txt"
$bulkCsvPath = Join-Path $OutputDir "kompas_bulk_binding.csv"
$bulkJsonPath = Join-Path $OutputDir "kompas_bulk_binding.json"

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("KOMPAS-3D bulk external-command binding")
$lines.Add("---------------------------------------")
$lines.Add("Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
$lines.Add("Repo root : $RepoRoot")
$lines.Add("Apps found: $($bindings.Count)")
$lines.Add("Apps ready: $readyCount")
$lines.Add("")

$i = 1
foreach ($item in $bindings) {
    $lines.Add("[$i] $($item.AppFolder)")
    $lines.Add("Name      : $($item.Name)")
    $lines.Add("Program   : $($item.Program)")
    $lines.Add("Arguments : $($item.Arguments)")
    $lines.Add("Icon      : $($item.Icon)")
    $lines.Add("Status    : $($item.Status)")
    $lines.Add("Source    : $($item.LauncherSource)")
    $lines.Add("App file  : $($item.BindingFile)")
    if (-not [string]::IsNullOrWhiteSpace($item.Note)) {
        $lines.Add("Note      : $($item.Note)")
    }
    $lines.Add("")
    $i++
}

$lines | Set-Content -LiteralPath $bulkTextPath -Encoding UTF8

$bindings |
    Select-Object AppFolder, Name, Program, Arguments, Icon, Status, LauncherSource, BindingFile, Note |
    Export-Csv -LiteralPath $bulkCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'

$bindings |
    Select-Object AppFolder, AppPath, Name, Program, Arguments, Icon, Status, LauncherSource, BindingFile, Note |
    ConvertTo-Json -Depth 4 |
    Set-Content -LiteralPath $bulkJsonPath -Encoding UTF8

$clipboardStatus = "SKIPPED (-NoClipboard)"
if (-not $NoClipboard) {
    if ($readyCount -eq 0) {
        $clipboardStatus = "SKIPPED (no READY apps)"
    }
    else {
        $clipboardText = ($readyBindings | ForEach-Object {
                "Name: $($_.Name)`r`nProgram: $($_.Program)`r`nArguments: $($_.Arguments)`r`nIcon: $($_.Icon)"
            }) -join "`r`n`r`n"
        try {
            Set-Clipboard -Value $clipboardText
            $clipboardStatus = "OK"
        }
        catch {
            $clipboardStatus = "FAIL: $($_.Exception.Message)"
        }
    }
}

Write-Host ""
Write-Host "=== BIND ALL APPS TO KOMPAS ===" -ForegroundColor Cyan
Write-Host "Repo root   : $RepoRoot"
Write-Host "Apps found  : $($bindings.Count)"
Write-Host "Apps ready  : $readyCount"
Write-Host "Bulk txt    : $bulkTextPath"
Write-Host "Bulk csv    : $bulkCsvPath"
Write-Host "Bulk json   : $bulkJsonPath"
Write-Host "Per-app out : $WritePerAppFiles"
Write-Host "Configurator: $configUpdateStatus"
if (-not [string]::IsNullOrWhiteSpace($configUpdatePath)) {
    Write-Host "Config path : $configUpdatePath"
}
if (-not [string]::IsNullOrWhiteSpace($configBackupPath)) {
    Write-Host "Config backup: $configBackupPath"
}
if ($configAdded -gt 0 -or $configUpdated -gt 0 -or $configRemoved -gt 0) {
    Write-Host "Config delta: +$configAdded / ~$configUpdated / -$configRemoved"
}
Write-Host "Clipboard   : $clipboardStatus"
Write-Host ""

foreach ($item in $bindings) {
    $statusColor = "Yellow"
    if ($item.Status -eq "READY") {
        $statusColor = "Green"
    }

    Write-Host "[$($item.Status)] $($item.AppFolder)" -ForegroundColor $statusColor
    Write-Host "  Program  : $($item.Program)"
    Write-Host "  Arguments: $($item.Arguments)"
    Write-Host "  Icon     : $($item.Icon)"
    Write-Host "  App file : $($item.BindingFile)"
    if (-not [string]::IsNullOrWhiteSpace($item.Note)) {
        Write-Host "  Note     : $($item.Note)" -ForegroundColor DarkYellow
    }
}
