[CmdletBinding()]
param()

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    $scriptPath = $MyInvocation.MyCommand.Path
    $argLine = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File ""$scriptPath"""
    Start-Process -FilePath "powershell.exe" -ArgumentList $argLine -WindowStyle Hidden | Out-Null
    exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

$script:AppRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$script:ScriptsDir = Join-Path $script:AppRoot "scripts"
$script:ConfigDir = Join-Path $script:AppRoot "config"
$script:OutDir = Join-Path $script:AppRoot "out"
$script:SettingsPath = Join-Path $script:ConfigDir "app_settings.json"
$script:SyncScriptPath = Join-Path $script:ScriptsDir "kompas_excel_text_sync.py"
$script:AppIconPath = Join-Path $script:AppRoot "assets\app.ico"
$script:CrashLogPath = Join-Path $script:OutDir "gui_sync_errors.log"
$script:SyncProcess = $null
$script:SyncWatchTimer = $null
$script:SyncLogRef = $null
$script:LastEngineLogPath = ""
$script:FormAppIcon = $null
$script:LastLayoutAuditSignature = ""

function Add-Log {
    param(
        [System.Windows.Forms.TextBox]$Box,
        [string]$Message
    )

    if ($null -eq $Box) {
        return
    }
    $ts = Get-Date -Format "HH:mm:ss"
    $Box.AppendText("[$ts] $Message`r`n")
}

function Ensure-AppIconFile {
    param(
        [string]$IconPath
    )

    if ([string]::IsNullOrWhiteSpace($IconPath)) {
        return
    }
    if (Test-Path -LiteralPath $IconPath -PathType Leaf) {
        return
    }

    $iconDir = Split-Path -Parent $IconPath
    if (-not [string]::IsNullOrWhiteSpace($iconDir) -and -not (Test-Path -LiteralPath $iconDir)) {
        New-Item -Path $iconDir -ItemType Directory -Force | Out-Null
    }

    $fs = $null
    try {
        $fs = [System.IO.File]::Create($IconPath)
        [System.Drawing.SystemIcons]::Shield.Save($fs)
    }
    catch {
    }
    finally {
        if ($null -ne $fs) {
            $fs.Dispose()
        }
    }
}

function Write-CrashLog {
    param([string]$Message)

    try {
        if (-not (Test-Path -LiteralPath $script:OutDir)) {
            New-Item -Path $script:OutDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        "[$stamp] $Message" | Add-Content -LiteralPath $script:CrashLogPath -Encoding UTF8
    }
    catch {
    }
}

function Get-DefaultSettings {
    return [ordered]@{
        corridor_mm = "1.5"
        poll_ms = "1200"
        sheet_name = "SyncData"
    }
}

function Read-Settings {
    if (-not (Test-Path -LiteralPath $script:SettingsPath)) {
        return Get-DefaultSettings
    }

    try {
        $raw = Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json
        $defaults = Get-DefaultSettings
        foreach ($k in $defaults.Keys) {
            if ($null -eq $raw.$k -or [string]::IsNullOrWhiteSpace([string]$raw.$k)) {
                $raw | Add-Member -NotePropertyName $k -NotePropertyValue $defaults[$k] -Force
            }
        }
        return $raw
    }
    catch {
        return Get-DefaultSettings
    }
}

function Save-Settings {
    param([System.Collections.IDictionary]$Values)

    if (-not (Test-Path -LiteralPath $script:ConfigDir)) {
        New-Item -Path $script:ConfigDir -ItemType Directory -Force | Out-Null
    }
    ($Values | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
}

function Apply-SettingsToUi {
    param(
        [object]$Settings,
        [hashtable]$Ctl
    )

    $defaults = Get-DefaultSettings
    $Ctl.corridor.Text = [string]($(if ($null -ne $Settings.corridor_mm) { $Settings.corridor_mm } else { $defaults.corridor_mm }))
    $Ctl.poll.Text = [string]($(if ($null -ne $Settings.poll_ms) { $Settings.poll_ms } else { $defaults.poll_ms }))
    $Ctl.sheet.Text = [string]($(if ($null -ne $Settings.sheet_name) { $Settings.sheet_name } else { $defaults.sheet_name }))
}

function Get-UiValues {
    param([hashtable]$Ctl)
    return [ordered]@{
        corridor_mm = $Ctl.corridor.Text.Trim()
        poll_ms = $Ctl.poll.Text.Trim()
        sheet_name = $Ctl.sheet.Text.Trim()
    }
}

function Validate-UiValues {
    param(
        [System.Collections.IDictionary]$Values,
        [System.Windows.Forms.TextBox]$LogBox
    )

    $corridor = 0.0
    if (-not [double]::TryParse(($Values.corridor_mm -replace ",", "."), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$corridor) -or $corridor -le 0) {
        Add-Log -Box $LogBox -Message "ERROR: corridor_mm должен быть положительным числом."
        return $false
    }

    $poll = 0
    if (-not [int]::TryParse([string]$Values.poll_ms, [ref]$poll) -or $poll -lt 250) {
        Add-Log -Box $LogBox -Message "ERROR: poll_ms должен быть целым числом >= 250."
        return $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$Values.sheet_name)) {
        Add-Log -Box $LogBox -Message "ERROR: sheet_name не должен быть пустым."
        return $false
    }

    return $true
}

function Get-PythonExe {
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($null -ne $python) {
        return $python.Source
    }

    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($null -ne $py) {
        return $py.Source
    }

    return $null
}

function Set-IconButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Icon]$Icon,
        [System.Windows.Forms.ToolTip]$ToolTip,
        [string]$TipText,
        [string]$FallbackText = "?"
    )

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 1
    $Button.UseVisualStyleBackColor = $true
    $Button.Text = $FallbackText
    $Button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $Button.Image = New-Object System.Drawing.Bitmap($Icon.ToBitmap(), (New-Object System.Drawing.Size(16, 16)))
    $Button.ImageAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $Button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $Button.TextImageRelation = [System.Windows.Forms.TextImageRelation]::Overlay
    if ($null -ne $ToolTip -and -not [string]::IsNullOrWhiteSpace($TipText)) {
        $ToolTip.SetToolTip($Button, $TipText)
    }
}

function Set-SyncStateUi {
    param(
        [System.Windows.Forms.Button]$ToggleButton,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.ToolTip]$ToolTip,
        [bool]$IsRunning
    )

    if ($null -eq $StatusLabel) {
        return
    }

    if ($IsRunning) {
        $StatusLabel.Text = "Статус: синхронизация включена"
        $StatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        if ($null -ne $ToolTip) {
            $ToolTip.SetToolTip($ToggleButton, "Остановить синхронизацию")
        }
        return
    }

    $StatusLabel.Text = "Статус: остановлено"
    $StatusLabel.ForeColor = [System.Drawing.Color]::DarkRed
    if ($null -ne $ToolTip) {
        $ToolTip.SetToolTip($ToggleButton, "Включить синхронизацию")
    }
}

function Stop-SyncProcess {
    param(
        [System.Windows.Forms.Button]$ToggleButton,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ToolTip]$ToolTip
    )

    if ($null -ne $script:SyncWatchTimer) {
        try {
            $script:SyncWatchTimer.Stop()
            $script:SyncWatchTimer.Dispose()
        }
        catch {
        }
        $script:SyncWatchTimer = $null
    }

    if ($null -eq $script:SyncProcess) {
        Set-SyncStateUi -ToggleButton $ToggleButton -StatusLabel $StatusLabel -ToolTip $ToolTip -IsRunning $false
        return
    }

    try {
        if (-not $script:SyncProcess.HasExited) {
            $script:SyncProcess.Kill()
            $script:SyncProcess.WaitForExit(1500) | Out-Null
        }
    }
    catch {
    }

    try {
        $script:SyncProcess.Dispose()
    }
    catch {
    }
    $script:SyncProcess = $null
    $script:SyncLogRef = $null
    Set-SyncStateUi -ToggleButton $ToggleButton -StatusLabel $StatusLabel -ToolTip $ToolTip -IsRunning $false
    if ($null -ne $LogBox) {
        Add-Log -Box $LogBox -Message "INFO: Синхронизация остановлена."
    }
}

function Start-SyncProcess {
    param(
        [hashtable]$Ctl,
        [System.Windows.Forms.Button]$ToggleButton,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.TextBox]$LogBox,
        [System.Windows.Forms.ToolTip]$ToolTip
    )

    try {
        if (-not (Test-Path -LiteralPath $script:SyncScriptPath)) {
            Add-Log -Box $LogBox -Message "ERROR: Python-скрипт не найден: $script:SyncScriptPath"
            return
        }

        $pythonExe = Get-PythonExe
        if ($null -eq $pythonExe) {
            Add-Log -Box $LogBox -Message "ERROR: Python не найден (python.exe или py.exe)."
            return
        }

        $values = Get-UiValues -Ctl $Ctl
        if (-not (Validate-UiValues -Values $values -LogBox $LogBox)) {
            return
        }
        Save-Settings -Values $values

        if (-not (Test-Path -LiteralPath $script:OutDir)) {
            New-Item -Path $script:OutDir -ItemType Directory -Force | Out-Null
        }

        $script:LastEngineLogPath = Join-Path $script:OutDir "sync_engine.log"
        $sheetName = [string]$values.sheet_name
        $arguments = @(
            '"' + $script:SyncScriptPath + '"'
            "--corridor-mm $($values.corridor_mm -replace ',', '.')"
            "--poll-ms $($values.poll_ms)"
            "--sheet-name `"$sheetName`""
            "--log-file `"$script:LastEngineLogPath`""
        ) -join " "

        if ([System.IO.Path]::GetFileName($pythonExe).ToLowerInvariant() -eq "py.exe") {
            $arguments = "-3 " + $arguments
        }

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $pythonExe
        $psi.Arguments = $arguments
        $psi.WorkingDirectory = $script:ScriptsDir
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardOutput = $false
        $psi.RedirectStandardError = $false

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        if (-not $process.Start()) {
            Add-Log -Box $LogBox -Message "ERROR: Не удалось запустить синхронизацию."
            return
        }

        if ($null -ne $script:SyncWatchTimer) {
            try {
                $script:SyncWatchTimer.Stop()
                $script:SyncWatchTimer.Dispose()
            }
            catch {
            }
            $script:SyncWatchTimer = $null
        }

        $script:SyncProcess = $process
        $script:SyncLogRef = $LogBox
        $watchToggleButton = $ToggleButton
        $watchStatusLabel = $StatusLabel
        $watchToolTip = $ToolTip

        $watchTimer = New-Object System.Windows.Forms.Timer
        $watchTimer.Interval = 1000
        $watchTimer.Add_Tick({
                try {
                    if ($null -eq $script:SyncProcess) {
                        return
                    }
                    if (-not $script:SyncProcess.HasExited) {
                        return
                    }

                    $exitCode = 0
                    try {
                        $exitCode = [int]$script:SyncProcess.ExitCode
                    }
                    catch {
                    }

                    if ($null -ne $script:SyncLogRef) {
                        Add-Log -Box $script:SyncLogRef -Message ("INFO: Процесс синхронизации завершён. Код: " + $exitCode)
                        if ($exitCode -eq 21) {
                            Add-Log -Box $script:SyncLogRef -Message "INFO: Excel закрыт, синхронизация отключена."
                        }
                        elseif ($exitCode -eq 22) {
                            Add-Log -Box $script:SyncLogRef -Message "ERROR: Excel-файл открыт только для чтения. Освободите файл и перезапустите синхронизацию."
                        }
                        if ($exitCode -ne 0) {
                            Add-Log -Box $script:SyncLogRef -Message ("ERROR: Проверьте лог движка: " + $script:LastEngineLogPath)
                        }
                    }

                    try {
                        $script:SyncProcess.Dispose()
                    }
                    catch {
                    }
                    $script:SyncProcess = $null
                    $script:SyncLogRef = $null

                    if ($null -ne $script:SyncWatchTimer) {
                        try {
                            $script:SyncWatchTimer.Stop()
                            $script:SyncWatchTimer.Dispose()
                        }
                        catch {
                        }
                        $script:SyncWatchTimer = $null
                    }

                    Set-SyncStateUi -ToggleButton $watchToggleButton -StatusLabel $watchStatusLabel -ToolTip $watchToolTip -IsRunning $false
                }
                catch {
                    Write-CrashLog -Message ("Timer watch error: " + $_.Exception.Message)
                }
            })
        $watchTimer.Start()
        $script:SyncWatchTimer = $watchTimer

        Set-SyncStateUi -ToggleButton $ToggleButton -StatusLabel $StatusLabel -ToolTip $ToolTip -IsRunning $true
        Add-Log -Box $LogBox -Message "INFO: Синхронизация запущена."
        Add-Log -Box $LogBox -Message ("INFO: Лог движка: " + $script:LastEngineLogPath)
        Add-Log -Box $LogBox -Message "INFO: JSON статуса создаётся автоматически рядом с .xlsx и автосохраняется."
    }
    catch {
        Add-Log -Box $LogBox -Message ("ERROR: Не удалось включить синхронизацию: " + $_.Exception.Message)
        Write-CrashLog -Message ("Start sync failed: " + $_.Exception.Message)
    }
}

function Get-ControlDebugName {
    param(
        [System.Windows.Forms.Control]$Control
    )

    if ($null -eq $Control) {
        return "<null>"
    }

    $name = [string]$Control.Name
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = $Control.GetType().Name
    }

    $text = [string]$Control.Text
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        if ($text.Length -gt 24) {
            $text = $text.Substring(0, 24) + "..."
        }
        return "$name('$text')"
    }

    return $name
}

function Get-LayoutOverflowIssues {
    param(
        [System.Windows.Forms.Control]$Parent,
        [string]$ParentPath = "Form"
    )

    $issues = New-Object System.Collections.ArrayList
    if ($null -eq $Parent) {
        return $issues.ToArray()
    }

    $parentWidth = [Math]::Max(0, $Parent.ClientSize.Width)
    $parentHeight = [Math]::Max(0, $Parent.ClientSize.Height)

    foreach ($child in @($Parent.Controls)) {
        if ($null -eq $child) {
            continue
        }
        if (-not $child.Visible) {
            continue
        }

        $childPath = "$ParentPath->$(Get-ControlDebugName -Control $child)"
        if ($child.Left -lt 0 -or $child.Top -lt 0 -or $child.Right -gt $parentWidth -or $child.Bottom -gt $parentHeight) {
            [void]$issues.Add(
                "$childPath bounds=[$($child.Left),$($child.Top),$($child.Right),$($child.Bottom)] parent=[$parentWidth,$parentHeight]"
            )
        }

        if ($child.Controls.Count -gt 0) {
            $nested = Get-LayoutOverflowIssues -Parent $child -ParentPath $childPath
            foreach ($entry in $nested) {
                [void]$issues.Add($entry)
            }
        }
    }

    return $issues.ToArray()
}

function Write-LayoutAuditLog {
    param(
        [System.Windows.Forms.TextBox]$LogBox,
        [switch]$Force
    )

    if ($null -eq $LogBox) {
        return
    }

    $issues = @(Get-LayoutOverflowIssues -Parent $form -ParentPath "Form")
    $issues = @(
        $issues |
            ForEach-Object { [string]$_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
    $signature = if ($issues.Count -eq 0) { "ok" } else { $issues -join "|" }
    if (-not $Force -and $signature -eq $script:LastLayoutAuditSignature) {
        return
    }

    $script:LastLayoutAuditSignature = $signature
    if ($issues.Count -eq 0) {
        Add-Log -Box $LogBox -Message "INFO: Проверка layout: все элементы в пределах окна."
        return
    }

    Add-Log -Box $LogBox -Message "WARN: Проверка layout: найдено выходов за границы: $($issues.Count)"
    foreach ($issue in $issues) {
        Add-Log -Box $LogBox -Message "WARN: $issue"
    }
}

function Update-CompactUiLayout {
    if ($null -eq $tabHome -or $null -eq $tabParams -or $null -eq $tabLog) {
        return
    }

    $homeMargin = 10
    $btnWidth = 34
    $btnGap = 4
    $homeWidth = [Math]::Max(220, $tabHome.ClientSize.Width - ($homeMargin * 2))
    $btnX = [Math]::Max($homeMargin + 220, $tabHome.ClientSize.Width - $homeMargin - $btnWidth)

    $lblStatus.Location = New-Object System.Drawing.Point($homeMargin, 10)
    $lblStatus.Size = New-Object System.Drawing.Size($homeWidth, 18)

    $lblExcelInfo.Location = New-Object System.Drawing.Point($homeMargin, 32)
    $lblExcelInfo.Size = New-Object System.Drawing.Size($homeWidth, 16)

    $lblJsonInfo.Location = New-Object System.Drawing.Point($homeMargin, 50)
    $lblJsonInfo.Size = New-Object System.Drawing.Size($homeWidth, 16)

    $lblHint.Location = New-Object System.Drawing.Point($homeMargin, 68)
    $lblHint.Size = New-Object System.Drawing.Size($homeWidth, 30)

    $actionsY = 102
    $lblActions.Location = New-Object System.Drawing.Point($homeMargin, ($actionsY + 8))

    $iconButtons = @($btnToggle, $btnSave, $btnOpenOut, $btnReset)
    $iconX = [Math]::Max($homeMargin, $lblActions.Right + 6)
    foreach ($actionBtn in $iconButtons) {
        $actionBtn.Location = New-Object System.Drawing.Point($iconX, $actionsY)
        $iconX += ($actionBtn.Width + $btnGap)
    }

    $leftActionsRight = $iconX - $btnGap
    $exitY = $actionsY
    if (($btnX - 8) -lt $leftActionsRight) {
        $exitY += 34
    }
    $btnExit.Location = New-Object System.Drawing.Point($btnX, $exitY)

    $grpMargin = 8
    $grpWidth = [Math]::Max(420, $tabParams.ClientSize.Width - ($grpMargin * 2))
    $grpHeight = [Math]::Max(240, $tabParams.ClientSize.Height - ($grpMargin * 2))
    $grp.Location = New-Object System.Drawing.Point($grpMargin, $grpMargin)
    $grp.Size = New-Object System.Drawing.Size($grpWidth, $grpHeight)

    $corrLabelX = 10
    $corrInputX = 140
    $pollLabelX = 240
    $pollInputX = 380
    $pollFits = (($pollInputX + 90 + 10) -le $grp.ClientSize.Width)

    $lblCorridor.Location = New-Object System.Drawing.Point($corrLabelX, 28)
    $txtCorridor.Location = New-Object System.Drawing.Point($corrInputX, 24)
    $txtCorridor.Size = New-Object System.Drawing.Size(84, 24)

    if ($pollFits) {
        $lblPoll.Location = New-Object System.Drawing.Point($pollLabelX, 28)
        $txtPoll.Location = New-Object System.Drawing.Point($pollInputX, 24)
        $txtPoll.Size = New-Object System.Drawing.Size(86, 24)
        $sheetRowY = 62
    }
    else {
        $lblPoll.Location = New-Object System.Drawing.Point($corrLabelX, 58)
        $txtPoll.Location = New-Object System.Drawing.Point($corrInputX, 54)
        $txtPoll.Size = New-Object System.Drawing.Size(84, 24)
        $sheetRowY = 90
    }

    $lblSheet.Location = New-Object System.Drawing.Point($corrLabelX, $sheetRowY)
    $sheetInputX = 140
    $sheetWidth = [Math]::Max(180, $grp.ClientSize.Width - $sheetInputX - 12)
    $txtSheet.Location = New-Object System.Drawing.Point($sheetInputX, ($sheetRowY - 4))
    $txtSheet.Size = New-Object System.Drawing.Size($sheetWidth, 24)

    $metaTop = $sheetRowY + 32
    $metaWidth = [Math]::Max(220, $grp.ClientSize.Width - 20)
    $lblStateMode.Location = New-Object System.Drawing.Point(10, $metaTop)
    $lblStateMode.Size = New-Object System.Drawing.Size($metaWidth, 32)
    $lblProfile.Location = New-Object System.Drawing.Point(10, ($metaTop + 34))
    $lblProfile.Size = New-Object System.Drawing.Size($metaWidth, 16)

    $txtLog.Location = New-Object System.Drawing.Point(8, 24)
    $txtLog.Size = New-Object System.Drawing.Size(
        [Math]::Max(220, $tabLog.ClientSize.Width - 16),
        [Math]::Max(140, $tabLog.ClientSize.Height - 32)
    )
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Синхронизация Excel <-> КОМПАС (тексты)"
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::None
$form.ClientSize = New-Object System.Drawing.Size(560, 430)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
$form.MaximizeBox = $false
$form.MinimizeBox = $true
$form.Padding = New-Object System.Windows.Forms.Padding(8)
$form.ShowIcon = $true
$form.ShowInTaskbar = $true
$form.TopMost = $false

Ensure-AppIconFile -IconPath $script:AppIconPath
try {
    if (Test-Path -LiteralPath $script:AppIconPath -PathType Leaf) {
        $script:FormAppIcon = New-Object System.Drawing.Icon($script:AppIconPath)
        $form.Icon = $script:FormAppIcon
    }
    else {
        $form.Icon = [System.Drawing.SystemIcons]::Shield
    }
}
catch {
}

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 6000
$toolTip.InitialDelay = 250
$toolTip.ReshowDelay = 150
$toolTip.ShowAlways = $true

$tabMain = New-Object System.Windows.Forms.TabControl
$tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($tabMain)

$tabHome = New-Object System.Windows.Forms.TabPage
$tabHome.Text = "Главное"
$tabHome.AutoScroll = $true
$tabMain.Controls.Add($tabHome)

$tabParams = New-Object System.Windows.Forms.TabPage
$tabParams.Text = "Параметры"
$tabParams.AutoScroll = $true
$tabMain.Controls.Add($tabParams)

$tabLog = New-Object System.Windows.Forms.TabPage
$tabLog.Text = "Журнал"
$tabLog.AutoScroll = $true
$tabMain.Controls.Add($tabLog)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(10, 10)
$lblStatus.AutoSize = $false
$lblStatus.Size = New-Object System.Drawing.Size(500, 18)
$lblStatus.AutoEllipsis = $true
$tabHome.Controls.Add($lblStatus)

$lblExcelInfo = New-Object System.Windows.Forms.Label
$lblExcelInfo.Text = "Excel: рядом с активным чертежом создаётся/открывается <ИмяЧертежа>.xlsx"
$lblExcelInfo.Location = New-Object System.Drawing.Point(10, 32)
$lblExcelInfo.AutoSize = $false
$lblExcelInfo.Size = New-Object System.Drawing.Size(500, 16)
$lblExcelInfo.AutoEllipsis = $true
$tabHome.Controls.Add($lblExcelInfo)

$lblJsonInfo = New-Object System.Windows.Forms.Label
$lblJsonInfo.Text = "JSON статуса: создаётся автоматически рядом с таблицей как <ИмяЧертежа>.json"
$lblJsonInfo.Location = New-Object System.Drawing.Point(10, 50)
$lblJsonInfo.AutoSize = $false
$lblJsonInfo.Size = New-Object System.Drawing.Size(500, 16)
$lblJsonInfo.AutoEllipsis = $true
$tabHome.Controls.Add($lblJsonInfo)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "При смене активной вкладки КОМПАС синхронизация автоматически переключается на соответствующие .xlsx/.json."
$lblHint.Location = New-Object System.Drawing.Point(10, 68)
$lblHint.AutoSize = $false
$lblHint.Size = New-Object System.Drawing.Size(500, 30)
$lblHint.AutoEllipsis = $true
$tabHome.Controls.Add($lblHint)

$lblActions = New-Object System.Windows.Forms.Label
$lblActions.Text = "Действия:"
$lblActions.Location = New-Object System.Drawing.Point(10, 110)
$lblActions.AutoSize = $true
$tabHome.Controls.Add($lblActions)

$btnToggle = New-Object System.Windows.Forms.Button
$btnToggle.Location = New-Object System.Drawing.Point(78, 102)
$btnToggle.Size = New-Object System.Drawing.Size(34, 30)
$tabHome.Controls.Add($btnToggle)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Location = New-Object System.Drawing.Point(116, 102)
$btnSave.Size = New-Object System.Drawing.Size(34, 30)
$tabHome.Controls.Add($btnSave)

$btnOpenOut = New-Object System.Windows.Forms.Button
$btnOpenOut.Location = New-Object System.Drawing.Point(154, 102)
$btnOpenOut.Size = New-Object System.Drawing.Size(34, 30)
$tabHome.Controls.Add($btnOpenOut)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Location = New-Object System.Drawing.Point(192, 102)
$btnReset.Size = New-Object System.Drawing.Size(34, 30)
$tabHome.Controls.Add($btnReset)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = New-Object System.Drawing.Point(462, 102)
$btnExit.Size = New-Object System.Drawing.Size(34, 30)
$btnExit.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$tabHome.Controls.Add($btnExit)

Set-IconButton -Button $btnToggle -Icon ([System.Drawing.SystemIcons]::Application) -ToolTip $toolTip -TipText "Включить синхронизацию" -FallbackText "SYNC"
Set-IconButton -Button $btnSave -Icon ([System.Drawing.SystemIcons]::Shield) -ToolTip $toolTip -TipText "Сохранить параметры" -FallbackText "S"
Set-IconButton -Button $btnOpenOut -Icon ([System.Drawing.SystemIcons]::Information) -ToolTip $toolTip -TipText "Открыть папку out" -FallbackText "O"
Set-IconButton -Button $btnReset -Icon ([System.Drawing.SystemIcons]::Warning) -ToolTip $toolTip -TipText "Сбросить параметры по умолчанию" -FallbackText "R"
Set-IconButton -Button $btnExit -Icon ([System.Drawing.SystemIcons]::Error) -ToolTip $toolTip -TipText "Закрыть окно" -FallbackText "X"

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Параметры синхронизации"
$grp.Location = New-Object System.Drawing.Point(8, 8)
$grp.Size = New-Object System.Drawing.Size(500, 264)
$grp.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabParams.Controls.Add($grp)

$lblCorridor = New-Object System.Windows.Forms.Label
$lblCorridor.Text = "Коридор уровня, мм:"
$lblCorridor.Location = New-Object System.Drawing.Point(10, 28)
$lblCorridor.AutoSize = $true
$grp.Controls.Add($lblCorridor)

$txtCorridor = New-Object System.Windows.Forms.TextBox
$txtCorridor.Location = New-Object System.Drawing.Point(140, 24)
$txtCorridor.Size = New-Object System.Drawing.Size(84, 24)
$grp.Controls.Add($txtCorridor)

$lblPoll = New-Object System.Windows.Forms.Label
$lblPoll.Text = "Период опроса, мс:"
$lblPoll.Location = New-Object System.Drawing.Point(240, 28)
$lblPoll.AutoSize = $true
$grp.Controls.Add($lblPoll)

$txtPoll = New-Object System.Windows.Forms.TextBox
$txtPoll.Location = New-Object System.Drawing.Point(380, 24)
$txtPoll.Size = New-Object System.Drawing.Size(86, 24)
$grp.Controls.Add($txtPoll)

$lblSheet = New-Object System.Windows.Forms.Label
$lblSheet.Text = "Лист Excel:"
$lblSheet.Location = New-Object System.Drawing.Point(10, 62)
$lblSheet.AutoSize = $true
$grp.Controls.Add($lblSheet)

$txtSheet = New-Object System.Windows.Forms.TextBox
$txtSheet.Location = New-Object System.Drawing.Point(140, 58)
$txtSheet.Size = New-Object System.Drawing.Size(326, 24)
$grp.Controls.Add($txtSheet)

$lblStateMode = New-Object System.Windows.Forms.Label
$lblStateMode.Text = "JSON состояния задаётся автоматически: путь = путь Excel, имя = имя Excel, расширение .json."
$lblStateMode.Location = New-Object System.Drawing.Point(10, 94)
$lblStateMode.AutoSize = $false
$lblStateMode.Size = New-Object System.Drawing.Size(480, 32)
$lblStateMode.AutoEllipsis = $true
$grp.Controls.Add($lblStateMode)

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Text = "Профиль настроек: $script:SettingsPath"
$lblProfile.Location = New-Object System.Drawing.Point(10, 128)
$lblProfile.Size = New-Object System.Drawing.Size(480, 16)
$lblProfile.AutoEllipsis = $true
$grp.Controls.Add($lblProfile)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Журнал:"
$lblLog.Location = New-Object System.Drawing.Point(8, 8)
$lblLog.AutoSize = $true
$tabLog.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(8, 24)
$txtLog.Size = New-Object System.Drawing.Size(500, 248)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$txtLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabLog.Controls.Add($txtLog)

$controls = @{
    corridor = $txtCorridor
    poll = $txtPoll
    sheet = $txtSheet
    log = $txtLog
}

try {
    [System.Windows.Forms.Application]::SetUnhandledExceptionMode([System.Windows.Forms.UnhandledExceptionMode]::CatchException)
}
catch {
}
[System.Windows.Forms.Application]::add_ThreadException({
        param($sender, $eventArgs)
        $message = "UI thread exception: " + $eventArgs.Exception.Message
        Write-CrashLog -Message $message
        Add-Log -Box $txtLog -Message ("ERROR: " + $message)
    })
[System.AppDomain]::CurrentDomain.add_UnhandledException({
        param($sender, $eventArgs)
        $obj = $eventArgs.ExceptionObject
        $message = "Unhandled exception: " + [string]$obj
        Write-CrashLog -Message $message
    })

$settings = Read-Settings
Apply-SettingsToUi -Settings $settings -Ctl $controls
Set-SyncStateUi -ToggleButton $btnToggle -StatusLabel $lblStatus -ToolTip $toolTip -IsRunning $false

Update-CompactUiLayout
$tabMain.Add_SelectedIndexChanged({
        Update-CompactUiLayout
        Write-LayoutAuditLog -LogBox $txtLog
    })
$tabHome.Add_Resize({
        Update-CompactUiLayout
        Write-LayoutAuditLog -LogBox $txtLog
    })
$tabParams.Add_Resize({
        Update-CompactUiLayout
        Write-LayoutAuditLog -LogBox $txtLog
    })
$tabLog.Add_Resize({
        Update-CompactUiLayout
        Write-LayoutAuditLog -LogBox $txtLog
    })
$grp.Add_Resize({
        Update-CompactUiLayout
        Write-LayoutAuditLog -LogBox $txtLog
    })
$form.Add_ClientSizeChanged({
        Update-CompactUiLayout
        Write-LayoutAuditLog -LogBox $txtLog
    })

$btnSave.Add_Click({
        $values = Get-UiValues -Ctl $controls
        if (Validate-UiValues -Values $values -LogBox $txtLog) {
            Save-Settings -Values $values
            Add-Log -Box $txtLog -Message "INFO: Настройки сохранены."
        }
    })

$btnReset.Add_Click({
        $defaults = Get-DefaultSettings
        Apply-SettingsToUi -Settings $defaults -Ctl $controls
        Save-Settings -Values (Get-UiValues -Ctl $controls)
        Add-Log -Box $txtLog -Message "INFO: Параметры сброшены по умолчанию."
    })

$btnOpenOut.Add_Click({
        if (-not (Test-Path -LiteralPath $script:OutDir)) {
            New-Item -Path $script:OutDir -ItemType Directory -Force | Out-Null
        }
        Start-Process explorer.exe $script:OutDir
    })

$btnToggle.Add_Click({
        try {
            if ($null -eq $script:SyncProcess) {
                Start-SyncProcess -Ctl $controls -ToggleButton $btnToggle -StatusLabel $lblStatus -LogBox $txtLog -ToolTip $toolTip
            }
            else {
                Stop-SyncProcess -ToggleButton $btnToggle -StatusLabel $lblStatus -LogBox $txtLog -ToolTip $toolTip
            }
        }
        catch {
            $message = $_.Exception.Message
            Add-Log -Box $txtLog -Message ("ERROR: toggle failed: " + $message)
            Write-CrashLog -Message ("Toggle failed: " + $message)
        }
    })

$btnExit.Add_Click({ $form.Close() })

$form.Add_Shown({
        Update-CompactUiLayout
        Write-LayoutAuditLog -LogBox $txtLog -Force
    })

$form.Add_FormClosing({
        Stop-SyncProcess -ToggleButton $btnToggle -StatusLabel $lblStatus -LogBox $null -ToolTip $toolTip
        $values = Get-UiValues -Ctl $controls
        if (Validate-UiValues -Values $values -LogBox $txtLog) {
            Save-Settings -Values $values
        }
        if ($null -ne $script:FormAppIcon) {
            try {
                $script:FormAppIcon.Dispose()
            }
            catch {
            }
            $script:FormAppIcon = $null
        }
    })

Add-Log -Box $txtLog -Message "Готово. Нажмите кнопку включения синхронизации."
Add-Log -Box $txtLog -Message "Excel таблица: <ИмяЧертежа>.xlsx рядом с активным файлом КОМПАС."
Add-Log -Box $txtLog -Message "JSON статус: <ИмяЧертежа>.json рядом с этой Excel-таблицей."
Add-Log -Box $txtLog -Message "Назначение пути JSON вручную отключено, автосохранение включено."

[System.Windows.Forms.Application]::Run($form)
