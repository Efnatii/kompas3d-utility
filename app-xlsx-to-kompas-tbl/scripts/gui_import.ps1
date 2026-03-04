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

if (-not ("Win32DragDrop" -as [type])) {
    $winFormsAssembly = [System.Windows.Forms.Form].Assembly.Location
    Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class Win32DragDrop {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool ChangeWindowMessageFilterEx(
        IntPtr hWnd,
        uint msg,
        uint action,
        IntPtr pChangeFilterStruct
    );

    [DllImport("shell32.dll")]
    public static extern void DragAcceptFiles(IntPtr hWnd, bool fAccept);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern uint DragQueryFile(IntPtr hDrop, uint iFile, StringBuilder lpszFile, int cch);

    [DllImport("shell32.dll")]
    public static extern void DragFinish(IntPtr hDrop);

    public static string[] ExtractDroppedFiles(IntPtr hDrop) {
        uint count = DragQueryFile(hDrop, 0xFFFFFFFF, null, 0);
        var files = new List<string>();
        for (uint i = 0; i < count; i++) {
            uint len = DragQueryFile(hDrop, i, null, 0);
            var sb = new StringBuilder((int)len + 1);
            DragQueryFile(hDrop, i, sb, sb.Capacity);
            files.Add(sb.ToString());
        }
        DragFinish(hDrop);
        return files.ToArray();
    }
}

public sealed class DropFilesWatcher : NativeWindow, IDisposable {
    private const int WM_DROPFILES = 0x0233;

    public event Action<string[]> FilesDropped;

    public DropFilesWatcher(IntPtr handle) {
        AssignHandle(handle);
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_DROPFILES) {
            var files = Win32DragDrop.ExtractDroppedFiles(m.WParam);
            if (files != null && files.Length > 0) {
                var handler = FilesDropped;
                if (handler != null) {
                    handler(files);
                }
            }
        }
        base.WndProc(ref m);
    }

    public void Dispose() {
        ReleaseHandle();
    }
}
"@ -ReferencedAssemblies @($winFormsAssembly)
}

$script:AppRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$script:RepoRoot = Split-Path -Parent $script:AppRoot
$script:ExporterVbs = Join-Path $script:RepoRoot "xlsx-to-kompas-tbl\src\create_tbl.vbs"
$script:InsertBridgePy = Join-Path $script:RepoRoot "xlsx-to-kompas-tbl\src\insert_tbl_bridge.py"
$script:DefaultInput = Join-Path $script:RepoRoot "xlsx-to-kompas-tbl\fixtures\table_M2.xlsx"
$script:LayoutConfigPath = Join-Path $script:RepoRoot "xlsx-to-kompas-tbl\config\table_layout.ini"
$script:SettingsPath = Join-Path $script:AppRoot "config\app_settings.json"
$script:AppIconPath = Join-Path $script:AppRoot "assets\app.ico"
$script:LastKompasDocResolveReason = ""

function Add-Log {
    param(
        [System.Windows.Forms.TextBox]$Box,
        [string]$Message
    )

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

function Format-Inv([double]$Value) {
    return $Value.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Test-IsElevated {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Convert-ToNonEmptyString {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        $asText = [string]$Value
        if ([string]::IsNullOrWhiteSpace($asText)) {
            return $null
        }
        return $asText.Trim()
    }
    catch {
        return $null
    }
}

function Get-ComObjectProperty {
    param(
        [object]$ComObject,
        [string[]]$PropertyNames
    )

    if ($null -eq $ComObject) {
        return $null
    }

    foreach ($propertyName in $PropertyNames) {
        if ([string]::IsNullOrWhiteSpace($propertyName)) {
            continue
        }

        try {
            $value = $ComObject.$propertyName
            if ($null -ne $value) {
                return $value
            }
        }
        catch {
        }

        try {
            $value = $ComObject.GetType().InvokeMember(
                $propertyName,
                [System.Reflection.BindingFlags]::GetProperty,
                $null,
                $ComObject,
                $null
            )
            if ($null -ne $value) {
                return $value
            }
        }
        catch {
        }
    }

    return $null
}

function Invoke-ComObjectMethod {
    param(
        [object]$ComObject,
        [string[]]$MethodNames,
        [object[]]$Arguments = @()
    )

    if ($null -eq $ComObject) {
        return $null
    }

    foreach ($methodName in $MethodNames) {
        if ([string]::IsNullOrWhiteSpace($methodName)) {
            continue
        }

        if ($Arguments.Count -eq 0) {
            try {
                $value = $ComObject.$methodName()
                if ($null -ne $value) {
                    return $value
                }
            }
            catch {
            }

            # Some COM members are exposed as property-get but semantically are methods.
            try {
                $value = $ComObject.$methodName
                if ($null -ne $value) {
                    return $value
                }
            }
            catch {
            }
        }

        try {
            $value = $ComObject.GetType().InvokeMember(
                $methodName,
                [System.Reflection.BindingFlags]::InvokeMethod,
                $null,
                $ComObject,
                $Arguments
            )
            if ($null -ne $value) {
                return $value
            }
        }
        catch {
        }
    }

    return $null
}

function Get-ComStringProperty {
    param(
        [object]$ComObject,
        [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        if ([string]::IsNullOrWhiteSpace($propertyName)) {
            continue
        }

        $value = $null
        try {
            $value = $ComObject.$propertyName
        }
        catch {
            $value = $null
        }

        if ($null -eq $value) {
            try {
                $value = $ComObject.GetType().InvokeMember(
                    $propertyName,
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null,
                    $ComObject,
                    $null
                )
            }
            catch {
                $value = $null
            }
        }

        $asText = Convert-ToNonEmptyString -Value $value
        if (-not [string]::IsNullOrWhiteSpace($asText)) {
            return $asText
        }
    }

    return $null
}

function Invoke-ComStringMethod {
    param(
        [object]$ComObject,
        [string[]]$MethodNames
    )

    foreach ($methodName in $MethodNames) {
        if ([string]::IsNullOrWhiteSpace($methodName)) {
            continue
        }

        $value = $null
        try {
            $value = $ComObject.GetType().InvokeMember(
                $methodName,
                [System.Reflection.BindingFlags]::InvokeMethod,
                $null,
                $ComObject,
                $null
            )
        }
        catch {
            $value = $null
        }

        $asText = Convert-ToNonEmptyString -Value $value
        if (-not [string]::IsNullOrWhiteSpace($asText)) {
            return $asText
        }
    }

    return $null
}

function Get-ComPositiveIntValue {
    param(
        [object]$ComObject,
        [string[]]$MemberNames
    )

    foreach ($memberName in $MemberNames) {
        if ([string]::IsNullOrWhiteSpace($memberName)) {
            continue
        }

        $value = Get-ComObjectProperty -ComObject $ComObject -PropertyNames @($memberName)
        if ($null -eq $value) {
            $value = Invoke-ComObjectMethod -ComObject $ComObject -MethodNames @($memberName)
        }

        if ($null -eq $value) {
            continue
        }

        try {
            $asInt = [int]$value
            if ($asInt -gt 0) {
                return $asInt
            }
        }
        catch {
        }
    }

    return $null
}

function Get-ComCollectionItems {
    param(
        [object]$Collection
    )

    if ($null -eq $Collection) {
        return @()
    }

    $items = New-Object System.Collections.ArrayList

    try {
        foreach ($item in $Collection) {
            if ($null -ne $item) {
                [void]$items.Add($item)
            }
        }
    }
    catch {
    }

    if ($items.Count -gt 0) {
        return ,$items.ToArray()
    }

    $count = Get-ComPositiveIntValue -ComObject $Collection -MemberNames @("Count", "Length", "DocumentCount")
    if ($null -eq $count) {
        return @()
    }

    for ($idx = 0; $idx -lt $count; $idx++) {
        $item = Invoke-ComObjectMethod -ComObject $Collection -MethodNames @("Item", "GetItem", "At", "Document") -Arguments @($idx)
        if ($null -eq $item) {
            $item = Invoke-ComObjectMethod -ComObject $Collection -MethodNames @("Item", "GetItem", "At", "Document") -Arguments @($idx + 1)
        }
        if ($null -ne $item) {
            [void]$items.Add($item)
        }
    }

    return ,$items.ToArray()
}

function Add-DocumentCandidate {
    param(
        [System.Collections.ArrayList]$Candidates,
        [object]$Document,
        [string]$Source
    )

    if ($null -eq $Candidates -or $null -eq $Document) {
        return
    }

    [void]$Candidates.Add([pscustomobject]@{
        Document = $Document
        Source = $Source
    })
}

function Add-DocumentsFromContainer {
    param(
        [System.Collections.ArrayList]$Candidates,
        [object]$Container,
        [string]$SourcePrefix
    )

    if ($null -eq $Container) {
        return
    }

    $activeDoc = Get-ComObjectProperty -ComObject $Container -PropertyNames @(
        "ActiveDocument",
        "CurrentDocument"
    )
    if ($null -eq $activeDoc) {
        $activeDoc = Invoke-ComObjectMethod -ComObject $Container -MethodNames @(
            "GetActiveDocument",
            "ActiveDocument"
        )
    }
    Add-DocumentCandidate -Candidates $Candidates -Document $activeDoc -Source "$SourcePrefix.ActiveDocument"

    $collectionCandidates = New-Object System.Collections.ArrayList
    [void]$collectionCandidates.Add([pscustomobject]@{Obj = $Container; Name = "$SourcePrefix.Collection" })

    $nestedCollection = Get-ComObjectProperty -ComObject $Container -PropertyNames @(
        "Documents",
        "OpenedDocuments",
        "Items",
        "List"
    )
    if ($null -ne $nestedCollection) {
        [void]$collectionCandidates.Add([pscustomobject]@{Obj = $nestedCollection; Name = "$SourcePrefix.Documents" })
    }

    $docIndex = 0
    foreach ($bucket in $collectionCandidates) {
        $items = Get-ComCollectionItems -Collection $bucket.Obj
        foreach ($item in $items) {
            Add-DocumentCandidate -Candidates $Candidates -Document $item -Source ($bucket.Name + "[" + $docIndex + "]")
            $docIndex++
        }
    }
}

function Resolve-DirectoryFromCandidatePath {
    param(
        [string]$CandidatePath
    )

    $path = Convert-ToNonEmptyString -Value $CandidatePath
    if ([string]::IsNullOrWhiteSpace($path)) {
        return $null
    }

    $path = $path.Trim('"').Trim()
    $path = $path.TrimEnd("*")
    $path = $path -replace "/", "\"
    if ([string]::IsNullOrWhiteSpace($path)) {
        return $null
    }

    try {
        if (-not [System.IO.Path]::IsPathRooted($path)) {
            return $null
        }

        if (Test-Path -LiteralPath $path -PathType Container) {
            return (Get-Item -LiteralPath $path).FullName
        }

        $parent = Split-Path -Parent $path
        if ([string]::IsNullOrWhiteSpace($parent)) {
            return $null
        }

        if (Test-Path -LiteralPath $parent -PathType Container) {
            return (Get-Item -LiteralPath $parent).FullName
        }
    }
    catch {
    }

    return $null
}

function Resolve-KompasDocumentDirectory {
    param(
        [object]$Document,
        [string]$Source,
        [ref]$Diagnostic
    )

    $Diagnostic.Value = ""

    if ($null -eq $Document) {
        $Diagnostic.Value = "${Source}: document is null."
        return $null
    }

    $candidatePaths = New-Object System.Collections.ArrayList

    $primaryPath = Get-ComStringProperty -ComObject $Document -PropertyNames @(
        "PathName",
        "DocumentPath",
        "FullFileName",
        "FullName",
        "FilePath",
        "Path",
        "PathAndName",
        "DocumentFileName",
        "FileName",
        "Name",
        "Title"
    )
    if (-not [string]::IsNullOrWhiteSpace($primaryPath)) {
        [void]$candidatePaths.Add($primaryPath)
    }

    $methodPath = Invoke-ComStringMethod -ComObject $Document -MethodNames @(
        "GetPathName",
        "GetDocumentPath",
        "GetFullName",
        "GetFileName",
        "GetName",
        "GetPath"
    )
    if (-not [string]::IsNullOrWhiteSpace($methodPath)) {
        [void]$candidatePaths.Add($methodPath)
    }

    $folderPart = Get-ComStringProperty -ComObject $Document -PropertyNames @(
        "Directory",
        "Folder",
        "FolderPath",
        "Path"
    )
    $filePart = Get-ComStringProperty -ComObject $Document -PropertyNames @(
        "FileName",
        "Name",
        "Title"
    )
    if (-not [string]::IsNullOrWhiteSpace($folderPart) -and -not [string]::IsNullOrWhiteSpace($filePart)) {
        try {
            [void]$candidatePaths.Add((Join-Path $folderPart $filePart))
        }
        catch {
        }
    }

    $nestedDoc = Get-ComObjectProperty -ComObject $Document -PropertyNames @("Document", "Doc", "Reference")
    if ($null -ne $nestedDoc -and -not [object]::ReferenceEquals($nestedDoc, $Document)) {
        $nestedPath = Get-ComStringProperty -ComObject $nestedDoc -PropertyNames @(
            "PathName",
            "DocumentPath",
            "FullFileName",
            "FullName",
            "FilePath",
            "Path"
        )
        if (-not [string]::IsNullOrWhiteSpace($nestedPath)) {
            [void]$candidatePaths.Add($nestedPath)
        }
    }

    foreach ($candidatePath in $candidatePaths) {
        $resolvedDir = Resolve-DirectoryFromCandidatePath -CandidatePath $candidatePath
        if (-not [string]::IsNullOrWhiteSpace($resolvedDir)) {
            $Diagnostic.Value = "${Source}: resolved from '$candidatePath'."
            return $resolvedDir
        }
    }

    if ($candidatePaths.Count -eq 0) {
        $Diagnostic.Value = "${Source}: no path-like values."
    }
    else {
        $Diagnostic.Value = "${Source}: path-like values found, but none resolved to existing folder."
    }

    return $null
}

function Get-RunningKompasObject {
    param(
        [ref]$ProgId
    )

    $ProgId.Value = ""
    foreach ($candidateProgId in @("Kompas.Application.5", "KOMPAS.Application.5", "Kompas.Application.7", "KOMPAS.Application.7")) {
        try {
            $obj = [Runtime.InteropServices.Marshal]::GetActiveObject($candidateProgId)
            if ($null -ne $obj) {
                $ProgId.Value = $candidateProgId
                return $obj
            }
        }
        catch {
        }
    }

    return $null
}

function Resolve-PythonExecutable {
    foreach ($candidate in @("python", "py")) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($null -ne $cmd) {
            return $cmd.Source
        }
    }

    return $null
}

function Get-KompasActiveDocumentDirectoryViaPython {
    $bridgeScript = Join-Path $script:AppRoot "scripts\resolve_kompas_doc_dir.py"
    if (-not (Test-Path -LiteralPath $bridgeScript -PathType Leaf)) {
        return [pscustomobject]@{
            Directory = $null
            Reason = "Python bridge script not found: $bridgeScript"
        }
    }

    $pythonExe = Resolve-PythonExecutable
    if ([string]::IsNullOrWhiteSpace($pythonExe)) {
        return [pscustomobject]@{
            Directory = $null
            Reason = "Python executable is not available."
        }
    }

    try {
        $outputLines = @(& $pythonExe $bridgeScript 2>&1)
        $exitCode = $LASTEXITCODE
    }
    catch {
        return [pscustomobject]@{
            Directory = $null
            Reason = "Python bridge execution failed: $($_.Exception.Message)"
        }
    }

    $outputText = ($outputLines | ForEach-Object { [string]$_ }) -join "`n"
    $outputText = $outputText.Trim()
    $jsonPayload = $null

    try {
        $jsonLine = $null
        foreach ($line in ($outputText -split "(`r`n|`n|`r)")) {
            $trimmed = $line.Trim()
            if ($trimmed.StartsWith("{") -and $trimmed.EndsWith("}")) {
                $jsonLine = $trimmed
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($jsonLine)) {
            $jsonPayload = $jsonLine | ConvertFrom-Json
        }
    }
    catch {
        $jsonPayload = $null
    }

    if ($null -ne $jsonPayload) {
        $payloadDirectory = Convert-ToNonEmptyString -Value $jsonPayload.directory
        $payloadReason = Convert-ToNonEmptyString -Value $jsonPayload.reason
        $payloadOk = $false
        try {
            $payloadOk = [bool]$jsonPayload.ok
        }
        catch {
            $payloadOk = $false
        }

        if ($payloadOk) {
            $resolved = Resolve-DirectoryFromCandidatePath -CandidatePath $payloadDirectory
            if (-not [string]::IsNullOrWhiteSpace($resolved)) {
                $reasonText = "Resolved via python bridge."
                if (-not [string]::IsNullOrWhiteSpace($payloadReason)) {
                    $reasonText = "Resolved via python bridge ($payloadReason)."
                }

                return [pscustomobject]@{
                    Directory = $resolved
                    Reason = $reasonText
                }
            }

            return [pscustomobject]@{
                Directory = $null
                Reason = "Python bridge returned unresolved path: $payloadDirectory"
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($payloadReason)) {
            return [pscustomobject]@{
                Directory = $null
                Reason = "Python bridge failed: $payloadReason"
            }
        }
    }

    if ($exitCode -eq 0) {
        $resolved = Resolve-DirectoryFromCandidatePath -CandidatePath $outputText
        if (-not [string]::IsNullOrWhiteSpace($resolved)) {
            return [pscustomobject]@{
                Directory = $resolved
                Reason = "Resolved via python bridge."
            }
        }

        return [pscustomobject]@{
            Directory = $null
            Reason = "Python bridge returned non-path output: $outputText"
        }
    }

    if ([string]::IsNullOrWhiteSpace($outputText)) {
        $outputText = "bridge exit code $exitCode"
    }

    return [pscustomobject]@{
        Directory = $null
        Reason = "Python bridge failed: $outputText"
    }
}

function Get-KompasActiveDocumentDirectory {
    $script:LastKompasDocResolveReason = ""
    $resolvedProgId = ""
    $runningKompas = Get-RunningKompasObject -ProgId ([ref]$resolvedProgId)
    if ($null -eq $runningKompas) {
        $script:LastKompasDocResolveReason = "KOMPAS instance is not available in current user context (or COM bitness mismatch)."
        return $null
    }

    $kompas5 = $null
    $kompas7 = $null

    if ($resolvedProgId.ToLowerInvariant().Contains(".5")) {
        $kompas5 = $runningKompas
        try {
            $kompas7 = $kompas5.ksGetApplication7
        }
        catch {
            $kompas7 = $null
        }

        if ($null -eq $kompas7) {
            try {
                $kompas7 = $kompas5.ksGetApplication7()
            }
            catch {
                $kompas7 = $null
            }
        }

        if ($null -eq $kompas7) {
            $kompas7 = Invoke-ComObjectMethod -ComObject $kompas5 -MethodNames @("ksGetApplication7")
        }
    }
    else {
        $kompas7 = $runningKompas
    }

    $docCandidates = New-Object System.Collections.ArrayList

    Add-DocumentsFromContainer -Candidates $docCandidates -Container $kompas7 -SourcePrefix "kompas7"
    Add-DocumentsFromContainer -Candidates $docCandidates -Container $kompas5 -SourcePrefix "kompas5"

    $activeWindow = Get-ComObjectProperty -ComObject $kompas7 -PropertyNames @("ActiveWindow")
    if ($null -eq $activeWindow) {
        $activeWindow = Get-ComObjectProperty -ComObject $kompas5 -PropertyNames @("ActiveWindow")
    }
    if ($null -ne $activeWindow) {
        $windowDoc = Get-ComObjectProperty -ComObject $activeWindow -PropertyNames @("Document", "ActiveDocument", "Doc")
        if ($null -eq $windowDoc) {
            $windowDoc = Invoke-ComObjectMethod -ComObject $activeWindow -MethodNames @("GetDocument")
        }
        Add-DocumentCandidate -Candidates $docCandidates -Document $windowDoc -Source "activeWindow.Document"
    }

    if ($docCandidates.Count -eq 0) {
        $script:LastKompasDocResolveReason = "No active/open documents exposed by KOMPAS COM ($resolvedProgId)."
        return $null
    }

    $diagnostics = New-Object System.Collections.ArrayList
    foreach ($candidate in $docCandidates) {
        $diagnostic = ""
        $resolvedDir = Resolve-KompasDocumentDirectory -Document $candidate.Document -Source $candidate.Source -Diagnostic ([ref]$diagnostic)
        if (-not [string]::IsNullOrWhiteSpace($resolvedDir)) {
            $script:LastKompasDocResolveReason = "Using active KOMPAS document directory via $($candidate.Source) [$resolvedProgId]: $resolvedDir"
            return $resolvedDir
        }

        if (-not [string]::IsNullOrWhiteSpace($diagnostic)) {
            [void]$diagnostics.Add($diagnostic)
        }
    }

    $diagPreview = ""
    if ($diagnostics.Count -gt 0) {
        $diagPreview = (($diagnostics | Select-Object -First 3) -join " | ")
    }

    if (-not [string]::IsNullOrWhiteSpace($diagPreview)) {
        $script:LastKompasDocResolveReason = "No saved path resolved from KOMPAS documents ($resolvedProgId). " + $diagPreview
    }
    else {
        $script:LastKompasDocResolveReason = "No saved path resolved from KOMPAS documents ($resolvedProgId)."
    }

    $pythonAttempt = Get-KompasActiveDocumentDirectoryViaPython
    if ($null -ne $pythonAttempt -and -not [string]::IsNullOrWhiteSpace($pythonAttempt.Directory)) {
        $script:LastKompasDocResolveReason = "Using active KOMPAS document directory via python bridge: $($pythonAttempt.Directory)"
        return $pythonAttempt.Directory
    }

    if ($null -ne $pythonAttempt -and -not [string]::IsNullOrWhiteSpace($pythonAttempt.Reason)) {
        $script:LastKompasDocResolveReason = $script:LastKompasDocResolveReason + " | " + $pythonAttempt.Reason
    }

    return $null
}

function Get-PreferredOutputDirectory {
    $kompasDir = Get-KompasActiveDocumentDirectory
    if (-not [string]::IsNullOrWhiteSpace($kompasDir)) {
        return $kompasDir
    }

    if ([string]::IsNullOrWhiteSpace($script:LastKompasDocResolveReason)) {
        $script:LastKompasDocResolveReason = "Using fallback output directory."
    }

    return (Join-Path $script:AppRoot "out")
}

function Build-DefaultOutputPath {
    param(
        [string]$InputPath,
        [string]$OutputDirectory
    )

    $dir = $OutputDirectory
    if ([string]::IsNullOrWhiteSpace($dir)) {
        $dir = Join-Path $script:AppRoot "out"
    }

    $baseName = "table"
    try {
        if (-not [string]::IsNullOrWhiteSpace($InputPath)) {
            $candidate = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
            if (-not [string]::IsNullOrWhiteSpace($candidate)) {
                $baseName = $candidate
            }
        }
    }
    catch {
    }

    return (Join-Path $dir ($baseName + ".tbl"))
}

function Is-LegacyDefaultOutputPath {
    param(
        [string]$OutputPath
    )

    if ([string]::IsNullOrWhiteSpace($OutputPath)) { return $true }

    $legacyDir = Join-Path $script:AppRoot "out"

    try {
        $normalizedOutput = [System.IO.Path]::GetFullPath($OutputPath).TrimEnd('\')
        $normalizedLegacy = [System.IO.Path]::GetFullPath($legacyDir).TrimEnd('\')

        return (
            $normalizedOutput.Equals($normalizedLegacy, [System.StringComparison]::OrdinalIgnoreCase) -or
            $normalizedOutput.StartsWith($normalizedLegacy + "\", [System.StringComparison]::OrdinalIgnoreCase)
        )
    }
    catch {
        return $false
    }
}

function Enable-DragDropForElevatedWindow {
    param(
        [System.Windows.Forms.Form]$TargetForm,
        [System.Windows.Forms.TextBox]$LogBox
    )

    try {
        $null = $TargetForm.Handle
        [Win32DragDrop]::DragAcceptFiles($TargetForm.Handle, $true)

        $MSGFLT_ALLOW = [uint32]1
        foreach ($msg in @(0x0233, 0x0049, 0x004A)) {
            [void][Win32DragDrop]::ChangeWindowMessageFilterEx(
                $TargetForm.Handle,
                [uint32]$msg,
                $MSGFLT_ALLOW,
                [IntPtr]::Zero
            )
        }

        Add-Log -Box $LogBox -Message "INFO: Включён фильтр сообщений drag/drop для повышенных прав."
    }
    catch {
        Add-Log -Box $LogBox -Message "WARN: Не удалось включить drag/drop фильтр для повышенных прав: $($_.Exception.Message)"
    }
}

function Enable-WmDropFilesWatcher {
    param(
        [System.Windows.Forms.Form]$TargetForm,
        [hashtable]$Ctl,
        [System.Windows.Forms.TextBox]$LogBox
    )

    try {
        $null = $TargetForm.Handle

        $script:DropCtlRef = $Ctl
        $script:DropLogRef = $LogBox
        $script:DropWatcher = New-Object DropFilesWatcher($TargetForm.Handle)
        $script:DropHandler = [System.Action[string[]]]{
            param([string[]]$files)
            Apply-DroppedXlsx -Files $files -Ctl $script:DropCtlRef -LogBox $script:DropLogRef
        }

        $script:DropWatcher.add_FilesDropped($script:DropHandler)
        Add-Log -Box $LogBox -Message "INFO: Включён обработчик WM_DROPFILES."
    }
    catch {
        Add-Log -Box $LogBox -Message "WARN: Не удалось включить обработчик WM_DROPFILES: $($_.Exception.Message)"
    }
}

function Get-DefaultSettings {
    $defaultOutDir = Get-PreferredOutputDirectory
    $defaultOutPath = Build-DefaultOutputPath -InputPath $script:DefaultInput -OutputDirectory $defaultOutDir

    return [ordered]@{
        input_path = $script:DefaultInput
        output_path = $defaultOutPath
        mode = "cell"
        cell_width_mm = "30"
        cell_height_mm = "8"
        table_width_mm = "390"
        table_height_mm = "64"
    }
}

function Read-AppSettings([string]$Path) {
    $defaults = Get-DefaultSettings
    if (-not (Test-Path -LiteralPath $Path)) { return $defaults }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw
        if ([string]::IsNullOrWhiteSpace($raw)) { return $defaults }
        $json = $raw | ConvertFrom-Json

        foreach ($k in $defaults.Keys) {
            if ($null -ne $json.$k -and -not [string]::IsNullOrWhiteSpace([string]$json.$k)) {
                $defaults[$k] = [string]$json.$k
            }
        }
    }
    catch {
        return (Get-DefaultSettings)
    }

    if ($defaults.mode -ne "cell" -and $defaults.mode -ne "table") {
        $defaults.mode = "cell"
    }

    return $defaults
}

function Write-AppSettings([string]$Path, [hashtable]$Settings) {
    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }

    ($Settings | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Try-ParsePositive([string]$Raw, [ref]$Value) {
    $normalized = $Raw.Trim().Replace(" ", "").Replace(",", ".")
    $tmp = 0.0
    if (-not [double]::TryParse($normalized, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$tmp)) {
        return $false
    }

    if ($tmp -le 0) { return $false }

    $Value.Value = $tmp
    return $true
}

function Normalize-Layout {
    param(
        [string]$Mode,
        [string]$CellW,
        [string]$CellH,
        [string]$TableW,
        [string]$TableH,
        [ref]$Error
    )

    $modeNorm = $Mode.Trim().ToLowerInvariant()
    if ($modeNorm -ne "cell" -and $modeNorm -ne "table") {
        $Error.Value = "Mode must be 'cell' or 'table'."
        return $null
    }

    $cw = 0.0; $ch = 0.0; $tw = 390.0; $th = 64.0

    if ($modeNorm -eq "cell") {
        if (-not (Try-ParsePositive -Raw $CellW -Value ([ref]$cw))) { $Error.Value = "Invalid Cell width"; return $null }
        if (-not (Try-ParsePositive -Raw $CellH -Value ([ref]$ch))) { $Error.Value = "Invalid Cell height"; return $null }

        if ([string]::IsNullOrWhiteSpace($TableW) -or -not (Try-ParsePositive -Raw $TableW -Value ([ref]$tw))) { $tw = 390.0 }
        if ([string]::IsNullOrWhiteSpace($TableH) -or -not (Try-ParsePositive -Raw $TableH -Value ([ref]$th))) { $th = 64.0 }
    }
    else {
        if (-not (Try-ParsePositive -Raw $TableW -Value ([ref]$tw))) { $Error.Value = "Invalid Table width"; return $null }
        if (-not (Try-ParsePositive -Raw $TableH -Value ([ref]$th))) { $Error.Value = "Invalid Table height"; return $null }

        if ([string]::IsNullOrWhiteSpace($CellW) -or -not (Try-ParsePositive -Raw $CellW -Value ([ref]$cw))) { $cw = 30.0 }
        if ([string]::IsNullOrWhiteSpace($CellH) -or -not (Try-ParsePositive -Raw $CellH -Value ([ref]$ch))) { $ch = 8.0 }
    }

    $Error.Value = ""
    return [ordered]@{
        mode = $modeNorm
        cell_width_mm = (Format-Inv $cw)
        cell_height_mm = (Format-Inv $ch)
        table_width_mm = (Format-Inv $tw)
        table_height_mm = (Format-Inv $th)
    }
}

function Write-LayoutConfig([string]$Path, [hashtable]$Layout) {
    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }

    $lines = @(
        "; Auto-generated by app-xlsx-to-kompas-tbl GUI"
        "mode=$($Layout.mode)"
        "cell_width_mm=$($Layout.cell_width_mm)"
        "cell_height_mm=$($Layout.cell_height_mm)"
        "table_width_mm=$($Layout.table_width_mm)"
        "table_height_mm=$($Layout.table_height_mm)"
    )

    Set-Content -LiteralPath $Path -Encoding ASCII -Value $lines
}

function Convert-ModeToUiValue([string]$Mode) {
    if ($Mode -eq "table") {
        return "Габариты таблицы"
    }
    return "Ячейки"
}

function Convert-UiToModeValue {
    param(
        [object]$UiValue,
        [int]$SelectedIndex = -1
    )

    if ($SelectedIndex -eq 1) {
        return "table"
    }

    $uiText = [string]$UiValue
    if ($uiText -eq "Габариты таблицы") {
        return "table"
    }

    return "cell"
}

function Update-ModeState([hashtable]$Ctl) {
    $modeValue = Convert-UiToModeValue -UiValue $Ctl.mode.SelectedItem -SelectedIndex $Ctl.mode.SelectedIndex
    $cellMode = ($modeValue -eq "cell")
    $Ctl.cellW.Enabled = $cellMode
    $Ctl.cellH.Enabled = $cellMode
    $Ctl.tableW.Enabled = -not $cellMode
    $Ctl.tableH.Enabled = -not $cellMode
}

function Apply-SettingsToUi([hashtable]$Settings, [hashtable]$Ctl) {
    $Ctl.input.Text = [string]$Settings.input_path
    $Ctl.output.Text = [string]$Settings.output_path

    $mode = [string]$Settings.mode
    if ($mode -ne "cell" -and $mode -ne "table") { $mode = "cell" }

    if ($mode -eq "table") {
        $Ctl.mode.SelectedIndex = 1
    }
    else {
        $Ctl.mode.SelectedIndex = 0
    }

    $Ctl.cellW.Text = [string]$Settings.cell_width_mm
    $Ctl.cellH.Text = [string]$Settings.cell_height_mm
    $Ctl.tableW.Text = [string]$Settings.table_width_mm
    $Ctl.tableH.Text = [string]$Settings.table_height_mm

    Update-ModeState -Ctl $Ctl
}

function Apply-DroppedXlsx {
    param(
        [string[]]$Files,
        [hashtable]$Ctl,
        [System.Windows.Forms.TextBox]$LogBox
    )

    $xlsxPath = $null
    foreach ($filePath in $Files) {
        if ($filePath.ToLowerInvariant().EndsWith(".xlsx")) {
            $xlsxPath = $filePath
            break
        }
    }

    if ($null -eq $xlsxPath) {
        Add-Log -Box $LogBox -Message "WARN: Перетаскивание проигнорировано. Нужен файл .xlsx."
        return
    }

    if (-not (Test-Path -LiteralPath $xlsxPath)) {
        Add-Log -Box $LogBox -Message "WARN: Перетащенный файл не существует: $xlsxPath"
        return
    }

    $Ctl.input.Text = $xlsxPath

    $preferredDir = Get-PreferredOutputDirectory
    $Ctl.output.Text = Build-DefaultOutputPath -InputPath $xlsxPath -OutputDirectory $preferredDir

    Add-Log -Box $LogBox -Message "Перетащен XLSX: $xlsxPath"
    Add-Log -Box $LogBox -Message "Определена папка результата: $preferredDir"
    if (-not [string]::IsNullOrWhiteSpace($script:LastKompasDocResolveReason)) {
        Add-Log -Box $LogBox -Message "Диагностика пути: $script:LastKompasDocResolveReason"
    }
}

function Register-XlsxDropTargets {
    param(
        [System.Windows.Forms.Control]$RootControl,
        [hashtable]$Ctl,
        [System.Windows.Forms.TextBox]$LogBox
    )

    if ($null -eq $RootControl) { return }

    $RootControl.AllowDrop = $true

    $RootControl.Add_DragEnter({
        param($sender, $e)

        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
        if (-not $e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { return }

        $files = [string[]]$e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        foreach ($filePath in $files) {
            if ($filePath.ToLowerInvariant().EndsWith(".xlsx")) {
                $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
                return
            }
        }
    })

    $RootControl.Add_DragDrop({
        param($sender, $e)

        if (-not $e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { return }
        $files = [string[]]$e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        Apply-DroppedXlsx -Files $files -Ctl $Ctl -LogBox $LogBox
    })

    foreach ($child in $RootControl.Controls) {
        Register-XlsxDropTargets -RootControl $child -Ctl $Ctl -LogBox $LogBox
    }
}

function Save-FullSettings {
    param(
        [hashtable]$Ctl,
        [System.Windows.Forms.TextBox]$LogBox,
        [switch]$Silent
    )

    $inputPath = $Ctl.input.Text.Trim()
    $outputPath = $Ctl.output.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        Add-Log -Box $LogBox -Message "ERROR: Не указан путь к входному файлу."
        if (-not $Silent) { [System.Windows.Forms.MessageBox]::Show("Не указан путь к входному файлу.", "Импорт XLSX в TBL") | Out-Null }
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($outputPath)) {
        Add-Log -Box $LogBox -Message "ERROR: Не указан путь к выходному файлу."
        if (-not $Silent) { [System.Windows.Forms.MessageBox]::Show("Не указан путь к выходному файлу.", "Импорт XLSX в TBL") | Out-Null }
        return $null
    }

    $err = ""
    $layoutMode = Convert-UiToModeValue -UiValue $Ctl.mode.SelectedItem -SelectedIndex $Ctl.mode.SelectedIndex
    $layout = Normalize-Layout -Mode $layoutMode -CellW $Ctl.cellW.Text -CellH $Ctl.cellH.Text -TableW $Ctl.tableW.Text -TableH $Ctl.tableH.Text -Error ([ref]$err)
    if ($null -eq $layout) {
        Add-Log -Box $LogBox -Message "ERROR: Некорректные параметры: $err"
        if (-not $Silent) { [System.Windows.Forms.MessageBox]::Show("Некорректные параметры:`n$err", "Импорт XLSX в TBL") | Out-Null }
        return $null
    }

    $settings = [ordered]@{
        input_path = $inputPath
        output_path = $outputPath
        mode = $layout.mode
        cell_width_mm = $layout.cell_width_mm
        cell_height_mm = $layout.cell_height_mm
        table_width_mm = $layout.table_width_mm
        table_height_mm = $layout.table_height_mm
    }

    try {
        Write-AppSettings -Path $script:SettingsPath -Settings $settings
        Write-LayoutConfig -Path $script:LayoutConfigPath -Layout $settings
        Apply-SettingsToUi -Settings $settings -Ctl $Ctl
        Add-Log -Box $LogBox -Message "Параметры сохранены (профиль + layout config)."
        return $settings
    }
    catch {
        Add-Log -Box $LogBox -Message "ERROR: Ошибка сохранения параметров: $($_.Exception.Message)"
        if (-not $Silent) { [System.Windows.Forms.MessageBox]::Show("Не удалось сохранить параметры.`n$($_.Exception.Message)", "Импорт XLSX в TBL") | Out-Null }
        return $null
    }
}

function Invoke-Export {
    param(
        [hashtable]$Settings,
        [System.Windows.Forms.TextBox]$LogBox
    )

    if (-not (Test-Path -LiteralPath $script:ExporterVbs)) {
        Add-Log -Box $LogBox -Message "ERROR: Exporter missing: $script:ExporterVbs"
        return
    }

    if (-not (Test-Path -LiteralPath $Settings.input_path)) {
        Add-Log -Box $LogBox -Message "ERROR: Input file missing: $($Settings.input_path)"
        return
    }

    if (-not (Test-Path -LiteralPath $script:LayoutConfigPath)) {
        Add-Log -Box $LogBox -Message "ERROR: Auto layout config missing: $script:LayoutConfigPath"
        return
    }

    $outDir = Split-Path -Parent $Settings.output_path
    if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
        New-Item -Path $outDir -ItemType Directory -Force | Out-Null
    }

    Add-Log -Box $LogBox -Message "Запуск экспорта..."
    Add-Log -Box $LogBox -Message "Input: $($Settings.input_path)"
    Add-Log -Box $LogBox -Message "Output: $($Settings.output_path)"
    Add-Log -Box $LogBox -Message "Mode: $($Settings.mode)"

    try {
        $outputLines = @(
            & cscript.exe //nologo $script:ExporterVbs $Settings.input_path $Settings.output_path $script:LayoutConfigPath 2>&1
        )
        $exitCode = $LASTEXITCODE

        foreach ($line in ($outputLines | ForEach-Object { [string]$_ })) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Add-Log -Box $LogBox -Message $line
            }
        }

        if ($exitCode -eq 0 -and (Test-Path -LiteralPath $Settings.output_path)) {
            $size = (Get-Item -LiteralPath $Settings.output_path).Length
            Add-Log -Box $LogBox -Message "DONE: Экспорт завершён, размер=$size байт."
        }
        else {
            Add-Log -Box $LogBox -Message "ERROR: cscript returned code $exitCode."
        }
    }
    catch {
        Add-Log -Box $LogBox -Message "ERROR: $($_.Exception.Message)"
    }
}

function Invoke-InsertTblToKompas {
    param(
        [string]$TblPath,
        [System.Windows.Forms.TextBox]$LogBox
    )

    if ([string]::IsNullOrWhiteSpace($TblPath)) {
        Add-Log -Box $LogBox -Message "ERROR: Не указан путь к .tbl для вставки."
        return
    }

    if (-not (Test-Path -LiteralPath $TblPath -PathType Leaf)) {
        Add-Log -Box $LogBox -Message "ERROR: Файл .tbl для вставки не найден: $TblPath"
        return
    }

    if (-not (Test-Path -LiteralPath $script:InsertBridgePy -PathType Leaf)) {
        Add-Log -Box $LogBox -Message "ERROR: Скрипт вставки не найден: $script:InsertBridgePy"
        return
    }

    $pythonExe = Resolve-PythonExecutable
    if ([string]::IsNullOrWhiteSpace($pythonExe)) {
        Add-Log -Box $LogBox -Message "ERROR: Python не найден. Невозможно выполнить вставку .tbl."
        return
    }

    Add-Log -Box $LogBox -Message "Запуск вставки .tbl в активный документ КОМПАС..."
    Add-Log -Box $LogBox -Message "Файл .tbl: $TblPath"

    try {
        $outputLines = @(
            & $pythonExe $script:InsertBridgePy $TblPath 2>&1
        )
        $exitCode = $LASTEXITCODE

        foreach ($line in ($outputLines | ForEach-Object { [string]$_ })) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                Add-Log -Box $LogBox -Message $line
            }
        }

        if ($exitCode -eq 0) {
            Add-Log -Box $LogBox -Message "DONE: Таблица .tbl вставлена в активный документ."
        }
        else {
            if ($exitCode -eq 42) {
                Add-Log -Box $LogBox -Message "WARN: Вставка не подтверждена визуально (число таблиц не изменилось). Проверьте активный вид/слой в КОМПАС."
            }
            Add-Log -Box $LogBox -Message "ERROR: Вставка .tbl завершилась с кодом $exitCode."
        }
    }
    catch {
        Add-Log -Box $LogBox -Message "ERROR: $($_.Exception.Message)"
    }
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

$form = New-Object System.Windows.Forms.Form
$form.Text = "Импорт XLSX в TBL (КОМПАС)"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(640, 500)
$form.MinimumSize = New-Object System.Drawing.Size(640, 500)
$form.ShowIcon = $true
$form.ShowInTaskbar = $true
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

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Файл Excel (.xlsx):"
$lblInput.Location = New-Object System.Drawing.Point(10, 8)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(10, 24)
$txtInput.Size = New-Object System.Drawing.Size(560, 24)
$form.Controls.Add($txtInput)

$btnInput = New-Object System.Windows.Forms.Button
$btnInput.Location = New-Object System.Drawing.Point(576, 23)
$btnInput.Size = New-Object System.Drawing.Size(34, 26)
$form.Controls.Add($btnInput)

$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Файл результата (.tbl):"
$lblOutput.Location = New-Object System.Drawing.Point(10, 52)
$lblOutput.AutoSize = $true
$form.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(10, 68)
$txtOutput.Size = New-Object System.Drawing.Size(560, 24)
$form.Controls.Add($txtOutput)

$btnOutput = New-Object System.Windows.Forms.Button
$btnOutput.Location = New-Object System.Drawing.Point(576, 67)
$btnOutput.Size = New-Object System.Drawing.Size(34, 26)
$form.Controls.Add($btnOutput)

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Параметры таблицы"
$grp.Location = New-Object System.Drawing.Point(10, 98)
$grp.Size = New-Object System.Drawing.Size(600, 124)
$form.Controls.Add($grp)

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Text = "Режим:"
$lblMode.Location = New-Object System.Drawing.Point(10, 24)
$lblMode.AutoSize = $true
$grp.Controls.Add($lblMode)

$cmbMode = New-Object System.Windows.Forms.ComboBox
$cmbMode.Location = New-Object System.Drawing.Point(64, 20)
$cmbMode.Size = New-Object System.Drawing.Size(144, 24)
$cmbMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cmbMode.Items.Add("Ячейки")
[void]$cmbMode.Items.Add("Габариты таблицы")
$cmbMode.SelectedIndex = 0
$grp.Controls.Add($cmbMode)

$lblCW = New-Object System.Windows.Forms.Label
$lblCW.Text = "W яч., мм:"
$lblCW.Location = New-Object System.Drawing.Point(222, 24)
$lblCW.AutoSize = $true
$grp.Controls.Add($lblCW)

$txtCW = New-Object System.Windows.Forms.TextBox
$txtCW.Location = New-Object System.Drawing.Point(292, 20)
$txtCW.Size = New-Object System.Drawing.Size(60, 24)
$grp.Controls.Add($txtCW)

$lblCH = New-Object System.Windows.Forms.Label
$lblCH.Text = "H яч., мм:"
$lblCH.Location = New-Object System.Drawing.Point(364, 24)
$lblCH.AutoSize = $true
$grp.Controls.Add($lblCH)

$txtCH = New-Object System.Windows.Forms.TextBox
$txtCH.Location = New-Object System.Drawing.Point(434, 20)
$txtCH.Size = New-Object System.Drawing.Size(60, 24)
$grp.Controls.Add($txtCH)

$lblTW = New-Object System.Windows.Forms.Label
$lblTW.Text = "W табл., мм:"
$lblTW.Location = New-Object System.Drawing.Point(222, 54)
$lblTW.AutoSize = $true
$grp.Controls.Add($lblTW)

$txtTW = New-Object System.Windows.Forms.TextBox
$txtTW.Location = New-Object System.Drawing.Point(292, 50)
$txtTW.Size = New-Object System.Drawing.Size(60, 24)
$grp.Controls.Add($txtTW)

$lblTH = New-Object System.Windows.Forms.Label
$lblTH.Text = "H табл., мм:"
$lblTH.Location = New-Object System.Drawing.Point(364, 54)
$lblTH.AutoSize = $true
$grp.Controls.Add($lblTH)

$txtTH = New-Object System.Windows.Forms.TextBox
$txtTH.Location = New-Object System.Drawing.Point(434, 50)
$txtTH.Size = New-Object System.Drawing.Size(60, 24)
$grp.Controls.Add($txtTH)

$btnApply = New-Object System.Windows.Forms.Button
$btnReset = New-Object System.Windows.Forms.Button

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Параметры сохраняются автоматически. Наведите курсор на иконки для подсказок."
$lblInfo.Location = New-Object System.Drawing.Point(10, 86)
$lblInfo.AutoSize = $true
$grp.Controls.Add($lblInfo)

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Text = "Профиль: $script:SettingsPath"
$lblProfile.Location = New-Object System.Drawing.Point(10, 104)
$lblProfile.Size = New-Object System.Drawing.Size(580, 16)
$lblProfile.AutoEllipsis = $true
$lblProfile.Visible = $false
$grp.Controls.Add($lblProfile)

$lblLayout = New-Object System.Windows.Forms.Label
$lblLayout.Text = "Автоконфиг разметки: $script:LayoutConfigPath"
$lblLayout.Location = New-Object System.Drawing.Point(10, 104)
$lblLayout.Size = New-Object System.Drawing.Size(580, 16)
$lblLayout.AutoEllipsis = $true
$lblLayout.Visible = $false
$grp.Controls.Add($lblLayout)

$lblActions = New-Object System.Windows.Forms.Label
$lblActions.Text = "Действия:"
$lblActions.Location = New-Object System.Drawing.Point(10, 228)
$lblActions.AutoSize = $true
$form.Controls.Add($lblActions)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(78, 222)
$btnRun.Size = New-Object System.Drawing.Size(34, 30)
$form.Controls.Add($btnRun)

$btnInsert = New-Object System.Windows.Forms.Button
$btnInsert.Location = New-Object System.Drawing.Point(116, 222)
$btnInsert.Size = New-Object System.Drawing.Size(34, 30)
$form.Controls.Add($btnInsert)

$btnOpenOut = New-Object System.Windows.Forms.Button
$btnOpenOut.Location = New-Object System.Drawing.Point(154, 222)
$btnOpenOut.Size = New-Object System.Drawing.Size(34, 30)
$form.Controls.Add($btnOpenOut)

$btnApply.Location = New-Object System.Drawing.Point(192, 222)
$btnApply.Size = New-Object System.Drawing.Size(34, 30)
$form.Controls.Add($btnApply)

$btnReset.Location = New-Object System.Drawing.Point(230, 222)
$btnReset.Size = New-Object System.Drawing.Size(34, 30)
$form.Controls.Add($btnReset)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = New-Object System.Drawing.Point(576, 222)
$btnExit.Size = New-Object System.Drawing.Size(34, 30)
$form.Controls.Add($btnExit)

Set-IconButton -Button $btnInput -Icon ([System.Drawing.SystemIcons]::Asterisk) -ToolTip $toolTip -TipText "Выбрать входной файл Excel (.xlsx)" -FallbackText "📥"
Set-IconButton -Button $btnOutput -Icon ([System.Drawing.SystemIcons]::Question) -ToolTip $toolTip -TipText "Выбрать путь выходного файла (.tbl)" -FallbackText "💾"
Set-IconButton -Button $btnRun -Icon ([System.Drawing.SystemIcons]::Application) -ToolTip $toolTip -TipText "Импорт XLSX в TBL" -FallbackText "▶"
Set-IconButton -Button $btnInsert -Icon ([System.Drawing.SystemIcons]::WinLogo) -ToolTip $toolTip -TipText "Вставить текущий .tbl в активный документ КОМПАС" -FallbackText "↳"
Set-IconButton -Button $btnOpenOut -Icon ([System.Drawing.SystemIcons]::Information) -ToolTip $toolTip -TipText "Открыть папку результата" -FallbackText "📂"
Set-IconButton -Button $btnApply -Icon ([System.Drawing.SystemIcons]::Shield) -ToolTip $toolTip -TipText "Применить и сохранить параметры" -FallbackText "✔"
Set-IconButton -Button $btnReset -Icon ([System.Drawing.SystemIcons]::Warning) -ToolTip $toolTip -TipText "Сбросить параметры по умолчанию" -FallbackText "↺"
Set-IconButton -Button $btnExit -Icon ([System.Drawing.SystemIcons]::Error) -ToolTip $toolTip -TipText "Закрыть окно" -FallbackText "✕"

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Журнал:"
$lblLog.Location = New-Object System.Drawing.Point(10, 258)
$lblLog.AutoSize = $true
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 274)
$txtLog.Size = New-Object System.Drawing.Size(600, 176)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

$controls = @{
    input = $txtInput
    output = $txtOutput
    mode = $cmbMode
    cellW = $txtCW
    cellH = $txtCH
    tableW = $txtTW
    tableH = $txtTH
}

Register-XlsxDropTargets -RootControl $form -Ctl $controls -LogBox $txtLog

$cmbMode.Add_SelectedIndexChanged({ Update-ModeState -Ctl $controls })

$btnInput.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*"
    $dialog.FileName = $txtInput.Text
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtInput.Text = $dialog.FileName
        $preferredDir = Get-PreferredOutputDirectory
        $txtOutput.Text = Build-DefaultOutputPath -InputPath $dialog.FileName -OutputDirectory $preferredDir
        Add-Log -Box $txtLog -Message "Выбран входной файл: $($dialog.FileName)"
        Add-Log -Box $txtLog -Message "Определена папка результата: $preferredDir"
        if (-not [string]::IsNullOrWhiteSpace($script:LastKompasDocResolveReason)) {
            Add-Log -Box $txtLog -Message "Диагностика пути: $script:LastKompasDocResolveReason"
        }
    }
})

$btnOutput.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Таблица КОМПАС (*.tbl)|*.tbl|Все файлы (*.*)|*.*"
    $dialog.FileName = $txtOutput.Text
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOutput.Text = $dialog.FileName
    }
})

$btnApply.Add_Click({
    $saved = Save-FullSettings -Ctl $controls -LogBox $txtLog
    if ($null -ne $saved) {
        [System.Windows.Forms.MessageBox]::Show("Параметры применены.", "Импорт XLSX в TBL") | Out-Null
    }
})

$btnReset.Add_Click({
    $defaults = Get-DefaultSettings
    Apply-SettingsToUi -Settings $defaults -Ctl $controls
    Add-Log -Box $txtLog -Message "Параметры интерфейса сброшены к значениям по умолчанию."
})

$btnRun.Add_Click({
    $saved = Save-FullSettings -Ctl $controls -LogBox $txtLog
    if ($null -eq $saved) { return }
    Invoke-Export -Settings $saved -LogBox $txtLog
})

$btnInsert.Add_Click({
    $saved = Save-FullSettings -Ctl $controls -LogBox $txtLog
    if ($null -eq $saved) { return }
    Invoke-InsertTblToKompas -TblPath $saved.output_path -LogBox $txtLog
})

$btnOpenOut.Add_Click({
    $dir = Split-Path -Parent $txtOutput.Text
    if ([string]::IsNullOrWhiteSpace($dir)) { $dir = Join-Path $script:AppRoot "out" }
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    Start-Process explorer.exe $dir
})

$btnExit.Add_Click({ $form.Close() })

$form.Add_FormClosing({
    if ($null -ne $script:DropWatcher) {
        try { $script:DropWatcher.Dispose() } catch {}
        $script:DropWatcher = $null
    }

    if ($null -ne $script:FormAppIcon) {
        try { $script:FormAppIcon.Dispose() } catch {}
        $script:FormAppIcon = $null
    }

    $null = Save-FullSettings -Ctl $controls -LogBox $txtLog -Silent
})

if (Test-Path -LiteralPath $script:ExporterVbs) {
    Add-Log -Box $txtLog -Message "Готово. Перед импортом держите активным 2D документ КОМПАС."
}
else {
    Add-Log -Box $txtLog -Message "WARN: Экспортёр не найден: $script:ExporterVbs"
}

$initial = Read-AppSettings -Path $script:SettingsPath
$preferredOutDir = Get-PreferredOutputDirectory
$initial.output_path = Build-DefaultOutputPath -InputPath $initial.input_path -OutputDirectory $preferredOutDir
Apply-SettingsToUi -Settings $initial -Ctl $controls
Add-Log -Box $txtLog -Message "Режим полных настроек включён. ini-файл генерируется автоматически."
Add-Log -Box $txtLog -Message "Папка результата по умолчанию: $preferredOutDir"
if (-not [string]::IsNullOrWhiteSpace($script:LastKompasDocResolveReason)) {
    Add-Log -Box $txtLog -Message "Диагностика пути: $script:LastKompasDocResolveReason"
}
Enable-DragDropForElevatedWindow -TargetForm $form -LogBox $txtLog
Enable-WmDropFilesWatcher -TargetForm $form -Ctl $controls -LogBox $txtLog
if (Test-IsElevated) {
    Add-Log -Box $txtLog -Message "WARN: Приложение запущено от администратора. Drag-and-drop из обычного Explorer может блокироваться Windows."
    Add-Log -Box $txtLog -Message "WARN: Если перетаскивание не работает, запустите приложение и Explorer с одинаковыми правами."
}

[void]$form.ShowDialog()



