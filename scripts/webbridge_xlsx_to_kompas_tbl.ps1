[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("status", "open-document", "export")]
    [string]$Action
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$KompasApi7Path = "C:\Program Files\ASCON\KOMPAS-3D v24\Libs\PolynomLib\Bin\Client\Interop.KompasAPI7.dll"
$KompasConstantsPath = "C:\Program Files\ASCON\KOMPAS-3D v24\Libs\PolynomLib\Bin\Client\Interop.Kompas6Constants.dll"

function Test-TypeLoaded {
    param([Parameter(Mandatory = $true)][string]$Name)
    return $null -ne ([System.Management.Automation.PSTypeName]$Name).Type
}

function Ensure-InteropAssembly {
    foreach ($path in @($KompasApi7Path, $KompasConstantsPath)) {
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Required KOMPAS interop assembly was not found: $path"
        }
    }

    if (-not (Test-TypeLoaded -Name "KompasAPI7.IApplication")) {
        Add-Type -Path @($KompasApi7Path, $KompasConstantsPath)
    }
}

function Ensure-BridgeType {
    if (Test-TypeLoaded -Name "KompasPagesBridge") {
        return
    }

    $typeDefinition = @"
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using KompasAPI7;
using Kompas6Constants;

public sealed class KompasBridgeStatus
{
    public KompasBridgeStatus()
    {
        DocumentName = string.Empty;
        DocumentPath = string.Empty;
        ViewName = string.Empty;
        ErrorMessage = string.Empty;
    }

    public bool Connected { get; set; }
    public bool HasActiveDocument { get; set; }
    public string DocumentName { get; set; }
    public string DocumentPath { get; set; }
    public string ViewName { get; set; }
    public string ErrorMessage { get; set; }
}

public sealed class KompasOpenDocumentResult
{
    public KompasOpenDocumentResult()
    {
        DocumentName = string.Empty;
        DocumentPath = string.Empty;
        ErrorCode = string.Empty;
        ErrorMessage = string.Empty;
    }

    public bool Success { get; set; }
    public string DocumentName { get; set; }
    public string DocumentPath { get; set; }
    public string ErrorCode { get; set; }
    public string ErrorMessage { get; set; }
}

public sealed class KompasExportResult
{
    public KompasExportResult()
    {
        SourceName = string.Empty;
        OutputPath = string.Empty;
        DocumentName = string.Empty;
        DocumentPath = string.Empty;
        ViewName = string.Empty;
        ErrorCode = string.Empty;
        ErrorMessage = string.Empty;
    }

    public bool Success { get; set; }
    public string SourceName { get; set; }
    public string OutputPath { get; set; }
    public int Rows { get; set; }
    public int Cols { get; set; }
    public bool FileExists { get; set; }
    public long FileSize { get; set; }
    public string DocumentName { get; set; }
    public string DocumentPath { get; set; }
    public string ViewName { get; set; }
    public string ErrorCode { get; set; }
    public string ErrorMessage { get; set; }
}

public static class KompasPagesBridge
{
    private static IApplication TryGetRunningApplication()
    {
        try
        {
            object instance = Marshal.GetActiveObject("KOMPAS.Application.7");
            return instance as IApplication;
        }
        catch
        {
            return null;
        }
    }

    private static IApplication GetOrStartApplication()
    {
        IApplication running = TryGetRunningApplication();
        if (running != null)
        {
            return running;
        }

        Type type = Type.GetTypeFromProgID("KOMPAS.Application.7");
        if (type == null)
        {
            throw new InvalidOperationException("KOMPAS.Application.7 ProgID was not found.");
        }

        object instance = Activator.CreateInstance(type);
        IApplication app = instance as IApplication;
        if (app == null)
        {
            throw new InvalidOperationException("Failed to create KOMPAS.Application.7 instance.");
        }

        return app;
    }

    private static IKompasDocument2D GetActiveDocument2D(IApplication app)
    {
        IKompasDocument2D doc2D = app.ActiveDocument as IKompasDocument2D;
        if (doc2D == null)
        {
            throw new InvalidOperationException("Active 2D document was not found.");
        }

        return doc2D;
    }

    private static IView GetActiveView(IKompasDocument2D doc2D)
    {
        IViews views = doc2D.ViewsAndLayersManager.Views;
        if (views == null)
        {
            throw new InvalidOperationException("Views collection was not found.");
        }

        IView view = views.ActiveView;
        if (view == null)
        {
            throw new InvalidOperationException("Active view was not found.");
        }

        return view;
    }

    private static string SafeString(Func<string> getter)
    {
        try
        {
            string value = getter();
            return value ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string BuildOutputPath(string sourceName, string outputPath)
    {
        string effective = (outputPath ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(effective))
        {
            string baseName = Path.GetFileNameWithoutExtension(sourceName ?? string.Empty);
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "table";
            }

            char[] invalid = Path.GetInvalidFileNameChars();
            string safeName = new string(baseName.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            string tempDir = Path.Combine(Path.GetTempPath(), "kompas-pages");
            Directory.CreateDirectory(tempDir);
            effective = Path.Combine(tempDir, safeName + ".tbl");
        }

        if (!string.Equals(Path.GetExtension(effective), ".tbl", StringComparison.OrdinalIgnoreCase))
        {
            effective = effective + ".tbl";
        }

        string directory = Path.GetDirectoryName(effective) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(directory))
        {
            throw new InvalidOperationException("Output directory is empty.");
        }

        Directory.CreateDirectory(directory);
        return effective;
    }

    public static KompasBridgeStatus GetStatus()
    {
        KompasBridgeStatus result = new KompasBridgeStatus();
        try
        {
            IApplication app = TryGetRunningApplication();
            if (app == null)
            {
                result.Connected = false;
                result.ErrorMessage = "KOMPAS is not running.";
                return result;
            }

            result.Connected = true;

            IKompasDocument2D doc2D = app.ActiveDocument as IKompasDocument2D;
            if (doc2D == null)
            {
                result.HasActiveDocument = false;
                result.ErrorMessage = "Active 2D document was not found.";
                return result;
            }

            result.HasActiveDocument = true;
            result.DocumentName = SafeString(delegate() { return doc2D.Name; });
            result.DocumentPath = SafeString(delegate() { return doc2D.PathName; });
            result.ViewName = SafeString(delegate() { return GetActiveView(doc2D).Name; });
            return result;
        }
        catch (Exception ex)
        {
            result.Connected = false;
            result.ErrorMessage = ex.Message;
            return result;
        }
    }

    public static KompasOpenDocumentResult OpenDocument(string documentPath)
    {
        if (string.IsNullOrWhiteSpace(documentPath))
        {
            throw new ArgumentException("documentPath must not be empty.", "documentPath");
        }

        if (!File.Exists(documentPath))
        {
            throw new FileNotFoundException("Source drawing was not found.", documentPath);
        }

        IApplication app = GetOrStartApplication();
        IKompasDocument2D doc2D = app.Documents.Open(documentPath, false, true) as IKompasDocument2D;
        if (doc2D == null)
        {
            throw new InvalidOperationException("KOMPAS failed to open the requested 2D document.");
        }

        KompasOpenDocumentResult result = new KompasOpenDocumentResult();
        result.Success = true;
        result.DocumentName = SafeString(delegate() { return doc2D.Name; });
        result.DocumentPath = SafeString(delegate() { return doc2D.PathName; });
        return result;
    }

    public static KompasExportResult ExportTable(
        string sourceName,
        string outputPath,
        double cellWidthMm,
        double cellHeightMm,
        string[][] matrix)
    {
        if (matrix == null || matrix.Length == 0)
        {
            throw new InvalidOperationException("Matrix is empty.");
        }

        int rows = matrix.Length;
        int cols = matrix.Max(delegate(string[] row) { return row == null ? 0 : row.Length; });
        if (cols <= 0)
        {
            throw new InvalidOperationException("Matrix has no columns.");
        }

        if (cellWidthMm <= 0 || cellHeightMm <= 0)
        {
            throw new InvalidOperationException("Cell width and height must be positive.");
        }

        IApplication app = GetOrStartApplication();
        IKompasDocument2D doc2D = GetActiveDocument2D(app);
        IView view = GetActiveView(doc2D);
        ISymbols2DContainer symbols = view as ISymbols2DContainer;
        if (symbols == null)
        {
            throw new InvalidOperationException("Active view does not expose ISymbols2DContainer.");
        }

        string effectiveOutputPath = BuildOutputPath(sourceName, outputPath);
        if (File.Exists(effectiveOutputPath))
        {
            File.Delete(effectiveOutputPath);
        }

        IDrawingTable drawingTable = symbols.DrawingTables.Add(
            cols,
            rows,
            cellWidthMm,
            cellHeightMm,
            ksTableTileLayoutEnum.ksTTLNotCreate);

        ITable table = (ITable)drawingTable;
        for (int rowIndex = 0; rowIndex < rows; rowIndex++)
        {
            string[] currentRow = matrix[rowIndex] ?? new string[0];
            for (int columnIndex = 0; columnIndex < cols; columnIndex++)
            {
                string value = columnIndex < currentRow.Length ? (currentRow[columnIndex] ?? string.Empty) : string.Empty;
                ITableCell cell = table.Cell[columnIndex, rowIndex];
                IText text = (IText)cell.Text;
                text.Str = value;
            }
        }

        bool saved = drawingTable.Save(effectiveOutputPath);
        FileInfo fileInfo = new FileInfo(effectiveOutputPath);

        KompasExportResult result = new KompasExportResult();
        result.Success = saved && fileInfo.Exists;
        result.SourceName = sourceName ?? string.Empty;
        result.OutputPath = effectiveOutputPath;
        result.Rows = rows;
        result.Cols = cols;
        result.FileExists = fileInfo.Exists;
        result.FileSize = fileInfo.Exists ? fileInfo.Length : 0;
        result.DocumentName = SafeString(delegate() { return doc2D.Name; });
        result.DocumentPath = SafeString(delegate() { return doc2D.PathName; });
        result.ViewName = SafeString(delegate() { return view.Name; });
        return result;
    }
}
"@

    Add-Type -ReferencedAssemblies @($KompasApi7Path, $KompasConstantsPath) -TypeDefinition $typeDefinition
}

function Read-JsonFromStdin {
    $raw = [Console]::In.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $null
    }

    return $raw | ConvertFrom-Json
}

function Convert-ToStringMatrix {
    param($Matrix)

    if ($null -eq $Matrix) {
        return [string[][]]@()
    }

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Matrix) {
        if ($null -eq $row) {
            [void]$rows.Add([string[]]@())
            continue
        }

        $values = New-Object System.Collections.Generic.List[string]
        foreach ($cell in $row) {
            if ($null -eq $cell) {
                [void]$values.Add("")
            }
            else {
                [void]$values.Add([string]$cell)
            }
        }

        [void]$rows.Add([string[]]$values.ToArray())
    }

    return [string[][]]$rows.ToArray()
}

function Convert-StatusPayload {
    param([Parameter(Mandatory = $true)]$Status)

    return [ordered]@{
        connected         = [bool]$Status.Connected
        hasActiveDocument = [bool]$Status.HasActiveDocument
        documentName      = [string]$Status.DocumentName
        documentPath      = [string]$Status.DocumentPath
        viewName          = [string]$Status.ViewName
        errorMessage      = [string]$Status.ErrorMessage
    }
}

function Convert-OpenDocumentPayload {
    param([Parameter(Mandatory = $true)]$Result)

    return [ordered]@{
        success      = [bool]$Result.Success
        documentName = [string]$Result.DocumentName
        documentPath = [string]$Result.DocumentPath
        errorCode    = [string]$Result.ErrorCode
        errorMessage = [string]$Result.ErrorMessage
    }
}

function Convert-ExportPayload {
    param([Parameter(Mandatory = $true)]$Result)

    return [ordered]@{
        success      = [bool]$Result.Success
        sourceName   = [string]$Result.SourceName
        outputPath   = [string]$Result.OutputPath
        rows         = [int]$Result.Rows
        cols         = [int]$Result.Cols
        fileExists   = [bool]$Result.FileExists
        fileSize     = [int64]$Result.FileSize
        documentName = [string]$Result.DocumentName
        documentPath = [string]$Result.DocumentPath
        viewName     = [string]$Result.ViewName
        errorCode    = [string]$Result.ErrorCode
        errorMessage = [string]$Result.ErrorMessage
    }
}

function Write-JsonPayload {
    param([Parameter(Mandatory = $true)]$Payload)
    $Payload | ConvertTo-Json -Depth 20 -Compress
}

Ensure-InteropAssembly
Ensure-BridgeType

try {
    switch ($Action) {
        "status" {
            $status = [KompasPagesBridge]::GetStatus()
            Write-Output (Write-JsonPayload -Payload (Convert-StatusPayload -Status $status))
            exit 0
        }

        "open-document" {
            $request = Read-JsonFromStdin
            $result = [KompasPagesBridge]::OpenDocument([string]$request.documentPath)
            Write-Output (Write-JsonPayload -Payload (Convert-OpenDocumentPayload -Result $result))
            exit 0
        }

        "export" {
            $request = Read-JsonFromStdin
            $matrix = Convert-ToStringMatrix -Matrix $request.matrix
            $result = [KompasPagesBridge]::ExportTable(
                [string]$request.sourceName,
                [string]$request.outputPath,
                [double]$request.cellWidthMm,
                [double]$request.cellHeightMm,
                $matrix)
            Write-Output (Write-JsonPayload -Payload (Convert-ExportPayload -Result $result))
            exit 0
        }
    }
}
catch {
    $payload = [ordered]@{
        success      = $false
        errorCode    = "UNHANDLED"
        errorMessage = $_.Exception.Message
    }
    Write-Output (Write-JsonPayload -Payload $payload)
    exit 50
}
