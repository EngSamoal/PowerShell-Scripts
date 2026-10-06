#Requires -Version 5.1
<#
.SYNOPSIS
    Extracts the Protection Group table from a Cohesity screenshot pasted into the Backup Summary
    workbook, using Windows' own built-in OCR engine (Windows.Media.Ocr) - no cloud, no external
    tool install, nothing leaves this machine.

.DESCRIPTION
    UNTESTED ON A REAL WINDOWS MACHINE - this was written and syntax-checked in a Linux
    environment with no access to Windows.Media.Ocr or a GUI session, so it could not be run
    end-to-end here. Treat the first run as a trial: it also writes a raw per-line OCR dump next
    to the output workbook so you can visually compare it against the real screenshot before
    trusting the structured result. A misread character (1 vs 7, GiB vs TiB) is possible with any
    OCR - this script cannot catch that on its own, only you comparing against the source image can.

    How it works:
      1. Extracts every embedded picture from each sheet of the Backup Summary workbook (a .xlsx
         is a zip - this unzips xl/media/*.png directly, no Excel COM needed for this part).
      2. Runs each image through Windows' on-device OCR engine (same one behind "Copy text from
         picture" in the Snipping Tool / PowerToys) - entirely local, no network call.
      3. Uses the header row's word X-positions ("Protection Group", "Source", "Runs", ...) as
         column boundaries, then buckets every other line's words into the nearest column by X
         position - a standard approach for reconstructing a table from OCR'd text, which has no
         native column structure of its own.
      4. Writes the result to an .xlsx in the exact shape BackupStorage_Weekly_HealthCheck.ps1
         expects (Site, ProtectionGroup, Source, Runs, LastRunObjects, DataRead, SLAViolation,
         LastRunStatus, Bandwidth), plus a .txt dump of the raw OCR per image for manual review.

.PARAMETER WorkbookPath
    Path to the Backup Summary workbook containing the pasted screenshots.
.PARAMETER SiteSheetMap
    Hashtable mapping each sheet name in that workbook to this dashboard's canonical site name,
    e.g. @{ AQ='AquaArabia'; SF='SixFlags'; TABUK='SEVEN Tabuk'; ABHA='SEVEN ABHA' }.
.PARAMETER OutputPath
    Where to write the structured .xlsx (and the raw-OCR .txt dump next to it).

.EXAMPLE
    .\Backup_OCR_Extract.ps1 -WorkbookPath "C:\...\Backup Summary - September 2026.xlsx" `
        -SiteSheetMap @{ AQ='AquaArabia'; SF='SixFlags'; TABUK='SEVEN Tabuk'; ABHA='SEVEN ABHA' } `
        -OutputPath "C:\temp\Backup_Summary_2026-09.xlsx"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$WorkbookPath,
    [Parameter(Mandatory)][hashtable]$SiteSheetMap,
    [Parameter(Mandatory)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }

# ============================================================================
# 1. EXTRACT EMBEDDED IMAGES (.xlsx is a zip - xl/media/*.png) + map sheet -> its image(s)
# ============================================================================
Add-Type -AssemblyName System.IO.Compression.FileSystem
$workDir = Join-Path $env:TEMP ("BackupOcr_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null
$zip = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
try {
    foreach ($entry in $zip.Entries) {
        $dest = Join-Path $workDir $entry.FullName
        $destDir = Split-Path -Parent $dest
        if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
        if (-not $entry.FullName.EndsWith('/')) {
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
        }
    }
} finally {
    $zip.Dispose()
}

# Map sheetN.xml (by workbook.xml's <sheet> order / r:id) -> sheet name -> its drawing -> its image(s).
[xml]$workbookXml = Get-Content (Join-Path $workDir 'xl\workbook.xml') -Raw
[xml]$workbookRels = Get-Content (Join-Path $workDir 'xl\_rels\workbook.xml.rels') -Raw
$sheetIdToTarget = @{}
foreach ($rel in $workbookRels.Relationships.Relationship) {
    if ($rel.Type -like '*worksheet*') { $sheetIdToTarget[$rel.Id] = $rel.Target }
}

$sheetImages = @{}  # sheet display name -> list of image file paths, in sheet order
foreach ($sheet in $workbookXml.workbook.sheets.sheet) {
    $sheetName = $sheet.name
    if (-not $SiteSheetMap.ContainsKey($sheetName)) { continue }
    $sheetFile = $sheetIdToTarget[$sheet.id]
    $sheetXmlPath = Join-Path $workDir "xl\$sheetFile"
    $sheetRelsPath = Join-Path $workDir ("xl\worksheets\_rels\" + (Split-Path -Leaf $sheetFile) + ".rels")
    if (-not (Test-Path $sheetRelsPath)) { continue }
    [xml]$sheetRels = Get-Content $sheetRelsPath -Raw
    $drawingTarget = ($sheetRels.Relationships.Relationship | Where-Object { $_.Type -like '*drawing*' }).Target
    if (-not $drawingTarget) { continue }
    $drawingPath = Join-Path $workDir ("xl\" + ($drawingTarget -replace '^\.\./', ''))
    $drawingRelsPath = Join-Path $workDir ("xl\drawings\_rels\" + (Split-Path -Leaf $drawingTarget) + ".rels")
    [xml]$drawingRels = Get-Content $drawingRelsPath -Raw
    $images = @()
    foreach ($imgRel in $drawingRels.Relationships.Relationship) {
        if ($imgRel.Type -like '*image*') {
            $images += (Join-Path $workDir ("xl\media\" + (Split-Path -Leaf $imgRel.Target)))
        }
    }
    $sheetImages[$sheetName] = $images
}

# ============================================================================
# 2. OCR EACH IMAGE (Windows.Media.Ocr - on-device, no network)
# ============================================================================
[Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Foundation, ContentType=WindowsRuntime] | Out-Null

# PowerShell 5.1 can't natively 'await' WinRT async operations - this helper bridges that.
Add-Type -TypeDefinition @"
using System;
using System.Threading.Tasks;
public static class AsyncHelper {
    public static T Await<T>(object winrtTask) {
        var asTask = (Task<T>)typeof(System.WindowsRuntimeSystemExtensions)
            .GetMethod("AsTask", new Type[] { winrtTask.GetType().GetInterfaces()[0] })
            .MakeGenericMethod(typeof(T)).Invoke(null, new object[] { winrtTask });
        asTask.Wait();
        return asTask.Result;
    }
}
"@ -ReferencedAssemblies 'System.Runtime.WindowsRuntime' -ErrorAction SilentlyContinue

function Invoke-Ocr {
    param([string]$ImagePath)
    $file = [Windows.Storage.StorageFile]::GetFileFromPathAsync($ImagePath)
    $storageFile = [AsyncHelper]::Await[Windows.Storage.StorageFile]($file)
    $streamOp = $storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)
    $stream = [AsyncHelper]::Await[Windows.Storage.Streams.IRandomAccessStream]($streamOp)
    $decoderOp = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
    $decoder = [AsyncHelper]::Await[Windows.Graphics.Imaging.BitmapDecoder]($decoderOp)
    $bitmapOp = $decoder.GetSoftwareBitmapAsync()
    $bitmap = [AsyncHelper]::Await[Windows.Graphics.Imaging.SoftwareBitmap]($bitmapOp)
    $ocrEngine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if (-not $ocrEngine) { throw "No OCR language pack available on this machine - install one under Settings > Time & Language > Language > Add a language (English)." }
    $resultOp = $ocrEngine.RecognizeAsync($bitmap)
    $result = [AsyncHelper]::Await[Windows.Media.Ocr.OcrResult]($resultOp)
    return $result
}

# ============================================================================
# 3. RECONSTRUCT THE TABLE FROM OCR LINES (cluster words into columns by X position,
#    using the header row's word positions as the column boundaries)
# ============================================================================
$expectedHeaders = @('Protection Group', 'Source', 'Runs', 'Data Read', 'SLA Violation', 'Last Run Status', 'Bandwidth')

function ConvertFrom-OcrResult {
    param($OcrResult, [string]$SiteName, [System.Collections.Generic.List[object]]$RowsOut, [System.IO.StreamWriter]$RawDumpWriter)

    $RawDumpWriter.WriteLine("===== $SiteName =====")
    foreach ($line in $OcrResult.Lines) {
        $RawDumpWriter.WriteLine($line.Text)
    }
    $RawDumpWriter.WriteLine("")

    # Find the header line (the one containing "Protection" and "Group" words close together).
    $headerLineIndex = -1
    for ($i = 0; $i -lt $OcrResult.Lines.Count; $i++) {
        if ($OcrResult.Lines[$i].Text -match 'Protection\s*Group') { $headerLineIndex = $i; break }
    }
    if ($headerLineIndex -lt 0) {
        Write-Warning "$SiteName - could not find the 'Protection Group' header line in the OCR output. Check the raw dump."
        return
    }

    # Column boundaries = the X (left) position of each word in the header line.
    $headerWords = @($OcrResult.Lines[$headerLineIndex].Words)
    $colBoundaries = @($headerWords | ForEach-Object { $_.BoundingRect.X } | Sort-Object)

    for ($i = $headerLineIndex + 1; $i -lt $OcrResult.Lines.Count; $i++) {
        $line = $OcrResult.Lines[$i]
        if (-not $line.Words -or $line.Words.Count -eq 0) { continue }
        # Bucket every word on this line into the nearest column boundary.
        $buckets = New-Object object[] ($colBoundaries.Count)
        for ($b = 0; $b -lt $buckets.Count; $b++) { $buckets[$b] = [System.Collections.Generic.List[string]]::new() }
        foreach ($word in $line.Words) {
            $wx = $word.BoundingRect.X
            $nearest = 0; $bestDist = [double]::MaxValue
            for ($b = 0; $b -lt $colBoundaries.Count; $b++) {
                $d = [Math]::Abs($wx - $colBoundaries[$b])
                if ($d -lt $bestDist) { $bestDist = $d; $nearest = $b }
            }
            $buckets[$nearest].Add($word.Text)
        }
        $cells = @($buckets | ForEach-Object { ($_ -join ' ').Trim() })
        if (-not $cells[0]) { continue }  # first bucket = Protection Group name; skip blank/decoration lines

        $RowsOut.Add([pscustomobject]@{
            Site            = $SiteName
            ProtectionGroup = $cells[0]
            Source          = ''   # OCR can't reliably separate this from a wrapped PG name - fill in manually if needed
            Runs            = $cells[1]
            LastRunObjects  = ''
            DataRead        = $cells[2]
            SLAViolation    = $cells[3]
            LastRunStatus   = $cells[4]
            Bandwidth       = $cells[5]
        })
    }
}

$allRows = [System.Collections.Generic.List[object]]::new()
$rawDumpPath = [System.IO.Path]::ChangeExtension($OutputPath, '.ocr-raw.txt')
$rawWriter = New-Object System.IO.StreamWriter($rawDumpPath, $false)
try {
    foreach ($sheetName in $sheetImages.Keys) {
        $siteName = $SiteSheetMap[$sheetName]
        Write-Host "OCR: $sheetName -> $siteName ($($sheetImages[$sheetName].Count) image(s))" -ForegroundColor Cyan
        foreach ($imgPath in $sheetImages[$sheetName]) {
            $result = Invoke-Ocr -ImagePath $imgPath
            ConvertFrom-OcrResult -OcrResult $result -SiteName $siteName -RowsOut $allRows -RawDumpWriter $rawWriter
        }
    }
} finally {
    $rawWriter.Close()
}

# ============================================================================
# 4. WRITE THE STRUCTURED WORKBOOK (same shape BackupStorage_Weekly_HealthCheck.ps1 reads)
# ============================================================================
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$wb = $Excel.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
$ws.Name = 'Backup'
$headers = @('Site', 'ProtectionGroup', 'Source', 'Runs', 'LastRunObjects', 'DataRead', 'SLAViolation', 'LastRunStatus', 'Bandwidth')
for ($i = 0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i + 1) = $headers[$i] }

$r = 2
foreach ($row in $allRows) {
    $ws.Cells.Item($r, 1) = $row.Site
    $ws.Cells.Item($r, 2) = $row.ProtectionGroup
    $ws.Cells.Item($r, 3) = $row.Source
    $ws.Cells.Item($r, 4) = $row.Runs
    $ws.Cells.Item($r, 5) = $row.LastRunObjects
    $ws.Cells.Item($r, 6) = $row.DataRead
    $ws.Cells.Item($r, 7) = $row.SLAViolation
    $ws.Cells.Item($r, 8) = $row.LastRunStatus
    $ws.Cells.Item($r, 9) = $row.Bandwidth
    $r++
}
$wb.SaveAs($OutputPath, 51)
$wb.Close($false)
$Excel.Quit()

Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`nDone. $($allRows.Count) row(s) extracted." -ForegroundColor Green
Write-Host "Structured workbook: $OutputPath" -ForegroundColor Cyan
Write-Host "Raw OCR dump (compare against the real screenshot before trusting the above): $rawDumpPath" -ForegroundColor Yellow
Write-Host "`nIMPORTANT: 'Source' is left blank for every row - OCR can't reliably tell it apart from a" -ForegroundColor Yellow
Write-Host "wrapped Protection Group name. Check $OutputPath and fill that column in, or leave it blank" -ForegroundColor Yellow
Write-Host "(the dashboard still works without it, it just won't show the vCenter source per card)." -ForegroundColor Yellow
