#Requires -Version 5.1
<#
.SYNOPSIS
    Discovers the real structure of the 4 source Excel workbooks (IDF Room Reports + DC
    RackWise Health Checks for AquaArabia and SixFlags) so a matching collection + HTML
    dashboard script - built the same way as VMware_Weekly_HealthCheck.ps1 - can be designed
    against their ACTUAL layout, without the real company data ever needing to leave this
    machine.

.DESCRIPTION
    Run this ON THE MACHINE THAT HAS THE REAL FILES (and Excel, or the ImportExcel module).
    It never uploads anything anywhere - it only writes a report to -OutputPath. Share THAT
    report (not the workbooks themselves) back so the real collection script can be built
    against the actual tab names, column headers, and value shapes.

    For each of the 4 workbooks, opens EVERY worksheet/tab and records, for every tab:
      - A RAW PREVIEW of the first -RawPreviewRows rows - every cell, by column letter,
        completely unprocessed (no header-row guessing). This is what settles ambiguous
        layouts (merged title bands, a header split across two rows, etc.) that an
        auto-detected "header row" can get wrong.
      - An auto-detected header row (first row with 2+ non-empty cells) and up to
        -SampleRows sample data rows under it - a quick best-guess view, useful for the many
        tabs that really are one flat table.
      - For any tab whose name is listed in -FullDumpTabs (e.g. one representative rack tab,
        the Summary/Dashboard tabs), ALL of its data rows instead of just -SampleRows, up to
        a 300-row safety cap - needed for tabs where the real per-item detail starts well
        past row 3 (e.g. a per-rack checklist below a Metric/Value block).

    Real values ARE included by default, since column names alone are often not enough to
    design a correct collection script (e.g. "Status" - is it text like "OK"/"Fail", or a raw
    sensor reading?). If some columns hold sensitive detail you'd rather not send even as a
    sample (e.g. exact IP addresses, serial numbers, asset tags), list their header names in
    -RedactColumns and their values are replaced with a type-only placeholder ("<text>",
    "<number>", "<date>", "<empty>") instead - this applies to the raw preview and full-dump
    sections too, matched by whichever header/column ends up above that cell.

    Prefers Excel itself via COM automation (same approach this project already uses for Word
    automation - needs Microsoft Excel installed, nothing extra to download). Falls back to
    the ImportExcel PowerShell module if Excel isn't available on this machine - install it
    first with: Install-Module ImportExcel -Scope CurrentUser

.PARAMETER AquaArabiaIdfPath
    Path to the AquaArabia_IDF_Room_Report workbook.
.PARAMETER AquaArabiaRackPath
    Path to the AquaArabi_DC_RackWise_HealthCheck_V1 workbook.
.PARAMETER SixFlagsIdfPath
    Path to the SixFlags_IDF_Room_Report workbook.
.PARAMETER SixFlagsRackPath
    Path to the SixFlags_DC RackWise_HealthCheck_V1 workbook.
.PARAMETER RawPreviewRows
    How many rows (from row 1) to dump completely raw, cell by cell, for EVERY tab (default
    6). This is the section that reveals true multi-row/merged header layouts.
.PARAMETER SampleRows
    How many sample data rows to capture under the auto-detected header, for tabs NOT listed
    in -FullDumpTabs (default 3).
.PARAMETER FullDumpTabs
    Tab names (case-insensitive, matched on any of the 4 workbooks) to capture completely -
    every data row, not just -SampleRows - up to a 300-row safety cap.
.PARAMETER RedactColumns
    Column/header names (matched case-insensitively) whose values are replaced with a type
    placeholder instead of the real value, everywhere in the report.
.PARAMETER OutputPath
    Folder to write the discovery report into (default: a Discovery_Report subfolder next to
    this script).

.EXAMPLE
    .\Infrastructure_Report_Discovery.ps1 `
        -AquaArabiaIdfPath  'C:\Reports\AquaArabia_IDF_Room_Report.xlsx' `
        -AquaArabiaRackPath 'C:\Reports\AquaArabi_DC_RackWise_HealthCheck_V1.xlsx' `
        -SixFlagsIdfPath    'C:\Reports\SixFlags_IDF_Room_Report.xlsx' `
        -SixFlagsRackPath   'C:\Reports\SixFlags_DC RackWise_HealthCheck_V1.xlsx' `
        -RawPreviewRows 8 `
        -FullDumpTabs 'Zone - 2','Summary','Dashboard','Lists','A01','B01','P1','S1' `
        -RedactColumns 'IP Address','Serial Number','Asset Tag'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AquaArabiaIdfPath,
    [Parameter(Mandatory)][string]$AquaArabiaRackPath,
    [Parameter(Mandatory)][string]$SixFlagsIdfPath,
    [Parameter(Mandatory)][string]$SixFlagsRackPath,

    [int]$RawPreviewRows = 6,
    [int]$SampleRows = 3,
    [string[]]$FullDumpTabs = @(),
    [string[]]$RedactColumns = @(),
    [string]$OutputPath = (Join-Path $PSScriptRoot 'Discovery_Report')
)

$ErrorActionPreference = 'Stop'
$FullDumpCap = 300
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

$Workbooks = @(
    [pscustomobject]@{ Label = 'AquaArabia_IDF_Room_Report';              Path = $AquaArabiaIdfPath }
    [pscustomobject]@{ Label = 'AquaArabi_DC_RackWise_HealthCheck_V1';    Path = $AquaArabiaRackPath }
    [pscustomobject]@{ Label = 'SixFlags_IDF_Room_Report';                Path = $SixFlagsIdfPath }
    [pscustomobject]@{ Label = 'SixFlags_DC RackWise_HealthCheck_V1';     Path = $SixFlagsRackPath }
)

foreach ($wbInfo in $Workbooks) {
    if (-not (Test-Path $wbInfo.Path)) {
        throw "Cannot find '$($wbInfo.Label)' at path: $($wbInfo.Path). Fix the path and re-run."
    }
}

# Classifies a raw cell value into a type-only placeholder, used both for -RedactColumns and
# as a fallback label whenever a value is empty - never leaks the actual value in either case.
function Get-ValuePlaceholder {
    param($Value)
    if ($null -eq $Value -or $Value -eq '') { return '<empty>' }
    if ($Value -is [datetime]) { return '<date>' }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [decimal]) { return '<number>' }
    return '<text>'
}

function Format-SampleValue {
    param($Value, [string]$Header, [string[]]$RedactColumns)
    if ($RedactColumns -and $Header -and ($RedactColumns | Where-Object { $_ -eq $Header })) {
        return Get-ValuePlaceholder $Value
    }
    if ($null -eq $Value -or $Value -eq '') { return '' }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    return [string]$Value
}

# A1-style column letters (1 -> A, 27 -> AA, ...) for the raw preview section, so cells that
# have no header text yet (or sit above the real header row) are still identifiable.
function Get-ColumnLetter {
    param([int]$Index)
    $letter = ''
    while ($Index -gt 0) {
        $rem = ($Index - 1) % 26
        $letter = [char](65 + $rem) + $letter
        $Index = [Math]::Floor(($Index - 1) / 26)
    }
    return $letter
}

# ============================================================================
# PRIMARY PATH: Excel via COM automation (same pattern this project already uses for Word)
# ============================================================================
function Read-WorkbookViaCom {
    param($Excel, [string]$Path, [string[]]$FullDumpTabs, [int]$RawPreviewRows, [int]$SampleRows, [string[]]$RedactColumns, [int]$FullDumpCap)
    $wb = $Excel.Workbooks.Open($Path, 0, $true)   # 0 = don't update links, $true = read-only
    try {
        $sheetsOut = @()
        foreach ($ws in $wb.Worksheets) {
            $used = $ws.UsedRange
            $rowCount = $used.Rows.Count
            $colCount = $used.Columns.Count
            $isFullDump = [bool]($FullDumpTabs | Where-Object { $_ -eq $ws.Name })

            # --- Best-guess header row: first of the first 10 rows with 2+ non-empty cells ---
            # (computed BEFORE the raw preview below, purely so -RedactColumns has a per-column
            # header name to match against there too - the raw preview always shows every cell
            # regardless of whether this guess turns out right.)
            $scanRows = [Math]::Min(10, $rowCount)
            $scanCols = [Math]::Min(30, $colCount)
            $headerRowIdx = 1
            for ($r = 1; $r -le $scanRows; $r++) {
                $nonEmpty = 0
                for ($c = 1; $c -le $scanCols; $c++) {
                    $v = $used.Cells.Item($r, $c).Value2
                    if ($null -ne $v -and "$v" -ne '') { $nonEmpty++ }
                }
                if ($nonEmpty -ge 2) { $headerRowIdx = $r; break }
            }

            $headers = @()
            for ($c = 1; $c -le $colCount; $c++) {
                $h = $used.Cells.Item($headerRowIdx, $c).Value2
                $headers += $(if ($null -ne $h -and "$h" -ne '') { "$h" } else { "Column$c" })
            }

            # --- Raw preview: first N rows, every cell, unprocessed - labeled by column letter
            # AND redacted against the best-guess header for that column (may be wrong for a
            # genuinely offset/split header, but still the best information available here). ---
            $previewRowCount = [Math]::Min($RawPreviewRows, $rowCount)
            $rawPreview = @()
            for ($r = 1; $r -le $previewRowCount; $r++) {
                $cells = @()
                for ($c = 1; $c -le $colCount; $c++) {
                    $raw = $used.Cells.Item($r, $c).Value2
                    $colLetter = Get-ColumnLetter $c
                    $cells += "$colLetter=$(Format-SampleValue -Value $raw -Header $headers[$c - 1] -RedactColumns $RedactColumns)"
                }
                $rawPreview += "Row $($r): " + ($cells -join '  |  ')
            }

            $rowsToCapture = if ($isFullDump) { $FullDumpCap } else { $SampleRows }
            $sampleRowsOut = @()
            $dataRowStart = $headerRowIdx + 1
            $captured = 0
            for ($r = $dataRowStart; $r -le $rowCount -and $captured -lt $rowsToCapture; $r++) {
                $rowHasData = $false
                $rowObj = [ordered]@{}
                for ($c = 1; $c -le $colCount; $c++) {
                    $raw = $used.Cells.Item($r, $c).Value2
                    if ($null -ne $raw -and "$raw" -ne '') { $rowHasData = $true }
                    $rowObj[$headers[$c - 1]] = Format-SampleValue -Value $raw -Header $headers[$c - 1] -RedactColumns $RedactColumns
                }
                if ($rowHasData) {
                    $sampleRowsOut += [pscustomobject]$rowObj
                    $captured++
                }
            }

            $sheetsOut += [pscustomobject]@{
                TabName       = $ws.Name
                IsFullDump    = $isFullDump
                RawPreview    = $rawPreview
                HeaderRow     = $headerRowIdx
                DataRowCount  = [Math]::Max(0, $rowCount - $headerRowIdx)
                ColumnCount   = $colCount
                Headers       = $headers
                SampleRows    = $sampleRowsOut
            }
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($used) | Out-Null
        }
        return $sheetsOut
    } finally {
        $wb.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
}

# ============================================================================
# FALLBACK PATH: ImportExcel module (no Excel installation required)
# ============================================================================
function Read-WorkbookViaImportExcel {
    param([string]$Path, [string[]]$FullDumpTabs, [int]$RawPreviewRows, [int]$SampleRows, [string[]]$RedactColumns, [int]$FullDumpCap)
    $sheetNames = (Get-ExcelSheetInfo -Path $Path) | Select-Object -ExpandProperty Name
    $sheetsOut = @()
    foreach ($name in $sheetNames) {
        $isFullDump = [bool]($FullDumpTabs | Where-Object { $_ -eq $name })

        # ImportExcel assumes row 1 is the header by default - fine as a best-guess view (and
        # gives -RedactColumns a per-column name to match against in the raw preview below);
        # the raw preview is what actually settles an ambiguous/offset header layout.
        $rows = @(Import-Excel -Path $Path -WorksheetName $name -ErrorAction SilentlyContinue)
        $headers = if ($rows.Count -gt 0) { $rows[0].PSObject.Properties.Name } else { @() }

        # Raw preview via -NoHeader so every cell (including a would-be header row) is visible.
        $rawRows = @(Import-Excel -Path $Path -WorksheetName $name -NoHeader -ErrorAction SilentlyContinue)
        $previewCount = [Math]::Min($RawPreviewRows, $rawRows.Count)
        $rawPreview = @()
        for ($i = 0; $i -lt $previewCount; $i++) {
            $props = @($rawRows[$i].PSObject.Properties)
            $cells = for ($p = 0; $p -lt $props.Count; $p++) {
                $colLetter = Get-ColumnLetter ($p + 1)
                $headerForCol = if ($p -lt $headers.Count) { $headers[$p] } else { $null }
                "$colLetter=$(Format-SampleValue -Value $props[$p].Value -Header $headerForCol -RedactColumns $RedactColumns)"
            }
            $rawPreview += "Row $($i + 1): " + ($cells -join '  |  ')
        }

        $rowsToCapture = if ($isFullDump) { $FullDumpCap } else { $SampleRows }
        $sampleRowsOut = @()
        for ($i = 0; $i -lt [Math]::Min($rowsToCapture, $rows.Count); $i++) {
            $rowObj = [ordered]@{}
            foreach ($h in $headers) {
                $rowObj[$h] = Format-SampleValue -Value $rows[$i].$h -Header $h -RedactColumns $RedactColumns
            }
            $sampleRowsOut += [pscustomobject]$rowObj
        }
        $sheetsOut += [pscustomobject]@{
            TabName       = $name
            IsFullDump    = $isFullDump
            RawPreview    = $rawPreview
            HeaderRow     = 1
            DataRowCount  = $rows.Count
            ColumnCount   = $headers.Count
            Headers       = $headers
            SampleRows    = $sampleRowsOut
        }
    }
    return $sheetsOut
}

# ============================================================================
# RUN
# ============================================================================
$UseCom = $false
$Excel = $null
try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $UseCom = $true
    Write-Host "Using Excel COM automation." -ForegroundColor Cyan
} catch {
    Write-Host "Excel COM automation not available ($($_.Exception.Message))." -ForegroundColor Yellow
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "Neither Excel nor the ImportExcel module is available on this machine. Install ImportExcel first: Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "Using the ImportExcel module instead." -ForegroundColor Cyan
}

$Results = @()
foreach ($wbInfo in $Workbooks) {
    Write-Host "Reading: $($wbInfo.Label)  ($($wbInfo.Path))" -ForegroundColor Green
    $sheets = if ($UseCom) {
        Read-WorkbookViaCom -Excel $Excel -Path $wbInfo.Path -FullDumpTabs $FullDumpTabs -RawPreviewRows $RawPreviewRows -SampleRows $SampleRows -RedactColumns $RedactColumns -FullDumpCap $FullDumpCap
    } else {
        Read-WorkbookViaImportExcel -Path $wbInfo.Path -FullDumpTabs $FullDumpTabs -RawPreviewRows $RawPreviewRows -SampleRows $SampleRows -RedactColumns $RedactColumns -FullDumpCap $FullDumpCap
    }
    $Results += [pscustomobject]@{
        Workbook = $wbInfo.Label
        Path     = $wbInfo.Path
        Tabs     = $sheets
    }
}

if ($UseCom) {
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

# ============================================================================
# WRITE REPORT (text - easy to read/paste back - and JSON - easy to re-parse)
# ============================================================================
$stamp = (Get-Date).ToString('yyyy-MM-dd_HHmm')
$txtPath  = Join-Path $OutputPath "Discovery_Report_$stamp.txt"
$jsonPath = Join-Path $OutputPath "Discovery_Report_$stamp.json"

$lines = @()
$lines += "Infrastructure Report Discovery - $((Get-Date).ToString('u'))"
$lines += "=" * 78
foreach ($wb in $Results) {
    $lines += ""
    $lines += "WORKBOOK: $($wb.Workbook)"
    $lines += "  Path: $($wb.Path)"
    $lines += "  Tabs found: $($wb.Tabs.Count)"
    foreach ($tab in $wb.Tabs) {
        $lines += ""
        $lines += "  --- Tab: '$($tab.TabName)' $(if ($tab.IsFullDump) { '[FULL DUMP]' }) ---"
        $lines += "    Raw preview (first $($tab.RawPreview.Count) rows, unprocessed):"
        foreach ($rp in $tab.RawPreview) { $lines += "      $rp" }
        $lines += "    Best-guess header row: $($tab.HeaderRow)  |  Data rows: $($tab.DataRowCount)  |  Columns: $($tab.ColumnCount)"
        $lines += "    Best-guess headers: $($tab.Headers -join ' | ')"
        if ($tab.SampleRows.Count -gt 0) {
            $label = if ($tab.IsFullDump) { "All captured rows ($($tab.SampleRows.Count)):" } else { "Sample rows:" }
            $lines += "    $label"
            $rowNum = 1
            foreach ($row in $tab.SampleRows) {
                $lines += "      [$rowNum] " + (($row.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '  |  ')
                $rowNum++
            }
        } else {
            $lines += "    (no data rows captured under the best-guess header)"
        }
    }
    $lines += ""
    $lines += ("-" * 78)
}
$lines | Out-File -FilePath $txtPath -Encoding UTF8
$Results | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8

Write-Host ""
Write-Host "Discovery report written:" -ForegroundColor Cyan
Write-Host "  $txtPath" -ForegroundColor Cyan
Write-Host "  $jsonPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next step: share the .txt file's contents back (paste it in, or attach the file) so the real collection + dashboard script can be built against this exact tab/column structure." -ForegroundColor Green
