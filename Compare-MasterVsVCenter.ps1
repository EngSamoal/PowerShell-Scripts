<#
.SYNOPSIS
    Compares the VM names in two Excel workbooks and writes a new workbook listing the names
    that exist in one file but are missing from the other.

.DESCRIPTION
    READ-ONLY on the two input files. Defaults:

        File 1 (Master)   C:\temp\Ayoub\Master_draft2.xlsx   tabs: Master, OT, Test, Template
        File 2 (vCenter)  C:\temp\Ayoub\vCenter_Full.xlsx    tabs: All, Template
        Column            "VM Name" in every tab

    The header row does not have to be row 1 - the first 20 rows of each tab are searched for
    the "VM Name" header (so title banners above the table are fine). If a tab has no
    "VM Name" header but has a "Name" header, that is used instead and a warning is shown.

    Names are compared trimmed and case-insensitive ("WEB01 " = "web01").

    Output workbook (default C:\temp\Ayoub\Compare_Result_<timestamp>.xlsx):
        Summary                - counts, plus any tab/column that could not be read and why
        Missing in vCenter     - in Master file, NOT in vCenter file (with the Master tab it came from)
        Missing in Master      - in vCenter file, NOT in Master file (with the vCenter tab it came from)

.PARAMETER MasterFile / VCenterFile
    Paths of the two workbooks. If no extension is given, .xlsx / .xlsm / .xls is tried.

.EXAMPLE
    .\Compare-MasterVsVCenter.ps1
.EXAMPLE
    .\Compare-MasterVsVCenter.ps1 -MasterFile D:\x\Master.xlsx -VCenterFile D:\x\vc.xlsx

.NOTES
    Requires Microsoft Excel (desktop) on this machine - it is used to read and write the files.
#>
[CmdletBinding()]
param(
    [string]  $MasterFile   = 'C:\temp\Ayoub\Master_draft2.xlsx',
    [string[]]$MasterTabs   = @('Master', 'OT', 'Test', 'Template'),
    [string]  $VCenterFile  = 'C:\temp\Ayoub\vCenter_Full.xlsx',
    [string[]]$VCenterTabs  = @('All', 'Template'),
    [string]  $ColumnName   = 'VM Name',
    [string]  $OutputFile   = ('C:\temp\Ayoub\Compare_Result_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmm'))
)

$ErrorActionPreference = 'Stop'
$ScriptBuild = 'Compare-MasterVsVCenter build 2026-10-10.1'
Write-Host $ScriptBuild -ForegroundColor Cyan

#region ---- Pure helpers (no Excel needed) ----

# Accepts a path with or without extension and returns the real file path, or throws.
function Resolve-Workbook([string]$Path) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) { return (Resolve-Path -LiteralPath $Path).Path }
    $base = [IO.Path]::ChangeExtension($Path, $null).TrimEnd('.')
    foreach ($ext in '.xlsx', '.xlsm', '.xls') {
        if (Test-Path -LiteralPath ($base + $ext) -PathType Leaf) { return (Resolve-Path -LiteralPath ($base + $ext)).Path }
    }
    throw "File not found: $Path (also tried .xlsx / .xlsm / .xls)"
}

# Finds the header cell in the first 20 rows of a 1-based 2D grid.
# Returns @{Row; Col; Header} or $null.
function Find-HeaderCell($Grid, [string]$Wanted) {
    $rows = $Grid.GetUpperBound(0); $cols = $Grid.GetUpperBound(1)
    foreach ($name in @($Wanted, 'Name')) {
        for ($r = 1; $r -le [Math]::Min($rows, 20); $r++) {
            for ($c = 1; $c -le $cols; $c++) {
                $v = "$($Grid[$r, $c])".Trim()
                if ($v -and $v -ieq $name) { return @{ Row = $r; Col = $c; Header = $v } }
            }
        }
    }
    return $null
}

# Returns the non-empty values under the header cell.
function Get-ColumnValues($Grid, [hashtable]$Header) {
    $out = New-Object System.Collections.Generic.List[string]
    for ($r = $Header.Row + 1; $r -le $Grid.GetUpperBound(0); $r++) {
        $v = "$($Grid[$r, $Header.Col])".Trim()
        if ($v) { $out.Add($v) }
    }
    , $out.ToArray()
}

# Turns a list of @{Name; Tab} into a case-insensitive map: lowercase name -> @{Name; Tabs}
function New-NameMap($Items) {
    $map = [ordered]@{}
    foreach ($i in $Items) {
        $key = $i.Name.ToLowerInvariant()
        if (-not $map.Contains($key)) { $map[$key] = @{ Name = $i.Name; Tabs = New-Object System.Collections.Generic.List[string] } }
        if (-not $map[$key].Tabs.Contains($i.Tab)) { $map[$key].Tabs.Add($i.Tab) }
    }
    $map
}
#endregion

#region ---- Excel read ----
function Read-WorkbookNames($Excel, [string]$Path, [string[]]$Tabs, [string]$Label, $Problems) {
    $items = New-Object System.Collections.Generic.List[object]
    $wb = $Excel.Workbooks.Open($Path, 0, $true)   # read-only
    try {
        $sheetNames = @($wb.Worksheets | ForEach-Object { $_.Name })
        foreach ($tab in $Tabs) {
            $real = $sheetNames | Where-Object { $_.Trim() -ieq $tab.Trim() } | Select-Object -First 1
            if (-not $real) {
                $Problems.Add("$Label : tab '$tab' not found. Tabs in file: $($sheetNames -join ', ')")
                Write-Warning "$Label : tab '$tab' not found"
                continue
            }
            $grid = $wb.Worksheets.Item($real).UsedRange.Value2
            if ($grid -isnot [array]) {
                $Problems.Add("$Label : tab '$real' is empty or has a single cell")
                continue
            }
            $hdr = Find-HeaderCell $grid $ColumnName
            if (-not $hdr) {
                $Problems.Add("$Label : tab '$real' has no '$ColumnName' (or 'Name') header in the first 20 rows")
                Write-Warning "$Label : tab '$real' has no '$ColumnName' column"
                continue
            }
            if ($hdr.Header -ine $ColumnName) {
                $Problems.Add("$Label : tab '$real' has no '$ColumnName' column - used '$($hdr.Header)' column instead")
                Write-Warning "$Label : tab '$real' - used '$($hdr.Header)' column instead of '$ColumnName'"
            }
            $vals = Get-ColumnValues $grid $hdr
            foreach ($v in $vals) { $items.Add([pscustomobject]@{ Name = $v; Tab = $real }) }
            Write-Host ("  {0,-8} {1,-12} {2,6} names" -f $Label, $real, $vals.Count)
        }
    } finally {
        $wb.Close($false)
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
    }
    , $items.ToArray()
}
#endregion

#region ---- Excel write ----
function Write-Sheet($Workbook, [string]$Name, [object[]]$Rows, [string[]]$Columns) {
    $ws = $Workbook.Worksheets.Add([Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
    $ws.Name = $Name
    $data = [object[,]]::new(($Rows.Count + 1), $Columns.Count)
    for ($c = 0; $c -lt $Columns.Count; $c++) { $data[0, $c] = $Columns[$c] }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $Columns.Count; $c++) { $data[($r + 1), $c] = "$($Rows[$r].($Columns[$c]))" }
    }
    $range = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($Rows.Count + 1, $Columns.Count))
    $range.NumberFormat = '@'          # text, so names like 0012 or 1E10 stay as typed
    $range.Value2 = $data
    $ws.Rows.Item(1).Font.Bold = $true
    $ws.Rows.Item(1).Interior.Color = 0xEBD9C6   # light blue (BGR)
    if ($Rows.Count -gt 0) { [void]$range.AutoFilter() }
    [void]$ws.UsedRange.Columns.AutoFit()
}
#endregion

#region ---- Main ----
$MasterFile  = Resolve-Workbook $MasterFile
$VCenterFile = Resolve-Workbook $VCenterFile
$outDir = Split-Path $OutputFile -Parent
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$problems = New-Object System.Collections.Generic.List[string]
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    Write-Host "Reading $MasterFile"
    $masterItems  = Read-WorkbookNames $excel $MasterFile  $MasterTabs  'Master'  $problems
    Write-Host "Reading $VCenterFile"
    $vcenterItems = Read-WorkbookNames $excel $VCenterFile $VCenterTabs 'vCenter' $problems

    $masterMap  = New-NameMap $masterItems
    $vcenterMap = New-NameMap $vcenterItems

    $missingInVc = @(@(foreach ($k in $masterMap.Keys)  { if (-not $vcenterMap.Contains($k)) {
        [pscustomobject]@{ 'VM Name' = $masterMap[$k].Name;  'Found in Master tab(s)'  = ($masterMap[$k].Tabs -join ', ') } } }) |
        Sort-Object 'VM Name')
    $missingInMaster = @(@(foreach ($k in $vcenterMap.Keys) { if (-not $masterMap.Contains($k)) {
        [pscustomobject]@{ 'VM Name' = $vcenterMap[$k].Name; 'Found in vCenter tab(s)' = ($vcenterMap[$k].Tabs -join ', ') } } }) |
        Sort-Object 'VM Name')
    $inBoth = @($masterMap.Keys | Where-Object { $vcenterMap.Contains($_) }).Count

    $summary = New-Object System.Collections.Generic.List[object]
    $summary.Add([pscustomobject]@{ Item = 'Script build';                         Value = $ScriptBuild })
    $summary.Add([pscustomobject]@{ Item = 'Master file';                          Value = "$MasterFile  (tabs: $($MasterTabs -join ', '))" })
    $summary.Add([pscustomobject]@{ Item = 'vCenter file';                         Value = "$VCenterFile  (tabs: $($VCenterTabs -join ', '))" })
    $summary.Add([pscustomobject]@{ Item = 'Unique VM names in Master';            Value = $masterMap.Count })
    $summary.Add([pscustomobject]@{ Item = 'Unique VM names in vCenter';           Value = $vcenterMap.Count })
    $summary.Add([pscustomobject]@{ Item = 'In both files';                        Value = $inBoth })
    $summary.Add([pscustomobject]@{ Item = 'In Master, missing in vCenter';        Value = @($missingInVc).Count })
    $summary.Add([pscustomobject]@{ Item = 'In vCenter, missing in Master';        Value = @($missingInMaster).Count })
    foreach ($p in $problems) { $summary.Add([pscustomobject]@{ Item = 'WARNING'; Value = $p }) }

    $out = $excel.Workbooks.Add()
    while ($out.Worksheets.Count -gt 1) { $out.Worksheets.Item($out.Worksheets.Count).Delete() }
    $out.Worksheets.Item(1).Name = 'tmp'
    Write-Sheet $out 'Summary'            $summary.ToArray()       @('Item', 'Value')
    Write-Sheet $out 'Missing in vCenter' @($missingInVc)     @('VM Name', 'Found in Master tab(s)')
    Write-Sheet $out 'Missing in Master'  @($missingInMaster) @('VM Name', 'Found in vCenter tab(s)')
    $out.Worksheets.Item('tmp').Delete()
    $out.Worksheets.Item('Summary').Activate()
    $out.SaveAs($OutputFile, 51)   # 51 = .xlsx
    $out.Close($false)
} finally {
    $excel.Quit()
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

Write-Host ''
$summary | Format-Table -AutoSize -Wrap | Out-Host
Write-Host "Result saved: $OutputFile" -ForegroundColor Green
#endregion
