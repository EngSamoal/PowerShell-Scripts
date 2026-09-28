# Add-SiteColumn.ps1
# Adds a "Site" column just before "Business Owner" in the consolidated master sheet and fills it from the
# VM name prefix (SF -> SixFlags, AQ -> AquaArabia ...). The title row becomes "All Sites".
# Other contents and colours are not changed.
# Safe: the input file is only read; the result is saved as a NEW file. Needs Microsoft Excel.
# Run:  .\Add-SiteColumn.ps1
param(
    [string]$InputFile = 'C:\temp\Master.xlsx',            # <-- consolidated master
    [string]$Output    = 'C:\temp\Master_withSite.xlsx'    # <-- result (new file)
)

#region ======================= SITE RULES (edit here) =======================
# VM name starts with  ->  Site
$SiteByPrefix = [ordered]@{
    'SF' = 'SixFlags'
    'AQ' = 'AquaArabia'
    'TB' = 'SEVEN Tabuk'
    'AB' = 'SEVEN ABHA'
    'HA' = 'SEVEN Alhamra'
}
$Title = 'All Sites  -  VM Inventory'      # new text for the title row (row above the header)
#endregion ==================================================================

$ErrorActionPreference = 'Stop'
function Test-JunkHeader($s) { ([string]$s).Trim() -match '^(Column\s*)?\d*$' }   # blank, "2", "Column8" = auto-named empty column
function Norm($s) { (([string]$s) -replace '[\s\u00A0]+', ' ').Trim().ToLower() }

if (-not (Test-Path $InputFile)) { Write-Host "Not found: $InputFile" -ForegroundColor Red; return }
$in = (Resolve-Path $InputFile).Path; $out = [IO.Path]::GetFullPath($Output)
if ($in -eq $out) { Write-Host 'Output must be a different file than the input.' -ForegroundColor Red; return }

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($in)

    # ---------- Find the sheet + header row that has "VM NAME" ----------
    $excel.ScreenUpdating = $false
    $excel.Calculation = -4135                                   # manual while working
    $ws = $null; $hdr = 0; $nameCol = 0
    foreach ($s in $wb.Worksheets) {
        # Excel's own search (fast) - exact "VM NAME" in the top 20 rows
        $hit = $s.Range('A1:ZZ20').Find('VM NAME', [Type]::Missing, -4163, 1)   # values, whole cell
        if ($hit) { $ws = $s; $hdr = $hit.Row; $nameCol = $hit.Column; break }
    }
    if (-not $ws) { throw "No 'VM NAME' header found in $InputFile" }
    # last real header (ignores blank / spaces / auto-named "Column8" headers)
    $hv = $ws.Range($ws.Cells.Item($hdr, 1), $ws.Cells.Item($hdr, $ws.Columns.Count)).Value2
    $lastCol = $nameCol
    for ($c = $hv.GetUpperBound(1); $c -gt $nameCol; $c--) { if (-not (Test-JunkHeader $hv[1, $c])) { $lastCol = $c; break } }
    Write-Host "Sheet '$($ws.Name)', header row $hdr, VM NAME in column $nameCol"

    # ---------- Site column: reuse if it exists, otherwise add it just before "Business Owner" ----------
    $siteCol = 0; $pos = 0
    $hdrVals = $ws.Range($ws.Cells.Item($hdr, 1), $ws.Cells.Item($hdr, $lastCol)).Value2
    for ($c = 1; $c -le $lastCol; $c++) {
        if (-not $siteCol -and (Norm $hdrVals[1, $c]) -eq 'site') { $siteCol = $c }
        if (-not $pos -and (Norm $hdrVals[1, $c]) -eq 'business owner') { $pos = $c }
    }
    if (-not $pos) { $pos = 2; Write-Host "No 'Business Owner' header found - Site goes in column 2" -ForegroundColor Yellow }
    if ($siteCol) { Write-Host "A 'Site' column already exists (column $siteCol) - it will be filled" -ForegroundColor Yellow }
    else {
        # Move the columns from "Business Owner" up to the last real column one step to the right.
        # The receiving column is the first column after the data that is empty - cells with only spaces and
        # auto-named table headers ("Column8", "Column9" ...) count as empty. Nothing else is touched.
        $lo = $null
        foreach ($t in $ws.ListObjects) { if ($t.Range.Row -le $hdr -and ($t.Range.Row + $t.Range.Rows.Count - 1) -ge $hdr) { $lo = $t } }
        $blockLastRow = $ws.Cells.Item($ws.Rows.Count, $nameCol).End(-4162).Row
        if ($lo) { $blockLastRow = [Math]::Max($blockLastRow, $lo.Range.Row + $lo.Range.Rows.Count - 1) }
        for ($c = $pos; $c -le $lastCol; $c++) { $blockLastRow = [Math]::Max($blockLastRow, $ws.Cells.Item($ws.Rows.Count, $c).End(-4162).Row) }

        $scanTo = [Math]::Min($ws.Columns.Count, $lastCol + 400)
        $grid = $ws.Range($ws.Cells.Item(1, $lastCol + 1), $ws.Cells.Item($blockLastRow, $scanTo)).Value2
        $recv = 0
        for ($k = 1; $k -le $grid.GetUpperBound(1) -and -not $recv; $k++) {
            $empty = $true
            for ($r = 1; $r -le $grid.GetUpperBound(0); $r++) {
                $v = ([string]$grid[$r, $k]).Trim()
                if ($v -eq '') { continue }
                if ($r -eq $hdr -and (Test-JunkHeader $v)) { continue }         # auto-named empty table column
                $empty = $false; break
            }
            if ($empty) { $recv = $lastCol + $k }
        }
        if (-not $recv) { throw "No free column found to the right of the data (checked up to column $scanTo)." }
        if ($recv -gt $lastCol + 1) { Write-Host "Columns between the data and column $recv hold values - they move one step right too" -ForegroundColor Yellow }

        $tblLast = $(if ($lo) { $lo.Range.Column + $lo.Range.Columns.Count - 1 } else { 0 })
        if ($lo) { $loRow = $lo.Range.Row; $loCol = $lo.Range.Column; $loLastRow = $loRow + $lo.Range.Rows.Count - 1 }
        [void]$ws.Range($ws.Cells.Item(1, $pos), $ws.Cells.Item($blockLastRow, $recv - 1)).Cut($ws.Cells.Item(1, $pos + 1))
        # table: keep its size, or grow it just enough to still hold all moved columns
        if ($lo -and $recv -gt $tblLast) { $lo.Resize($ws.Range($ws.Cells.Item($loRow, $loCol), $ws.Cells.Item($loLastRow, $recv))) }
        $blockLastCol = $recv - 1
        # New column gets the same look as the "Business Owner" column next to it
        [void]$ws.Range($ws.Cells.Item(1, $pos + 1), $ws.Cells.Item($blockLastRow, $pos + 1)).Copy()
        [void]$ws.Range($ws.Cells.Item(1, $pos), $ws.Cells.Item($blockLastRow, $pos)).PasteSpecial(-4122)   # formats only
        $ws.Cells.Item($hdr, $pos).Value2 = 'Site'
        $ws.Columns.Item($pos).ColumnWidth = 18
        $siteCol = $pos
        if ($nameCol -ge $pos) { $nameCol++ }
        $lastCol++
    }
    # ---------- Fill Site from the VM name ----------
    $lastRow = $ws.Cells.Item($ws.Rows.Count, $nameCol).End(-4162).Row       # last filled VM NAME cell
    $names = $ws.Range($ws.Cells.Item($hdr, $nameCol), $ws.Cells.Item([Math]::Max($lastRow, $hdr + 1), $nameCol)).Value2   # read all names at once
    Write-Host "Rows to check: $($lastRow - $hdr)"
    $count = [ordered]@{}; foreach ($v in $SiteByPrefix.Values) { $count[$v] = 0 }
    $noMatch = New-Object System.Collections.Generic.List[string]
    for ($r = $hdr + 1; $r -le $lastRow; $r++) {
        $vm = ([string]$names[($r - $hdr + 1), 1]).Trim()
        if (-not $vm -or (Norm $vm) -eq 'vm name') { continue }          # empty row or a repeated header
        $site = $null
        foreach ($p in $SiteByPrefix.Keys) { if ($vm.StartsWith($p, [StringComparison]::OrdinalIgnoreCase)) { $site = $SiteByPrefix[$p]; break } }
        if ($site) { $ws.Cells.Item($r, $siteCol).Value2 = $site; $count[$site]++ }
        else { $noMatch.Add("row ${r}: $vm") }
    }

    # ---------- Title row(s) above the header: all sites, not one site ----------
    if ($hdr -ge 2) {
        $ws.Cells.Item(1, 1).Value2 = $Title
        if ($hdr -ge 3) {
            $parts = @($count.Keys | ForEach-Object { '{0} {1}' -f $_, $count[$_] })
            $total = ($count.Values | Measure-Object -Sum).Sum + $noMatch.Count
            $ws.Cells.Item(2, 1).Value2 = ('{0} VMs  |  {1}  |  Generated {2}' -f $total, ($parts -join '  |  '), (Get-Date -Format 'yyyy-MM-dd HH:mm'))
        }
    }

    $excel.Calculation = -4105                                   # automatic again
    $excel.ScreenUpdating = $true
    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 51)
    Write-Host ''
    foreach ($k in $count.Keys) { Write-Host ("{0,-16} {1}" -f $k, $count[$k]) }
    Write-Host ("{0,-16} {1}" -f 'No site (blank)', $noMatch.Count) -ForegroundColor $(if ($noMatch.Count) { 'Yellow' } else { 'Gray' })
    $noMatch | Select-Object -First 15 | ForEach-Object { Write-Host "   $_" -ForegroundColor Yellow }
    if ($noMatch.Count -gt 15) { Write-Host "   ... and $($noMatch.Count - 15) more" -ForegroundColor Yellow }
    Write-Host "`nSaved: $out" -ForegroundColor Green
}
catch { Write-Host "FAILED: $($_.Exception.Message)  (input file not changed)" -ForegroundColor Red }
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
