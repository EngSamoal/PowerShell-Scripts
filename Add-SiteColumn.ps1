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
    $lastCol = $ws.Cells.Item($hdr, $ws.Columns.Count).End(-4159).Column     # last filled header cell ...
    $hv = $ws.Range($ws.Cells.Item($hdr, 1), $ws.Cells.Item($hdr, $lastCol)).Value2
    while ($lastCol -gt $nameCol -and ([string]$hv[1, $lastCol]).Trim() -eq '') { $lastCol-- }   # ... ignoring cells with only spaces
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
        # Move the block from "Business Owner" to the last column one column to the right (only these cells move)
        $lo = $null
        foreach ($t in $ws.ListObjects) { if ($t.Range.Row -le $hdr -and ($t.Range.Row + $t.Range.Rows.Count - 1) -ge $hdr) { $lo = $t } }
        $blockLastRow = $ws.Cells.Item($ws.Rows.Count, $nameCol).End(-4162).Row
        $blockLastCol = $lastCol
        if ($lo) {
            $blockLastRow = [Math]::Max($blockLastRow, $lo.Range.Row + $lo.Range.Rows.Count - 1)
            $blockLastCol = $lo.Range.Column + $lo.Range.Columns.Count - 1
        }
        # Extend the block to the right up to the first completely empty column (that column receives the shift).
        # One bulk read of the area; looks at most 200 columns to the right.
        for ($c = $pos; $c -le $blockLastCol; $c++) { $blockLastRow = [Math]::Max($blockLastRow, $ws.Cells.Item($ws.Rows.Count, $c).End(-4162).Row) }
        $startCol = $blockLastCol
        for ($pass = 1; $pass -le 3; $pass++) {
            $scanTo = [Math]::Min($ws.Columns.Count - 1, $startCol + 200)
            $grid = $ws.Range($ws.Cells.Item(1, $startCol + 1), $ws.Cells.Item($blockLastRow, $scanTo)).Value2
            $found = $false
            for ($k = 1; $k -le $grid.GetUpperBound(1); $k++) {
                $empty = $true
                for ($r = 1; $r -le $grid.GetUpperBound(0); $r++) { if (([string]$grid[$r, $k]).Trim() -ne '') { $empty = $false; break } }
                if ($empty) { $blockLastCol = $startCol + $k - 1; $found = $true; break }
            }
            if (-not $found) { throw "No empty column found within 200 columns to the right of the data (rows 1-$blockLastRow). Clear some cells there and run again." }
            # Columns added to the block may go further down - include those rows and check again
            $grow = $blockLastRow
            for ($c = $startCol + 1; $c -le $blockLastCol; $c++) { $grow = [Math]::Max($grow, $ws.Cells.Item($ws.Rows.Count, $c).End(-4162).Row) }
            if ($grow -eq $blockLastRow) { break }
            $blockLastRow = $grow
        }
        if ($lo) {   # remember the table size before the move
            $loRow = $lo.Range.Row; $loCol = $lo.Range.Column
            $loLastRow = $loRow + $lo.Range.Rows.Count - 1; $loLastCol = $loCol + $lo.Range.Columns.Count - 1
        }
        [void]$ws.Range($ws.Cells.Item(1, $pos), $ws.Cells.Item($blockLastRow, $blockLastCol)).Cut($ws.Cells.Item(1, $pos + 1))
        # the table grows by exactly one column (the new Site column) - nothing else is added to it
        if ($lo) { $lo.Resize($ws.Range($ws.Cells.Item($loRow, $loCol), $ws.Cells.Item($loLastRow, $loLastCol + 1))) }

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
