# New-MasterWithSite.ps1
# Builds a NEW master file: the same columns (values, colours, formats) as the current master sheet,
# plus a "Site" column just before "Business Owner", filled from the VM name prefix.
# The original file is only read. Needs Microsoft Excel.
# Run:  .\New-MasterWithSite.ps1
param(
    [string]$InputFile = 'C:\temp\Master.xlsx',            # <-- current master
    [string]$Output    = 'C:\temp\Master_withSite.xlsx'    # <-- new master
)

#region ======================= SITE RULES (edit here) =======================
$SiteByPrefix = [ordered]@{
    'SF' = 'SixFlags'
    'AQ' = 'AquaArabia'
    'TB' = 'SEVEN Tabuk'
    'AB' = 'SEVEN ABHA'
    'HA' = 'SEVEN Alhamra'
}
$Title = 'All Sites  -  VM Inventory'      # text for the title row above the header
#endregion ==================================================================

$ErrorActionPreference = 'Stop'
function Norm($s) { (([string]$s) -replace '[\s\u00A0]+', ' ').Trim().ToLower() }
function Test-JunkHeader($s) { ([string]$s).Trim() -match '^(Column\s*)?\d*$' }   # blank / "2" / "Column8"

if (-not (Test-Path $InputFile)) { Write-Host "Not found: $InputFile" -ForegroundColor Red; return }
$in = (Resolve-Path $InputFile).Path; $out = [IO.Path]::GetFullPath($Output)
if ($in -eq $out) { Write-Host 'Output must be a different file than the input.' -ForegroundColor Red; return }

$excel = $null; $wbS = $null; $wbN = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $wbS = $excel.Workbooks.Open($in, 0, $true)                  # read-only

    # ---------- Source sheet, header row, real columns / rows ----------
    $src = $null; $hdr = 0; $nameCol = 0
    foreach ($s in $wbS.Worksheets) {
        $hit = $s.Range('A1:ZZ20').Find('VM NAME', [Type]::Missing, -4163, 1)
        if ($hit) { $src = $s; $hdr = $hit.Row; $nameCol = $hit.Column; break }
    }
    if (-not $src) { throw "No 'VM NAME' header found in $InputFile" }
    $hv = $src.Range($src.Cells.Item($hdr, 1), $src.Cells.Item($hdr, $src.Columns.Count)).Value2
    $lastCol = $nameCol
    for ($c = $hv.GetUpperBound(1); $c -gt $nameCol; $c--) { if (-not (Test-JunkHeader $hv[1, $c])) { $lastCol = $c; break } }
    $pos = 0
    for ($c = 1; $c -le $lastCol; $c++) { if ((Norm $hv[1, $c]) -eq 'business owner') { $pos = $c; break } }
    if (-not $pos) { $pos = 2; Write-Host "No 'Business Owner' header - Site goes in column 2" -ForegroundColor Yellow }
    $lastRow = $hdr + 1
    for ($c = 1; $c -le $lastCol; $c++) { $lastRow = [Math]::Max($lastRow, $src.Cells.Item($src.Rows.Count, $c).End(-4162).Row) }
    $lo = $null
    foreach ($t in $src.ListObjects) { if ($t.Range.Row -le $hdr -and ($t.Range.Row + $t.Range.Rows.Count - 1) -ge $hdr) { $lo = $t } }
    Write-Host "Sheet '$($src.Name)': header row $hdr, columns 1-$lastCol, rows 1-$lastRow, Site before column $pos"

    # ---------- New sheet (inside the opened copy, so the file keeps its sensitivity label) ----------
    # copy columns 1..pos-1, then pos..lastCol one step right
    $srcName = $src.Name
    $dst = $wbS.Worksheets.Add([Type]::Missing, $wbS.Worksheets.Item($wbS.Worksheets.Count))
    $dst.Cells.Font.Name = $src.Cells.Item($hdr + 1, $nameCol).Font.Name
    for ($c = 1; $c -le $lastCol; $c++) {
        $to = $(if ($c -lt $pos) { $c } else { $c + 1 })
        [void]$src.Range($src.Cells.Item(1, $c), $src.Cells.Item($lastRow, $c)).Copy($dst.Cells.Item(1, $to))
        $dst.Columns.Item($to).ColumnWidth = $src.Columns.Item($c).ColumnWidth
    }
    # Site column: same look as the Business Owner column, then its own values
    [void]$src.Range($src.Cells.Item(1, $pos), $src.Cells.Item($lastRow, $pos)).Copy($dst.Cells.Item(1, $pos))
    [void]$dst.Range($dst.Cells.Item(1, $pos), $dst.Cells.Item($lastRow, $pos)).ClearContents()
    $dst.Columns.Item($pos).ColumnWidth = 18
    $dst.Cells.Item($hdr, $pos).Value2 = 'Site'
    for ($r = 1; $r -le $lastRow; $r++) { $dst.Rows.Item($r).RowHeight = $src.Rows.Item($r).RowHeight }

    # Table (filters + banding) with the same style as the original
    if ($lo) {
        $nt = $dst.ListObjects.Add(1, $dst.Range($dst.Cells.Item($hdr, 1), $dst.Cells.Item($lastRow, $lastCol + 1)), $null, 1)
        try { $nt.TableStyle = $lo.TableStyle.Name } catch { }
        $nt.ShowTableStyleRowStripes = $lo.ShowTableStyleRowStripes
        $nt.ShowAutoFilter = $lo.ShowAutoFilter
    }
    elseif ($src.AutoFilterMode) { [void]$dst.Range($dst.Cells.Item($hdr, 1), $dst.Cells.Item($lastRow, $lastCol + 1)).AutoFilter() }

    # ---------- Fill Site ----------
    $count = [ordered]@{}; foreach ($v in $SiteByPrefix.Values) { $count[$v] = 0 }
    $noMatch = New-Object System.Collections.Generic.List[string]
    $vmCol = $(if ($nameCol -lt $pos) { $nameCol } else { $nameCol + 1 })
    $names = $dst.Range($dst.Cells.Item($hdr, $vmCol), $dst.Cells.Item($lastRow, $vmCol)).Value2
    for ($r = $hdr + 1; $r -le $lastRow; $r++) {
        $vm = ([string]$names[($r - $hdr + 1), 1]).Trim()
        if (-not $vm -or (Norm $vm) -eq 'vm name') { continue }
        $site = $null
        foreach ($p in $SiteByPrefix.Keys) { if ($vm.StartsWith($p, [StringComparison]::OrdinalIgnoreCase)) { $site = $SiteByPrefix[$p]; break } }
        if ($site) { $dst.Cells.Item($r, $pos).Value2 = $site; $count[$site]++ } else { $noMatch.Add("row ${r}: $vm") }
    }

    # ---------- Title row(s): all sites ----------
    if ($hdr -ge 2) {
        $dst.Cells.Item(1, 1).Value2 = $Title
        if ($hdr -ge 3) {
            $total = ($count.Values | Measure-Object -Sum).Sum + $noMatch.Count
            $dst.Cells.Item(2, 1).Value2 = ('{0} VMs  |  {1}  |  Generated {2}' -f $total, (@($count.Keys | ForEach-Object { '{0} {1}' -f $_, $count[$_] }) -join '  |  '), (Get-Date -Format 'yyyy-MM-dd HH:mm'))
        }
    }

    # ---------- Keep only the new sheet, with the original name ----------
    $tabColor = $src.Tab.Color
    foreach ($w in @($wbS.Worksheets | Where-Object { $_.Name -ne $dst.Name })) { $w.Delete() }
    $dst.Name = $srcName
    $dst.Tab.Color = $tabColor
    $excel.ScreenUpdating = $true
    $dst.Activate()
    $excel.ActiveWindow.ScrollRow = 1; $excel.ActiveWindow.ScrollColumn = 1
    $excel.ActiveWindow.SplitRow = $hdr; $excel.ActiveWindow.SplitColumn = 0
    $excel.ActiveWindow.FreezePanes = $true
    $excel.ActiveWindow.DisplayGridlines = $false

    if (Test-Path $out) { Remove-Item $out -Force }
    $wbS.SaveAs($out, 51)
    if (-not (Test-Path $out)) { throw "Excel did not write $out (a save prompt, e.g. a sensitivity-label prompt, may have been blocked)." }
    Write-Host ''
    foreach ($k in $count.Keys) { Write-Host ("{0,-16} {1}" -f $k, $count[$k]) }
    Write-Host ("{0,-16} {1}" -f 'No site (blank)', $noMatch.Count) -ForegroundColor $(if ($noMatch.Count) { 'Yellow' } else { 'Gray' })
    $noMatch | Select-Object -First 15 | ForEach-Object { Write-Host "   $_" -ForegroundColor Yellow }
    if ($noMatch.Count -gt 15) { Write-Host "   ... and $($noMatch.Count - 15) more" -ForegroundColor Yellow }
    Write-Host "`nSaved: $out" -ForegroundColor Green
}
catch { Write-Host "FAILED: $($_.Exception.Message)  (original file not changed)" -ForegroundColor Red }
finally {
    foreach ($w in @($wbN, $wbS)) { if ($w) { try { $w.Close($false) } catch { } } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wbN, $wbS, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
