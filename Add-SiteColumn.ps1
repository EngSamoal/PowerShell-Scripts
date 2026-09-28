# Add-SiteColumn.ps1
# Adds a "Site" column as the 2nd column of the consolidated master sheet and fills it from the
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
    $ws = $null; $hdr = 0; $nameCol = 0
    foreach ($s in $wb.Worksheets) {
        $used = $s.UsedRange
        $maxCol = $used.Column + $used.Columns.Count - 1
        for ($r = 1; $r -le 20 -and -not $hdr; $r++) {
            for ($c = 1; $c -le $maxCol; $c++) {
                if ((Norm $s.Cells.Item($r, $c).Value2) -eq 'vm name') { $ws = $s; $hdr = $r; $nameCol = $c; break }
            }
        }
        if ($ws) { break }
    }
    if (-not $ws) { throw "No 'VM NAME' header found in $InputFile" }
    $lastCol = $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1
    Write-Host "Sheet '$($ws.Name)', header row $hdr, VM NAME in column $nameCol"

    # ---------- Site column: reuse if it exists, otherwise insert as column 2 ----------
    $siteCol = 0
    for ($c = 1; $c -le $lastCol; $c++) { if ((Norm $ws.Cells.Item($hdr, $c).Value2) -eq 'site') { $siteCol = $c; break } }
    if ($siteCol) { Write-Host "A 'Site' column already exists (column $siteCol) - it will be filled" -ForegroundColor Yellow }
    else {
        [void]$ws.Columns.Item(2).Insert(-4161, 1)            # shift right, format taken from the column on the right
        $siteCol = 2
        if ($nameCol -ge 2) { $nameCol++ }
        $ws.Cells.Item($hdr, 2).Value2 = 'Site'
        $ws.Columns.Item(2).ColumnWidth = 18
        $lastCol++
    }

    # ---------- Fill Site from the VM name ----------
    $lastRow = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    $count = [ordered]@{}; foreach ($v in $SiteByPrefix.Values) { $count[$v] = 0 }
    $noMatch = New-Object System.Collections.Generic.List[string]
    for ($r = $hdr + 1; $r -le $lastRow; $r++) {
        $vm = ([string]$ws.Cells.Item($r, $nameCol).Value2).Trim()
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
