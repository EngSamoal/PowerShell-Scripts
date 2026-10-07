# Set-MissingOwnersByAppCode.ps1
# Fills EMPTY "Business Owner" / "Technical Owner" cells from other VMs of the same application.
# Application = 2nd part of the VM name:  DX-BC-AP02 -> BC   (not case-sensitive)
#   - Existing owner values are never changed. Other columns are not touched.
#   - If the VMs of an application have 2+ different owners, nothing is filled for that application
#     (listed as CONFLICT so you can check it by hand).
#   - Uses Microsoft Excel itself to edit the file, so colours and formats stay exactly as they are.
#
# Run:  .\Set-MissingOwnersByAppCode.ps1            -> only SHOWS what would be filled (file not changed)
#       .\Set-MissingOwnersByAppCode.ps1 -Apply     -> makes a backup copy, fills the owners, saves the file
# Close the Excel file before running.
param(
    [string]$File = 'C:\temp\master_draft.xlsx',
    [switch]$Apply
)
$ErrorActionPreference = 'Stop'
Write-Host 'Set-MissingOwnersByAppCode build 2026-10-07.6' -ForegroundColor Cyan

function Clean($v) { (([string]$v) -replace '[\s\u00A0]+', ' ').Trim() }

function Get-AppCode([string]$vm) {
    $parts = @((Clean $vm) -split '-')
    if ($parts.Count -ge 2 -and $parts[0].Trim() -and $parts[1].Trim()) { return $parts[1].Trim().ToUpper() }
    return $null
}

# Decides what to fill. $rows: objects with Row, VM, 'Business Owner', 'Technical Owner'.
function Get-OwnerFills($rows) {
    $fields = 'Business Owner', 'Technical Owner'
    # Owners already known per application: app|field -> @{ UPPER value -> value as written }
    $known = @{}
    foreach ($r in $rows) {
        $app = Get-AppCode $r.VM; if (-not $app) { continue }
        foreach ($f in $fields) {
            $v = Clean $r.$f; if (-not $v) { continue }
            $k = "$app|$f"
            if (-not $known.ContainsKey($k)) { $known[$k] = @{} }
            if (-not $known[$k].ContainsKey($v.ToUpper())) { $known[$k][$v.ToUpper()] = $v }
        }
    }
    foreach ($r in $rows) {
        $app = Get-AppCode $r.VM
        foreach ($f in $fields) {
            if (Clean $r.$f) { continue }                                     # already has an owner
            $k = "$app|$f"
            $res = [pscustomobject]@{ Row = $r.Row; VM = $r.VM; App = $app; Field = $f; NewValue = ''; Result = '' }
            if (-not $app)                       { $res.Result = 'SKIPPED - VM name is blank or has no 2nd part' }
            elseif (-not $known.ContainsKey($k)) { $res.Result = "NOT FOUND - no other $app VM has a $f" }
            elseif ($known[$k].Count -gt 1)      { $res.Result = "CONFLICT - $app has different owners: $(@($known[$k].Values) -join ' / ')" }
            else                                 { $res.NewValue = @($known[$k].Values)[0]; $res.Result = 'FILLED' }
            $res
        }
    }
}

if ($MyInvocation.InvocationName -eq '.') { return }    # dot-sourced (testing): only load the functions

if (-not (Test-Path -LiteralPath $File)) { Write-Host "File not found: $File" -ForegroundColor Red; exit 1 }
$File = (Resolve-Path -LiteralPath $File).ProviderPath

$excel = $null; $wb = $null
try {
    if ($Apply) {
        $backup = Join-Path (Split-Path $File) ('{0}_backup_{1}{2}' -f [IO.Path]::GetFileNameWithoutExtension($File), (Get-Date -Format 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($File))
        Copy-Item -LiteralPath $File -Destination $backup
        Write-Host "Backup copy : $backup" -ForegroundColor Green
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($File, 0, (-not $Apply))     # read-only unless -Apply

    # Find the sheet + header row with "VM Name", then the two owner columns in that row.
    $ws = $null; $hdr = 0; $cols = @{}
    foreach ($s in $wb.Worksheets) {
        $hit = $s.Range('A1:ZZ30').Find('VM Name', [Type]::Missing, -4163, 1)    # whole-cell match on values
        if (-not $hit) { continue }
        $cols = @{ 'VM Name' = $hit.Column }
        $lastCol = $s.Cells.Item($hit.Row, $s.Columns.Count).End(-4159).Column
        for ($c = 1; $c -le $lastCol; $c++) {
            $h = Clean $s.Cells.Item($hit.Row, $c).Value2
            if ($h -eq 'Business Owner' -or $h -eq 'Technical Owner') { $cols[$h] = $c }
        }
        if ($cols.Count -eq 3) { $ws = $s; $hdr = $hit.Row; break }
    }
    if (-not $ws) { throw "Could not find the headers 'VM Name', 'Business Owner' and 'Technical Owner' in $File" }

    $lastRow = 0
    foreach ($c in $cols.Values) { $lastRow = [Math]::Max($lastRow, $ws.Cells.Item($ws.Rows.Count, $c).End(-4162).Row) }
    Write-Host "Sheet '$($ws.Name)': header row $hdr, rows $($hdr + 1)-$lastRow"

    $rows = for ($r = $hdr + 1; $r -le $lastRow; $r++) {
        $vm = Clean $ws.Cells.Item($r, $cols['VM Name']).Value2
        $bo = Clean $ws.Cells.Item($r, $cols['Business Owner']).Value2
        $to = Clean $ws.Cells.Item($r, $cols['Technical Owner']).Value2
        if (-not ($vm -or $bo -or $to)) { continue }                            # empty row
        [pscustomobject]@{ Row = $r; VM = $vm; 'Business Owner' = $bo; 'Technical Owner' = $to }
    }
    $results = @(Get-OwnerFills @($rows))
    $fills = @($results | Where-Object { $_.Result -eq 'FILLED' })

    if ($Apply) {
        foreach ($x in $fills) { $ws.Cells.Item($x.Row, $cols[$x.Field]).Value2 = $x.NewValue }   # value only, format kept
        $wb.Save()
    }

    # ---- Report ----
    $csv = Join-Path (Split-Path $File) ('OwnerFill_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $results | Select-Object VM, App, Field, NewValue, Result, Row | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
    $results | Select-Object VM, App, Field, NewValue, Result | Format-Table -AutoSize -Wrap | Out-String -Width 300 | Write-Host
    Write-Host "VMs checked        : $(@($rows).Count)"
    Write-Host "Applications       : $(@($rows | ForEach-Object { Get-AppCode $_.VM } | Where-Object { $_ } | Sort-Object -Unique).Count)"
    Write-Host "Missing owners     : $($results.Count)  (Business: $(@($results | Where-Object Field -eq 'Business Owner').Count), Technical: $(@($results | Where-Object Field -eq 'Technical Owner').Count))"
    Write-Host "Filled             : $($fills.Count)" -ForegroundColor Green
    Write-Host "Conflicts (check)  : $(@($results | Where-Object { $_.Result -like 'CONFLICT*' }).Count)" -ForegroundColor Yellow
    Write-Host "Not found          : $(@($results | Where-Object { $_.Result -like 'NOT FOUND*' }).Count)"
    Write-Host "Bad/blank VM names : $(@($results | Where-Object { $_.Result -like 'SKIPPED*' } | Select-Object -Unique Row).Count)"
    Write-Host "Report             : $csv"
    if ($Apply) { Write-Host "SAVED              : $File" -ForegroundColor Green }
    else        { Write-Host 'Preview only - nothing was changed. Run again with -Apply to fill the owners.' -ForegroundColor Yellow }
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($Apply) { Write-Host 'The file was not saved. The backup copy is untouched.' -ForegroundColor Yellow }
    exit 1
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
