# Fill-OwnersByService.ps1
# Fills EMPTY owner cells (every column whose header contains "Owner") using the owners of the same
# service at another site. Service = 2nd part of the VM name: SF-UFM-APP01 -> UFM.
# Existing values are never changed. Columns, colours and formats are not touched.
# The input file is only read; the result is saved as a NEW file. Needs Microsoft Excel.
# Run:  .\Fill-OwnersByService.ps1
param(
    [string]$InputFile = 'C:\temp\Master_withSite.xlsx',          # <-- master with Site column
    [string]$Output    = 'C:\temp\Master_ownersFilled.xlsx',      # <-- result (new file)
    [string]$Report    = 'C:\temp\Owners_filled_report.csv'       # <-- what was filled / still missing
)
$ErrorActionPreference = 'Stop'
function Norm($s) { (([string]$s) -replace '[\s\u00A0]+', ' ').Trim() }
function Get-Service([string]$vm) {
    $p = @($vm.Trim() -split '[-_ ]+' | Where-Object { $_ })
    if ($p.Count -ge 3) { return $p[1].ToUpper() }        # SF-UFM-APP01 -> UFM
    return $null                                          # names that do not follow the standard
}

if (-not (Test-Path $InputFile)) { Write-Host "Not found: $InputFile" -ForegroundColor Red; return }
$in = (Resolve-Path $InputFile).Path; $out = [IO.Path]::GetFullPath($Output)
if ($in -eq $out) { Write-Host 'Output must be a different file than the input.' -ForegroundColor Red; return }

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $wb = $excel.Workbooks.Open($in, 0, $true)            # read-only; saved under the new name

    # ---------- Sheet, header row, columns ----------
    $ws = $null; $hdr = 0; $nameCol = 0
    foreach ($s in $wb.Worksheets) {
        $hit = $s.Range('A1:ZZ20').Find('VM NAME', [Type]::Missing, -4163, 1)
        if ($hit) { $ws = $s; $hdr = $hit.Row; $nameCol = $hit.Column; break }
    }
    if (-not $ws) { throw "No 'VM NAME' header found in $InputFile" }
    $hv = $ws.Range($ws.Cells.Item($hdr, 1), $ws.Cells.Item($hdr, 200)).Value2
    $ownerCols = @(for ($c = 1; $c -le 200; $c++) { if ((Norm $hv[1, $c]) -match 'owner') { $c } })
    if ($ownerCols.Count -eq 0) { throw "No column with 'Owner' in its header" }
    $lastRow = $ws.Cells.Item($ws.Rows.Count, $nameCol).End(-4162).Row
    Write-Host ("Sheet '{0}': {1} VMs, owner columns: {2}" -f $ws.Name, ($lastRow - $hdr), (($ownerCols | ForEach-Object { Norm $hv[1, $_] }) -join ', '))

    # ---------- Read everything at once ----------
    $names = $ws.Range($ws.Cells.Item($hdr + 1, $nameCol), $ws.Cells.Item($lastRow, $nameCol)).Value2
    $own = @{}
    foreach ($c in $ownerCols) { $own[$c] = $ws.Range($ws.Cells.Item($hdr + 1, $c), $ws.Cells.Item($lastRow, $c)).Value2 }
    $n = $lastRow - $hdr

    # ---------- Known owners per service (most common value if they differ) ----------
    $votes = @{}      # "SERVICE|col" -> @{ value = count }
    for ($i = 1; $i -le $n; $i++) {
        $svc = Get-Service ([string]$names[$i, 1]); if (-not $svc) { continue }
        foreach ($c in $ownerCols) {
            $v = Norm $own[$c][$i, 1]; if (-not $v) { continue }
            $k = "$svc|$c"
            if (-not $votes.ContainsKey($k)) { $votes[$k] = @{} }
            if (-not $votes[$k].ContainsKey($v)) { $votes[$k][$v] = 0 }
            $votes[$k][$v]++
        }
    }
    # One clear winner -> used. A tie (e.g. 1 vs 1) -> not filled, listed in the report for you to decide.
    $best = @{}; $tie = @{}
    foreach ($k in $votes.Keys) {
        $top = @($votes[$k].GetEnumerator() | Sort-Object Value -Descending)
        if ($top.Count -gt 1 -and $top[0].Value -eq $top[1].Value) { $tie[$k] = (($top | ForEach-Object { $_.Key }) -join ' / ') }
        else { $best[$k] = $top[0].Key }
    }

    # ---------- Fill only the empty cells ----------
    $rows = New-Object System.Collections.Generic.List[object]
    $filled = 0; $missing = 0
    for ($i = 1; $i -le $n; $i++) {
        $vm = Norm $names[$i, 1]; if (-not $vm) { continue }
        $svc = Get-Service $vm
        foreach ($c in $ownerCols) {
            if (Norm $own[$c][$i, 1]) { continue }                          # already has a value - keep it
            $col = Norm $hv[1, $c]
            if ($svc -and $best.ContainsKey("$svc|$c")) {
                $val = $best["$svc|$c"]
                $ws.Cells.Item($hdr + $i, $c).Value2 = [string]$val
                $filled++
                $rows.Add([pscustomobject]@{ Row = $hdr + $i; 'VM NAME' = $vm; Service = $svc; Column = $col; Result = 'Filled'; Value = $val })
            }
            else {
                $missing++
                $why = $(if (-not $svc) { 'Missing - name not in SITE-SERVICE-xxx format' }
                         elseif ($tie.ContainsKey("$svc|$c")) { 'Missing - other sites disagree, choose one' }
                         else { 'Missing - no other VM of this service has it' })
                $rows.Add([pscustomobject]@{ Row = $hdr + $i; 'VM NAME' = $vm; Service = $svc; Column = $col; Result = $why
                    Value = $(if ($svc -and $tie.ContainsKey("$svc|$c")) { $tie["$svc|$c"] } else { '' }) })
            }
        }
    }

    $excel.ScreenUpdating = $true
    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 51)
    if (-not (Test-Path $out)) { throw "Excel did not write $out" }
    $rows | Export-Csv $Report -NoTypeInformation -Encoding UTF8

    $conflicts = @($tie.Keys)
    Write-Host ''
    Write-Host "Cells filled        : $filled" -ForegroundColor Green
    Write-Host "Cells still empty   : $missing" -ForegroundColor $(if ($missing) { 'Yellow' } else { 'Gray' })
    if ($conflicts.Count) { Write-Host "Services where sites disagree (left empty, see report): $(($conflicts | ForEach-Object { ($_ -split '\|')[0] } | Select-Object -Unique) -join ', ')" -ForegroundColor Yellow }
    Write-Host "`nSaved : $out" -ForegroundColor Green
    Write-Host "Report: $Report"
}
catch { Write-Host "FAILED: $($_.Exception.Message)  (input file not changed)" -ForegroundColor Red }
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
