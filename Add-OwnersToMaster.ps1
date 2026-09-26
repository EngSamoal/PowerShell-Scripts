# Add-OwnersToMaster.ps1
# Copies the owner columns (Business Owner / Technical Owner ...) from Master2 into Master1,
# placed just before the "VM NAME" column, for every VM found in both files.
# Match: VM NAME + IP. If the IP differs but the VM name is unique, it is matched by name only.
# Safe: Master1 and Master2 are only read; the result is saved as a NEW file.
# Needs Microsoft Excel. Run:  .\Add-OwnersToMaster.ps1
param(
    [string]$Master1 = 'C:\temp\Master1.xlsx',               # <-- your master file
    [string]$Master2 = 'C:\temp\Master2.xlsx',               # <-- file from management (has the owners)
    [string]$Output  = 'C:\temp\Master1_withOwners.xlsx',    # <-- result (new file)
    [string]$Report  = 'C:\temp\Owners_match_report.csv'     # <-- which VM matched / not matched
)

$OwnerPattern     = 'Owner'                                  # Master2 columns whose header contains this are copied
$DefaultOwnerCols = @('Business Owner', 'Technical Owner')   # used for tabs that have no same-named tab in Master2
$ErrorActionPreference = 'Stop'

function Norm($s) { ([string]$s).Trim().ToLower() }
function Get-IPs($s) { @(([string]$s) -split '[,;\s]+' | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' }) }

function Read-Sheet($ws) {
    # Finds the header row (the one containing "VM NAME") and returns the grid + column positions
    $used = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $lastCol = $used.Column + $used.Columns.Count - 1
    $g = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item([Math]::Max($lastRow, 2), [Math]::Max($lastCol, 2))).Value2
    for ($r = 1; $r -le [Math]::Min(15, $g.GetUpperBound(0)); $r++) {
        for ($c = 1; $c -le $g.GetUpperBound(1); $c++) {
            if ((Norm $g[$r, $c]) -eq 'vm name') {
                $headers = @(for ($k = 1; $k -le $g.GetUpperBound(1); $k++) { ([string]$g[$r, $k]).Trim() })
                $ipCol = 0
                for ($k = 1; $k -le $headers.Count; $k++) { if ((Norm $headers[$k - 1]) -eq 'ip') { $ipCol = $k; break } }
                return @{ Grid = $g; HeaderRow = $r; NameCol = $c; IpCol = $ipCol; Headers = $headers }
            }
        }
    }
    return $null
}

if (-not (Test-Path $Master1)) { Write-Host "Not found: $Master1" -ForegroundColor Red; return }
if (-not (Test-Path $Master2)) { Write-Host "Not found: $Master2" -ForegroundColor Red; return }
if ((Resolve-Path $Master1).Path -eq [IO.Path]::GetFullPath($Output)) { Write-Host 'Output must be a different file than Master1.' -ForegroundColor Red; return }

$excel = $null; $wb1 = $null; $wb2 = $null
$reportRows = New-Object System.Collections.Generic.List[object]
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # ---------- 1. Read owners from Master2 ----------
    $wb2 = $excel.Workbooks.Open((Resolve-Path $Master2).Path, 0, $true)      # read-only
    $index = @{}              # vm name -> list of entries
    $ownerColsBySheet = @{}   # master2 tab -> owner headers (in order)
    foreach ($ws in $wb2.Worksheets) {
        $s = Read-Sheet $ws
        if (-not $s) { continue }
        $oc = @(for ($k = 1; $k -le $s.Headers.Count; $k++) { if ($s.Headers[$k - 1] -match $OwnerPattern) { $k } })
        if ($oc.Count -eq 0) { continue }
        $ownerColsBySheet[(Norm $ws.Name)] = @($oc | ForEach-Object { $s.Headers[$_ - 1] })
        for ($r = $s.HeaderRow + 1; $r -le $s.Grid.GetUpperBound(0); $r++) {
            $name = Norm $s.Grid[$r, $s.NameCol]
            if (-not $name) { continue }
            $entry = @{
                Sheet   = $ws.Name
                IPs     = $(if ($s.IpCol) { Get-IPs $s.Grid[$r, $s.IpCol] } else { @() })
                Headers = @($oc | ForEach-Object { $s.Headers[$_ - 1] })
                Values  = @($oc | ForEach-Object { ([string]$s.Grid[$r, $_]).Trim() })
            }
            if (-not $index.ContainsKey($name)) { $index[$name] = New-Object System.Collections.Generic.List[object] }
            $index[$name].Add($entry)
        }
        Write-Host "Master2 tab '$($ws.Name)': owner columns = $(($ownerColsBySheet[(Norm $ws.Name)]) -join ', ')"
    }
    $wb2.Close($false); $wb2 = $null
    if ($index.Count -eq 0) { throw "No columns containing '$OwnerPattern' found in $Master2" }

    # ---------- 2. Add the columns to Master1 (saved as a new file) ----------
    $wb1 = $excel.Workbooks.Open((Resolve-Path $Master1).Path)
    foreach ($ws in $wb1.Worksheets) {
        $s = Read-Sheet $ws
        if (-not $s) { Write-Host "Tab '$($ws.Name)': no 'VM NAME' column - skipped" -ForegroundColor Yellow; continue }
        $tab = $ws.Name
        $targetHeaders = $(if ($ownerColsBySheet.ContainsKey((Norm $tab))) { $ownerColsBySheet[(Norm $tab)] } else { $DefaultOwnerCols })
        $n = $targetHeaders.Count

        # Columns already there (script run before)? then reuse them, otherwise insert before VM NAME
        $existing = @(for ($k = 1; $k -le $s.Headers.Count; $k++) { if ((Norm $s.Headers[$k - 1]) -eq (Norm $targetHeaders[0])) { $k } })
        if ($existing.Count -gt 0) { $first = $existing[0] }
        else {
            $first = $s.NameCol
            for ($i = 0; $i -lt $n; $i++) { [void]$ws.Columns.Item($first).Insert(-4161, 1) }   # shift right, format from VM NAME column
            for ($i = 0; $i -lt $n; $i++) {
                $ws.Cells.Item($s.HeaderRow, $first + $i).Value2 = [string]$targetHeaders[$i]
                $ws.Columns.Item($first + $i).ColumnWidth = 24
            }
        }

        $cnt = @{ 'Name+IP' = 0; 'Name only' = 0; 'Not found' = 0; 'Ambiguous' = 0 }
        for ($r = $s.HeaderRow + 1; $r -le $s.Grid.GetUpperBound(0); $r++) {
            $vm = ([string]$s.Grid[$r, $s.NameCol]).Trim()
            if (-not $vm) { continue }
            $ips = $(if ($s.IpCol) { Get-IPs $s.Grid[$r, $s.IpCol] } else { @() })
            $cands = @($index[(Norm $vm)] | Where-Object { $_ })
            $same = @($cands | Where-Object { (Norm $_.Sheet) -eq (Norm $tab) })
            $byIp = @($cands | Where-Object { $e = $_; @($e.IPs | Where-Object { $ips -contains $_ }).Count -gt 0 })
            $match = $null; $how = 'Not found'
            if (@($byIp | Where-Object { (Norm $_.Sheet) -eq (Norm $tab) }).Count -gt 0) { $match = @($byIp | Where-Object { (Norm $_.Sheet) -eq (Norm $tab) })[0]; $how = 'Name+IP' }
            elseif ($byIp.Count -gt 0) { $match = $byIp[0]; $how = 'Name+IP' }
            elseif ($same.Count -eq 1) { $match = $same[0]; $how = 'Name only' }
            elseif ($cands.Count -eq 1) { $match = $cands[0]; $how = 'Name only' }
            elseif ($cands.Count -gt 1) { $how = 'Ambiguous' }
            $cnt[$how]++

            if ($match) {
                for ($i = 0; $i -lt $n; $i++) {
                    # same tab -> by position; other tab -> by header name
                    $val = ''
                    if ((Norm $match.Sheet) -eq (Norm $tab) -and $i -lt $match.Values.Count) { $val = $match.Values[$i] }
                    else {
                        for ($k = 0; $k -lt $match.Headers.Count; $k++) { if ((Norm $match.Headers[$k]) -eq (Norm $targetHeaders[$i])) { $val = $match.Values[$k]; break } }
                    }
                    if ($val) { $ws.Cells.Item($r, $first + $i).Value2 = [string]$val }
                }
            }
            $reportRows.Add([pscustomobject]@{ Tab = $tab; 'VM NAME' = $vm; IP = ($ips -join ', '); Result = $how; 'Master2 Tab' = $(if ($match) { $match.Sheet } else { '' }) })
        }
        Write-Host ("Tab '{0}': +{1} columns | Name+IP {2} | Name only {3} | Not found {4} | Ambiguous {5}" -f $tab, $(if ($existing.Count) { 0 } else { $n }), $cnt['Name+IP'], $cnt['Name only'], $cnt['Not found'], $cnt['Ambiguous'])
    }

    $out = [IO.Path]::GetFullPath($Output)
    if (Test-Path $out) { Remove-Item $out -Force }
    $wb1.SaveAs($out, 51)
    $wb1.Close($false); $wb1 = $null
    $reportRows | Export-Csv $Report -NoTypeInformation -Encoding UTF8
    Write-Host "`nSaved : $out" -ForegroundColor Green
    Write-Host "Report: $Report  (see 'Not found' / 'Ambiguous' / 'Name only' rows)" -ForegroundColor Green
}
catch { Write-Host "FAILED: $($_.Exception.Message)  (Master1 was not changed)" -ForegroundColor Red }
finally {
    foreach ($w in @($wb2, $wb1)) { if ($w) { try { $w.Close($false) } catch { } } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb2, $wb1, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
