# Add-OwnersToMaster.ps1
# Copies the owner columns (every column whose header contains "Owner") from Master2 into Master1,
# placed just before the "VM NAME" column. Match: VM name + IP (or VM name only when it is unique).
# Safe: Master1 and Master2 are only read; the result is saved as a NEW file.
# Needs Microsoft Excel.
# Run:            .\Add-OwnersToMaster.ps1
# Check only:     .\Add-OwnersToMaster.ps1 -Check     (shows what is detected, writes nothing)
param(
    [string]$Master1 = 'C:\temp\Master1.xlsx',               # <-- your master file
    [string]$Master2 = 'C:\temp\Master2.xlsx',               # <-- file from management (has the owners)
    [string]$Output  = 'C:\temp\Master1_withOwners.xlsx',    # <-- result (new file)
    [string]$Report  = 'C:\temp\Owners_match_report.csv',    # <-- which VM matched / not matched
    [switch]$Check
)

#region ======================= TAB MAPPING (edit here) =======================
# First word of the Master2 tab name  ->  Master1 tab name.
# (First word = text before the first space, "-" or "_": "SF-Windows" -> SF, "OT_on_ VMware platform" -> OT)
# Master2 tabs whose first word is not listed here are ignored.
$TabMap = @{
    'SF'  = 'SixFlags'
    'AQ'  = 'SixFlags'        # change to 'AquaArabia' if AQ tabs belong to AquaArabia
    'TB'  = 'SEVEN Tabuk'
    'AMC' = 'AMC'
    'OT'  = 'OT'
}
#endregion ===================================================================

$ErrorActionPreference = 'Stop'

function Norm($s) { (([string]$s) -replace '[\s\u00A0]+', ' ').Trim().ToLower() }
function Get-IPs($s) { @(([string]$s) -split '[,;\s]+' | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' }) }
function Get-FirstWord([string]$s) { (($s.Trim()) -split '[\s_\-]+')[0] }

$NameHeaders = @('vm name', 'vmname', 'vm_name', 'vm', 'name', 'server name', 'servername', 'hostname', 'host name')
$IpHeaders   = @('ip', 'ip address', 'ipaddress', 'ip addresses', 'primary ip', 'ip addr')

function Read-Sheet($ws) {
    # Finds the header row (first row, within the top 20, that has a VM-name header) and returns the grid
    $used = $ws.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $lastCol = $used.Column + $used.Columns.Count - 1
    $g = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item([Math]::Max($lastRow, 2), [Math]::Max($lastCol, 2))).Value2
    for ($r = 1; $r -le [Math]::Min(20, $g.GetUpperBound(0)); $r++) {
        $headers = @(for ($k = 1; $k -le $g.GetUpperBound(1); $k++) { (([string]$g[$r, $k]) -replace '[\s\u00A0]+', ' ').Trim() })
        $nameCol = 0; $ipCol = 0
        foreach ($h in $NameHeaders) {
            for ($k = 1; $k -le $headers.Count; $k++) { if ((Norm $headers[$k - 1]) -eq $h) { $nameCol = $k; break } }
            if ($nameCol) { break }
        }
        if (-not $nameCol) { continue }
        foreach ($h in $IpHeaders) {
            for ($k = 1; $k -le $headers.Count; $k++) { if ((Norm $headers[$k - 1]) -eq $h) { $ipCol = $k; break } }
            if ($ipCol) { break }
        }
        $owners = @(for ($k = 1; $k -le $headers.Count; $k++) { if ((Norm $headers[$k - 1]) -like '*owner*') { $k } })
        return @{ Grid = $g; HeaderRow = $r; NameCol = $nameCol; IpCol = $ipCol; Headers = $headers; OwnerCols = $owners }
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

    $wb1 = $excel.Workbooks.Open((Resolve-Path $Master1).Path, 0, [bool]$Check)
    $m1Tabs = @{}
    foreach ($ws in $wb1.Worksheets) { $m1Tabs[(Norm $ws.Name)] = $ws.Name }

    # ---------- 1. Read owners from Master2 ----------
    $wb2 = $excel.Workbooks.Open((Resolve-Path $Master2).Path, 0, $true)      # read-only
    $index = @{}              # vm name -> list of entries
    $targetHeaders = @{}      # Master1 tab -> owner headers to add
    Write-Host "`n--- Master2 ---"
    foreach ($ws in $wb2.Worksheets) {
        $fw = Get-FirstWord $ws.Name
        $target = $null
        foreach ($k in $TabMap.Keys) { if ((Norm $k) -eq (Norm $fw)) { $target = $TabMap[$k] } }
        if (-not $target -and $m1Tabs.ContainsKey((Norm $ws.Name))) { $target = $m1Tabs[(Norm $ws.Name)] }
        if (-not $target -or -not $m1Tabs.ContainsKey((Norm $target))) { Write-Host ("{0,-28} ignored (no matching Master1 tab)" -f $ws.Name) -ForegroundColor DarkGray; continue }
        $target = $m1Tabs[(Norm $target)]

        $s = Read-Sheet $ws
        if (-not $s) { Write-Host ("{0,-28} -> {1,-14} NO 'VM Name' header found in the first 20 rows - skipped" -f $ws.Name, $target) -ForegroundColor Yellow; continue }
        if ($s.OwnerCols.Count -eq 0) { Write-Host ("{0,-28} -> {1,-14} no 'Owner' columns - skipped" -f $ws.Name, $target) -ForegroundColor Yellow; continue }

        $oh = @($s.OwnerCols | ForEach-Object { $s.Headers[$_ - 1] })
        if (-not $targetHeaders.ContainsKey($target) -or $oh.Count -gt $targetHeaders[$target].Count) { $targetHeaders[$target] = $oh }
        $rows = 0
        for ($r = $s.HeaderRow + 1; $r -le $s.Grid.GetUpperBound(0); $r++) {
            $name = Norm $s.Grid[$r, $s.NameCol]
            if (-not $name) { continue }
            $entry = @{
                Sheet   = $ws.Name
                Target  = $target
                IPs     = $(if ($s.IpCol) { Get-IPs $s.Grid[$r, $s.IpCol] } else { @() })
                Headers = $oh
                Values  = @($s.OwnerCols | ForEach-Object { ([string]$s.Grid[$r, $_]).Trim() })
            }
            if (-not $index.ContainsKey($name)) { $index[$name] = New-Object System.Collections.Generic.List[object] }
            $index[$name].Add($entry); $rows++
        }
        Write-Host ("{0,-28} -> {1,-14} header row {2}, name col '{3}', IP col '{4}', {5} VMs, owners: {6}" -f $ws.Name, $target, $s.HeaderRow,
            $s.Headers[$s.NameCol - 1], $(if ($s.IpCol) { $s.Headers[$s.IpCol - 1] } else { '-' }), $rows, ($oh -join ' | ')) -ForegroundColor Green
    }
    $wb2.Close($false); $wb2 = $null
    if ($index.Count -eq 0) { throw 'Nothing usable found in Master2 (see the lines above).' }

    # ---------- 2. Add the columns to Master1 ----------
    Write-Host "`n--- Master1 ---"
    foreach ($ws in $wb1.Worksheets) {
        $tab = $ws.Name
        if (-not $targetHeaders.ContainsKey($tab)) { Write-Host ("{0,-16} no Master2 tab mapped - not changed" -f $tab) -ForegroundColor DarkGray; continue }
        $s = Read-Sheet $ws
        if (-not $s) { Write-Host ("{0,-16} no 'VM NAME' column - skipped" -f $tab) -ForegroundColor Yellow; continue }
        $th = $targetHeaders[$tab]; $n = $th.Count

        # Columns already there (script run on this file before)? reuse them, otherwise insert before VM NAME
        $existing = @(for ($k = 1; $k -le $s.Headers.Count; $k++) { if ((Norm $s.Headers[$k - 1]) -eq (Norm $th[0])) { $k } })
        $first = $(if ($existing.Count) { $existing[0] } else { $s.NameCol })
        if (-not $Check -and $existing.Count -eq 0) {
            for ($i = 0; $i -lt $n; $i++) { [void]$ws.Columns.Item($first).Insert(-4161, 1) }   # shift right, format from VM NAME column
            for ($i = 0; $i -lt $n; $i++) {
                $ws.Cells.Item($s.HeaderRow, $first + $i).Value2 = [string]$th[$i]
                $ws.Columns.Item($first + $i).ColumnWidth = 26
            }
        }

        $cnt = @{ 'Name+IP' = 0; 'Name only' = 0; 'Not found' = 0; 'Ambiguous' = 0 }; $filled = 0
        for ($r = $s.HeaderRow + 1; $r -le $s.Grid.GetUpperBound(0); $r++) {
            $vm = ([string]$s.Grid[$r, $s.NameCol]).Trim()
            if (-not $vm) { continue }
            $ips = $(if ($s.IpCol) { Get-IPs $s.Grid[$r, $s.IpCol] } else { @() })
            $all = @($index[(Norm $vm)] | Where-Object { $_ })
            $mine = @($all | Where-Object { $_.Target -eq $tab })
            $ipOf = { param($list) @($list | Where-Object { $e = $_; @($e.IPs | Where-Object { $ips -contains $_ }).Count -gt 0 }) }
            $match = $null; $how = 'Not found'
            $x = @(& $ipOf $mine); if ($x.Count) { $match = $x[0]; $how = 'Name+IP' }
            if (-not $match) { $x = @(& $ipOf $all); if ($x.Count) { $match = $x[0]; $how = 'Name+IP' } }
            if (-not $match -and $mine.Count -eq 1) { $match = $mine[0]; $how = 'Name only' }
            if (-not $match -and $all.Count -eq 1) { $match = $all[0]; $how = 'Name only' }
            if (-not $match -and $all.Count -gt 1) { $how = 'Ambiguous' }
            $cnt[$how]++

            if ($match -and -not $Check) {
                for ($i = 0; $i -lt $n; $i++) {
                    $val = ''
                    if ($match.Values.Count -eq $n) { $val = $match.Values[$i] }                  # same layout -> by position
                    else { for ($k = 0; $k -lt $match.Headers.Count; $k++) { if ((Norm $match.Headers[$k]) -eq (Norm $th[$i])) { $val = $match.Values[$k]; break } } }
                    if ($val) { $ws.Cells.Item($r, $first + $i).Value2 = [string]$val; $filled++ }
                }
            }
            $reportRows.Add([pscustomobject]@{ Tab = $tab; 'VM NAME' = $vm; IP = ($ips -join ', '); Result = $how; 'Master2 Tab' = $(if ($match) { $match.Sheet } else { '' }) })
        }
        Write-Host ("{0,-16} {1} columns ({2}) | Name+IP {3} | Name only {4} | Not found {5} | Ambiguous {6} | cells filled {7}" -f $tab, $n, ($th -join ' | '),
            $cnt['Name+IP'], $cnt['Name only'], $cnt['Not found'], $cnt['Ambiguous'], $filled)
    }

    $reportRows | Export-Csv $Report -NoTypeInformation -Encoding UTF8
    if ($Check) { Write-Host "`nCHECK ONLY - nothing written. Report: $Report" -ForegroundColor Cyan; return }
    $out = [IO.Path]::GetFullPath($Output)
    if (Test-Path $out) { Remove-Item $out -Force }
    $wb1.SaveAs($out, 51)
    Write-Host "`nSaved : $out" -ForegroundColor Green
    Write-Host "Report: $Report" -ForegroundColor Green
}
catch { Write-Host "FAILED: $($_.Exception.Message)  (Master1 was not changed)" -ForegroundColor Red }
finally {
    foreach ($w in @($wb2, $wb1)) { if ($w) { try { $w.Close($false) } catch { } } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb2, $wb1, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
