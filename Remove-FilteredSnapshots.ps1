# Remove-FilteredSnapshots.ps1
# Lists every VM snapshot on all connected vCenters, filters them by a word in the snapshot
# NAME and a word in the snapshot DESCRIPTION, and (only when asked) deletes the matches in
# batches of 10 - waiting for each batch to finish in vCenter before starting the next one.
#
# SAFE BY DEFAULT: without -Delete the script only lists/filters and writes CSV files.
#
# Step 1 - list + filter (read-only, nothing is changed):
#   .\Remove-FilteredSnapshots.ps1 -NameKeyword 'patch' -DescriptionKeyword 'CHG12345'
#     -> AllSnapshots_<time>.csv      every snapshot found
#     -> SelectedSnapshots_<time>.csv only the ones that match (review this file!)
#
# Step 2 - delete what step 1 selected (asks you to type DELETE first):
#   .\Remove-FilteredSnapshots.ps1 -NameKeyword 'patch' -DescriptionKeyword 'CHG12345' -Delete
#   or delete exactly the rows left in a reviewed CSV (you may remove rows from it in Excel):
#   .\Remove-FilteredSnapshots.ps1 -FromCsv 'C:\temp\SnapshotCleanup\SelectedSnapshots_xxx.csv' -Delete
#     -> DeleteResults_<time>.csv     result of every deletion (written after every batch)
#     -> SnapshotCleanup_<time>.log   full transcript
#
# Safety rules built in:
#   * Nothing is deleted without -Delete AND typing DELETE at the prompt.
#   * Both keywords must match by default (-MatchMode Or to accept either one).
#   * A batch never contains two snapshots of the same VM (vCenter rejects parallel snapshot
#     tasks on one VM); the other snapshots of that VM go into a later batch.
#   * Each snapshot is looked up again right before deletion; if it is already gone it is skipped.
#   * Child snapshots are never removed implicitly (Remove-Snapshot without -RemoveChildren).
#   * If any deletion in a batch fails or times out, the script stops and asks before continuing.
#   * Each successful deletion is verified (snapshot no longer exists) and the VM is checked for
#     "disk consolidation needed".
param(
    [string]$NameKeyword,                                   # word that must appear in the snapshot name
    [string]$DescriptionKeyword,                            # word that must appear in the snapshot description
    [ValidateSet('And', 'Or')][string]$MatchMode = 'And',   # And = both words must match, Or = either one
    [switch]$WholeWord,                                     # match whole words only ('patch' will not match 'patching')
    [string]$FromCsv,                                       # delete exactly the rows of a (reviewed) SelectedSnapshots CSV
    [switch]$Delete,                                        # actually delete; without it the script is read-only
    [ValidateRange(1, 50)][int]$BatchSize = 10,             # snapshots deleted in parallel per batch
    [ValidateRange(1, 1440)][int]$TaskTimeoutMinutes = 180, # max wait for one batch
    [ValidateRange(2, 300)][int]$PollSeconds = 10,          # how often task progress is checked
    [string]$OutputFolder = 'C:\temp\SnapshotCleanup'
)

$ScriptBuild = 'Remove-FilteredSnapshots 2026.10.24-r1'
Write-Host $ScriptBuild -ForegroundColor Cyan

$ErrorActionPreference = 'Stop'

#region ---------- Helpers ----------
function Test-KeywordMatch([string]$Text, [string]$Keyword, [bool]$Whole) {
    # Case-insensitive. Keyword is matched literally (no wildcards / regex characters).
    if ([string]::IsNullOrEmpty($Text)) { return $false }
    if ($Whole) { return [regex]::IsMatch($Text, '(?<![\w])' + [regex]::Escape($Keyword) + '(?![\w])', 'IgnoreCase') }
    return ($Text.IndexOf($Keyword, [StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Test-SnapshotSelected($Row, [string]$NameKw, [string]$DescKw, [string]$Mode, [bool]$Whole) {
    $checks = @()
    if ($NameKw) { $checks += (Test-KeywordMatch $Row.SnapshotName $NameKw $Whole) }
    if ($DescKw) { $checks += (Test-KeywordMatch $Row.Description $DescKw $Whole) }
    if ($checks.Count -eq 0) { return $false }
    if ($Mode -eq 'Or') { return ($checks -contains $true) }
    return (-not ($checks -contains $false))
}

function Get-RowKey($Row) { '{0}|{1}|{2}' -f $Row.VCenter, $Row.VMId, $Row.SnapshotId }

function New-DeleteBatches([object[]]$Items, [int]$Size) {
    # Splits items into batches of at most $Size with at most ONE snapshot per VM per batch.
    # Items keep their order (oldest first), so a VM's older snapshot is removed before its newer one.
    $remaining = New-Object System.Collections.ArrayList
    foreach ($i in @($Items)) { [void]$remaining.Add($i) }
    $batches = New-Object System.Collections.ArrayList
    while ($remaining.Count -gt 0) {
        $batch = New-Object System.Collections.ArrayList
        $vmsInBatch = @{}
        foreach ($item in @($remaining)) {
            if ($batch.Count -ge $Size) { break }
            $vmKey = '{0}|{1}' -f $item.VCenter, $item.VMId
            if ($vmsInBatch.ContainsKey($vmKey)) { continue }
            $vmsInBatch[$vmKey] = $true
            [void]$batch.Add($item)
        }
        foreach ($item in $batch) { $remaining.Remove($item) }
        [void]$batches.Add($batch.ToArray())   # added as ONE element, even when it holds 1 item
    }
    return , @($batches.ToArray())
}

function New-ResultRow($Row, [int]$BatchNo, [string]$Status, [string]$Notes) {
    [pscustomobject]@{
        Time         = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Batch        = $BatchNo
        Status       = $Status
        VCenter      = $Row.VCenter
        VM           = $Row.VM
        SnapshotName = $Row.SnapshotName
        Description  = $Row.Description
        Created      = $Row.Created
        SizeGB       = $Row.SizeGB
        VMId         = $Row.VMId
        SnapshotId   = $Row.SnapshotId
        Notes        = $Notes
    }
}

function Save-Results($Results, [string]$Path) {
    # Rewritten after every batch so a crash/Ctrl+C still leaves an accurate record
    @($Results) | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}
#endregion

#region ---------- Main ----------
function Invoke-SnapshotCleanup {
    # ---- Validate parameters before touching vCenter ----
    if ($FromCsv -and ($NameKeyword -or $DescriptionKeyword)) {
        Write-Host 'Use either -FromCsv or the keyword filters, not both.' -ForegroundColor Red; return
    }
    if ($FromCsv -and -not (Test-Path -LiteralPath $FromCsv)) {
        Write-Host "CSV not found: $FromCsv" -ForegroundColor Red; return
    }
    if ($Delete -and -not $FromCsv -and -not $NameKeyword -and -not $DescriptionKeyword) {
        Write-Host 'Refusing to delete: give -NameKeyword and/or -DescriptionKeyword, or -FromCsv.' -ForegroundColor Red; return
    }
    if ($NameKeyword -and $NameKeyword.Trim().Length -lt 3) { Write-Host '-NameKeyword must be at least 3 characters.' -ForegroundColor Red; return }
    if ($DescriptionKeyword -and $DescriptionKeyword.Trim().Length -lt 3) { Write-Host '-DescriptionKeyword must be at least 3 characters.' -ForegroundColor Red; return }
    if ($NameKeyword) { $NameKeyword = $NameKeyword.Trim() }
    if ($DescriptionKeyword) { $DescriptionKeyword = $DescriptionKeyword.Trim() }

    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    $stamp      = Get-Date -Format 'yyyyMMdd_HHmmss'
    $allCsv     = Join-Path $OutputFolder "AllSnapshots_$stamp.csv"
    $selCsv     = Join-Path $OutputFolder "SelectedSnapshots_$stamp.csv"
    $resultsCsv = Join-Path $OutputFolder "DeleteResults_$stamp.csv"

    # ---- PowerCLI / connections ----
    if (-not (Get-Command Get-Snapshot -ErrorAction SilentlyContinue)) {
        try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue }
        catch { Write-Host 'VMware PowerCLI is not installed.' -ForegroundColor Red; return }
    }
    $servers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
    if ($servers.Count -eq 0) { Write-Host 'Not connected to any vCenter. Run Connect-VIServer first.' -ForegroundColor Red; return }
    Write-Host ("Connected vCenters: {0}" -f (($servers | ForEach-Object { $_.Name }) -join ', '))

    # ---- STEP 1: list all snapshots ----
    $all = New-Object System.Collections.ArrayList
    $scanErrors = 0
    foreach ($srv in $servers) {
        Write-Host "Reading snapshots from $($srv.Name) ..." -ForegroundColor Yellow
        try { $vms = @(Get-VM -Server $srv -ErrorAction Stop) }
        catch { Write-Host "  Cannot list VMs on $($srv.Name): $($_.Exception.Message)" -ForegroundColor Red; $scanErrors++; continue }
        $n = 0
        foreach ($vm in $vms) {
            $n++
            if ($n % 50 -eq 0) { Write-Progress -Activity "Scanning $($srv.Name)" -Status "$n / $($vms.Count) VMs" -PercentComplete ($n * 100 / $vms.Count) }
            try { $snaps = @(Get-Snapshot -VM $vm -Server $srv -ErrorAction Stop) }
            catch { Write-Host "  Cannot read snapshots of $($vm.Name): $($_.Exception.Message)" -ForegroundColor Red; $scanErrors++; continue }
            foreach ($s in $snaps) {
                $parentName = ''
                if ($s.Parent) { $parentName = $s.Parent.Name }
                [void]$all.Add([pscustomobject]@{
                    VCenter      = $srv.Name
                    VM           = $vm.Name
                    PowerState   = [string]$vm.PowerState
                    SnapshotName = [string]$s.Name
                    Description  = [string]$s.Description
                    Created      = $s.Created
                    AgeDays      = [math]::Round(((Get-Date) - $s.Created).TotalDays, 1)
                    SizeGB       = [math]::Round([double]$s.SizeGB, 2)
                    IsCurrent    = [bool]$s.IsCurrent
                    ParentName   = $parentName
                    HasChildren  = (@($s.Children).Count -gt 0)
                    VMId         = $vm.Id
                    SnapshotId   = $s.Id
                })
            }
        }
        Write-Progress -Activity "Scanning $($srv.Name)" -Completed
    }
    $all = @($all | Sort-Object Created)
    $all | Export-Csv -Path $allCsv -NoTypeInformation -Encoding UTF8
    Write-Host ("Total snapshots found: {0}  (saved to {1})" -f $all.Count, $allCsv) -ForegroundColor Green
    if ($scanErrors -gt 0) { Write-Host "WARNING: $scanErrors VM(s)/vCenter(s) could not be read - see messages above." -ForegroundColor Red }

    # ---- STEP 2: filter ----
    if ($FromCsv) {
        $wanted = @(Import-Csv -LiteralPath $FromCsv)
        foreach ($col in 'VCenter', 'VMId', 'SnapshotId') {
            if ($wanted.Count -gt 0 -and -not ($wanted[0].PSObject.Properties.Name -contains $col)) {
                Write-Host "CSV is missing column '$col'. Use a SelectedSnapshots_*.csv produced by this script." -ForegroundColor Red; return
            }
        }
        $live = @{}
        foreach ($r in $all) { $live[(Get-RowKey $r)] = $r }
        $selected = New-Object System.Collections.ArrayList
        $notFound = @()
        $seen = @{}
        foreach ($w in $wanted) {
            $k = Get-RowKey $w
            if ($seen.ContainsKey($k)) { continue }
            $seen[$k] = $true
            if ($live.ContainsKey($k)) { [void]$selected.Add($live[$k]) } else { $notFound += $w }
        }
        $selected = @($selected | Sort-Object Created)
        Write-Host ("Rows in CSV: {0}   still exist: {1}   already gone / not found: {2}" -f $wanted.Count, $selected.Count, $notFound.Count)
        foreach ($nf in $notFound) { Write-Host "  not found (skipped): $($nf.VM) / $($nf.SnapshotName)" -ForegroundColor DarkYellow }
    }
    elseif ($NameKeyword -or $DescriptionKeyword) {
        $selected = @($all | Where-Object { Test-SnapshotSelected $_ $NameKeyword $DescriptionKeyword $MatchMode $WholeWord.IsPresent })
        $desc = @()
        if ($NameKeyword) { $desc += "name contains '$NameKeyword'" }
        if ($DescriptionKeyword) { $desc += "description contains '$DescriptionKeyword'" }
        $joiner = ' AND '
        if ($MatchMode -eq 'Or') { $joiner = ' OR ' }
        $wholeTxt = ''
        if ($WholeWord) { $wholeTxt = ' (whole word)' }
        Write-Host ("Filter: {0}{1}" -f ($desc -join $joiner), $wholeTxt)
    }
    else {
        Write-Host 'No filter given - list only. Add -NameKeyword / -DescriptionKeyword to select snapshots.' -ForegroundColor Yellow
        return
    }

    $selected | Export-Csv -Path $selCsv -NoTypeInformation -Encoding UTF8
    $totalGB = [math]::Round((@($selected | Measure-Object SizeGB -Sum).Sum) + 0, 2)
    Write-Host ("Selected snapshots: {0} on {1} VM(s), {2} GB  (saved to {3})" -f $selected.Count,
        @($selected | Select-Object -Property VCenter, VMId -Unique).Count, $totalGB, $selCsv) -ForegroundColor Green
    if ($selected.Count -eq 0) { Write-Host 'Nothing matched - nothing to delete.' -ForegroundColor Yellow; return }
    $selected | Format-Table VCenter, VM, SnapshotName, Description, Created, SizeGB -AutoSize | Out-String -Width 250 | Write-Host

    if (-not $Delete) {
        Write-Host 'READ-ONLY run: nothing was deleted. Review the CSV above, then run again with -Delete' -ForegroundColor Cyan
        Write-Host "  (or: -FromCsv '$selCsv' -Delete  after removing any rows you want to keep)" -ForegroundColor Cyan
        return
    }

    # ---- STEP 3: delete in batches ----
    $batches = New-DeleteBatches $selected $BatchSize
    Write-Host ''
    Write-Host ("About to DELETE {0} snapshot(s) in {1} batch(es) of up to {2}." -f $selected.Count, $batches.Count, $BatchSize) -ForegroundColor Red
    $answer = Read-Host 'Type DELETE to continue (anything else cancels)'
    if ($answer -cne 'DELETE') { Write-Host 'Cancelled - nothing was deleted.' -ForegroundColor Yellow; return }

    $results = New-Object System.Collections.ArrayList
    $serverByName = @{}
    foreach ($srv in $servers) { $serverByName[$srv.Name] = $srv }
    $batchNo = 0

    foreach ($batch in $batches) {
        $batchNo++
        Write-Host ''
        Write-Host ("===== Batch {0} of {1} ({2} snapshot(s)) =====" -f $batchNo, $batches.Count, $batch.Count) -ForegroundColor Cyan
        $running = @()   # @{ Row; Task; Server }

        foreach ($row in $batch) {
            $srv = $serverByName[$row.VCenter]
            if (-not $srv -or -not $srv.IsConnected) { [void]$results.Add((New-ResultRow $row $batchNo 'Failed' 'vCenter no longer connected')); continue }
            try {
                # Look the snapshot up again right now - it may have been removed since the scan
                $vmNow = Get-VM -Id $row.VMId -Server $srv -ErrorAction Stop
                $snapNow = @(Get-Snapshot -VM $vmNow -Server $srv -ErrorAction Stop | Where-Object { $_.Id -eq $row.SnapshotId })
                if ($snapNow.Count -eq 0) {
                    Write-Host "  [skip] $($row.VM) / $($row.SnapshotName) - already gone" -ForegroundColor DarkYellow
                    [void]$results.Add((New-ResultRow $row $batchNo 'Skipped' 'Snapshot no longer exists'))
                    continue
                }
                $task = Remove-Snapshot -Snapshot $snapNow[0] -RunAsync -Confirm:$false -ErrorAction Stop
                Write-Host "  [start] $($row.VM) / $($row.SnapshotName)"
                $running += @{ Row = $row; Task = $task; Server = $srv }
            }
            catch {
                Write-Host "  [FAIL] $($row.VM) / $($row.SnapshotName): $($_.Exception.Message)" -ForegroundColor Red
                [void]$results.Add((New-ResultRow $row $batchNo 'Failed' "Could not start removal: $($_.Exception.Message)"))
            }
        }

        # Wait until every task of this batch is finished (or timeout)
        $deadline = (Get-Date).AddMinutes($TaskTimeoutMinutes)
        $pending = @($running)
        $batchProblem = $false
        while ($pending.Count -gt 0) {
            $still = @()
            foreach ($p in $pending) {
                try { $t = Get-Task -Id $p.Task.Id -Server $p.Server -ErrorAction Stop }
                catch { $still += $p; continue }   # transient lookup error - try again next poll
                $state = [string]$t.State
                if ($state -eq 'Success') {
                    $notes = 'Removed'
                    try {
                        $vmAfter = Get-VM -Id $p.Row.VMId -Server $p.Server -ErrorAction Stop
                        $left = @(Get-Snapshot -VM $vmAfter -Server $p.Server -ErrorAction Stop | Where-Object { $_.Id -eq $p.Row.SnapshotId })
                        if ($left.Count -gt 0) { $notes = 'Task succeeded but snapshot still listed - check manually' }
                        elseif ($vmAfter.ExtensionData.Runtime.ConsolidationNeeded) { $notes = 'Removed - VM reports DISK CONSOLIDATION NEEDED' }
                    }
                    catch { $notes = "Removed (verification failed: $($_.Exception.Message))" }
                    $status = 'Deleted'
                    if ($notes -ne 'Removed') { $status = 'Deleted-CheckVM'; $batchProblem = $true }
                    Write-Host "  [done] $($p.Row.VM) / $($p.Row.SnapshotName) - $notes" -ForegroundColor Green
                    [void]$results.Add((New-ResultRow $p.Row $batchNo $status $notes))
                }
                elseif ($state -eq 'Error') {
                    $msg = ''
                    try { $msg = [string]$t.ExtensionData.Info.Error.LocalizedMessage } catch { }
                    if (-not $msg) { $msg = [string]$t.Description }
                    Write-Host "  [FAIL] $($p.Row.VM) / $($p.Row.SnapshotName): $msg" -ForegroundColor Red
                    [void]$results.Add((New-ResultRow $p.Row $batchNo 'Failed' "vCenter task error: $msg"))
                    $batchProblem = $true
                }
                else { $still += $p }   # Queued / Running
            }
            $pending = @($still)
            if ($pending.Count -eq 0) { break }
            if ((Get-Date) -gt $deadline) {
                foreach ($p in $pending) {
                    Write-Host "  [TIMEOUT] $($p.Row.VM) / $($p.Row.SnapshotName) - task still running in vCenter" -ForegroundColor Red
                    [void]$results.Add((New-ResultRow $p.Row $batchNo 'StillRunning' "Not finished after $TaskTimeoutMinutes min - task continues in vCenter, check Recent Tasks"))
                }
                $batchProblem = $true
                break
            }
            Write-Host ("  waiting for {0} task(s) ..." -f $pending.Count) -ForegroundColor DarkGray
            Start-Sleep -Seconds $PollSeconds
        }

        Save-Results $results $resultsCsv
        $fails = @($results | Where-Object { $_.Batch -eq $batchNo -and $_.Status -ne 'Deleted' -and $_.Status -ne 'Skipped' }).Count
        if ($fails -gt 0) { $batchProblem = $true }
        if ($batchProblem -and $batchNo -lt $batches.Count) {
            Write-Host "Batch $batchNo had $fails problem(s) - see above." -ForegroundColor Red
            $go = Read-Host 'Continue with the next batch? (Y/N)'
            if ($go -notmatch '^(y|yes)$') {
                Write-Host 'Stopped by user. Remaining snapshots were NOT touched.' -ForegroundColor Yellow
                $done = @{}
                foreach ($r in $results) { $done[(Get-RowKey $r)] = $true }
                foreach ($b in $batches) { foreach ($r in $b) { if (-not $done.ContainsKey((Get-RowKey $r))) { [void]$results.Add((New-ResultRow $r 0 'NotAttempted' 'Run stopped before this batch')) } } }
                break
            }
        }
    }

    Save-Results $results $resultsCsv
    Write-Host ''
    Write-Host '===== Summary =====' -ForegroundColor Cyan
    $results | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("  {0,-16} {1}" -f $_.Name, $_.Count) }
    Write-Host "Results saved to $resultsCsv" -ForegroundColor Green
}

$logFile = $null
if ($OutputFolder) {
    try {
        if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
        $logFile = Join-Path $OutputFolder ("SnapshotCleanup_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        Start-Transcript -Path $logFile -Force | Out-Null
    } catch { $logFile = $null }
}
try { Invoke-SnapshotCleanup }
catch { Write-Host "UNEXPECTED ERROR: $($_.Exception.Message)" -ForegroundColor Red; Write-Host $_.ScriptStackTrace }
finally { if ($logFile) { try { Stop-Transcript | Out-Null; Write-Host "Log: $logFile" } catch { } } }
#endregion
