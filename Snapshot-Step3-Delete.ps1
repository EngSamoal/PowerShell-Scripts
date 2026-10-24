# Snapshot-Step3-Delete.ps1
# STEP 3 of 3 - Deletes the snapshots listed in the Step 2 file, 10 at a time.
#
# Run     : .\Snapshot-Step3-Delete.ps1
#           It shows what will be deleted and asks you to type DELETE before doing anything.
# Input   : C:\temp\SnapshotCleanup\2-SnapshotsToDelete.csv   (from Step 2, after your review)
# Output  : C:\temp\SnapshotCleanup\3-DeleteResults_<time>.csv  (result of every snapshot)
#           C:\temp\SnapshotCleanup\3-DeleteLog_<time>.log      (full screen log)
#
# How it works:
#   * Only snapshots in the CSV are deleted. Each one is looked up again in vCenter first;
#     if it no longer exists it is skipped.
#   * 10 deletions run at the same time; the script waits until all 10 are finished before
#     starting the next 10. Two snapshots of the same VM are never deleted at the same time.
#   * Child snapshots are never deleted together with their parent.
#   * If a deletion fails, the script stops and asks if it should continue.
#   * After each deletion it checks the snapshot is really gone and if the VM needs disk consolidation.
param(
    [ValidateRange(1, 50)][int]$BatchSize = 10,             # deletions at the same time
    [ValidateRange(1, 1440)][int]$TaskTimeoutMinutes = 180, # max wait for one batch
    [ValidateRange(2, 300)][int]$PollSeconds = 10,          # how often progress is checked
    [string]$OutputFolder = 'C:\temp\SnapshotCleanup'
)

$ScriptBuild = 'Snapshot-Step3-Delete 2026.10.24-r2'
Write-Host $ScriptBuild -ForegroundColor Cyan
$ErrorActionPreference = 'Stop'

function Get-RowKey($Row) { '{0}|{1}|{2}' -f $Row.VCenter, $Row.VMId, $Row.SnapshotId }

function New-DeleteBatches([object[]]$Items, [int]$Size) {
    # Batches of at most $Size, never two snapshots of the same VM in one batch.
    # Order is kept (oldest first), so a VM's older snapshot goes before its newer one.
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

function Invoke-SnapshotDelete {
    $inCsv = Join-Path $OutputFolder '2-SnapshotsToDelete.csv'
    if (-not (Test-Path -LiteralPath $inCsv)) { Write-Host "Not found: $inCsv  - run Step 1 and Step 2 first." -ForegroundColor Red; return }
    $resultsCsv = Join-Path $OutputFolder ("3-DeleteResults_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    # ---- Read and check the list ----
    $rows = @(Import-Csv -LiteralPath $inCsv)
    if ($rows.Count -eq 0) { Write-Host "$inCsv is empty - nothing to delete." -ForegroundColor Yellow; return }
    foreach ($col in 'VCenter', 'VM', 'SnapshotName', 'VMId', 'SnapshotId') {
        if (-not ($rows[0].PSObject.Properties.Name -contains $col)) {
            Write-Host "The CSV is missing column '$col'. When editing in Excel only delete whole rows, and save as CSV." -ForegroundColor Red; return
        }
    }
    $list = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($r in $rows) {
        if (-not $r.VCenter -or -not $r.VMId -or -not $r.SnapshotId) { continue }   # blank rows left by Excel
        $k = Get-RowKey $r
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true
        [void]$list.Add($r)
    }
    $list = @($list | Sort-Object Created)
    if ($list.Count -eq 0) { Write-Host 'No valid rows in the CSV - nothing to delete.' -ForegroundColor Yellow; return }

    # ---- vCenter connections ----
    if (-not (Get-Command Remove-Snapshot -ErrorAction SilentlyContinue)) {
        try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue }
        catch { Write-Host 'VMware PowerCLI is not installed.' -ForegroundColor Red; return }
    }
    $serverByName = @{}
    foreach ($srv in @($global:DefaultVIServers | Where-Object { $_.IsConnected })) { $serverByName[$srv.Name] = $srv }
    if ($serverByName.Count -eq 0) { Write-Host 'Not connected to any vCenter. Run Connect-VIServer first.' -ForegroundColor Red; return }
    $missing = @($list | ForEach-Object { $_.VCenter } | Select-Object -Unique | Where-Object { -not $serverByName.ContainsKey($_) })
    if ($missing.Count -gt 0) {
        Write-Host ("Not connected to: {0}. Connect to it (Connect-VIServer) and run again." -f ($missing -join ', ')) -ForegroundColor Red; return
    }

    # ---- Confirm ----
    $batches = New-DeleteBatches $list $BatchSize
    $list | Format-Table VCenter, VM, SnapshotName, Description, Created, SizeGB -AutoSize | Out-String -Width 250 | Write-Host
    Write-Host ("About to DELETE {0} snapshot(s) in {1} batch(es) of up to {2}." -f $list.Count, $batches.Count, $BatchSize) -ForegroundColor Red
    $answer = Read-Host 'Type DELETE to continue (anything else cancels)'
    if ($answer -cne 'DELETE') { Write-Host 'Cancelled - nothing was deleted.' -ForegroundColor Yellow; return }

    # ---- Delete batch by batch ----
    $results = New-Object System.Collections.ArrayList
    $batchNo = 0
    foreach ($batch in $batches) {
        $batchNo++
        Write-Host ''
        Write-Host ("===== Batch {0} of {1} ({2} snapshot(s)) =====" -f $batchNo, $batches.Count, $batch.Count) -ForegroundColor Cyan
        $running = @()

        foreach ($row in $batch) {
            $srv = $serverByName[$row.VCenter]
            try {
                # Look the snapshot up again right now - it may have been removed since Step 1
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

        # Wait until every task of this batch has finished (or timeout)
        $deadline = (Get-Date).AddMinutes($TaskTimeoutMinutes)
        $pending = @($running)
        while ($pending.Count -gt 0) {
            $still = @()
            foreach ($p in $pending) {
                try { $t = Get-Task -Id $p.Task.Id -Server $p.Server -ErrorAction Stop }
                catch { $still += $p; continue }   # temporary lookup error - try again next time
                $state = [string]$t.State
                if ($state -eq 'Success') {
                    $notes = 'Removed'
                    try {
                        $vmAfter = Get-VM -Id $p.Row.VMId -Server $p.Server -ErrorAction Stop
                        $left = @(Get-Snapshot -VM $vmAfter -Server $p.Server -ErrorAction Stop | Where-Object { $_.Id -eq $p.Row.SnapshotId })
                        if ($left.Count -gt 0) { $notes = 'Task succeeded but snapshot still listed - check manually' }
                        elseif ($vmAfter.ExtensionData.Runtime.ConsolidationNeeded) { $notes = 'Removed - VM reports DISK CONSOLIDATION NEEDED' }
                    }
                    catch { $notes = "Removed (could not verify: $($_.Exception.Message))" }
                    $status = 'Deleted'
                    if ($notes -ne 'Removed') { $status = 'Deleted-CheckVM' }
                    Write-Host "  [done] $($p.Row.VM) / $($p.Row.SnapshotName) - $notes" -ForegroundColor Green
                    [void]$results.Add((New-ResultRow $p.Row $batchNo $status $notes))
                }
                elseif ($state -eq 'Error') {
                    $msg = ''
                    try { $msg = [string]$t.ExtensionData.Info.Error.LocalizedMessage } catch { }
                    if (-not $msg) { $msg = [string]$t.Description }
                    Write-Host "  [FAIL] $($p.Row.VM) / $($p.Row.SnapshotName): $msg" -ForegroundColor Red
                    [void]$results.Add((New-ResultRow $p.Row $batchNo 'Failed' "vCenter task error: $msg"))
                }
                else { $still += $p }   # Queued / Running
            }
            $pending = @($still)
            if ($pending.Count -eq 0) { break }
            if ((Get-Date) -gt $deadline) {
                foreach ($p in $pending) {
                    Write-Host "  [TIMEOUT] $($p.Row.VM) / $($p.Row.SnapshotName) - still running in vCenter" -ForegroundColor Red
                    [void]$results.Add((New-ResultRow $p.Row $batchNo 'StillRunning' "Not finished after $TaskTimeoutMinutes min - task continues in vCenter, check Recent Tasks"))
                }
                break
            }
            Write-Host ("  waiting for {0} task(s) ..." -f $pending.Count) -ForegroundColor DarkGray
            Start-Sleep -Seconds $PollSeconds
        }

        # Saved after every batch so the file is correct even if the run is interrupted
        @($results) | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8
        $problems = @($results | Where-Object { $_.Batch -eq $batchNo -and $_.Status -ne 'Deleted' -and $_.Status -ne 'Skipped' }).Count
        if ($problems -gt 0 -and $batchNo -lt $batches.Count) {
            Write-Host "Batch $batchNo had $problems problem(s) - see above." -ForegroundColor Red
            $go = Read-Host 'Continue with the next batch? (Y/N)'
            if ($go -notmatch '^(y|yes)$') {
                Write-Host 'Stopped. The remaining snapshots were NOT touched.' -ForegroundColor Yellow
                $done = @{}
                foreach ($r in $results) { $done[(Get-RowKey $r)] = $true }
                foreach ($b in $batches) { foreach ($r in $b) { if (-not $done.ContainsKey((Get-RowKey $r))) { [void]$results.Add((New-ResultRow $r 0 'NotAttempted' 'Run stopped before this batch')) } } }
                break
            }
        }
    }

    @($results) | Export-Csv -Path $resultsCsv -NoTypeInformation -Encoding UTF8
    Write-Host ''
    Write-Host '===== Summary =====' -ForegroundColor Cyan
    $results | Group-Object Status | Sort-Object Name | ForEach-Object { Write-Host ("  {0,-16} {1}" -f $_.Name, $_.Count) }
    Write-Host "Results saved to: $resultsCsv" -ForegroundColor Green
}

$logFile = $null
try {
    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
    $logFile = Join-Path $OutputFolder ("3-DeleteLog_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Start-Transcript -Path $logFile -Force | Out-Null
} catch { $logFile = $null }
try { Invoke-SnapshotDelete }
catch { Write-Host "UNEXPECTED ERROR: $($_.Exception.Message)" -ForegroundColor Red; Write-Host $_.ScriptStackTrace }
finally { if ($logFile) { try { Stop-Transcript | Out-Null; Write-Host "Log: $logFile" } catch { } } }
