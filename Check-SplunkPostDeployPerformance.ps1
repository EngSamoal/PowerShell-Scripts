<#
Check-SplunkPostDeployPerformance.ps1

Read-only sanity check to run shortly after Deploy-SplunkUniversalForwarder.ps1 has
installed/upgraded the Splunk Universal Forwarder on a batch of VMs. It never installs,
starts/stops, or changes anything - it only reads:
  - vCenter-side VM performance counters (CPU/memory %) via Get-Stat.
  - Guest-side splunkd process CPU/memory and service status via Invoke-VMScript
    (same guest-credential/guest-ops approach as the deploy script - no WinRM).

This is a SANITY CHECK, not a full performance audit: it catches an obvious problem
(service not running, process missing, VM-level CPU/memory pegged) in the few minutes
right after a deployment. Real performance impact from a UF is normally negligible and,
if it does show up, tends to do so over hours - for real assurance, keep an eye on these
VMs over the following day through vCenter/your normal monitoring, not just this script.

Requirements:
  - Already connected to the target vCenter via Connect-VIServer before running this script.
  - Same vmlist.txt and guest credential (wincred.xml, via Export-Clixml) used by the
    deploy script.
#>

param(
    [string]$VmListPath = 'C:\temp\vmlist.txt',
    [string]$GuestCredentialPath = 'C:\temp\wincred.xml',
    [int]$MinutesBack = 15,
    [double]$CpuWarnPercent = 80,
    [double]$MemWarnPercent = 90,
    [string]$SplunkProcessName = 'splunkd',
    [string]$SplunkServiceName = 'SplunkForwarder',
    [string]$ReportFolder = 'C:\temp'
)

$ScriptBuild = '2026.09.22-1'
Write-Host "Check-SplunkPostDeployPerformance.ps1 - build $ScriptBuild" -ForegroundColor Cyan
Write-Host "READ-ONLY sanity check - nothing on any VM will be installed, started, stopped, or changed.`n" -ForegroundColor Cyan

if (-not $global:DefaultVIServer -or -not $global:DefaultVIServer.IsConnected) {
    throw "No active vCenter connection found. Connect with Connect-VIServer before running this script."
}
if (-not (Test-Path -LiteralPath $VmListPath)) {
    throw "VM list not found: $VmListPath"
}
if (-not (Test-Path -LiteralPath $GuestCredentialPath)) {
    throw "Guest credential file not found: $GuestCredentialPath"
}
if (-not (Test-Path -LiteralPath $ReportFolder)) {
    New-Item -ItemType Directory -Path $ReportFolder -Force -ErrorAction Stop | Out-Null
}

function Add-RowNote {
    param($Row, [string]$Note)
    $Row.Notes = if ($Row.Notes) { "$($Row.Notes) | $Note" } else { $Note }
}

$guestCred = Import-Clixml -Path $GuestCredentialPath
$vmNames = Get-Content -Path $VmListPath | Where-Object { $_.Trim() -ne '' }
if (-not $vmNames) {
    throw "VM list at $VmListPath contains no VM names."
}

$statStart = (Get-Date).AddMinutes(-1 * [Math]::Abs($MinutesBack))

# Guest-side check is read-only: Get-Process / Get-Service only, nothing is installed,
# started, stopped, or modified.
$GuestCheckTemplate = @'
$result = [ordered]@{
    ProcessFound  = $false
    ProcessCpuSec = $null
    ProcessMemMB  = $null
    ServiceExists = $false
    ServiceStatus = $null
    Error         = $null
}
try {
    $proc = Get-Process -Name '__PROCNAME__' -ErrorAction SilentlyContinue
    if ($proc) {
        $result.ProcessFound  = $true
        $result.ProcessCpuSec = [Math]::Round((($proc | Measure-Object -Property CPU -Sum).Sum), 1)
        $result.ProcessMemMB  = [Math]::Round((($proc | Measure-Object -Property WorkingSet64 -Sum).Sum / 1MB), 1)
    }
    $svc = Get-Service -Name '__SVCNAME__' -ErrorAction SilentlyContinue
    if ($svc) {
        $result.ServiceExists = $true
        $result.ServiceStatus = $svc.Status.ToString()
    }
} catch {
    $result.Error = $_.Exception.Message
}
[PSCustomObject]$result | ConvertTo-Json -Compress
'@
$guestCheckScript = $GuestCheckTemplate.Replace('__PROCNAME__', $SplunkProcessName).Replace('__SVCNAME__', $SplunkServiceName)

$results = [System.Collections.Generic.List[object]]::new()

foreach ($vmName in $vmNames) {
    Write-Host "=== $vmName ===" -ForegroundColor Cyan
    $row = [ordered]@{
        VMName        = $vmName
        PowerState    = $null
        CpuAvgPercent = $null
        MemAvgPercent = $null
        ProcessFound  = $null
        ProcessCpuSec = $null
        ProcessMemMB  = $null
        ServiceStatus = $null
        Flag          = 'OK'
        Notes         = ''
    }

    try {
        $vm = Get-VM -Name $vmName -ErrorAction Stop
        $row.PowerState = $vm.PowerState.ToString()

        if ($vm.PowerState -ne 'PoweredOn') {
            $row.Flag = 'Skipped'
            $row.Notes = 'VM is not powered on - cannot read live performance stats.'
            $results.Add([PSCustomObject]$row)
            Write-Host "  Skipped (not powered on)." -ForegroundColor DarkGray
            continue
        }

        # vCenter-side performance counters - read-only, does not touch the guest.
        $cpuStat = Get-Stat -Entity $vm -Stat 'cpu.usage.average' -Start $statStart -Realtime -ErrorAction SilentlyContinue
        $memStat = Get-Stat -Entity $vm -Stat 'mem.usage.average' -Start $statStart -Realtime -ErrorAction SilentlyContinue

        if ($cpuStat) { $row.CpuAvgPercent = [Math]::Round((($cpuStat | Measure-Object -Property Value -Average).Average), 1) }
        else { Add-RowNote -Row $row -Note 'Could not read vCenter CPU performance counter for this VM.' }

        if ($memStat) { $row.MemAvgPercent = [Math]::Round((($memStat | Measure-Object -Property Value -Average).Average), 1) }
        else { Add-RowNote -Row $row -Note 'Could not read vCenter memory performance counter for this VM.' }

        # Guest-side splunkd process + service check - reuses the same guest credential
        # the deploy script uses.
        $guestResult = Invoke-VMScript -VM $vm -ScriptText $guestCheckScript -ScriptType Powershell -GuestCredential $guestCred -ErrorAction Stop
        $parsed = $null
        try { $parsed = $guestResult.ScriptOutput | ConvertFrom-Json } catch {}

        if ($parsed) {
            $row.ProcessFound  = $parsed.ProcessFound
            $row.ProcessCpuSec = $parsed.ProcessCpuSec
            $row.ProcessMemMB  = $parsed.ProcessMemMB
            $row.ServiceStatus = $parsed.ServiceStatus
            if ($parsed.Error) { Add-RowNote -Row $row -Note "Guest-side check error: $($parsed.Error)" }
        } else {
            Add-RowNote -Row $row -Note "Could not parse guest-side process/service check output. Raw: $($guestResult.ScriptOutput)"
        }

        # Simple threshold flags - meant to catch a gross problem, not fine-tune anything.
        if ($null -ne $row.CpuAvgPercent -and $row.CpuAvgPercent -ge $CpuWarnPercent) {
            $row.Flag = 'Review'
            Add-RowNote -Row $row -Note "VM-level CPU averaged $($row.CpuAvgPercent)% over the last $MinutesBack min (threshold $CpuWarnPercent%)."
        }
        if ($null -ne $row.MemAvgPercent -and $row.MemAvgPercent -ge $MemWarnPercent) {
            $row.Flag = 'Review'
            Add-RowNote -Row $row -Note "VM-level memory averaged $($row.MemAvgPercent)% over the last $MinutesBack min (threshold $MemWarnPercent%)."
        }
        if ($row.ServiceStatus -and $row.ServiceStatus -ne 'Running') {
            $row.Flag = 'Review'
            Add-RowNote -Row $row -Note "Splunk service status is '$($row.ServiceStatus)', not Running."
        }
        if ($row.ProcessFound -eq $false) {
            $row.Flag = 'Review'
            Add-RowNote -Row $row -Note "Process '$SplunkProcessName' was not found running on this VM."
        }

    } catch {
        $row.Flag = 'Error'
        $row.Notes = "Check failed: $($_.Exception.Message)"
    }

    $color = switch ($row.Flag) {
        'Review'  { 'Yellow' }
        'Error'   { 'Red' }
        'Skipped' { 'DarkGray' }
        default   { 'Green' }
    }
    Write-Host "  Flag: $($row.Flag)  CPU avg: $($row.CpuAvgPercent)%  Mem avg: $($row.MemAvgPercent)%  Service: $($row.ServiceStatus)" -ForegroundColor $color
    if ($row.Notes) { Write-Host "  Notes: $($row.Notes)" -ForegroundColor Yellow }

    $results.Add([PSCustomObject]$row)
}

$reportPath = Join-Path $ReportFolder ("SplunkPostDeployCheck_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nReport written to: $reportPath" -ForegroundColor Cyan

# Plain loop counters, not "(Where-Object {...}).Count" - a single matching object in
# Windows PowerShell 5.1 returns .Count = $null instead of 1 (same bug class documented
# in this repo for VMware_Weekly_HealthCheck.ps1 and hit live in Deploy-SplunkUniversalForwarder.ps1).
$okCount = 0; $reviewCount = 0; $errorCount = 0; $skippedCount = 0
foreach ($r in $results) {
    switch ($r.Flag) {
        'OK'      { $okCount++ }
        'Review'  { $reviewCount++ }
        'Error'   { $errorCount++ }
        'Skipped' { $skippedCount++ }
    }
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
Write-Host ("{0,-20}: {1}" -f 'Total VMs checked', $results.Count)
Write-Host ("{0,-20}: {1}" -f 'OK', $okCount)
Write-Host ("{0,-20}: {1}" -f 'Flagged for review', $reviewCount)
Write-Host ("{0,-20}: {1}" -f 'Errors', $errorCount)
Write-Host ("{0,-20}: {1}" -f 'Skipped', $skippedCount)

if ($reviewCount -gt 0 -or $errorCount -gt 0) {
    Write-Host "`nOne or more VMs need a look - see the Flag/Notes columns in $reportPath." -ForegroundColor Yellow
}
