<#
Check-SplunkPostDeployPerformance.ps1

Read-only sanity check to run shortly after Deploy-SplunkUniversalForwarder.ps1 has
installed/upgraded the Splunk Universal Forwarder on a batch of VMs. It never installs,
starts/stops, or changes anything - it only reads:
  - vCenter-side VM performance counters (CPU/memory %) via Get-Stat.
  - Guest-side splunkd process CPU/memory and service status via Invoke-VMScript
    (same guest-credential/guest-ops approach as the deploy script - no WinRM).

This is a SANITY CHECK, not a full performance audit: it catches an obvious problem
(service not running, process missing, VM-level CPU/memory pegged) in the hours
right after a deployment (3 hours back by default - see -MinutesBack). Real performance
impact from a UF is normally negligible and, if it does show up, tends to do so over
longer than that - for real assurance, keep an eye on these VMs over the following day
through vCenter/your normal monitoring, not just this script.

Requirements:
  - Already connected to the target vCenter via Connect-VIServer before running this script.
  - Same vmlist.txt and guest credential (wincred.xml, via Export-Clixml) used by the
    deploy script.
#>

param(
    [string]$VmListPath = 'C:\temp\vmlist.txt',
    [string]$GuestCredentialPath = 'C:\temp\wincred.xml',
    [int]$MinutesBack = 180,
    [double]$CpuWarnPercent = 80,
    [double]$MemWarnPercent = 90,
    [string]$SplunkProcessName = 'splunkd',
    [string]$SplunkServiceName = 'SplunkForwarder',
    [string]$ReportFolder = 'C:\temp'
)

$ScriptBuild = '2026.09.22-3'
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

function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
}

# Same visual language as Deploy-SplunkUniversalForwarder.ps1 / Generate-PatchingDashboard.ps1
# elsewhere in this repo (green gradient header, KPI cards, colored status badges) so reports
# across this repo look consistent.
function Build-PostDeployDashboardHtml {
    param(
        [Parameter(Mandatory)] [System.Collections.IEnumerable] $Results,
        [Parameter(Mandatory)] [int] $OkCount,
        [Parameter(Mandatory)] [int] $ReviewCount,
        [Parameter(Mandatory)] [int] $ErrorCount,
        [Parameter(Mandatory)] [int] $SkippedCount,
        [Parameter(Mandatory)] [int] $MinutesBack,
        [Parameter(Mandatory)] [string] $ScriptBuild
    )

    $css = @'
  :root{--ink:#1F2933;--line:#E3E8EF;--muted:#6B7683;--bg:#F4F6F9;}
  *{box-sizing:border-box;}
  body{margin:0;background:var(--bg);color:var(--ink);
       font-family:-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;line-height:1.45;}
  .wrap{max-width:1280px;margin:0 auto;padding:28px;}
  header.hero{background:linear-gradient(135deg,#0B2A1B 0%,#146339 60%,#1FA85C 100%);color:#fff;
       border-radius:14px;padding:34px 40px;box-shadow:0 10px 30px rgba(0,0,0,.25);}
  .report-h1{font-size:32px;font-weight:800;letter-spacing:.4px;margin:0;line-height:1.15;}
  .report-sub{font-size:16px;font-weight:600;margin:10px 0 0;color:#CFEFD9;}
  .mode-badge{display:inline-block;margin-top:14px;padding:5px 14px;border-radius:20px;
       font-size:12px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;background:#fff;color:#146339;}
  .kpi{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin:22px 0 10px;}
  .kpi-card{position:relative;background:#fff;border:1px solid var(--line);border-radius:12px;
       padding:16px 14px 14px;overflow:hidden;box-shadow:0 2px 6px rgba(31,41,51,.04);
       display:flex;flex-direction:column;}
  .kpi-accent{position:absolute;top:0;left:0;right:0;height:5px;}
  .kpi-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);
       min-height:28px;line-height:14px;}
  .kpi-value{font-size:30px;font-weight:800;margin-top:6px;}
  .panel{background:#fff;border:1px solid var(--line);border-radius:12px;padding:20px 22px;margin-top:18px;
       box-shadow:0 2px 6px rgba(31,41,51,.04);}
  h2{font-size:15px;text-transform:uppercase;letter-spacing:.7px;margin:0 0 16px;color:#334155;}
  table.ex-table{width:100%;border-collapse:collapse;font-size:12.5px;}
  .ex-table th{background:#0F2A43;color:#fff;text-align:left;padding:9px 10px;font-weight:600;white-space:nowrap;}
  .ex-table td{padding:8px 10px;border-bottom:1px solid var(--line);}
  .ex-table tr:nth-child(even){background:#FAFBFD;}
  .badge{display:inline-block;width:100px;text-align:center;padding:4px 0;border-radius:4px;font-size:11.5px;font-weight:700;color:#fff;white-space:nowrap;}
  .badge.b-green{background:#1F8A4C;} .badge.b-red{background:#C0392B;} .badge.b-amber{background:#C77700;} .badge.b-grey{background:#6B7683;}
  .notes-cell{color:var(--muted);font-size:11.5px;max-width:280px;}
  footer{margin-top:24px;font-size:12px;color:var(--muted);}
  @media (max-width:1100px){.kpi{grid-template-columns:repeat(3,1fr);}}
  @media (max-width:640px){.kpi{grid-template-columns:repeat(2,1fr);}}
'@

    function Get-KpiCard($label, $value, $color) {
        return "  <div class=`"kpi-card`"><div class=`"kpi-accent`" style=`"background:$color`"></div><div class=`"kpi-label`">$(ConvertTo-HtmlSafe $label)</div><div class=`"kpi-value`" style=`"color:$color`">$value</div></div>`n"
    }

    $totalVMs = $OkCount + $ReviewCount + $ErrorCount + $SkippedCount
    $kpi = ''
    $kpi += Get-KpiCard 'Total VMs' $totalVMs '#1D4E79'
    $kpi += Get-KpiCard 'OK' $OkCount '#1F8A4C'
    $kpi += Get-KpiCard 'Flagged for Review' $ReviewCount '#C77700'
    $kpi += Get-KpiCard 'Errors' $ErrorCount '#C0392B'
    $kpi += Get-KpiCard 'Skipped' $SkippedCount '#6B7683'

    function Get-FlagBadgeClass($flag) {
        switch ($flag) {
            'Review'  { return 'b-amber' }
            'Error'   { return 'b-red' }
            'Skipped' { return 'b-grey' }
            default   { return 'b-green' }
        }
    }

    $rows = ''
    foreach ($r in $Results) {
        $badgeClass = Get-FlagBadgeClass $r.Flag
        $cpuText = if ($null -ne $r.CpuAvgPercent) { "$($r.CpuAvgPercent)%" } else { '-' }
        $memText = if ($null -ne $r.MemAvgPercent) { "$($r.MemAvgPercent)%" } else { '-' }
        $procFoundText = if ($null -eq $r.ProcessFound) { '-' } elseif ($r.ProcessFound) { 'Yes' } else { 'No' }
        $procCpuText = if ($null -ne $r.ProcessCpuSec) { $r.ProcessCpuSec } else { '-' }
        $procMemText = if ($null -ne $r.ProcessMemMB) { "$($r.ProcessMemMB) MB" } else { '-' }
        $rows += "          <tr>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.VMName)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.PowerState)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $cpuText)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $memText)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $procFoundText)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $procCpuText)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $procMemText)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.ServiceStatus)</td>`n" +
            "            <td><span class=`"badge $badgeClass`">$(ConvertTo-HtmlSafe $r.Flag)</span></td>`n" +
            "            <td class=`"notes-cell`">$(ConvertTo-HtmlSafe $r.Notes)</td>`n" +
            "          </tr>`n"
    }

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Splunk Post-Deploy Performance Check</title>
<style>
$css
</style>
</head>
<body>
<div class="wrap">

  <header class="hero">
    <h1 class="report-h1">Splunk Post-Deployment Performance Check</h1>
    <p class="report-sub">Read-only sanity check - VM/process stats over the last $MinutesBack minutes</p>
  </header>

  <div class="kpi">
$kpi  </div>

  <div class="panel">
    <h2>Per-VM Results</h2>
    <div style="overflow-x:auto;">
    <table class="ex-table">
      <thead>
        <tr>
          <th>VM Name</th><th>Power State</th><th>CPU Avg</th><th>Mem Avg</th>
          <th>Process Found</th><th>Process CPU (s)</th><th>Process Mem</th>
          <th>Service Status</th><th>Flag</th><th>Notes</th>
        </tr>
      </thead>
      <tbody>
$rows      </tbody>
    </table>
    </div>
  </div>

  <footer>
    Generated by Check-SplunkPostDeployPerformance.ps1 (build $ScriptBuild). This is a
    sanity check, not a full performance audit - real impact tends to show up over hours,
    not minutes.
  </footer>

</div>
</body>
</html>
"@
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
        # -Realtime data (20-second samples) is only retained by vCenter for about the last
        # hour; anything further back has to come from the "Past Day" rollup (5-minute
        # samples, retained for 24h) instead, or Get-Stat returns nothing for that range.
        if ($MinutesBack -le 60) {
            $cpuStat = Get-Stat -Entity $vm -Stat 'cpu.usage.average' -Start $statStart -Realtime -ErrorAction SilentlyContinue
            $memStat = Get-Stat -Entity $vm -Stat 'mem.usage.average' -Start $statStart -Realtime -ErrorAction SilentlyContinue
        } else {
            $cpuStat = Get-Stat -Entity $vm -Stat 'cpu.usage.average' -Start $statStart -IntervalMins 5 -ErrorAction SilentlyContinue
            $memStat = Get-Stat -Entity $vm -Stat 'mem.usage.average' -Start $statStart -IntervalMins 5 -ErrorAction SilentlyContinue
        }

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

$reportTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$reportPath = Join-Path $ReportFolder ("SplunkPostDeployCheck_{0}.csv" -f $reportTimestamp)
$reportHtmlPath = Join-Path $ReportFolder ("SplunkPostDeployCheck_{0}.html" -f $reportTimestamp)
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

# HTML dashboard - best-effort, never fails the run if it can't be written.
try {
    $dashboardHtml = Build-PostDeployDashboardHtml -Results $results -OkCount $okCount -ReviewCount $reviewCount `
        -ErrorCount $errorCount -SkippedCount $skippedCount -MinutesBack $MinutesBack -ScriptBuild $ScriptBuild
    [System.IO.File]::WriteAllText($reportHtmlPath, $dashboardHtml, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Dashboard written to: $reportHtmlPath" -ForegroundColor Cyan
} catch {
    Write-Host "Could not write HTML dashboard: $($_.Exception.Message)" -ForegroundColor Yellow
}
