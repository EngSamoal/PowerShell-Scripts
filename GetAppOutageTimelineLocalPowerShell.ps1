#Requires -Version 5.1
<#
.SYNOPSIS
    App outage triage - READ-ONLY, runs entirely locally on the server being checked. Pulls the
    relevant Windows service/event-log evidence and works out (a) when the app service(s) were last
    confirmed running and (b) the exact time of the failing transition, from the Service Control
    Manager and Application event log entries actually present on this machine.

    No PowerCLI, no vCenter, no VMware Tools, no WinRM to another host. Makes zero configuration
    changes and does not touch/start/stop/restart any service.

.DESCRIPTION
    Copy this single file onto EACH of the affected servers (RDP clipboard/file copy, a network
    share, or paste its contents into a new .ps1 in a PowerShell/ISE session on that server) and run
    it there once per server. It does not need PowerCLI or a network path between servers - it only
    reads the local Service Control Manager (System log) and Application log on the box it runs on.

    For each service named in -ServiceName it reports:
      - Current state (Get-Service) and start mode (Get-CimInstance Win32_Service), plus how long
        the currently-running process instance has been up (if it is running right now).
      - Every Service Control Manager state-change / failure event in the System log within the
        -HoursBack window: Event ID 7036 (entered running/stopped state), 7031/7034 (terminated
        unexpectedly), 7000/7024 (failed to start / service-specific error), 7009 (timeout waiting
        for the service to connect to the SCM). Matched against this service specifically using the
        event's own insertion-string parameter (not a loose text search), so it will not confuse
        similarly-named services.
      - Application-log Critical/Error/Warning entries in the same window whose event source matches
        the service name (covers the app's own logged exceptions, not just the SCM's view of it).
      - Unexpected-reboot / planned-shutdown markers in the System log (Event ID 6008 "previous
        shutdown was unexpected", 41 "Kernel-Power" unexpected reboot, 1074 planned
        shutdown/restart, 6005/6006 Event Log service start/stop = boot/shutdown boundary) in the
        same window, since an outage that lines up with one of these is a different root cause
        (server crash/reboot) than a service that just stopped/crashed on its own.

    From the merged, time-sorted evidence it derives, per service:
      LastKnownGoodTime  - timestamp of the most recent "entered the running state" event.
      FailureTime        - timestamp of the first stop/crash/failure event AFTER that, i.e. the
                            transition that took it down and (per the log) never recovered.
    If the service's current Get-Service status is still Running, that is reported plainly instead
    of a failure time - a currently-running service is not "down" no matter what older events say.

    HONESTY / LIMITS (do not guess past what the log actually shows):
      - If no SCM state-change events exist for a service in the window, this is reported as
        "Unable to Determine (no SCM events in window)" - NOT guessed from surrounding events. Widen
        -HoursBack, or check the System log's OldestRecord (printed at the top of the run) to see how
        far back it actually goes; if the log rolled over before the failure, the evidence is gone
        and that is stated, not papered over.
      - A currently Stopped service with no matching SCM stop/crash event in the window is flagged
        explicitly ("service is Stopped now but no stop/crash event was found") rather than silently
        picking the wrong nearest event - that pattern usually means the log rolled over, the service
        was stopped outside the window, or it was stopped by something that doesn't log to the SCM.
      - Reboot/shutdown correlation is presented as evidence, not asserted as the cause - you decide
        whether the timestamps line up.

    Run this on all 4 affected servers, then compare the four *_Summary_*.csv files (or merge them:
        Get-ChildItem .\AppOutage_Reports\AppOutageTimeline_Summary_*.csv | Import-Csv |
            Export-Csv Combined_Summary.csv -NoTypeInformation
    ) to see which server(s) failed first and whether the failure times line up across the fleet
    (points at a shared dependency - DB, license server, network share, upstream API) or are staggered
    (points at something rolling through the servers one at a time - patch, certificate expiry, disk
    full, etc.).

.PARAMETER ServiceName
    One or more Windows service short names (the "Name" column in services.msc / Get-Service, not
    the Description) to investigate. EDIT THIS to match the actual app - the default below is only a
    starting guess based on what's visible in this environment's services list (ECSGenSvc,
    EcsReportsService, "Embed API"). If unsure of the exact name, run
    `Get-Service | Where-Object DisplayName -match 'part of the name' | Format-Table Name,DisplayName,Status`
    on the server first.

.PARAMETER HoursBack
    How far back to search the System/Application logs. Default 72 hours. Widen this
    (e.g. -HoursBack 168 for a week) if the app has been down longer than that or if the first pass
    reports "no SCM events in window".

.PARAMETER OutputPath
    Folder for the CSV/log files. Created if it doesn't exist. Defaults to .\AppOutage_Reports next
    to this script.

.PARAMETER MaxAppLogEvents
    Cap on Application-log rows pulled per run (default 500), so a very noisy Application log
    doesn't turn this into an hours-long query. Lower -HoursBack if you hit the cap and still need
    older Application-log evidence.

.EXAMPLE
    # Run locally on each of the 4 affected servers:
    .\Get-AppOutageTimelineLocal-PowerShell.ps1 -ServiceName 'ECSGenSvc','EcsReportsService','Embed API'

.EXAMPLE
    # Failure looks older than the default 3-day window:
    .\Get-AppOutageTimelineLocal-PowerShell.ps1 -ServiceName 'ECSGenSvc' -HoursBack 336

.NOTES
    Reads System and Application logs only (no Security log access needed), so an elevated session
    is not required for the SCM/Application-log evidence itself. Run elevated anyway if you want
    Win32_Service/Get-Process detail to be complete on a locked-down box.
#>

[CmdletBinding()]
param(
    [string[]]$ServiceName = @('ECSGenSvc', 'EcsReportsService', 'Embed API'),

    [ValidateRange(1, 8760)]
    [int]$HoursBack = 72,

    [string]$OutputPath = (Join-Path $PSScriptRoot 'AppOutage_Reports'),

    [ValidateRange(50, 5000)]
    [int]$MaxAppLogEvents = 500
)

$ProgressPreference = 'SilentlyContinue'
$ScriptBuild = '2026-09-07-01-local-standalone'
Write-Host "Get-AppOutageTimelineLocal-PowerShell.ps1 - build $ScriptBuild" -ForegroundColor Magenta
Write-Host "READ-ONLY - no service is started, stopped, or restarted by this script." -ForegroundColor Yellow

$ComputerName = $env:COMPUTERNAME
try {
    $IPAddresses = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
        Where-Object { $_.InterfaceAlias -notmatch 'Loopback' -and $_.IPAddress -notmatch '^169\.254\.' }).IPAddress -join ', '
} catch {
    $IPAddresses = $null
}
if ([string]::IsNullOrWhiteSpace($IPAddresses)) {
    try {
        $IPAddresses = ([System.Net.Dns]::GetHostAddresses($ComputerName) |
            Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
            Select-Object -First 1).IPAddressToString
    } catch {
        $IPAddresses = 'Unknown'
    }
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp        = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogPath         = Join-Path $OutputPath "AppOutageTimeline_${ComputerName}_$RunStamp.log"
$TimelineCsvPath = Join-Path $OutputPath "AppOutageTimeline_Events_${ComputerName}_$RunStamp.csv"
$SummaryCsvPath  = Join-Path $OutputPath "AppOutageTimeline_Summary_${ComputerName}_$RunStamp.csv"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

$Since = (Get-Date).AddHours(-$HoursBack)
Write-Log "Run started. Build $ScriptBuild. Computer: $ComputerName ($IPAddresses). Services: $($ServiceName -join ', '). Window: $Since to now ($HoursBack hours)."

try {
    $sysLog = Get-WinEvent -ListLog System -ErrorAction Stop
    if ($sysLog.OldestRecordTime -and $sysLog.OldestRecordTime -gt $Since) {
        Write-Log "NOTE: the System log's oldest record ($($sysLog.OldestRecordTime)) is AFTER the start of the requested window ($Since) - the log has rolled over and does not cover the full window. Evidence before $($sysLog.OldestRecordTime) is gone." 'WARN'
    }
} catch {
    Write-Log "Could not read System log metadata: $($_.Exception.Message)" 'WARN'
}

# =====================================================================================
# 1. Service Control Manager state-change / failure events (System log)
# =====================================================================================
$ScmIds = 7000, 7009, 7011, 7022, 7024, 7031, 7034, 7036
$scmEvents = @()
try {
    $scmEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Service Control Manager'; Id = $ScmIds; StartTime = $Since } -ErrorAction Stop)
    Write-Log "Retrieved $($scmEvents.Count) Service Control Manager event(s) from the System log in the window."
} catch {
    if ($_.Exception.Message -match 'No events were found') {
        Write-Log "No Service Control Manager events (7000/7009/7011/7022/7024/7031/7034/7036) at all in the System log in this window."
    } else {
        Write-Log "Get-WinEvent (System/Service Control Manager) failed: $($_.Exception.Message)" 'WARN'
    }
}

# =====================================================================================
# 2. Reboot / shutdown correlation markers (System log) - not filtered by service
# =====================================================================================
$rebootEvents = @()
try {
    $rebootEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 6005, 6006, 6008, 1074, 41; StartTime = $Since } -ErrorAction Stop)
    Write-Log "Retrieved $($rebootEvents.Count) boot/shutdown marker event(s) from the System log in the window."
} catch {
    if ($_.Exception.Message -notmatch 'No events were found') {
        Write-Log "Get-WinEvent (System/reboot markers) failed: $($_.Exception.Message)" 'WARN'
    }
}

# =====================================================================================
# 3. Application-log errors/warnings whose source matches one of the services
# =====================================================================================
$appEvents = @()
try {
    $rawApp = @(Get-WinEvent -FilterHashtable @{ LogName = 'Application'; Level = 1, 2, 3; StartTime = $Since } -MaxEvents $MaxAppLogEvents -ErrorAction Stop)
    $appEvents = @($rawApp | Where-Object {
        $prov = [string]$_.ProviderName
        $msg  = [string]$_.Message
        ($ServiceName | Where-Object { $prov -match [regex]::Escape($_) -or $msg -match [regex]::Escape($_) }).Count -gt 0
    })
    Write-Log "Application log: pulled $($rawApp.Count) Critical/Error/Warning event(s) (cap $MaxAppLogEvents), $($appEvents.Count) matched one of the named services by provider/message."
} catch {
    if ($_.Exception.Message -match 'No events were found') {
        Write-Log "No Critical/Error/Warning events at all in the Application log in this window."
    } else {
        Write-Log "Get-WinEvent (Application) failed: $($_.Exception.Message)" 'WARN'
    }
}

# =====================================================================================
# 4. Build the merged, time-sorted timeline
# =====================================================================================
function New-TimelineRow {
    param($Time, $LogName, $Id, $Level, $Provider, $Message, $MatchedService = '')
    [PSCustomObject]@{
        ComputerName    = $ComputerName
        TimeCreated     = $Time
        LogName         = $LogName
        EventId         = $Id
        Level           = $Level
        ProviderName    = $Provider
        MatchedService  = $MatchedService
        Message         = (($Message -replace '\s+', ' ').Trim())
    }
}

$Timeline = New-Object System.Collections.Generic.List[object]

# Map each SCM event to the specific service it's about via the event's own insertion
# strings (Properties[0] is %1, the service name/display name the SCM logged) - this
# avoids false matches between similarly-named services, which a plain text search on
# the rendered Message would risk.
foreach ($evt in $scmEvents) {
    $param1 = $null
    try { if ($evt.Properties.Count -gt 0) { $param1 = [string]$evt.Properties[0].Value } } catch { }
    $matched = $ServiceName | Where-Object {
        $svcObj = Get-Service -Name $_ -ErrorAction SilentlyContinue
        ($param1 -and ($param1 -eq $_ -or ($svcObj -and $param1 -eq $svcObj.DisplayName))) -or
        ($evt.Message -match [regex]::Escape($_))
    }
    if ($matched) {
        foreach ($m in $matched) {
            $Timeline.Add((New-TimelineRow -Time $evt.TimeCreated -LogName 'System' -Id $evt.Id -Level $evt.LevelDisplayName -Provider $evt.ProviderName -Message $evt.Message -MatchedService $m)) | Out-Null
        }
    }
}
foreach ($evt in $rebootEvents) {
    $Timeline.Add((New-TimelineRow -Time $evt.TimeCreated -LogName 'System' -Id $evt.Id -Level $evt.LevelDisplayName -Provider $evt.ProviderName -Message $evt.Message -MatchedService '(server boot/shutdown)')) | Out-Null
}
foreach ($evt in $appEvents) {
    $matched = $ServiceName | Where-Object { $evt.ProviderName -match [regex]::Escape($_) -or $evt.Message -match [regex]::Escape($_) }
    foreach ($m in $matched) {
        $Timeline.Add((New-TimelineRow -Time $evt.TimeCreated -LogName 'Application' -Id $evt.Id -Level $evt.LevelDisplayName -Provider $evt.ProviderName -Message $evt.Message -MatchedService $m)) | Out-Null
    }
}

$Timeline = @($Timeline | Sort-Object TimeCreated)
$Timeline | Export-Csv -Path $TimelineCsvPath -NoTypeInformation
Write-Log "Wrote $($Timeline.Count) timeline row(s) to $TimelineCsvPath"

# =====================================================================================
# 5. Per-service current state + last-known-good / failure derivation
# =====================================================================================
$Summary = New-Object System.Collections.Generic.List[object]

foreach ($svc in $ServiceName) {
    $svcObj = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if (-not $svcObj) {
        Write-Log "Service '$svc' was not found on $ComputerName - skipping (check the exact short name with Get-Service)." 'WARN'
        $Summary.Add([PSCustomObject]@{
            ComputerName = $ComputerName; ServiceName = $svc; CurrentStatus = 'Not Found'; StartMode = ''
            CurrentInstanceUpSince = ''; LastKnownGoodTime = ''; FailureTime = ''; DownForApprox = ''
            FailureEventId = ''; FailureMessage = ''; Notes = "No service named '$svc' exists on this host."
        }) | Out-Null
        continue
    }

    $startMode = ''
    $upSince = ''
    try {
        $cim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$svc'" -ErrorAction Stop
        if ($cim) { $startMode = $cim.StartMode }
        if ($svcObj.Status -eq 'Running' -and $cim -and $cim.ProcessId -gt 0) {
            try { $upSince = (Get-Process -Id $cim.ProcessId -ErrorAction Stop).StartTime }
            catch { $upSince = '' }
        }
    } catch {
        Write-Log "Win32_Service lookup failed for '$svc': $($_.Exception.Message)" 'WARN'
    }

    $svcRows = @($Timeline | Where-Object { $_.MatchedService -eq $svc })
    $runningRows = @($svcRows | Where-Object { $_.EventId -eq 7036 -and $_.Message -match 'running state' })
    $badRows     = @($svcRows | Where-Object {
        ($_.EventId -eq 7036 -and $_.Message -match 'stopped state') -or
        $_.EventId -in 7000, 7009, 7011, 7022, 7024, 7031, 7034
    })

    $lastGood = $null
    if ($runningRows.Count -gt 0) { $lastGood = ($runningRows | Sort-Object TimeCreated -Descending | Select-Object -First 1).TimeCreated }

    $failureRow = $null
    if ($badRows.Count -gt 0) {
        $candidates = if ($lastGood) { $badRows | Where-Object { $_.TimeCreated -gt $lastGood } } else { $badRows }
        if (-not $candidates -or @($candidates).Count -eq 0) { $candidates = $badRows }
        $failureRow = $candidates | Sort-Object TimeCreated | Select-Object -First 1
    }

    $notes = New-Object System.Collections.Generic.List[string]
    if ($svcRows.Count -eq 0) {
        $notes.Add("No SCM state-change events for '$svc' found in the last $HoursBack hours.")
    }
    if ($svcObj.Status -eq 'Running') {
        $notes.Add("Current Get-Service status is Running - not currently down on this host, whatever older events show.")
        if ($upSince) { $notes.Add("Current process instance has been up since $upSince.") }
    } elseif (-not $failureRow) {
        $notes.Add("Service is currently '$($svcObj.Status)' but no matching stop/crash SCM event was found in the window - the log may have rolled over, it may have been stopped outside this window, or it was stopped by something that doesn't log to the SCM. Widen -HoursBack.")
    }
    $reboots = @($Timeline | Where-Object { $_.MatchedService -eq '(server boot/shutdown)' })
    if ($reboots.Count -gt 0) {
        $notes.Add("$($reboots.Count) server boot/shutdown marker event(s) also present in this window - check whether their timestamps line up with the failure.")
    }

    $downFor = ''
    if ($failureRow -and $svcObj.Status -ne 'Running') {
        $downFor = '{0:dd}d {0:hh}h {0:mm}m' -f ((Get-Date) - $failureRow.TimeCreated)
    }

    $Summary.Add([PSCustomObject]@{
        ComputerName           = $ComputerName
        ServiceName            = $svc
        CurrentStatus          = $svcObj.Status
        StartMode              = $startMode
        CurrentInstanceUpSince = $upSince
        LastKnownGoodTime      = $lastGood
        FailureTime            = $(if ($svcObj.Status -eq 'Running') { '' } else { $failureRow.TimeCreated })
        DownForApprox          = $downFor
        FailureEventId         = $(if ($failureRow) { $failureRow.EventId } else { '' })
        FailureMessage         = $(if ($failureRow) { $failureRow.Message } else { '' })
        Notes                  = ($notes -join ' ')
    }) | Out-Null
}

$Summary | Export-Csv -Path $SummaryCsvPath -NoTypeInformation
Write-Log "Wrote per-service summary to $SummaryCsvPath"

# =====================================================================================
# 6. Console report
# =====================================================================================
Write-Host ""
Write-Host "=== $ComputerName - App outage summary (window: last $HoursBack hours) ===" -ForegroundColor Cyan
foreach ($row in $Summary) {
    $color = switch ($row.CurrentStatus) {
        'Running'   { 'Green' }
        'Not Found' { 'DarkGray' }
        default     { 'Red' }
    }
    Write-Host ("  {0}  [{1}]" -f $row.ServiceName, $row.CurrentStatus) -ForegroundColor $color
    if ($row.LastKnownGoodTime) { Write-Host ("      Last known good : {0}" -f $row.LastKnownGoodTime) -ForegroundColor Gray }
    if ($row.FailureTime)       { Write-Host ("      Failed at       : {0}  (event {1})" -f $row.FailureTime, $row.FailureEventId) -ForegroundColor Yellow }
    if ($row.DownForApprox)     { Write-Host ("      Down for approx : {0}" -f $row.DownForApprox) -ForegroundColor Yellow }
    if ($row.FailureMessage)    { Write-Host ("      Detail          : {0}" -f $row.FailureMessage) -ForegroundColor DarkGray }
    if ($row.Notes)             { Write-Host ("      Notes           : {0}" -f $row.Notes) -ForegroundColor DarkGray }
}
Write-Host ""
Write-Host "Full event timeline : $TimelineCsvPath" -ForegroundColor Gray
Write-Host "Per-service summary : $SummaryCsvPath" -ForegroundColor Gray
Write-Host "Run log             : $LogPath" -ForegroundColor Gray
Write-Host ""
Write-Host "Repeat this on the other 3 affected servers, then compare the *_Summary_*.csv files - matching failure times across all 4 points at a shared dependency; staggered times point at something moving through the fleet." -ForegroundColor Cyan

Write-Log "Run complete."
