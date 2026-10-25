#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only Windows OS health check for one or many servers, results written to Excel (.xlsx).

.DESCRIPTION
    Collects, per server:
        OS caption / version / build (incl. UBR), last boot + uptime, CPU load,
        RAM total/free, every fixed disk's free space, last installed KB + patch age,
        pending reboot (CBS, Windows Update, pending file renames, pending computer
        rename, SCCM client if present), Windows activation status, Automatic
        services that are not running, and Critical/Error events from the System log.

    Each server gets an overall Status (OK / Warning / Critical) and a plain-text
    Issues column explaining exactly why. A check that cannot be run (access denied,
    WMI broken, etc.) is reported as "Unable to check ..." with the reason - it is
    never silently left blank or treated as healthy.

    STRICTLY READ-ONLY: only Get-* cmdlets, CIM queries and registry reads are used.
    Nothing is changed on the target servers.

    Output (in -OutputPath):
        OSHealthCheck_<SERVER>_<timestamp>.xlsx   Summary / Disks / StoppedServices / Events sheets
        OSHealthCheck_<SERVER>_<timestamp>.log    Run log (kept out of the report)
        OSHealthCheck_<SERVER>_<timestamp>.csv    Only written if the .xlsx could not be

    Typical use: run it locally on each VM, copy every VM's .xlsx into one folder, then run
    it once with -ConsolidatePath <folder> to merge them into one report
    (OSHealthCheck_Consolidated_<timestamp>.xlsx). If the same server appears in more than
    one report, only its newest report is used.

    The .xlsx is written natively (Open XML) - no Excel installation and no
    ImportExcel module are needed on the machine running the script.

.PARAMETER ComputerName
    One or more servers. Defaults to the local machine. The local machine is checked
    directly; remote servers are checked over WinRM (Invoke-Command), in parallel.

.PARAMETER ComputerListPath
    Text file with one server name per line ('#' comment lines and blanks ignored).
    Combined with -ComputerName if both are given.

.PARAMETER Credential
    Credential for remote servers (not used for the local machine).

.PARAMETER OutputPath
    Folder for the .xlsx/.log (created if missing). Default C:\Temp\OSHealthCheck.
    With -ConsolidatePath, defaults to the consolidate folder itself.

.PARAMETER ConsolidatePath
    Folder (searched recursively) holding the per-VM .xlsx reports. Runs no health checks;
    only merges the reports into one workbook.

.EXAMPLE
    .\Invoke-OSHealthCheck.ps1
    Checks the local server.

.EXAMPLE
    .\Invoke-OSHealthCheck.ps1 -ConsolidatePath C:\Temp\AllReports
    Merges every per-VM report in C:\Temp\AllReports into one consolidated report.

.EXAMPLE
    .\Invoke-OSHealthCheck.ps1 -ComputerListPath C:\Temp\servers.txt -Credential (Get-Credential)

.NOTES
    Run elevated (as Administrator) for complete results. Remote servers need WinRM
    enabled and admin rights for the supplied account.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromPipeline = $true)]
    [string[]]$ComputerName,

    [string]$ComputerListPath,

    [System.Management.Automation.PSCredential]$Credential,

    [string]$OutputPath = 'C:\Temp\OSHealthCheck',

    # Merge mode: folder of per-VM reports to consolidate (no checks are run)
    [string]$ConsolidatePath,

    # Look-back window for System log Critical/Error events
    [ValidateRange(1, 90)][int]$EventDays = 7,

    # Disk free % below which a disk is Warning / Critical
    [ValidateRange(1, 99)][int]$DiskWarnPct = 15,
    [ValidateRange(1, 99)][int]$DiskCritPct = 5,

    # Free RAM % below which the server is flagged
    [ValidateRange(1, 99)][int]$RamWarnPct = 10,

    # Days since last installed KB before flagging
    [ValidateRange(1, 3650)][int]$PatchWarnDays = 45,

    # Days of uptime before flagging (long uptime = patches likely not applied)
    [ValidateRange(1, 3650)][int]$UptimeWarnDays = 60,

    # Number of System log Error events in the window before flagging
    [ValidateRange(1, 1000000)][int]$ErrorEventWarn = 50,

    # Max events read per server (protects against huge logs)
    [ValidateRange(100, 100000)][int]$MaxEvents = 5000,

    # Automatic services that are normally stopped (trigger/delayed start). Wildcards allowed.
    [string[]]$ServiceExclusions = @(
        'gupdate', 'gupdatem', 'GoogleUpdater*', 'edgeupdate', 'edgeupdatem',
        'MapsBroker', 'RemoteRegistry', 'sppsvc', 'CDPSvc', 'tiledatamodelsvc',
        'ShellHWDetection', 'WbioSrvc', 'clr_optimization_*', 'OneSyncSvc_*',
        'CDPUserSvc_*', 'BITS', 'wuauserv', 'UsoSvc', 'TrustedInstaller'
    ),

    [ValidateRange(1, 256)][int]$ThrottleLimit = 32,

    # Also return the summary objects to the pipeline
    [switch]$PassThru
)

begin {
    $pipedNames = New-Object System.Collections.Generic.List[string]
}

process {
    foreach ($n in $ComputerName) { if ($n) { $pipedNames.Add($n) } }
}

end {

$ScriptBuild = '2026-10-25.1'
Write-Host "Invoke-OSHealthCheck.ps1 - build $ScriptBuild" -ForegroundColor Magenta

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------------------------
# Collector - runs ON the target server (locally or via Invoke-Command). Must stay PS 5.1
# compatible and must return only plain values (strings/numbers/dates) so it deserializes
# cleanly over remoting.
# ---------------------------------------------------------------------------------------------
$collector = {
    param([int]$EventDays, [string[]]$ServiceExclusions, [int]$MaxEvents)

    $ErrorActionPreference = 'Stop'
    $checkErrors = New-Object System.Collections.Generic.List[string]
    $r = [ordered]@{
        ComputerName = $env:COMPUTERNAME
        CollectedAt  = Get-Date
        IsAdmin      = $null
        CheckErrors  = $null
    }

    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $r.IsAdmin = ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $r.IsAdmin = $null }

    # --- OS / memory / uptime --------------------------------------------------------------
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -OperationTimeoutSec 60
        $r.OS           = [string]$os.Caption
        $r.Version      = [string]$os.Version
        $r.BuildNumber  = [string]$os.BuildNumber
        $r.LastBoot     = [datetime]$os.LastBootUpTime
        $r.TotalRAM_KB  = [double]$os.TotalVisibleMemorySize
        $r.FreeRAM_KB   = [double]$os.FreePhysicalMemory
        try {
            $ubr = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name UBR).UBR
            if ($null -ne $ubr) { $r.UBR = [int]$ubr }
        } catch { }   # UBR is cosmetic; absent on older OS
    } catch { $checkErrors.Add("OS info: $($_.Exception.Message)") }

    # --- CPU -------------------------------------------------------------------------------
    try {
        $loads = @(Get-CimInstance -ClassName Win32_Processor -OperationTimeoutSec 60 |
                   Where-Object { $null -ne $_.LoadPercentage } | ForEach-Object { [double]$_.LoadPercentage })
        if ($loads.Count -gt 0) { $r.CPULoadPct = [math]::Round(($loads | Measure-Object -Average).Average, 0) }
    } catch { $checkErrors.Add("CPU load: $($_.Exception.Message)") }

    # --- Disks (all local fixed disks, not just C:) ----------------------------------------
    try {
        $r.Disks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -OperationTimeoutSec 60 | ForEach-Object {
            [pscustomobject]@{
                Drive  = [string]$_.DeviceID
                Label  = [string]$_.VolumeName
                SizeB  = [double]$_.Size
                FreeB  = [double]$_.FreeSpace
            }
        })
    } catch { $checkErrors.Add("Disks: $($_.Exception.Message)") }

    # --- Last hotfix -----------------------------------------------------------------------
    # InstalledOn can be empty / unparsable on some systems - only dated entries are used.
    try {
        $hf = @(Get-HotFix | ForEach-Object {
            $d = $_.InstalledOn -as [datetime]
            if ($d) { [pscustomobject]@{ HotFixID = [string]$_.HotFixID; InstalledOn = $d } }
        } | Sort-Object InstalledOn -Descending)
        $r.HotFixCount = $hf.Count
        if ($hf.Count -gt 0) {
            $r.LastKB     = $hf[0].HotFixID
            $r.LastKBDate = $hf[0].InstalledOn
        }
    } catch { $checkErrors.Add("Hotfixes: $($_.Exception.Message)") }

    # --- Pending reboot --------------------------------------------------------------------
    try {
        $reasons = New-Object System.Collections.Generic.List[string]
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons.Add('CBS') }
        if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons.Add('WindowsUpdate') }
        $sm = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ErrorAction SilentlyContinue
        if ($sm -and $sm.PSObject.Properties['PendingFileRenameOperations'] -and @($sm.PendingFileRenameOperations | Where-Object { $_ }).Count -gt 0) {
            $reasons.Add('PendingFileRename')
        }
        $active  = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -Name ComputerName -ErrorAction SilentlyContinue).ComputerName
        $pending = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -Name ComputerName -ErrorAction SilentlyContinue).ComputerName
        if ($active -and $pending -and $active -ne $pending) { $reasons.Add('ComputerRename') }
        try {
            $ccm = Invoke-CimMethod -Namespace 'root\ccm\ClientSDK' -ClassName CCM_ClientUtilities -MethodName DetermineIfRebootPending -OperationTimeoutSec 30
            if ($ccm -and ($ccm.RebootPending -or $ccm.IsHardRebootPending)) { $reasons.Add('SCCM') }
        } catch { }   # SCCM client not installed on most servers - not an error
        $r.PendingReboot       = ($reasons.Count -gt 0)
        $r.PendingRebootReason = ($reasons -join ', ')
    } catch { $checkErrors.Add("Pending reboot: $($_.Exception.Message)") }

    # --- Activation ------------------------------------------------------------------------
    # Filter on the Windows ApplicationID (much faster than LIKE 'Windows%') and handle
    # multiple products with keys: prefer a licensed one.
    try {
        $actMap = @{ 0 = 'Unlicensed'; 1 = 'Licensed'; 2 = 'OOB Grace'; 3 = 'OOT Grace'; 4 = 'Non-Genuine Grace'; 5 = 'Notification'; 6 = 'Extended Grace' }
        $lic = @(Get-CimInstance -ClassName SoftwareLicensingProduct -OperationTimeoutSec 120 `
                    -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL")
        if ($lic.Count -eq 0) {
            $r.Activation = 'No Windows product key found'
        } else {
            $best = $lic | Sort-Object @{ Expression = { if ([int]$_.LicenseStatus -eq 1) { 0 } else { 1 } } } | Select-Object -First 1
            $code = [int]$best.LicenseStatus
            $r.Activation = if ($actMap.ContainsKey($code)) { $actMap[$code] } else { "Unknown ($code)" }
            $r.ActivationChannel = [string]$best.ProductKeyChannel
        }
    } catch { $checkErrors.Add("Activation: $($_.Exception.Message)") }

    # --- Stopped automatic services --------------------------------------------------------
    try {
        $svc = @(Get-CimInstance -ClassName Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" -OperationTimeoutSec 60)
        $r.StoppedServices = @($svc | Where-Object {
            $name = $_.Name
            -not ($ServiceExclusions | Where-Object { $name -like $_ })
        } | Sort-Object Name | ForEach-Object {
            [pscustomobject]@{
                Name             = [string]$_.Name
                DisplayName      = [string]$_.DisplayName
                State            = [string]$_.State
                ExitCode         = $_.ExitCode
                DelayedAutoStart = if ($_.PSObject.Properties['DelayedAutoStart']) { $_.DelayedAutoStart } else { $null }
            }
        })
        $r.ExcludedStoppedCount = $svc.Count - $r.StoppedServices.Count
    } catch { $checkErrors.Add("Services: $($_.Exception.Message)") }

    # --- System log critical / error events ------------------------------------------------
    try {
        $events = @()
        try {
            $events = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2; StartTime = (Get-Date).AddDays(-$EventDays) } -MaxEvents $MaxEvents)
        } catch {
            # "No events found" is a healthy result, not an error.
            if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { throw }
        }
        $r.EventsCapped   = ($events.Count -ge $MaxEvents)
        $r.CriticalEvents = @($events | Where-Object { $_.Level -eq 1 }).Count
        $r.ErrorEvents    = @($events | Where-Object { $_.Level -eq 2 }).Count
        # Reduce to a top-N summary on the server (keeps remoting payload small).
        $r.TopEvents = @($events | Group-Object Level, ProviderName, Id |
            Sort-Object @{ Expression = { $_.Group[0].Level } }, @{ Expression = 'Count'; Descending = $true } |
            Select-Object -First 15 | ForEach-Object {
                $e = $_.Group[0]   # Get-WinEvent returns newest first
                $msg = ''
                try { $msg = [string]$e.Message } catch { }
                if (-not $msg) { $msg = '(message unavailable)' }
                $msg = ($msg -replace '\s+', ' ').Trim()
                if ($msg.Length -gt 400) { $msg = $msg.Substring(0, 400) + '...' }
                [pscustomobject]@{
                    Level    = if ($e.Level -eq 1) { 'Critical' } else { 'Error' }
                    Source   = [string]$e.ProviderName
                    EventId  = [int]$e.Id
                    Count    = $_.Count
                    LastSeen = [datetime]$e.TimeCreated
                    Message  = $msg
                }
            })
    } catch { $checkErrors.Add("System event log: $($_.Exception.Message)") }

    $r.CheckErrors = @($checkErrors)
    [pscustomobject]$r
}

# ---------------------------------------------------------------------------------------------
# Helpers (run on the admin machine)
# ---------------------------------------------------------------------------------------------
$script:LogFile = $null
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
    if ($script:LogFile) { try { Add-Content -Path $script:LogFile -Value $line -Encoding UTF8 } catch { } }
}

function Test-IsLocalComputer {
    param([string]$Name)
    $n = $Name.Trim().ToLowerInvariant()
    $me = ([string]$env:COMPUTERNAME).ToLowerInvariant()
    return ($n -in @('.', 'localhost', '127.0.0.1', '::1') -or ($me -and ($n -eq $me -or $n.StartsWith("$me."))))
}

function Get-WorstStatus {
    param([string[]]$Statuses)
    if ($Statuses -contains 'Critical') { return 'Critical' }
    if ($Statuses -contains 'Warning')  { return 'Warning' }
    return 'OK'
}

function ConvertTo-HealthRows {
    # Turns one collector result (or an unreachable marker) into Summary/Disk/Service/Event rows.
    param([string]$Target, $Raw, [string]$ConnectError)

    $issues   = New-Object System.Collections.Generic.List[string]
    $statuses = New-Object System.Collections.Generic.List[string]
    $diskRows = @(); $svcRows = @(); $evtRows = @()
    $critCol  = "CritEvents_$($EventDays)d"
    $errCol   = "ErrEvents_$($EventDays)d"

    if (-not $Raw) {
        $row = [ordered]@{
            Server = $Target; Status = 'Critical'; Issues = "Unreachable / collection failed: $ConnectError"
        }
        return [pscustomobject]@{ Summary = [pscustomobject]$row; Disks = @(); Services = @(); Events = @() }
    }

    $now = $Raw.CollectedAt

    # Unable-to-check items are surfaced, never hidden.
    foreach ($ce in @($Raw.CheckErrors)) { if ($ce) { $issues.Add("Unable to check $ce"); $statuses.Add('Warning') } }
    if ($Raw.IsAdmin -eq $false) { $issues.Add('Not run as Administrator - some results may be incomplete'); $statuses.Add('Warning') }

    # Uptime
    $uptimeDays = $null
    if ($Raw.LastBoot) {
        $uptimeDays = [math]::Round(($now - $Raw.LastBoot).TotalDays, 1)
        if ($uptimeDays -ge $UptimeWarnDays) { $issues.Add("Uptime $uptimeDays days (>= $UptimeWarnDays)"); $statuses.Add('Warning') }
    }

    # Memory
    $totalRamGB = $null; $freeRamGB = $null; $freeRamPct = $null
    if ($Raw.TotalRAM_KB -gt 0) {
        $totalRamGB = [math]::Round($Raw.TotalRAM_KB / 1MB, 1)
        $freeRamGB  = [math]::Round($Raw.FreeRAM_KB / 1MB, 1)
        $freeRamPct = [math]::Round(100 * $Raw.FreeRAM_KB / $Raw.TotalRAM_KB, 1)
        if ($freeRamPct -lt $RamWarnPct) { $issues.Add("Free RAM $freeRamPct% (< $RamWarnPct%)"); $statuses.Add('Warning') }
    }

    # Disks
    $cFreeGB = $null; $cFreePct = $null
    foreach ($d in @($Raw.Disks)) {
        if (-not $d) { continue }
        $sizeGB = [math]::Round($d.SizeB / 1GB, 1)
        $freeGB = [math]::Round($d.FreeB / 1GB, 1)
        $pct    = if ($d.SizeB -gt 0) { [math]::Round(100 * $d.FreeB / $d.SizeB, 1) } else { $null }
        $ds = 'OK'
        if ($null -ne $pct) {
            if ($pct -lt $DiskCritPct)     { $ds = 'Critical' }
            elseif ($pct -lt $DiskWarnPct) { $ds = 'Warning' }
        }
        if ($ds -ne 'OK') { $issues.Add("Disk $($d.Drive) free $freeGB GB ($pct%)"); $statuses.Add($ds) }
        if ($d.Drive -eq 'C:') { $cFreeGB = $freeGB; $cFreePct = $pct }
        $diskRows += [pscustomobject][ordered]@{
            Server = $Target; Drive = $d.Drive; Label = $d.Label
            Size_GB = $sizeGB; Free_GB = $freeGB; Free_Pct = $pct; Status = $ds
        }
    }

    # Patching
    $patchAge = $null
    if ($Raw.LastKBDate) {
        $patchAge = [int]($now.Date - $Raw.LastKBDate.Date).TotalDays
        if ($patchAge -gt $PatchWarnDays) { $issues.Add("Last KB installed $patchAge days ago (> $PatchWarnDays)"); $statuses.Add('Warning') }
    } elseif ($null -ne $Raw.HotFixCount) {
        $issues.Add('No dated hotfix found - verify patch status manually'); $statuses.Add('Warning')
    }

    # Pending reboot
    if ($Raw.PendingReboot) { $issues.Add("Pending reboot ($($Raw.PendingRebootReason))"); $statuses.Add('Warning') }

    # Activation
    if ($Raw.Activation -and $Raw.Activation -ne 'Licensed') { $issues.Add("Windows activation: $($Raw.Activation)"); $statuses.Add('Warning') }

    # Services
    $stopped = @($Raw.StoppedServices | Where-Object { $_ })
    if ($stopped.Count -gt 0) {
        $issues.Add("$($stopped.Count) automatic service(s) not running: $(($stopped.Name) -join ', ')"); $statuses.Add('Warning')
    }
    foreach ($s in $stopped) {
        $svcRows += [pscustomobject][ordered]@{
            Server = $Target; Service = $s.Name; DisplayName = $s.DisplayName; State = $s.State
            ExitCode = $s.ExitCode; DelayedAutoStart = $s.DelayedAutoStart
        }
    }

    # Events
    $capNote = if ($Raw.EventsCapped) { " (first $MaxEvents read)" } else { '' }
    if ($Raw.CriticalEvents -gt 0) { $issues.Add("$($Raw.CriticalEvents) Critical System event(s) in $EventDays d$capNote"); $statuses.Add('Warning') }
    if ($Raw.ErrorEvents -ge $ErrorEventWarn) { $issues.Add("$($Raw.ErrorEvents) Error System events in $EventDays d$capNote (>= $ErrorEventWarn)"); $statuses.Add('Warning') }
    foreach ($e in @($Raw.TopEvents | Where-Object { $_ })) {
        $evtRows += [pscustomobject][ordered]@{
            Server = $Target; Level = $e.Level; Source = $e.Source; EventId = $e.EventId
            Count = $e.Count; LastSeen = $e.LastSeen; Message = $e.Message
        }
    }

    $version = $Raw.Version
    if ($null -ne $Raw.UBR -and $version) { $version = "$version.$($Raw.UBR)" }   # e.g. 10.0.20348.2700
    $lowest = @($diskRows | Where-Object { $null -ne $_.Free_Pct } | Sort-Object Free_Pct | Select-Object -First 1)

    $row = [ordered]@{
        Server              = $Target
        Status              = Get-WorstStatus -Statuses $statuses
        Issues              = if ($issues.Count) { $issues -join '; ' } else { 'None' }
        OS                  = $Raw.OS
        Version             = $version
        LastBoot            = $Raw.LastBoot
        Uptime_Days         = $uptimeDays
        CPULoad_Pct         = $Raw.CPULoadPct
        TotalRAM_GB         = $totalRamGB
        FreeRAM_GB          = $freeRamGB
        FreeRAM_Pct         = $freeRamPct
        C_Free_GB           = $cFreeGB
        C_Free_Pct          = $cFreePct
        LowestDisk          = if ($lowest) { "$($lowest[0].Drive) $($lowest[0].Free_Pct)%" } else { $null }
        LastKB              = $Raw.LastKB
        LastPatch           = $Raw.LastKBDate
        PatchAge_Days       = $patchAge
        PendingReboot       = $Raw.PendingReboot
        PendingRebootReason = $Raw.PendingRebootReason
        Activation          = $Raw.Activation
        StoppedAutoSvc      = if ($null -ne $Raw.StoppedServices) { $stopped.Count } else { $null }
        StoppedAutoSvcNames = ($stopped.Name -join ', ')
        $critCol            = $Raw.CriticalEvents
        $errCol             = $Raw.ErrorEvents
        CollectedAt         = $Raw.CollectedAt
    }
    [pscustomobject]@{ Summary = [pscustomobject]$row; Disks = $diskRows; Services = $svcRows; Events = $evtRows }
}

# ---------------------------------------------------------------------------------------------
# Minimal native .xlsx writer (Open XML via System.IO.Compression) - no Excel / module needed.
# Sheets: array of @{ Name = 'x'; Rows = <objects> }. Columns come from the first row's properties.
# ---------------------------------------------------------------------------------------------
function ConvertTo-XlsxColumnName {
    param([int]$Index)
    $s = ''
    while ($Index -gt 0) { $m = ($Index - 1) % 26; $s = [char](65 + $m) + $s; $Index = [int][math]::Floor(($Index - 1) / 26) }
    $s
}

function ConvertTo-XmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    $s = [string]$Value -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
    if ($s.Length -gt 32000) { $s = $s.Substring(0, 32000) + '...' }
    [System.Security.SecurityElement]::Escape($s)
}

function Export-SimpleXlsx {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][object[]]$Sheets)

    Add-Type -AssemblyName System.IO.Compression
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $statusStyle = @{ 'OK' = 3; 'Warning' = 4; 'Critical' = 5 }
    $numTypes = @([int16], [int32], [int64], [uint16], [uint32], [uint64], [byte], [sbyte], [double], [single], [decimal])
    $hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $nsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $nsRel  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    $parts = [ordered]@{}
    $sheetEntries = New-Object System.Text.StringBuilder
    $definedNames = New-Object System.Text.StringBuilder
    $sheetIndex = 0

    foreach ($sheet in $Sheets) {
        $sheetIndex++
        $rows = @($sheet.Rows | Where-Object { $null -ne $_ })
        if ($rows.Count -eq 0) { $rows = @([pscustomobject]@{ Info = 'No items found' }) }
        $cols = @($rows[0].PSObject.Properties.Name)

        # Column widths from header + first 500 rows
        $widths = @{}
        for ($c = 0; $c -lt $cols.Count; $c++) { $widths[$c] = $cols[$c].Length + 3 }
        foreach ($row in ($rows | Select-Object -First 500)) {
            for ($c = 0; $c -lt $cols.Count; $c++) {
                $v = $row.($cols[$c])
                $len = if ($v -is [datetime]) { 16 } elseif ($null -ne $v) { ([string]$v).Length } else { 0 }
                if ($len + 2 -gt $widths[$c]) { $widths[$c] = $len + 2 }
            }
        }

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append($hdr).Append("<worksheet xmlns=`"$nsMain`" xmlns:r=`"$nsRel`">")
        [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><cols>')
        for ($c = 0; $c -lt $cols.Count; $c++) {
            $w = [math]::Min([math]::Max($widths[$c], 8), 80)
            [void]$sb.Append("<col min=`"$($c+1)`" max=`"$($c+1)`" width=`"$w`" customWidth=`"1`"/>")
        }
        [void]$sb.Append('</cols><sheetData><row r="1">')
        for ($c = 0; $c -lt $cols.Count; $c++) {
            [void]$sb.Append("<c r=`"$(ConvertTo-XlsxColumnName ($c+1))1`" s=`"1`" t=`"inlineStr`"><is><t>$(ConvertTo-XmlText $cols[$c])</t></is></c>")
        }
        [void]$sb.Append('</row>')

        $r = 1
        foreach ($row in $rows) {
            $r++
            [void]$sb.Append("<row r=`"$r`">")
            for ($c = 0; $c -lt $cols.Count; $c++) {
                $v = $row.($cols[$c])
                if ($null -eq $v -or ($v -is [string] -and $v -eq '')) { continue }
                $ref = "$(ConvertTo-XlsxColumnName ($c+1))$r"
                if ($v -is [datetime]) {
                    [void]$sb.Append("<c r=`"$ref`" s=`"2`"><v>$($v.ToOADate().ToString('R', $inv))</v></c>")
                } elseif ($v -is [bool]) {
                    [void]$sb.Append("<c r=`"$ref`" t=`"inlineStr`"><is><t>$(if ($v) { 'Yes' } else { 'No' })</t></is></c>")
                } elseif ($numTypes -contains $v.GetType() -and -not ($v -is [double] -and ([double]::IsNaN($v) -or [double]::IsInfinity($v)))) {
                    [void]$sb.Append("<c r=`"$ref`"><v>$(([double]$v).ToString('R', $inv))</v></c>")
                } else {
                    $text = [string]$v
                    $s = if ($cols[$c] -eq 'Status' -and $statusStyle.ContainsKey($text)) { " s=`"$($statusStyle[$text])`"" } else { '' }
                    [void]$sb.Append("<c r=`"$ref`"$s t=`"inlineStr`"><is><t xml:space=`"preserve`">$(ConvertTo-XmlText $text)</t></is></c>")
                }
            }
            [void]$sb.Append('</row>')
        }
        $lastRef = "$(ConvertTo-XlsxColumnName $cols.Count)$r"
        [void]$sb.Append("</sheetData><autoFilter ref=`"A1:$lastRef`"/></worksheet>")
        $parts["xl/worksheets/sheet$sheetIndex.xml"] = $sb.ToString()

        $safeName = ($sheet.Name -replace '[\[\]\:\*\?\/\\]', '_')
        if ($safeName.Length -gt 31) { $safeName = $safeName.Substring(0, 31) }
        $xmlName = ConvertTo-XmlText $safeName
        [void]$sheetEntries.Append("<sheet name=`"$xmlName`" sheetId=`"$sheetIndex`" r:id=`"rId$sheetIndex`"/>")
        $absRef = '$A$1:$' + (ConvertTo-XlsxColumnName $cols.Count) + '$' + $r
        [void]$definedNames.Append("<definedName name=`"_xlnm._FilterDatabase`" localSheetId=`"$($sheetIndex-1)`" hidden=`"1`">'$($xmlName -replace "'", "''")'!$absRef</definedName>")
    }

    $ct = New-Object System.Text.StringBuilder
    [void]$ct.Append($hdr).Append('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">')
    [void]$ct.Append('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>')
    [void]$ct.Append('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>')
    [void]$ct.Append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    $wbRels = New-Object System.Text.StringBuilder
    [void]$wbRels.Append($hdr).Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 1; $i -le $sheetIndex; $i++) {
        [void]$ct.Append("<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>")
        [void]$wbRels.Append("<Relationship Id=`"rId$i`" Type=`"$nsRel/worksheet`" Target=`"worksheets/sheet$i.xml`"/>")
    }
    [void]$ct.Append('</Types>')
    [void]$wbRels.Append("<Relationship Id=`"rId$($sheetIndex+1)`" Type=`"$nsRel/styles`" Target=`"styles.xml`"/></Relationships>")

    $parts['[Content_Types].xml'] = $ct.ToString()
    $parts['_rels/.rels'] = $hdr + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    $parts['xl/workbook.xml'] = $hdr + "<workbook xmlns=`"$nsMain`" xmlns:r=`"$nsRel`"><sheets>$sheetEntries</sheets><definedNames>$definedNames</definedNames></workbook>"
    $parts['xl/_rels/workbook.xml.rels'] = $wbRels.ToString()
    $parts['xl/styles.xml'] = $hdr + "<styleSheet xmlns=`"$nsMain`">" +
        '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm"/></numFmts>' +
        '<fonts count="5">' +
            '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/><family val="2"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FF9C5700"/><name val="Calibri"/><family val="2"/></font>' +
            '<font><b/><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/><family val="2"/></font>' +
        '</fonts><fills count="6">' +
            '<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill>' +
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>' +
        '</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="6">' +
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
            '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
            '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
            '<xf numFmtId="0" fontId="4" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

    # Write to a temp file first so a failure never leaves a half-written .xlsx behind.
    $tmp = "$Path.tmp"
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $fs = [System.IO.File]::Open($tmp, [System.IO.FileMode]::Create)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in $parts.Keys) {
                $sw = New-Object System.IO.StreamWriter($zip.CreateEntry($name).Open(), $utf8)
                try { $sw.Write($parts[$name]) } finally { $sw.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
    Move-Item -LiteralPath $tmp -Destination $Path -Force
}

function ConvertFrom-XlsxColumnName {
    param([string]$Letters)
    $n = 0
    foreach ($ch in $Letters.ToUpperInvariant().ToCharArray()) { $n = $n * 26 + ([int]$ch - 64) }
    $n
}

function Import-SimpleXlsx {
    # Reads every sheet of an .xlsx into objects keyed by the header row. Handles both files
    # written by this script and files re-saved by Excel (shared strings, date styles).
    param([Parameter(Mandatory)][string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $nsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    # FileShare.ReadWrite so a report that is open in Excel can still be read.
    $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        try {
            $readXml = {
                param([string]$Name)
                $e = $zip.GetEntry($Name)
                if (-not $e) { return $null }
                $sr = New-Object System.IO.StreamReader($e.Open())
                try { [xml]$sr.ReadToEnd() } finally { $sr.Dispose() }
            }

            $shared = New-Object System.Collections.Generic.List[string]
            $ss = & $readXml 'xl/sharedStrings.xml'
            if ($ss) {
                foreach ($si in $ss.GetElementsByTagName('si')) {
                    $shared.Add((@($si.GetElementsByTagName('t') | Where-Object { $_.ParentNode.LocalName -ne 'rPh' } | ForEach-Object { $_.InnerText }) -join ''))
                }
            }

            # Which cell styles are dates (built-in date formats or custom ones with d/m/y/h/s)
            $dateXf = @{}
            $st = & $readXml 'xl/styles.xml'
            if ($st) {
                $custom = @{}
                foreach ($nf in $st.GetElementsByTagName('numFmt')) { $custom[[int]$nf.GetAttribute('numFmtId')] = $nf.GetAttribute('formatCode') }
                $cellXfs = $st.GetElementsByTagName('cellXfs')
                if ($cellXfs.Count -gt 0) {
                    $i = 0
                    foreach ($xf in $cellXfs[0].ChildNodes) {
                        if ($xf.LocalName -ne 'xf') { continue }
                        $id = 0; [void][int]::TryParse($xf.GetAttribute('numFmtId'), [ref]$id)
                        $dateXf[$i] = (($id -ge 14 -and $id -le 22) -or ($id -ge 45 -and $id -le 47) -or
                                       ($custom.ContainsKey($id) -and (($custom[$id] -replace '"[^"]*"|\[[^\]]*\]|\\.', '') -match '[dmyhs]')))
                        $i++
                    }
                }
            }

            $wb   = & $readXml 'xl/workbook.xml'
            $rels = & $readXml 'xl/_rels/workbook.xml.rels'
            $relMap = @{}
            foreach ($rel in $rels.GetElementsByTagName('Relationship')) { $relMap[$rel.GetAttribute('Id')] = $rel.GetAttribute('Target') }

            $book = [ordered]@{}
            foreach ($sh in $wb.GetElementsByTagName('sheet')) {
                $target = [string]$relMap[$sh.GetAttribute('id', $nsRel)]
                if (-not $target) { continue }
                $entry = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target }
                $doc = & $readXml $entry
                if (-not $doc) { continue }

                $grid = New-Object System.Collections.Generic.List[object]
                foreach ($row in $doc.GetElementsByTagName('row')) {
                    $cells = @{}; $col = 0
                    foreach ($c in $row.ChildNodes) {
                        if ($c.LocalName -ne 'c') { continue }
                        $ref = $c.GetAttribute('r')
                        $col = if ($ref -match '^([A-Za-z]+)') { ConvertFrom-XlsxColumnName $Matches[1] } else { $col + 1 }
                        $t = $c.GetAttribute('t')
                        $vNode = @($c.GetElementsByTagName('v'))
                        $v = if ($vNode.Count) { $vNode[0].InnerText } else { $null }
                        $val = switch ($t) {
                            'inlineStr' { $is = @($c.GetElementsByTagName('is')); if ($is.Count) { $is[0].InnerText } else { $null } }
                            's'         { if ($null -ne $v) { $shared[[int]$v] } else { $null } }
                            'b'         { if ($v -eq '1') { 'Yes' } else { 'No' } }
                            'str'       { $v }
                            'e'         { $v }
                            default {
                                if ($null -eq $v -or $v -eq '') { $null }
                                else {
                                    $d = [double]::Parse($v, [System.Globalization.NumberStyles]::Float, $inv)
                                    $sIdx = $c.GetAttribute('s')
                                    if ($sIdx -and $dateXf[[int]$sIdx]) { [datetime]::FromOADate($d) }
                                    elseif ($d -eq [math]::Truncate($d) -and [math]::Abs($d) -lt 1e15) { [long]$d }
                                    else { $d }
                                }
                            }
                        }
                        $cells[$col] = $val
                    }
                    $grid.Add($cells)
                }
                if ($grid.Count -eq 0) { $book[$sh.GetAttribute('name')] = @(); continue }

                $header = $grid[0]
                $colIdx = @($header.Keys | Where-Object { $header[$_] } | Sort-Object)
                $rows = New-Object System.Collections.Generic.List[object]
                for ($r = 1; $r -lt $grid.Count; $r++) {
                    $o = [ordered]@{}; $any = $false
                    foreach ($ci in $colIdx) {
                        $o[[string]$header[$ci]] = $grid[$r][$ci]
                        if ($null -ne $grid[$r][$ci] -and [string]$grid[$r][$ci] -ne '') { $any = $true }
                    }
                    if ($any) { $rows.Add([pscustomobject]$o) }
                }
                $book[$sh.GetAttribute('name')] = $rows.ToArray()
            }
            $book
        } finally { $zip.Dispose() }
    } finally { $fs.Dispose() }
}

function Join-RowsUnion {
    # Gives every row the same columns (union, first-seen order) so they can share one sheet.
    param([object[]]$Rows)
    $cols = New-Object System.Collections.Generic.List[string]
    foreach ($r in $Rows) { foreach ($p in $r.PSObject.Properties.Name) { if (-not $cols.Contains($p)) { $cols.Add($p) } } }
    foreach ($r in $Rows) {
        $o = [ordered]@{}
        foreach ($c in $cols) { $o[$c] = if ($r.PSObject.Properties[$c]) { $r.$c } else { $null } }
        [pscustomobject]$o
    }
}

function Invoke-Consolidation {
    param([string]$SourcePath, [string]$DestPath)

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) { throw "Consolidate folder not found: $SourcePath" }
    if (-not (Test-Path -LiteralPath $DestPath)) { New-Item -ItemType Directory -Path $DestPath -Force | Out-Null }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $outXlsx = Join-Path $DestPath "OSHealthCheck_Consolidated_$stamp.xlsx"
    $script:LogFile = Join-Path $DestPath "OSHealthCheck_Consolidated_$stamp.log"

    $files = @(Get-ChildItem -LiteralPath $SourcePath -Filter '*.xlsx' -File -Recurse |
               Where-Object { $_.Name -notlike 'OSHealthCheck_Consolidated_*' -and $_.Name -notlike '~$*' } |
               Sort-Object FullName)
    Write-Log "Consolidating from $SourcePath - $($files.Count) .xlsx file(s) found"

    $books = @{}
    $best  = @{}   # SERVER -> @{ When; File; Row }
    $skipped = 0
    foreach ($f in $files) {
        try { $book = Import-SimpleXlsx -Path $f.FullName }
        catch { Write-Log "Skipped $($f.Name): cannot read ($($_.Exception.Message))" 'WARN'; $skipped++; continue }
        $sum = @($book['Summary'] | Where-Object { $_ })
        if ($sum.Count -eq 0 -or -not $sum[0].PSObject.Properties['Server'] -or -not $sum[0].PSObject.Properties['Status']) {
            Write-Log "Skipped $($f.Name): not an OS health check report" 'WARN'; $skipped++; continue
        }
        $books[$f.FullName] = $book
        foreach ($s in $sum) {
            if (-not $s.Server) { continue }
            $key  = ([string]$s.Server).ToUpperInvariant()
            $when = $s.CollectedAt -as [datetime]
            if (-not $when) { $when = $f.LastWriteTime }
            if (-not $best.ContainsKey($key)) {
                $best[$key] = @{ When = $when; File = $f.FullName; Row = $s }
            } elseif ($when -gt $best[$key].When) {
                Write-Log "$($s.Server): newer report $($f.Name) replaces $(Split-Path $best[$key].File -Leaf)" 'WARN'
                $best[$key] = @{ When = $when; File = $f.FullName; Row = $s }
            } else {
                Write-Log "$($s.Server): older duplicate in $($f.Name) ignored" 'WARN'
            }
        }
    }
    if ($best.Count -eq 0) { Write-Log 'No valid reports found - nothing written' 'ERROR'; return }

    $rank = @{ 'Critical' = 0; 'Warning' = 1; 'OK' = 2 }
    $servers = @($best.Keys | Sort-Object @{ Expression = { $r = $rank[[string]$best[$_].Row.Status]; if ($null -eq $r) { 3 } else { $r } } }, @{ Expression = { $_ } })

    $sheets = @(@{ Name = 'Summary'; Rows = @(Join-RowsUnion -Rows @($servers | ForEach-Object { $best[$_].Row })) })
    foreach ($name in 'Disks', 'StoppedServices', 'Events') {
        $rows = foreach ($key in $servers) {
            $book = $books[$best[$key].File]
            @($book[$name]) | Where-Object { $_ -and $_.PSObject.Properties['Server'] -and ([string]$_.Server).ToUpperInvariant() -eq $key }
        }
        $sheets += @{ Name = $name; Rows = @(Join-RowsUnion -Rows @($rows)) }
    }

    Export-SimpleXlsx -Path $outXlsx -Sheets $sheets
    $st = @($servers | ForEach-Object { [string]$best[$_].Row.Status })
    Write-Log ("Consolidated {0} server(s) from {1} report(s), {2} file(s) skipped. OK={3} Warning={4} Critical={5}" -f `
        $servers.Count, $books.Count, $skipped, @($st | Where-Object { $_ -eq 'OK' }).Count,
        @($st | Where-Object { $_ -eq 'Warning' }).Count, @($st | Where-Object { $_ -eq 'Critical' }).Count)
    Write-Log "Excel written: $outXlsx"
}

# ---------------------------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------------------------
if ($ConsolidatePath) {
    $dest = if ($PSBoundParameters.ContainsKey('OutputPath')) { $OutputPath } else { $ConsolidatePath }
    Invoke-Consolidation -SourcePath $ConsolidatePath -DestPath $dest
    return
}

$targets = New-Object System.Collections.Generic.List[string]
foreach ($n in $pipedNames) { $targets.Add($n) }
if ($ComputerListPath) {
    if (-not (Test-Path -LiteralPath $ComputerListPath)) { throw "Computer list not found: $ComputerListPath" }
    Get-Content -LiteralPath $ComputerListPath | ForEach-Object { $targets.Add($_) }
}
$targets = @($targets | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') } | ForEach-Object {
    if ($_ -in @('.', 'localhost', '127.0.0.1', '::1')) { $env:COMPUTERNAME } else { $_ }   # report the real name
} | Sort-Object -Unique)
if ($targets.Count -eq 0) { $targets = @($env:COMPUTERNAME) }

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$stamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
# One server -> its name goes in the file name, so per-VM reports never collide when collected.
$baseName = if ($targets.Count -eq 1) { "OSHealthCheck_$($targets[0] -replace '[^\w\.-]', '_')_$stamp" } else { "OSHealthCheck_$stamp" }
$xlsxPath = Join-Path $OutputPath "$baseName.xlsx"
$csvPath  = Join-Path $OutputPath "$baseName.csv"
$script:LogFile = Join-Path $OutputPath "$baseName.log"

Write-Log "Build $ScriptBuild - run by $env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME"
Write-Log "Targets ($($targets.Count)): $($targets -join ', ')"

$collectorArgs = @($EventDays, $ServiceExclusions, $MaxEvents)
$rawByTarget   = @{}
$errByTarget   = @{}

$localTargets  = @($targets | Where-Object { Test-IsLocalComputer $_ })
$remoteTargets = @($targets | Where-Object { -not (Test-IsLocalComputer $_) })

foreach ($t in $localTargets) {
    Write-Log "Checking $t (local)"
    try   { $rawByTarget[$t] = & $collector @collectorArgs }
    catch { $errByTarget[$t] = $_.Exception.Message; Write-Log "$t : $($_.Exception.Message)" 'ERROR' }
}

if ($remoteTargets.Count -gt 0) {
    Write-Log "Checking $($remoteTargets.Count) remote server(s) over WinRM (throttle $ThrottleLimit)"
    $icParams = @{
        ComputerName  = $remoteTargets
        ScriptBlock   = $collector
        ArgumentList  = $collectorArgs
        ThrottleLimit = $ThrottleLimit
        ErrorAction   = 'SilentlyContinue'
        ErrorVariable = 'remoteErrors'
    }
    # Fail fast on dead servers instead of hanging (default open timeout is 3 minutes).
    try { $icParams.SessionOption = New-PSSessionOption -OpenTimeout 30000 -OperationTimeout 600000 } catch { }
    if ($Credential) { $icParams.Credential = $Credential }
    $remoteErrors = $null
    $remoteResults = @(Invoke-Command @icParams)

    foreach ($res in $remoteResults) {
        $match = $remoteTargets | Where-Object { $_ -eq $res.PSComputerName } | Select-Object -First 1
        if ($match) { $rawByTarget[$match] = $res }
    }
    foreach ($t in $remoteTargets) {
        if ($rawByTarget.ContainsKey($t)) { continue }
        $msg = @($remoteErrors | Where-Object {
                    ($_.TargetObject -and [string]$_.TargetObject -eq $t) -or
                    ($_.OriginInfo -and $_.OriginInfo.PSComputerName -eq $t) -or
                    ($_.ToString() -match "\b$([regex]::Escape($t))\b")
               } | Select-Object -First 1 | ForEach-Object { $_.Exception.Message })
        $errByTarget[$t] = if ($msg) { ($msg[0] -replace '\s+', ' ').Trim() } else { 'No data returned (unknown remoting error)' }
        Write-Log "$t : $($errByTarget[$t])" 'ERROR'
    }
}

$summary = @(); $disks = @(); $services = @(); $events = @()
foreach ($t in $targets) {
    $rows = ConvertTo-HealthRows -Target $t -Raw $rawByTarget[$t] -ConnectError $errByTarget[$t]
    $summary  += $rows.Summary
    $disks    += $rows.Disks
    $services += $rows.Services
    $events   += $rows.Events
}

try {
    Export-SimpleXlsx -Path $xlsxPath -Sheets @(
        @{ Name = 'Summary';         Rows = $summary }
        @{ Name = 'Disks';           Rows = $disks }
        @{ Name = 'StoppedServices'; Rows = $services }
        @{ Name = 'Events';          Rows = $events }
    )
    Write-Log "Excel written: $xlsxPath"
} catch {
    Write-Log "Excel export failed: $($_.Exception.Message) - writing CSV instead" 'ERROR'
    try {
        $summary | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV written: $csvPath"
    } catch { Write-Log "CSV export failed: $($_.Exception.Message)" 'ERROR' }
}

Write-Host ''
$summary | Format-Table Server, Status, Issues -AutoSize -Wrap | Out-String -Width 220 | Write-Host
Write-Log ("Done. OK={0} Warning={1} Critical={2}" -f @($summary | Where-Object Status -eq 'OK').Count,
           @($summary | Where-Object Status -eq 'Warning').Count, @($summary | Where-Object Status -eq 'Critical').Count)

if ($PassThru) { $summary }
}
