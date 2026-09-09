#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only Windows Update / servicing diagnostics across a list of guest VMs,
    collected entirely through VMware Tools Guest Operations (Invoke-VMScript). No
    WinRM / PsExec / SMB / RDP guest network path required - only vCenter + Guest
    Operations privileges, from a machine that already has an active Connect-VIServer
    session.

.DESCRIPTION
    Companion diagnostic script for troubleshooting:
        "Windows Update - Error encountered - We could not complete the install because
         an update service was shutting down."

    Runs against every VM listed in -VMListPath (one name per line, '#' comments
    ignored). Each VM is fully isolated - a VM that can't be found, is powered off,
    has no VMware Tools, or throws mid-collection is recorded as "Unable to Check" /
    "Skipped" and the batch continues with the next VM.

    Collects, WITHOUT CHANGING ANYTHING on any guest:
      - Time sync: current time, time zone, w32time service/status/source/config, and
        whether VMware Tools host time sync and Windows Time (w32time) are both active -
        a combination Microsoft/VMware guidance flags for domain-joined VMs.
      - Service state for wuauserv, BITS, CryptSvc, UsoSvc, WaaSMedicSvc, DoSvc, plus
        recent unexpected-stop/crash events (Service Control Manager 7031/7034/7036)
        for each, so a service dying mid-install is visible as evidence, not guesswork.
      - Windows Update event log evidence: System log WU install failures (Event ID 20),
        the Microsoft-Windows-WindowsUpdateClient/Operational log, and a keyword tail of
        CBS.log around servicing/shutdown activity.
      - Pending-reboot indicators (CBS RebootPending/RebootInProgress, WU Auto Update
        RebootRequired, PendingFileRenameOperations).
      - DISM component-store health (read-only CheckHealth flag only - never RestoreHealth).
      - Windows Update policy/GPO configuration: WSUS vs Windows Update for Business vs
        Microsoft Update, DisableDualScan, DoNotConnectToWindowsUpdateInternetLocations,
        UseWUServer, TargetGroup.
      - Real DNS + TLS handshake tests (not just TCP) to the Windows Update endpoints.
      - Azure Arc / Azure Update Manager footprint (himds service, agent folder, GC
        patch-extension presence) so an AUM-vs-local-WU orchestration conflict is visible
        as evidence rather than assumed.
      - Exact failed KB / HResult from Windows Update history (COM
        Microsoft.Update.Session), with a plain-English meaning for well-documented
        HResults only - unrecognised codes are reported as such, never guessed.
      - Installed VMware Tools version/status (informational - VMware Tools does not own
        or manage any Windows Update service).

    Anything this script cannot determine reliably from inside the guest (for example,
    the Azure-side AUM patch-orchestration mode, which is an Azure/Arc resource property,
    not a guest artifact) is reported as "Manual/External Required" with an explanation -
    never guessed or left blank.

.PARAMETER VMListPath     Text file of VM names, one per line (default C:\temp\vmlist.txt).
                           '#' lines and blanks are ignored; names are de-duplicated.
.PARAMETER CredentialPath Export-Clixml PSCredential for the guest admin. If omitted or
                           not found, you are prompted with Get-Credential instead.
.PARAMETER OutputPath     Folder on the admin machine for the consolidated CSV/JSON/log.
.PARAMETER EventDays      How many days back to pull WU/service-control events (default 30).
.PARAMETER HistoryCount   How many Windows Update history records to pull (default 40).
.PARAMETER ToolsWaitSecs  Invoke-VMScript VMware Tools wait, seconds (default 180).

.EXAMPLE
    Connect-VIServer vcenter01
    "TB-KRTN-APP02`nTB-KRTN-APP03" | Set-Content C:\temp\vmlist.txt
    .\Invoke-WindowsUpdateDiagnostics.ps1

.EXAMPLE
    Connect-VIServer vcenter01
    Get-Credential | Export-Clixml C:\temp\wincred.xml
    .\Invoke-WindowsUpdateDiagnostics.ps1 -VMListPath C:\temp\wsus_servers.txt -EventDays 45

.NOTES
    100% read-only. Requires PowerCLI, an existing Connect-VIServer session, and vCenter
    "Guest Operation Program Execution/Query" privileges on every target VM.

    Run the companion Invoke-WindowsUpdateRemediation.ps1 SEPARATELY, never in the same
    pass, so a diagnostic run can never accidentally change any server.
#>

[CmdletBinding()]
param(
    [string] $VMListPath     = 'C:\temp\vmlist.txt',
    [string] $CredentialPath = 'C:\temp\wincred.xml',
    [string] $OutputPath     = (Join-Path $PSScriptRoot 'WindowsUpdate_Diagnostics'),

    [ValidateRange(1, 365)]
    [int]    $EventDays = 30,

    [ValidateRange(1, 500)]
    [int]    $HistoryCount = 40,

    [ValidateRange(30, 3600)]
    [int]    $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

Write-Host "Invoke-WindowsUpdateDiagnostics.ps1 - read-only Windows Update / servicing evidence across a VM list  [build 2026-09-09b]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray

# =====================================================================================
# 1. Pre-flight (read-only)
# =====================================================================================
if (-not (Get-Command Invoke-VMScript -ErrorAction SilentlyContinue)) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop }
    catch { throw "VMware PowerCLI (VMware.VimAutomation.Core) is not available. Install PowerCLI and retry." }
}
$connectedServers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($connectedServers.Count -eq 0) { throw "No connected vCenter session. Run Connect-VIServer <vcenter> first." }
Write-Host ("Connected vCenter(s): {0}" -f (($connectedServers | ForEach-Object { $_.Name }) -join ', ')) -ForegroundColor Gray

if (-not (Test-Path -LiteralPath $VMListPath)) { throw "VM list not found: $VMListPath" }
$vmNames = @(Get-Content -LiteralPath $VMListPath |
             ForEach-Object { $_.Trim() } |
             Where-Object { $_ -and -not $_.StartsWith('#') } |
             Select-Object -Unique)
if ($vmNames.Count -eq 0) { throw "VM list '$VMListPath' contained no usable VM names." }
Write-Host ("VMs to process: {0}" -f $vmNames.Count) -ForegroundColor Gray

if (Test-Path -LiteralPath $CredentialPath) {
    try {
        $GuestCredential = Import-Clixml -LiteralPath $CredentialPath
        if (-not ($GuestCredential -is [pscredential])) { throw "did not deserialise to a PSCredential" }
    } catch {
        throw "Failed to import guest credential from '$CredentialPath': $($_.Exception.Message)"
    }
} else {
    $GuestCredential = Get-Credential -Message "Guest administrator credential (used for every VM in the list)"
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$CsvPath  = Join-Path $OutputPath "WU_Diagnostics_$RunStamp.csv"
$JsonPath = Join-Path $OutputPath "WU_Diagnostics_$RunStamp.json"
$LogPath  = Join-Path $OutputPath "WU_Diagnostics_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Diagnostics run started. VMs=$($vmNames.Count). EventDays=$EventDays HistoryCount=$HistoryCount."

# =====================================================================================
# 2. In-guest READ-ONLY payload (unchanged per-VM logic; identical to the single-VM
#    version, just invoked once per VM in the loop below)
#    {{EVENTDAYS}} / {{HISTORYCOUNT}} are replaced with integer literals before each run.
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'
$EventDays    = {{EVENTDAYS}}
$HistoryCount = {{HISTORYCOUNT}}
$Since        = (Get-Date).AddDays(-$EventDays)

$rows = New-Object System.Collections.Generic.List[object]
function Add-Row {
    param($Section, $Check, $Status, $Value, $Detail, $NextStep = '')
    $rows.Add([pscustomobject]@{
        Section = $Section; Check = $Check; Status = $Status
        Value = "$Value"; Detail = $Detail; NextStep = $NextStep
    }) | Out-Null
}
function Get-RegValue {
    param([string]$Path, [string]$Name)
    try {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return $item.$Name
    } catch { return $null }
}

# --- Machine identity -----------------------------------------------------------------
$ComputerName = $env:COMPUTERNAME
try { $OSInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop } catch { $OSInfo = $null }
try { $CSInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { $CSInfo = $null }
$OSCaption   = if ($OSInfo) { $OSInfo.Caption } else { 'Unknown' }
$OSBuild     = if ($OSInfo) { $OSInfo.BuildNumber } else { 'Unknown' }
$LastBoot    = if ($OSInfo) { $OSInfo.LastBootUpTime } else { $null }
$IsDomain    = if ($CSInfo) { [bool]$CSInfo.PartOfDomain } else { $false }

# =======================================================================================
# SECTION: Time Sync
# =======================================================================================
Add-Row 'Time Sync' 'Guest local time / UTC' 'Info' ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')) "UTC now: $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))"
try { $tz = (Get-TimeZone).Id } catch { $tz = 'Unknown' }
Add-Row 'Time Sync' 'Time zone' 'Info' $tz ''

try {
    $w32svc = Get-Service -Name W32Time -ErrorAction Stop
    $w32Status = if ($w32svc.Status -eq 'Running') { 'Healthy' } else { 'Warning' }
    $w32Detail = if ($w32svc.Status -eq 'Running') { 'Windows Time service is running.' } else { "Windows Time service is '$($w32svc.Status)'. Time sync is relying solely on VMware Tools host sync, which is a coarser/periodic correction, not domain-hierarchy aware." }
    Add-Row 'Time Sync' 'W32Time service state' $w32Status "$($w32svc.Status) (StartType: $($w32svc.StartType))" $w32Detail
} catch {
    Add-Row 'Time Sync' 'W32Time service state' 'Unable to Check' 'n/a' "Could not query W32Time service: $($_.Exception.Message)"
}

try {
    $w32qs = & w32tm /query /status 2>&1 | Out-String
    Add-Row 'Time Sync' 'w32tm /query /status' 'Info' 'see Detail' $w32qs.Trim()
} catch {
    Add-Row 'Time Sync' 'w32tm /query /status' 'Unable to Check' 'n/a' "w32tm query failed: $($_.Exception.Message)"
}
try {
    $w32src = & w32tm /query /source 2>&1 | Out-String
    Add-Row 'Time Sync' 'w32tm /query /source' 'Info' ($w32src.Trim()) ''
} catch { }
try {
    $w32cfg = & w32tm /query /configuration 2>&1 | Out-String
    $cfgType = ($w32cfg -split "`r?`n" | Where-Object { $_ -match '^\s*Type:' } | Select-Object -First 1)
    Add-Row 'Time Sync' 'w32tm configured Type' 'Info' ($(if ($cfgType) { $cfgType.Trim() } else { 'n/a' })) 'NT5DS = follow domain hierarchy (expected on domain-joined machines). NTP = explicit server list. NoSync = time provider disabled.'
} catch { }

if ($IsDomain -and $w32svc -and $w32svc.Status -eq 'Running') {
    Add-Row 'Time Sync' 'VMware host sync vs W32Time conflict check' 'Warning' 'Domain-joined + W32Time running' "This is a domain-joined machine with Windows Time active (normal - it should follow the domain hierarchy). If VMware Tools 'Synchronize guest time with host' (SyncTimeWithHost) is ALSO enabled at the VM/VMX level, Microsoft and VMware both document this dual-source combination as not recommended for domain members: the two correction sources can fight, producing log noise and occasional larger jumps. It is a real hygiene finding but is NOT expected to produce a 'service is shutting down' Windows Update error by itself - track it separately from the WU failure." 'Confirm the current VMX SyncTimeWithHost setting from vCenter for this VM. If domain-joined, the general guidance is to disable VMware Tools periodic time sync and let w32time follow the PDC emulator hierarchy instead. Do this as a follow-up, not as the fix for the WU error.'
} elseif (-not $IsDomain) {
    Add-Row 'Time Sync' 'VMware host sync vs W32Time conflict check' 'Healthy' 'Workgroup machine' 'Not domain-joined, so VMware host time sync is the expected/recommended time source here. No conflict.' ''
}

# =======================================================================================
# SECTION: Update Services
# =======================================================================================
$serviceNames = @('wuauserv','BITS','CryptSvc','UsoSvc','WaaSMedicSvc','DoSvc')
$scmEvents = $null
try {
    $scmEvents = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 7031,7034,7036; StartTime = $Since } -ErrorAction Stop)
} catch {
    if ($_.Exception.Message -notmatch 'No events were found') { $scmEvents = $null }
    else { $scmEvents = @() }
}

foreach ($svcName in $serviceNames) {
    try {
        $svc = Get-Service -Name $svcName -ErrorAction Stop
        $startType = $svc.StartType
        $expectedRunning = $svcName -in @('wuauserv','CryptSvc')  # BITS/UsoSvc/WaaSMedicSvc/DoSvc are normally demand-start
        $status = 'Healthy'
        $detail = "Status=$($svc.Status), StartType=$startType."
        if ($svcName -eq 'wuauserv' -and $svc.Status -ne 'Running' -and $startType -eq 'Disabled') {
            $status = 'Problem'; $detail += ' wuauserv is Disabled - Windows Update cannot function at all until this is changed.'
        } elseif ($startType -eq 'Disabled' -and $svcName -in @('BITS','CryptSvc')) {
            $status = 'Problem'; $detail += " $svcName is Disabled but is required by Windows Update (BITS for transfer, CryptSvc for signature verification)."
        } elseif ($expectedRunning -and $svc.Status -ne 'Running') {
            $status = 'Warning'; $detail += " Expected to be Running (or will auto-start on demand); currently $($svc.Status)."
        }
        # Recent unexpected-stop / crash events referencing this service.
        $svcEvts = @()
        if ($scmEvents) { $svcEvts = @($scmEvents | Where-Object { $_.Message -match [regex]::Escape($svcName) -or ($svc.DisplayName -and $_.Message -match [regex]::Escape($svc.DisplayName)) }) }
        if ($svcEvts.Count -gt 0) {
            $latest = ($svcEvts | Sort-Object TimeCreated -Descending | Select-Object -First 1)
            $status = if ($status -eq 'Healthy') { 'Warning' } else { $status }
            $detail += " $($svcEvts.Count) unexpected stop/crash/restart event(s) (SCM 7031/7034/7036) in the last $EventDays day(s); most recent $($latest.TimeCreated): $($latest.Message -replace '\s+',' ')."
            Add-Row 'Update Services' "$svcName - recent SCM stop/crash events" 'Warning' "$($svcEvts.Count) event(s)" ($latest.Message -replace '\s+',' ') 'Correlate this timestamp against the failed update install time in Windows Update history (below) - a service dying here at the same moment is direct evidence for the "service is shutting down" error.'
        }
        Add-Row 'Update Services' "$svcName state" $status "$($svc.Status)/$startType" $detail ''
    } catch {
        Add-Row 'Update Services' "$svcName state" 'Unable to Check' 'n/a' "Service not found or query failed: $($_.Exception.Message)" ''
    }
}
if ($null -eq $scmEvents) {
    Add-Row 'Update Services' 'SCM 7031/7034/7036 event query' 'Unable to Check' 'n/a' 'Could not query the System log for service-control-manager stop/crash events.' ''
}

# =======================================================================================
# SECTION: Windows Update Event Log Evidence
# =======================================================================================
try {
    $wu20 = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; ProviderName = 'Microsoft-Windows-WindowsUpdateClient'; Id = 20; StartTime = $Since } -ErrorAction Stop)
    if ($wu20.Count -gt 0) {
        $sample = $wu20 | Sort-Object TimeCreated -Descending | Select-Object -First 10 | ForEach-Object { "$($_.TimeCreated): $($_.Message -replace '\s+',' ')" }
        Add-Row 'WU Event Logs' 'System log Event ID 20 (WU install failure)' 'Problem' "$($wu20.Count) event(s)" ($sample -join ' || ') 'These are Windows Update Agent install-failure records with the HResult in the message text - match the timestamp of the most recent one to the update history entries below.'
    } else {
        Add-Row 'WU Event Logs' 'System log Event ID 20 (WU install failure)' 'Healthy' '0 events' "No install-failure events in the last $EventDays day(s)." ''
    }
} catch {
    if ($_.Exception.Message -match 'No events were found') {
        Add-Row 'WU Event Logs' 'System log Event ID 20 (WU install failure)' 'Healthy' '0 events' "No install-failure events in the last $EventDays day(s)." ''
    } else {
        Add-Row 'WU Event Logs' 'System log Event ID 20 (WU install failure)' 'Unable to Check' 'n/a' "Query failed: $($_.Exception.Message)" ''
    }
}

try {
    $wuOp = @(Get-WinEvent -LogName 'Microsoft-Windows-WindowsUpdateClient/Operational' -ErrorAction Stop |
              Where-Object { $_.TimeCreated -ge $Since -and $_.LevelDisplayName -in @('Error','Warning') })
    if ($wuOp.Count -gt 0) {
        $sample = $wuOp | Sort-Object TimeCreated -Descending | Select-Object -First 15 | ForEach-Object { "$($_.TimeCreated) [$($_.LevelDisplayName)] Id=$($_.Id): $($_.Message -replace '\s+',' ')" }
        Add-Row 'WU Event Logs' 'WindowsUpdateClient/Operational errors+warnings' 'Warning' "$($wuOp.Count) event(s)" ($sample -join ' || ') 'Look for Event ID 25/31/34/35 around the failure time; these show exactly which phase (download, install, finalize) failed and why.'
    } else {
        Add-Row 'WU Event Logs' 'WindowsUpdateClient/Operational errors+warnings' 'Healthy' '0 events' "No Error/Warning entries in the last $EventDays day(s)." ''
    }
} catch {
    Add-Row 'WU Event Logs' 'WindowsUpdateClient/Operational errors+warnings' 'Unable to Check' 'n/a' "Log not present or query failed: $($_.Exception.Message)" 'This operational log can be disabled by GPO/local policy; if absent, rely on System log Event ID 20 and Windows Update history instead.'
}

$cbsPath = 'C:\Windows\Logs\CBS\CBS.log'
if (Test-Path -LiteralPath $cbsPath) {
    try {
        $cbsTail = Get-Content -LiteralPath $cbsPath -Tail 4000 -ErrorAction Stop
        $hits = @($cbsTail | Select-String -Pattern 'shutting down|shutdown|CBS_E_|servicing stack|component store' -SimpleMatch:$false | Select-Object -Last 25)
        if ($hits.Count -gt 0) {
            Add-Row 'WU Event Logs' 'CBS.log keyword tail (last 4000 lines)' 'Warning' "$($hits.Count) matching line(s)" (($hits | ForEach-Object { $_.Line.Trim() }) -join ' || ') 'Correlate line timestamps (prefix of each CBS.log line) with the failure time reported in the Windows Update UI/history.'
        } else {
            Add-Row 'WU Event Logs' 'CBS.log keyword tail (last 4000 lines)' 'Healthy' '0 matches' 'No shutdown/CBS error/servicing-stack keywords in the tail of CBS.log.' ''
        }
    } catch {
        Add-Row 'WU Event Logs' 'CBS.log keyword tail' 'Unable to Check' 'n/a' "Could not read CBS.log: $($_.Exception.Message)" ''
    }
} else {
    Add-Row 'WU Event Logs' 'CBS.log keyword tail' 'Unable to Check' 'n/a' 'CBS.log was not found at the expected path.' ''
}

# =======================================================================================
# SECTION: Pending Reboot
# =======================================================================================
$rebootFlags = New-Object System.Collections.Generic.List[string]
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')    { $rebootFlags.Add('CBS RebootPending') }
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress') { $rebootFlags.Add('CBS RebootInProgress') }
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')   { $rebootFlags.Add('WU Auto Update RebootRequired') }
$pfro = Get-RegValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations'
if ($pfro) { $rebootFlags.Add('PendingFileRenameOperations present') }
if ($rebootFlags.Count -gt 0) {
    Add-Row 'Pending Reboot' 'Pending reboot indicators' 'Problem' ($rebootFlags -join '; ') 'One or more pending-reboot flags are set. Windows Update will refuse or fail to start a new install session while a prior update is waiting on a reboot - this alone can produce install-aborted / service-busy style errors.' 'Schedule and perform a reboot during a maintenance window, THEN retry Windows Update. Do this before any Windows Update component reset.'
} else {
    Add-Row 'Pending Reboot' 'Pending reboot indicators' 'Healthy' 'None found' 'No CBS/WU/PendingFileRename reboot flags set.' ''
}

# =======================================================================================
# SECTION: Component Store Health (read-only CheckHealth only)
# =======================================================================================
try {
    $dismOut = & dism.exe /online /cleanup-image /checkhealth 2>&1 | Out-String
    if ($dismOut -match 'No component store corruption detected') {
        Add-Row 'Servicing Stack' 'DISM /CheckHealth (read-only)' 'Healthy' 'No corruption detected' $dismOut.Trim() ''
    } elseif ($dismOut -match 'repairable') {
        Add-Row 'Servicing Stack' 'DISM /CheckHealth (read-only)' 'Warning' 'Repairable corruption flagged' $dismOut.Trim() 'Run the deeper (still read-only) DISM /ScanHealth to confirm before considering DISM /RestoreHealth, which downloads repair source content and does modify the system.'
    } else {
        Add-Row 'Servicing Stack' 'DISM /CheckHealth (read-only)' 'Unable to Check' 'n/a' $dismOut.Trim() ''
    }
} catch {
    Add-Row 'Servicing Stack' 'DISM /CheckHealth (read-only)' 'Unable to Check' 'n/a' "dism.exe /checkhealth failed: $($_.Exception.Message)" ''
}

# =======================================================================================
# SECTION: WU Policy / GPO Configuration
# =======================================================================================
$wuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$auPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
$wuServer        = Get-RegValue -Path $auPolicyPath -Name 'WUServer'
$wuStatusServer  = Get-RegValue -Path $auPolicyPath -Name 'WUStatusServer'
$useWUServer     = Get-RegValue -Path $auPolicyPath -Name 'UseWUServer'
$noAutoUpdate    = Get-RegValue -Path $auPolicyPath -Name 'NoAutoUpdate'
$auOptions       = Get-RegValue -Path $auPolicyPath -Name 'AUOptions'
$disableDualScan = Get-RegValue -Path $wuPolicyPath -Name 'DisableDualScan'
$doNotConnect    = Get-RegValue -Path $wuPolicyPath -Name 'DoNotConnectToWindowsUpdateInternetLocations'
$targetGroup     = Get-RegValue -Path $wuPolicyPath -Name 'TargetGroup'
$deferQuality    = Get-RegValue -Path $wuPolicyPath -Name 'DeferQualityUpdatesPeriodInDays'
$deferFeature    = Get-RegValue -Path $wuPolicyPath -Name 'DeferFeatureUpdatesPeriodInDays'
$manageBuilds    = Get-RegValue -Path $wuPolicyPath -Name 'ManagePreviewBuilds'

if ($wuServer -and $useWUServer -eq 1) {
    $sourceMode = "WSUS ($wuServer)"
} elseif ($deferQuality -or $deferFeature -or $manageBuilds) {
    $sourceMode = 'Windows Update for Business (deferral policies set, no WSUS server)'
} elseif ($doNotConnect -eq 1) {
    $sourceMode = 'BLOCKED - DoNotConnectToWindowsUpdateInternetLocations=1 with no WSUS server configured'
} else {
    $sourceMode = 'Microsoft Update / default (no WSUS, no WU-for-Business deferral policy detected)'
}
$sourceStatus = if ($sourceMode -match '^BLOCKED') { 'Problem' } else { 'Info' }
Add-Row 'WU Policy' 'Configured update source' $sourceStatus $sourceMode "WUServer='$wuServer' WUStatusServer='$wuStatusServer' UseWUServer=$useWUServer NoAutoUpdate=$noAutoUpdate AUOptions=$auOptions" ''

if ($doNotConnect -eq 1) {
    Add-Row 'WU Policy' 'DoNotConnectToWindowsUpdateInternetLocations' 'Problem' '1 (blocked)' 'This policy blocks the client from contacting Windows Update / Microsoft Update over the internet, regardless of WSUS settings.' 'If this server should be able to reach Windows Update directly (e.g. for AUM/Arc-driven patching), remove or set this value to 0 via GPO/registry.'
} else {
    Add-Row 'WU Policy' 'DoNotConnectToWindowsUpdateInternetLocations' 'Healthy' $(if ($null -eq $doNotConnect) { 'Not set' } else { $doNotConnect }) 'Direct internet connections to Windows Update are not blocked by this policy.' ''
}

if ($wuServer -and $useWUServer -eq 1 -and ($null -eq $disableDualScan -or $disableDualScan -eq 0)) {
    Add-Row 'WU Policy' 'DisableDualScan vs WSUS' 'Warning' "WUServer set, DisableDualScan=$disableDualScan" 'A WSUS server is configured but DisableDualScan is not set to 1. On builds that support dual scan this can let the client simultaneously evaluate WSUS content AND Windows Update for Business / Microsoft Update content, which is exactly the kind of overlapping-orchestration condition that can produce competing update sessions.' 'If this server is meant to be governed solely by WSUS or solely by AUM/cloud-based update policy, set DisableDualScan explicitly (1 = WSUS/ConfigMgr only, no cloud dual-scan) to remove the ambiguity.'
}

# =======================================================================================
# SECTION: Connectivity (DNS + real TLS handshake, not just TCP)
# =======================================================================================
function Test-HttpsEndpoint {
    param([string]$TargetHost)
    $r = [ordered]@{ Host = $TargetHost; DnsResolved = $false; DnsAddresses = ''; TlsOk = $false; HttpStatus = $null; Error = $null }
    try {
        $addrs = [System.Net.Dns]::GetHostAddresses($TargetHost)
        $r.DnsResolved = $true
        $r.DnsAddresses = ($addrs | ForEach-Object { $_.IPAddressToString }) -join ', '
    } catch {
        $r.Error = "DNS resolution failed: $($_.Exception.Message)"
        return [pscustomobject]$r
    }
    try {
        $resp = Invoke-WebRequest -Uri "https://$TargetHost/" -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        $r.TlsOk = $true
        $r.HttpStatus = [int]$resp.StatusCode
    } catch [System.Net.WebException] {
        if ($_.Exception.Response) {
            # Any HTTP-level response - even 403/404/405 - proves the TLS handshake succeeded.
            $r.TlsOk = $true
            $r.HttpStatus = [int]$_.Exception.Response.StatusCode
        } else {
            $r.Error = $_.Exception.Message
        }
    } catch {
        $r.Error = $_.Exception.Message
    }
    [pscustomobject]$r
}

$wuEndpoints = @('www.update.microsoft.com','fe2.update.microsoft.com','fe3.delivery.mp.microsoft.com',
                 'download.windowsupdate.com','sls.update.microsoft.com','ctldl.windowsupdate.com',
                 'settings-win.data.microsoft.com')
foreach ($ep in $wuEndpoints) {
    $t = Test-HttpsEndpoint -TargetHost $ep
    if (-not $t.DnsResolved) {
        Add-Row 'Connectivity' "HTTPS to $ep" 'Problem' 'DNS failed' $t.Error 'Check DNS resolution from this VM (correct resolvers, no split-DNS/black-hole for Microsoft domains).'
    } elseif ($t.TlsOk) {
        Add-Row 'Connectivity' "HTTPS to $ep" 'Healthy' "TLS OK, HTTP $($t.HttpStatus)" "Resolved to $($t.DnsAddresses)." ''
    } else {
        Add-Row 'Connectivity' "HTTPS to $ep" 'Problem' 'TLS/HTTPS failed' "Resolved to $($t.DnsAddresses). Error: $($t.Error)" 'A TCP:443 success with a failed TLS handshake here usually means TLS interception (a proxy/firewall doing SSL inspection with an untrusted cert), a missing/expired root CA on the client, or a blocked SNI - not a simple "no route" problem.'
    }
}

try {
    $proxyOut = & netsh winhttp show proxy 2>&1 | Out-String
    Add-Row 'Connectivity' 'WinHTTP proxy (machine-wide, used by wuauserv/SYSTEM)' 'Info' 'see Detail' $proxyOut.Trim() ''
} catch { }

# =======================================================================================
# SECTION: Azure Arc / Azure Update Manager footprint
# =======================================================================================
$himds = Get-Service -Name himds -ErrorAction SilentlyContinue
$arcAgentPresent = Test-Path -LiteralPath 'C:\Program Files\AzureConnectedMachineAgent\azcmagent.exe'
if ($himds -or $arcAgentPresent) {
    $himdsDetail = if ($himds) { "himds service: $($himds.Status)/$($himds.StartType)." } else { 'himds service not found, but azcmagent.exe is present.' }
    Add-Row 'Azure Arc / AUM' 'Azure Connected Machine Agent presence' 'Info' $(if ($himds) { "$($himds.Status)" } else { 'Agent files present, service not found' }) $himdsDetail 'This server is (or was) Arc-enabled. Confirm from the Azure/Arc portal side whether Patch Orchestration is set to "Azure Update Manager (Customer Managed Schedules)" vs "Image Default" - that setting lives in the Azure resource, not in the guest, and cannot be read from here.'

    $pluginsPath = 'C:\Packages\Plugins'
    if (Test-Path -LiteralPath $pluginsPath) {
        $patchExt = Get-ChildItem -LiteralPath $pluginsPath -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match 'SoftwareUpdateManagement|WindowsPatchExtension|LinuxPatchExtension' }
        if ($patchExt) {
            Add-Row 'Azure Arc / AUM' 'AUM patch-management GC extension' 'Warning' (($patchExt.Name) -join ', ') 'An Azure Update Manager guest-config/patch extension is installed. If a manual Windows Update session is run at the same time AUM triggers an assessment/install cycle, the two orchestrators can each try to control wuauserv/UsoSvc, and one stopping its session can surface to the user as "an update service was shutting down".' 'Do not run manual Windows Update installs and AUM-triggered patch runs against this VM at the same time. Check the Azure Update Manager / Arc portal for any currently scheduled or in-progress assessment/install operation on this machine before retrying manually.'
        } else {
            Add-Row 'Azure Arc / AUM' 'AUM patch-management GC extension' 'Info' 'Not found in C:\Packages\Plugins' 'No AUM patch-orchestration extension folder detected locally.' ''
        }
    }
} else {
    Add-Row 'Azure Arc / AUM' 'Azure Connected Machine Agent presence' 'Not Applicable' 'Not found' 'No Azure Arc agent detected on this machine - Azure Update Manager is not orchestrating patches here (or is managed as a native Azure VM instead of Arc; that distinction is an Azure-side property, not a guest artifact).' ''
}

# =======================================================================================
# SECTION: Windows Update History (exact failed KB / HResult)
# =======================================================================================
$hresultMeanings = @{
    '0x80070005' = 'Access denied (E_ACCESSDENIED)'
    '0x80070003' = 'Path not found'
    '0x8007000E' = 'Out of memory / insufficient resources'
    '0x80070422' = 'The service is disabled (ERROR_SERVICE_DISABLED) - a required Windows Update service was disabled at the time of the operation'
    '0x8007041D' = 'The service did not respond in a timely fashion (ERROR_SERVICE_REQUEST_TIMEOUT) - consistent with a service being stopped/killed mid-operation'
    '0x800705B4' = 'The operation timed out (ERROR_TIMEOUT)'
    '0x8024402C' = 'WU_E_PT_TIMEOUT - a communication timeout occurred with the Windows Update / delivery endpoint'
    '0x80240438' = 'WU_E_SETUP_SKIP_UPDATE - setup determined the update should be skipped'
    '0x8024200D' = 'WU_E_UH_INSTALLERHUNG - the installer hung and was terminated'
}
try {
    $session  = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $total    = $searcher.GetTotalHistoryCount()
    $take     = [Math]::Min($HistoryCount, $total)
    if ($take -gt 0) {
        $entries = $searcher.QueryHistory(0, $take)
        $failed = @()
        foreach ($e in $entries) {
            $rc = switch ($e.ResultCode) { 2 {'Succeeded'} 3 {'SucceededWithErrors'} 4 {'Failed'} 5 {'Aborted'} 1 {'InProgress'} default {'NotStarted'} }
            if ($e.ResultCode -in 4,5) {
                $hres = '0x{0:X8}' -f $e.HResult
                $meaning = if ($hresultMeanings.ContainsKey($hres)) { $hresultMeanings[$hres] } else { 'Unrecognised HResult - look up in the Microsoft Windows Update error-code reference before acting on it.' }
                $failed += "$($e.Date) [$rc] '$($e.Title)' HResult=$hres ($meaning)"
            }
        }
        if ($failed.Count -gt 0) {
            Add-Row 'Update History' "Failed/Aborted entries (last $take history records)" 'Problem' "$($failed.Count) failure(s)" ($failed -join ' || ') 'Match the most recent failure timestamp against the Update Services / WU Event Logs sections above for a direct cause-and-effect link.'
        } else {
            Add-Row 'Update History' "Failed/Aborted entries (last $take history records)" 'Healthy' '0 failures' 'No Failed/Aborted entries in the sampled history.' ''
        }
    } else {
        Add-Row 'Update History' 'Windows Update history' 'Unable to Check' 'n/a' 'GetTotalHistoryCount() returned 0 - no history recorded yet on this machine.' ''
    }
} catch {
    Add-Row 'Update History' 'Windows Update history (COM)' 'Unable to Check' 'n/a' "Microsoft.Update.Session COM query failed: $($_.Exception.Message)" ''
}

# =======================================================================================
# SECTION: VMware Tools (informational only)
# =======================================================================================
$vmToolsVer = Get-RegValue -Path 'HKLM:\SOFTWARE\VMware, Inc.\VMware Tools' -Name 'ProductVersion'
if (-not $vmToolsVer) { $vmToolsVer = Get-RegValue -Path 'HKLM:\SOFTWARE\WOW6432Node\VMware, Inc.\VMware Tools' -Name 'ProductVersion' }
Add-Row 'VMware Tools' 'Installed VMware Tools version (informational)' 'Info' $(if ($vmToolsVer) { $vmToolsVer } else { 'Unknown - registry value not found' }) 'VMware Tools does not own, start, stop, or otherwise manage wuauserv/BITS/CryptSvc/UsoSvc/WaaSMedicSvc. An outdated VMware Tools build is not a realistic direct cause of the "update service was shutting down" error. The one indirect overlap worth keeping in mind is host time sync (covered above) and, separately, VSS quiescing during a VM snapshot/backup taken while an update install is in progress, which CAN interrupt services mid-operation.' 'Treat VMware Tools upgrade as routine hygiene, not as a fix for this specific error. Separately, check whether any scheduled VM snapshot/backup job overlaps the patch window.'

# =======================================================================================
# Emit envelope
# =======================================================================================
$meta = [pscustomobject]@{
    Hostname   = $ComputerName
    OS         = $OSCaption
    Build      = $OSBuild
    LastBoot   = $(if ($LastBoot) { $LastBoot.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Unknown' })
    IsDomain   = $IsDomain
    CollectedAtUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
}
$envelope = [pscustomobject]@{ Schema = 'wu-diagnostics-1'; Meta = $meta; Rows = $rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<WU-DIAG-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-WU-DIAG-ENVELOPE>>>'
'@

$Payload = $PayloadTemplate.Replace('{{EVENTDAYS}}', "$EventDays").Replace('{{HISTORYCOUNT}}', "$HistoryCount")

# =====================================================================================
# 3. Admin-side helpers
# =====================================================================================
function Read-EnvelopeFromScriptOutput {
    param([string]$Output, [string]$StartMarker, [string]$EndMarker)
    if ([string]::IsNullOrEmpty($Output)) { return $null }
    $capture = $false
    $sb = New-Object System.Text.StringBuilder
    foreach ($line in ($Output -split "`r?`n")) {
        $t = $line.Trim()
        if ($t -eq $StartMarker) { $capture = $true; continue }
        if ($t -eq $EndMarker)   { break }
        if ($capture) { [void]$sb.Append($t) }
    }
    if ($sb.Length -eq 0) { return $null }
    try {
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($sb.ToString()))
        return $json | ConvertFrom-Json
    } catch { return $null }
}

# Classifies the common VIX/guest-ops Invoke-VMScript failures into an actionable
# explanation instead of surfacing the raw (often misleading) VMware error text.
function Get-FriendlyVixError {
    param([string]$Message)
    switch -Regex ($Message) {
        'Could not locate .?Powershell.? script interpreter' {
            return "VMware Tools authenticated the guest credential but could not confirm it has a full administrator token in-guest. Almost always one of: (1) the credential is a LOCAL (non-domain, non-built-in-Administrator) account and Windows' UAC remote token filtering silently downgrades it for network-style logons - fix on the guest with: reg add `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System`" /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f  (no reboot needed, then re-run); or (2) VMware Tools on this VM is out of date and isn't correctly registering the PowerShell interpreter path - upgrade VMware Tools on this VM. Isolate which one with: Invoke-VMScript -VM <vm> -ScriptText 'echo hi' -ScriptType Bat -GuestCredential `$cred - if that also fails, it's the credential/UAC issue; if that works, it's Tools."
        }
        'InvalidGuestLogin|[Ff]ailed to authenticate|authentication fail|incorrect user name or password|InvalidLogin|guest permissions|permission to perform this operation' {
            return "Guest authentication failed - VMware Tools rejected the supplied credential. Verify the username/password in the credential file/prompt and that the account exists and isn't locked out on this VM."
        }
        'Tools are not running|GuestOperationsUnavailable|VMware Tools is not|not installed|VIX' {
            return "VMware Tools guest operations are unavailable on this VM (Tools not running/installed, or too old to support guest operations). Check Tools status in vCenter."
        }
        'timed out|timeout' {
            return "Timed out waiting for VMware Tools / guest response. The guest may be under load, booting, or Tools may be unresponsive - retry, or increase -ToolsWaitSecs."
        }
        default { return $null }
    }
}

# Runs a (potentially large) PowerShell payload in the guest WITHOUT ever asking VMware
# Tools to locate a PowerShell interpreter for -ScriptType Powershell. Some VMs (older/
# stale VMware Tools builds) fail that lookup with a misleading "Could not locate
# 'Powershell' script interpreter... Probably you do not have enough permissions" error
# even though the credential is perfectly valid - Get-FriendlyVixError above explains it,
# but the real fix is to stop depending on that lookup at all. Instead: stage the payload
# as a .ps1 file in the guest (Copy-VMGuestFile - a plain VIX file transfer, no script
# interpreter involved at all) and execute it via "-ScriptType Bat" invoking
# "powershell.exe -File" directly - cmd.exe is always locatable, so this works on every
# VM regardless of Tools' PowerShell-path registration. Falls back to chunked staging
# (also Bat/certutil-only, no PowerShell dependency) if Copy-VMGuestFile is unavailable.
function Invoke-GuestPowerShellFile {
    param(
        $VM, $Server, [pscredential]$Credential,
        [string]$PayloadText, [int]$ToolsWaitSecs,
        [int]$ChunkSize = 1500
    )
    $tag  = [guid]::NewGuid().ToString('N').Substring(0, 12)
    $gB64 = "C:\Windows\Temp\wu$tag.b64"
    $gPs1 = "C:\Windows\Temp\wu$tag.ps1"
    $staged = $false

    try {
        if (Get-Command Copy-VMGuestFile -ErrorAction SilentlyContinue) {
            $local = Join-Path $env:TEMP "wu$tag.ps1"
            try {
                [System.IO.File]::WriteAllText($local, $PayloadText, (New-Object System.Text.UTF8Encoding($false)))
                Copy-VMGuestFile -Source $local -Destination $gPs1 -VM $VM -Server $Server `
                                 -GuestCredential $Credential -LocalToGuest -Force -ErrorAction Stop | Out-Null
                $staged = $true
            } catch {
                Write-Log "  Copy-VMGuestFile not usable ($(($_.Exception.Message -split "`n")[0].Trim())); using chunked Bat staging." 'INFO'
            } finally {
                Remove-Item -LiteralPath $local -Force -ErrorAction SilentlyContinue
            }
        }

        if (-not $staged) {
            $b64   = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($PayloadText))
            $total = [Math]::Ceiling($b64.Length / $ChunkSize)
            for ($i = 0; $i -lt $total; $i++) {
                $part     = $b64.Substring($i * $ChunkSize, [Math]::Min($ChunkSize, $b64.Length - ($i * $ChunkSize)))
                $redirect = if ($i -eq 0) { '>' } else { '>>' }
                $st = "cmd /c echo $part $redirect `"$gB64`""
                Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $st `
                                -ScriptType Bat -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop | Out-Null
            }
            # certutil is a stock Windows binary (no PowerShell involved) that decodes base64.
            $decode = "certutil -decode `"$gB64`" `"$gPs1`""
            Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $decode `
                            -ScriptType Bat -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop | Out-Null
        }

        $runner = "powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File `"$gPs1`""
        $res = Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $runner `
                               -ScriptType Bat -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop
        return [string]$res.ScriptOutput
    }
    finally {
        $cleanup = "del `"$gB64`" `"$gPs1`" 2>nul"
        try {
            Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $cleanup `
                            -ScriptType Bat -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}

function Get-VMInventory {
    param([object[]]$Servers)
    $inv = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Servers) {
        try { $vms = @(Get-VM -Server $s -ErrorAction Stop) }
        catch { Write-Log "Get-VM failed on vCenter '$($s.Name)': $($_.Exception.Message)" 'WARN'; $vms = @() }
        $inv.Add([pscustomobject]@{ Server = $s; VMs = $vms }) | Out-Null
    }
    return $inv
}

function Resolve-AndValidateVM {
    param([string]$Name, [object[]]$Inventory)
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $Inventory) {
        foreach ($v in $entry.VMs) {
            if ($v.Name -eq $Name) { $hits.Add([pscustomobject]@{ Server = $entry.Server; VM = $v }) | Out-Null }
        }
    }
    $serverList = ($Inventory | ForEach-Object { $_.Server.Name }) -join ', '
    if ($hits.Count -eq 0) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VM lookup'; ServerName = ''
            Reason = "VM '$Name' was not found in any connected vCenter ($serverList)."; VM = $null; Server = $null }
    }
    if ($hits.Count -gt 1) {
        $where = ($hits | ForEach-Object { "$($_.Server.Name):$($_.VM.Id)" }) -join '; '
        return [pscustomobject]@{ Ok = $false; Stage = 'VM lookup'; ServerName = $hits[0].Server.Name
            Reason = "VM name '$Name' is ambiguous - $($hits.Count) matching VMs ($where). Skipped for safety."; VM = $null; Server = $null }
    }
    $vm = $hits[0].VM; $srv = $hits[0].Server
    if ($vm.PowerState -ne 'PoweredOn') {
        return [pscustomobject]@{ Ok = $false; Stage = 'Power state'; ServerName = $srv.Name
            Reason = "VM is not powered on (PowerState=$($vm.PowerState))."; VM = $null; Server = $null }
    }
    $g = $vm.ExtensionData.Guest
    $toolsOk = ($g.ToolsRunningStatus -eq 'guestToolsRunning') -or ($g.ToolsStatus -in 'toolsOk', 'toolsOld')
    if (-not $toolsOk) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VMware Tools'; ServerName = $srv.Name
            Reason = "VMware Tools not installed/running (ToolsStatus=$($g.ToolsStatus), ToolsRunningStatus=$($g.ToolsRunningStatus))."; VM = $null; Server = $null }
    }
    $isWin = ($g.GuestFamily -eq 'windowsGuest') -or ($g.GuestId -match 'windows') -or ($vm.Guest.OSFullName -match 'Windows')
    if (-not $isWin) {
        return [pscustomobject]@{ Ok = $false; Stage = 'Guest OS'; ServerName = $srv.Name
            Reason = "Guest OS is not Windows (GuestFamily=$($g.GuestFamily), OS=$($vm.Guest.OSFullName))."; VM = $null; Server = $null }
    }
    return [pscustomobject]@{ Ok = $true; Stage = 'OK'; ServerName = $srv.Name; Reason = ''; VM = $vm; Server = $srv }
}

function New-CentralRow {
    param($vCenter, $VMName, $GuestHostname, $Section, $Check, $Status, $Value, $Detail, $NextStep)
    [PSCustomObject]([ordered]@{
        vCenter = $vCenter; VMName = $VMName; GuestHostname = $GuestHostname
        Section = $Section; Check = $Check; Status = $Status; Value = $Value
        Detail = $Detail; NextStep = $NextStep; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

function Get-StatusColor {
    param([string]$Status)
    switch ($Status) {
        'Healthy'                 { 'Green' }
        'Info'                    { 'Gray' }
        'Warning'                 { 'Yellow' }
        'Problem'                 { 'Red' }
        'Unable to Check'         { 'DarkYellow' }
        'Manual/External Required'{ 'Cyan' }
        'Not Applicable'          { 'DarkGray' }
        default                   { 'White' }
    }
}

function Write-VmDetail {
    param([string]$Header, $Rows)
    Write-Host ""
    Write-Host ("=== $Header ===") -ForegroundColor Cyan
    $bySection = $Rows | Group-Object Section
    foreach ($grp in $bySection) {
        Write-Host "-- $($grp.Name) --" -ForegroundColor White
        foreach ($row in $grp.Group) {
            $color = Get-StatusColor $row.Status
            Write-Host ("  [{0,-9}] {1}: {2}" -f $row.Status, $row.Check, $row.Value) -ForegroundColor $color
            if ($row.Detail)   { Write-Host "             $($row.Detail)"   -ForegroundColor DarkGray }
            if ($row.NextStep) { Write-Host "             NEXT: $($row.NextStep)" -ForegroundColor DarkCyan }
        }
    }
}

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows = New-Object System.Collections.Generic.List[object]
$envelopes   = New-Object System.Collections.Generic.List[object]
$summary = [ordered]@{ Total = 0; Checked = 0; WithProblems = 0; WithWarningsOnly = 0; Clean = 0; SkippedFailed = 0 }
$inventory = Get-VMInventory -Servers $connectedServers

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Windows Update diagnostics for VM '$vmName' ==="
    try {
        $val = Resolve-AndValidateVM -Name $vmName -Inventory $inventory
        if (-not $val.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -GuestHostname '' `
                -Section $val.Stage -Check 'VM validation' -Status 'Unable to Check' -Value '' -Detail $val.Reason -NextStep 'Fix the VM/credential/tools issue and re-run for this VM.'))
            Write-Log "SKIP '$vmName' [$($val.Stage)]: $($val.Reason)" 'WARN'
            Write-VmDetail -Header "$vmName  [SKIPPED - not processed]" -Rows @($centralRows[$centralRows.Count - 1])
            continue
        }
        $vm = $val.VM; $srv = $val.Server

        Write-Log "'$vmName': invoking guest diagnostics (ToolsWaitSecs=$ToolsWaitSecs)..."
        $scriptOutput = Invoke-GuestPowerShellFile -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $Payload -ToolsWaitSecs $ToolsWaitSecs

        $envelope = Read-EnvelopeFromScriptOutput -Output $scriptOutput -StartMarker '<<<WU-DIAG-ENVELOPE-B64>>>' -EndMarker '<<<END-WU-DIAG-ENVELOPE>>>'
        if (-not $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($scriptOutput) { ($scriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $vm.Guest.HostName `
                -Section 'Collection' -Check 'Guest payload result' -Status 'Unable to Check' -Value '' -Detail "No parseable result envelope. Output start: $snippet" -NextStep 'Re-run for this VM; if it repeats, check VMware Tools guest-ops health.'))
            Write-Log "'$vmName': no parseable envelope returned." 'ERROR'
            Write-VmDetail -Header "$vmName  [SKIPPED - not processed]" -Rows @($centralRows[$centralRows.Count - 1])
            continue
        }

        $envelopes.Add([pscustomobject]@{ VMName = $vmName; vCenter = $srv.Name; Envelope = $envelope }) | Out-Null
        $summary.Checked++
        $vmRowStart = $centralRows.Count
        foreach ($r in @($envelope.Rows)) {
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $envelope.Meta.Hostname `
                -Section $r.Section -Check $r.Check -Status $r.Status -Value $r.Value -Detail $r.Detail -NextStep $r.NextStep))
        }
        $vmRows = $centralRows.GetRange($vmRowStart, $centralRows.Count - $vmRowStart)
        $problemCount = @($vmRows | Where-Object { $_.Status -eq 'Problem' }).Count
        $warningCount = @($vmRows | Where-Object { $_.Status -eq 'Warning' }).Count
        if ($problemCount -gt 0) { $summary.WithProblems++ }
        elseif ($warningCount -gt 0) { $summary.WithWarningsOnly++ }
        else { $summary.Clean++ }

        Write-VmDetail -Header "$vmName @ $($srv.Name)  ($($envelope.Meta.OS), build $($envelope.Meta.Build))  -  $problemCount Problem, $warningCount Warning" -Rows $vmRows
        Write-Log "'$vmName': collection complete. Problems=$problemCount Warnings=$warningCount."
    }
    catch {
        $summary.SkippedFailed++
        $friendly = Get-FriendlyVixError -Message $_.Exception.Message
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        if ($friendly) { Write-Log "'$vmName': diagnosis - $friendly" 'WARN' }
        try {
            $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -GuestHostname '' `
                -Section 'Collection' -Check 'Processing' -Status 'Unable to Check' -Value '' -Detail $_.Exception.Message `
                -NextStep $(if ($friendly) { $friendly } else { 'Re-run for this VM after resolving the error.' })))
            Write-VmDetail -Header "$vmName  [SKIPPED - error]" -Rows @($centralRows[$centralRows.Count - 1])
        } catch { }
        continue
    }
}

# =====================================================================================
# 5. Consolidated report + console summary
# =====================================================================================
$centralRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV written: $CsvPath ($($centralRows.Count) rows across $($vmNames.Count) VM(s))."
$envelopes | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $JsonPath -Encoding UTF8
Write-Log "JSON written: $JsonPath"

Write-Host ""
Write-Host "================================================================================" -ForegroundColor DarkCyan
Write-Host " WINDOWS UPDATE DIAGNOSTICS SUMMARY" -ForegroundColor Green
Write-Host ("  Total VMs in list          : {0}" -f $summary.Total)
Write-Host ("  Successfully checked       : {0}" -f $summary.Checked)
Write-Host ("  VMs with >=1 Problem       : {0}" -f $summary.WithProblems) -ForegroundColor $(if ($summary.WithProblems) { 'Red' } else { 'Gray' })
Write-Host ("  VMs with Warning only      : {0}" -f $summary.WithWarningsOnly) -ForegroundColor $(if ($summary.WithWarningsOnly) { 'Yellow' } else { 'Gray' })
Write-Host ("  Clean (Healthy/Info only)  : {0}" -f $summary.Clean) -ForegroundColor Green
Write-Host ("  Skipped / failed to process: {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "================================================================================" -ForegroundColor DarkCyan
Write-Host "CSV : $CsvPath"
Write-Host "JSON: $JsonPath"
Write-Host "Log : $LogPath"
Write-Host "Review every 'Problem' row's NEXT step per VM before touching Invoke-WindowsUpdateRemediation.ps1." -ForegroundColor Gray

Write-Log "Diagnostics run complete."
