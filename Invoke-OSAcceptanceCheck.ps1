#Requires -Version 5.1
<#
.SYNOPSIS
    Centralised (PowerCLI) Windows OS Acceptance Checklist runner.

    Strictly READ-ONLY assessment, run against every VM in a list from a single
    admin workstation over VMware Tools Guest Operations. No WinRM / PsExec /
    SMB / RDP / guest network path is used to reach the target - only
    Invoke-VMScript guest ops through the already-connected vCenter session.

.DESCRIPTION
    Implements the "Guest OS" checks from windows-os-acceptance-checklist.html,
    categories 1-6 (Windows OS Health, Network & DNS, Roles & Features, Security
    & Hardening, Installed Agents, Patching & Vulnerability Mgmt), plus a
    host-side Category 7 (Configuration Consistency) pass computed by comparing
    the collected fleet data once every VM has been processed.

    Category 1 also includes "VMware Tools status" (checked once a VM is
    uniquely resolved in vCenter, before any guest command is attempted). This
    is the one item pulled from the VM object rather than from inside the
    guest - every other check depends on Invoke-VMScript, which itself depends
    on Tools, so this row is recorded even for a VM that gets skipped for
    being powered off or Tools-less, not just for VMs that pass validation.

    The in-guest payload contains ONLY Get-*/Test-*/Resolve-* cmdlets, registry
    reads, and read-only query switches (slmgr /dli, w32tm /query, auditpol
    /get, azcmagent show/check, net accounts). It never calls Set-*, New-*,
    Remove-*, Stop-*, Start-*, Restart-*, Disable-* or Add-*. It changes
    nothing on the guest.

    Items the checklist marks type "Team" (security baseline/GPO compliance,
    Azure Update Manager enrollment, Qualys vulnerability posture, fleet-wide
    patch-source consistency) cannot be assessed from inside the guest OS and
    are recorded as "Manual Verification Required" with a pointer to the
    console/team that owns the answer - they are not silently skipped.

    Per VM, per checklist item, the CSV records one of:
        Compliant
        Non-Compliant
        Manual Verification Required
        Unable to Check
    together with the detected value and the expected value, so the CSV is
    auditable against the original checklist.

.PARAMETER VMListPath
    Text file of VM names (default C:\temp\vmlist.txt). Lines starting with '#'
    are ignored.

.PARAMETER CredentialPath
    Export-Clixml PSCredential for the guest admin (default C:\temp\wincred.xml).
    Loaded in-memory only, never written back out.

.PARAMETER OutputPath
    Folder on the admin machine for the CSV + log.

.PARAMETER DnsTestNames
    FQDNs to resolve from inside each guest as the DNS resolution test (item
    2.3). If omitted, the script falls back to the guest's own AD domain name
    when the VM is domain-joined; workgroup VMs with no names supplied are
    marked Manual Verification Required for that item.

.PARAMETER RequiredConnectivity
    Array of "host:port" strings to Test-NetConnection from inside each guest
    (item 2.6 - required connectivity paths: AD, monitoring, backup, PAM jump,
    patch servers, etc). If omitted, that item is recorded as Manual
    Verification Required (no target list supplied).

.PARAMETER ExpectedAgentServices
    Extra Windows service names, beyond the five named agents the script
    always checks (Qualys, Binalyze, Splunk UF, Defender for Endpoint/Sense,
    Azure Arc), that your org mandates on every server (item 5.6).

.PARAMETER EventLookbackDays
    Lookback window in days for the Critical/Error event log review and the
    unexpected-shutdown (6008/41) review. Default 7.

.PARAMETER DiskFreePercentThreshold
    Volumes below this free-space percentage are flagged Non-Compliant.
    Default 15.

.PARAMETER UptimeWarnDays
    Uptime beyond this many days is flagged Non-Compliant (missed patch
    cycles). Default 60.

.PARAMETER PatchAgeWarnDays
    Latest installed update older than this many days is flagged
    Non-Compliant. Default 60.

.PARAMETER MinPasswordLength / LockoutThreshold
    Local password-policy baseline compared against `net accounts` output.
    Defaults 14 / 5.

.PARAMETER AllowedRdpSources
    Array of expected RDP-source CIDRs/hosts (e.g. your PAM jump box). If
    supplied, the RDP firewall-scoping check is auto-evaluated; if omitted it
    is recorded as Manual Verification Required.

.PARAMETER SkipRolesInventory
    Skip the Get-WindowsFeature inventory (category 3) - useful for a fleet of
    non-Server (client) guests where the cmdlet is not available anyway.

.PARAMETER ToolsWaitSecs
    Invoke-VMScript VMware Tools wait, seconds (default 180).

.EXAMPLE
    .\Invoke-OSAcceptanceCheck.ps1
    .\Invoke-OSAcceptanceCheck.ps1 -VMListPath C:\temp\vmlist.txt -DnsTestNames 'contoso.local' `
        -RequiredConnectivity '10.0.0.10:389','10.0.0.20:443' -OutputPath D:\Reports

.NOTES
    Prerequisites: PowerCLI imported; already Connect-VIServer'd to the target
    vCenter(s); the vCenter account holds the three "Guest Operation ..."
    privileges. This script issues no vSphere write of any kind and performs
    no in-guest remediation - it is an assessment tool only.
#>

[CmdletBinding()]
param(
    [string]   $VMListPath              = 'C:\temp\vmlist.txt',
    [string]   $CredentialPath          = 'C:\temp\wincred.xml',
    [string]   $OutputPath              = (Join-Path $PSScriptRoot 'OSAcceptance_Reports'),

    [string[]] $DnsTestNames            = @(),
    [string[]] $RequiredConnectivity    = @(),
    [string[]] $ExpectedAgentServices   = @(),
    [string[]] $AllowedRdpSources       = @(),

    [ValidateRange(1, 90)]   [int] $EventLookbackDays         = 7,
    [ValidateRange(1, 90)]   [int] $DiskFreePercentThreshold  = 15,
    [ValidateRange(1, 365)]  [int] $UptimeWarnDays             = 60,
    [ValidateRange(1, 365)]  [int] $PatchAgeWarnDays           = 60,
    [ValidateRange(4, 32)]   [int] $MinPasswordLength          = 14,
    [ValidateRange(1, 50)]   [int] $LockoutThreshold           = 5,

    [switch]   $SkipRolesInventory,

    [ValidateRange(30, 3600)] [int] $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'

Write-Host "Invoke-OSAcceptanceCheck.ps1 - READ-ONLY Windows OS Acceptance assessment  [build 2026-09-13a]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray

# =====================================================================================
# 1. Pre-flight
# =====================================================================================
if (-not (Get-Command Invoke-VMScript -ErrorAction SilentlyContinue)) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop }
    catch { throw "VMware PowerCLI (VMware.VimAutomation.Core) is not available. Install PowerCLI and retry." }
}
$connectedServers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($connectedServers.Count -eq 0) {
    throw "No connected vCenter session. Run Connect-VIServer <vcenter> first."
}
Write-Host ("Connected vCenter(s): {0}" -f (($connectedServers | ForEach-Object { $_.Name }) -join ', ')) -ForegroundColor Gray

# Invoke-VMScript fetches script output via DownloadFileFromGuest, which goes straight to the
# ESXi host by its registered name - not through vCenter. If this workstation can't resolve
# that name, every single VM on that host fails identically with a generic "An error occurred
# while sending the request." Check that up front instead of discovering it VM-by-VM.
$unresolvedHosts = New-Object System.Collections.Generic.List[pscustomobject]
foreach ($srv in $connectedServers) {
    foreach ($esxHost in (Get-VMHost -Server $srv -ErrorAction SilentlyContinue)) {
        try { [void][System.Net.Dns]::GetHostAddresses($esxHost.Name) }
        catch {
            $mgmtIp = ($esxHost | Get-VMHostNetworkAdapter -VMKernel -ErrorAction SilentlyContinue |
                       Where-Object ManagementTrafficEnabled | Select-Object -First 1 -ExpandProperty IP)
            $unresolvedHosts.Add([pscustomobject]@{ Name = $esxHost.Name; IP = $mgmtIp }) | Out-Null
        }
    }
}
if ($unresolvedHosts.Count -gt 0) {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $hostsFile = "$env:WINDIR\System32\drivers\etc\hosts"
    Write-Host ""
    Write-Host "WARNING: $($unresolvedHosts.Count) ESXi host(s) do not resolve via DNS from this machine - every VM on them will fail at the guest-ops probe stage:" -ForegroundColor Yellow
    $unresolvedHosts | ForEach-Object { Write-Host ("  {0,-16} {1}" -f $_.IP, $_.Name) -ForegroundColor Yellow }
    if ($isAdmin) {
        $existing = Get-Content -LiteralPath $hostsFile -ErrorAction SilentlyContinue
        $added = 0
        foreach ($h in $unresolvedHosts) {
            if ($h.IP -and -not ($existing -match [regex]::Escape($h.Name))) {
                Add-Content -LiteralPath $hostsFile -Value ("{0}`t{1}" -f $h.IP, $h.Name)
                $added++
            }
        }
        Write-Host "Added $added entr$(if ($added -eq 1) {'y'} else {'ies'}) to $hostsFile - re-run this script now." -ForegroundColor Green
        exit 1
    } else {
        Write-Host "Re-run this PowerShell session as Administrator so it can add these to $hostsFile automatically, or add them yourself, then re-run." -ForegroundColor Yellow
        exit 1
    }
}

if (-not (Test-Path -LiteralPath $VMListPath)) { throw "VM list not found: $VMListPath" }
$vmNames = @(Get-Content -LiteralPath $VMListPath |
             ForEach-Object { $_.Trim() } |
             Where-Object { $_ -and -not $_.StartsWith('#') } |
             Select-Object -Unique)
if ($vmNames.Count -eq 0) { throw "VM list '$VMListPath' contained no usable VM names." }

if (-not (Test-Path -LiteralPath $CredentialPath)) { throw "Guest credential file not found: $CredentialPath" }
try {
    $GuestCredential = Import-Clixml -LiteralPath $CredentialPath
} catch {
    throw ("Failed to import guest credential from '$CredentialPath'. Export-Clixml credentials are DPAPI-protected " +
           "and only readable by the same Windows account on the same machine that created them. Recreate with: " +
           "Get-Credential | Export-Clixml '$CredentialPath'. Underlying: $($_.Exception.Message)")
}
if (-not ($GuestCredential -is [pscredential])) { throw "'$CredentialPath' did not deserialise to a PSCredential." }

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$CsvPath  = Join-Path $OutputPath "OSAcceptance_$RunStamp.csv"
$LogPath  = Join-Path $OutputPath "OSAcceptance_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "OS Acceptance run started. VMs=$($vmNames.Count)."

function ConvertTo-PSArrayLiteral {
    param([string[]]$Items)
    if (-not $Items -or $Items.Count -eq 0) { return '@()' }
    $esc = $Items | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }
    return '@(' + ($esc -join ',') + ')'
}

# =====================================================================================
# 2. In-guest READ-ONLY payload
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'

$EventLookbackDays      = {{EVENT_LOOKBACK_DAYS}}
$DiskFreePctThreshold   = {{DISK_FREE_PCT}}
$UptimeWarnDays         = {{UPTIME_WARN_DAYS}}
$PatchAgeWarnDays       = {{PATCH_AGE_DAYS}}
$MinPasswordLength      = {{MIN_PW_LENGTH}}
$LockoutThreshold       = {{LOCKOUT_THRESHOLD}}
$SkipRolesInventory     = {{SKIP_ROLES}}
$DnsTestNames           = {{DNS_TEST_NAMES}}
$RequiredConnectivity   = {{REQUIRED_CONNECTIVITY}}
$ExpectedAgentServices  = {{EXPECTED_AGENT_SERVICES}}
$AllowedRdpSources      = {{ALLOWED_RDP_SOURCES}}

$Rows = New-Object System.Collections.Generic.List[object]
function Add-Row {
    param(
        [string]$Category, [string]$Item, [string]$NormStatus,
        [string]$Detected = '', [string]$Expected = '', [string]$Detail = ''
    )
    $Rows.Add([PSCustomObject]@{
        Category = $Category; Item = $Item; Status = $NormStatus
        DetectedValue = $Detected; ExpectedValue = $Expected; Detail = $Detail
    }) | Out-Null
}
function Get-RegVal {
    param([string]$Path, [string]$Name)
    try {
        $ip = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return [pscustomobject]@{ Found = $true; Value = $ip.$Name }
    } catch { return [pscustomobject]@{ Found = $false; Value = $null } }
}
function Test-AgentService {
    param([string]$Name, [string]$Label, [string]$Category = '5. Installed Agents')
    try {
        $svc = Get-Service -Name $Name -ErrorAction Stop
        if ($svc.Status -eq 'Running') {
            Add-Row -Category $Category -Item $Label -NormStatus 'Compliant' `
                -Detected "Service $Name Status=Running" -Expected 'Installed and Running'
        } else {
            Add-Row -Category $Category -Item $Label -NormStatus 'Non-Compliant' `
                -Detected "Service $Name Status=$($svc.Status)" -Expected 'Installed and Running'
        }
    } catch {
        Add-Row -Category $Category -Item $Label -NormStatus 'Non-Compliant' `
            -Detected "Service $Name not found" -Expected 'Installed and Running'
    }
}
function Test-PendingRebootFlags {
    $hits = New-Object System.Collections.Generic.List[string]
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $hits.Add('CBS RebootPending') }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $hits.Add('WU RebootRequired') }
    $pfro = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' 'PendingFileRenameOperations'
    if ($pfro.Found -and $pfro.Value) { $hits.Add('PendingFileRenameOperations') }
    return $hits
}

$ComputerName = $env:COMPUTERNAME
try {
    $IPAddresses = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
        Where-Object { $_.InterfaceAlias -notmatch 'Loopback' -and $_.IPAddress -notmatch '^169\.254\.' }).IPAddress -join ', '
} catch { $IPAddresses = $null }
if ([string]::IsNullOrWhiteSpace($IPAddresses)) {
    try { $IPAddresses = ([System.Net.Dns]::GetHostAddresses($ComputerName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString } catch { $IPAddresses = 'Unknown' }
}
try { $OSCaption = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Caption } catch { $OSCaption = 'Unknown' }
try { $CsInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop } catch { $CsInfo = $null }

# Fleet data, consumed on the admin side for the Category 7 consistency pass.
$Fleet = [ordered]@{
    DnsServers = ''; TimeSource = ''; AdminMembers = ''; FirewallAllEnabled = $null; AgentSet = ''
}

$fatal = $null
try {

    # ============================ 1. WINDOWS OS HEALTH ============================
    $Cat = '1. Windows OS Health'

    if ($CsInfo) {
        $domTxt = if ($CsInfo.PartOfDomain) { "Domain-joined: $($CsInfo.Domain)" } else { "Workgroup: $($CsInfo.Workgroup)" }
        Add-Row -Category $Cat -Item 'Domain vs. workgroup membership' -NormStatus 'Manual Verification Required' `
            -Detected $domTxt -Expected 'Matches intended design for this VM' `
            -Detail 'Confirm domain/workgroup membership matches the build design document.'
    } else {
        Add-Row -Category $Cat -Item 'Domain vs. workgroup membership' -NormStatus 'Unable to Check' `
            -Detected 'Win32_ComputerSystem query failed' -Expected 'Matches intended design'
    }

    try {
        $ci = Get-ComputerInfo -Property OsName, OsVersion, OsBuildNumber -ErrorAction Stop
        Add-Row -Category $Cat -Item 'OS edition/version/build' -NormStatus 'Manual Verification Required' `
            -Detected ("{0} | Version={1} Build={2}" -f $ci.OsName, $ci.OsVersion, $ci.OsBuildNumber) `
            -Expected 'Supported, currently-serviced Windows Server version/build' `
            -Detail 'Cross-check the build number against the Microsoft servicing/lifecycle page.'
    } catch {
        Add-Row -Category $Cat -Item 'OS edition/version/build' -NormStatus 'Unable to Check' `
            -Detected "Get-ComputerInfo failed: $($_.Exception.Message)" -Expected 'Supported, currently-serviced build'
    }

    try {
        $dli = (& cscript.exe //nologo "$env:WINDIR\System32\slmgr.vbs" /dli) -join ' '
        if ($dli -match 'License Status:\s*Licensed') {
            Add-Row -Category $Cat -Item 'Activation/licensing status' -NormStatus 'Compliant' -Detected $dli -Expected 'Licensed'
        } else {
            Add-Row -Category $Cat -Item 'Activation/licensing status' -NormStatus 'Non-Compliant' -Detected $dli -Expected 'Licensed'
        }
    } catch {
        Add-Row -Category $Cat -Item 'Activation/licensing status' -NormStatus 'Unable to Check' `
            -Detected "slmgr /dli failed: $($_.Exception.Message)" -Expected 'Licensed'
    }

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $lastBoot = $os.LastBootUpTime
        $upDays = [math]::Round(((Get-Date) - $lastBoot).TotalDays, 1)
        $status = if ($upDays -gt $UptimeWarnDays) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Uptime / last reboot' -NormStatus $status `
            -Detected "LastBootUpTime=$lastBoot ($upDays days uptime)" -Expected "Uptime <= $UptimeWarnDays days, consistent with patch cadence"
    } catch {
        Add-Row -Category $Cat -Item 'Uptime / last reboot' -NormStatus 'Unable to Check' `
            -Detected "Win32_OperatingSystem query failed: $($_.Exception.Message)" -Expected "Uptime <= $UptimeWarnDays days"
    }

    try {
        $cpu = (Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 3 -ErrorAction Stop).CounterSamples.CookedValue
        $avgCpu = [math]::Round(($cpu | Measure-Object -Average).Average, 1)
        $mem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $freeMemPct = [math]::Round(($mem.FreePhysicalMemory / $mem.TotalVisibleMemorySize) * 100, 1)
        $status = if ($avgCpu -gt 85 -or $freeMemPct -lt 10) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'CPU/memory utilization baseline' -NormStatus $status `
            -Detected "AvgCPU=$avgCpu% FreeMemory=$freeMemPct%" -Expected 'CPU <= 85% sustained, free memory >= 10%'
    } catch {
        Add-Row -Category $Cat -Item 'CPU/memory utilization baseline' -NormStatus 'Unable to Check' `
            -Detected "Get-Counter failed: $($_.Exception.Message)" -Expected 'CPU <= 85%, free memory >= 10%'
    }

    try {
        $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop
        $low = New-Object System.Collections.Generic.List[string]
        $summary = New-Object System.Collections.Generic.List[string]
        foreach ($d in $disks) {
            if ($d.Size -gt 0) {
                $pct = [math]::Round(($d.FreeSpace / $d.Size) * 100, 1)
                $summary.Add("$($d.DeviceID) $pct% free")
                if ($pct -lt $DiskFreePctThreshold) { $low.Add("$($d.DeviceID) $pct%") }
            }
        }
        $status = if ($low.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Disk usage / low-space conditions' -NormStatus $status `
            -Detected ($summary -join '; ') -Expected "All volumes >= $DiskFreePercentThreshold`% free" `
            -Detail $(if ($low.Count -gt 0) { "Below threshold: $($low -join ', ')" } else { '' })
    } catch {
        Add-Row -Category $Cat -Item 'Disk usage / low-space conditions' -NormStatus 'Unable to Check' `
            -Detected "Win32_LogicalDisk query failed: $($_.Exception.Message)" -Expected "All volumes >= $DiskFreePercentThreshold`% free"
    }

    try {
        $stopped = Get-Service -ErrorAction Stop | Where-Object { $_.StartType -eq 'Automatic' -and $_.Status -ne 'Running' }
        $status = if (@($stopped).Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Critical Windows services status' -NormStatus $status `
            -Detected $(if (@($stopped).Count -gt 0) { ($stopped | ForEach-Object { $_.Name }) -join ', ' } else { 'All Automatic services running' }) `
            -Expected 'No Automatic-start service is stopped'
    } catch {
        Add-Row -Category $Cat -Item 'Critical Windows services status' -NormStatus 'Unable to Check' `
            -Detected "Get-Service failed: $($_.Exception.Message)" -Expected 'No Automatic-start service is stopped'
    }

    try {
        $evts = @(Get-WinEvent -FilterHashtable @{ LogName = 'System', 'Application'; Level = 1, 2; StartTime = (Get-Date).AddDays(-$EventLookbackDays) } -ErrorAction Stop)
        if ($evts.Count -eq 0) {
            Add-Row -Category $Cat -Item 'Event Viewer critical/error review' -NormStatus 'Compliant' `
                -Detected "0 Critical/Error events in last $EventLookbackDays days" -Expected 'No unexplained recurring Critical/Error events'
        } else {
            $top = $evts | Group-Object Id | Sort-Object Count -Descending | Select-Object -First 8 |
                   ForEach-Object { "Id=$($_.Name) x$($_.Count)" }
            Add-Row -Category $Cat -Item 'Event Viewer critical/error review' -NormStatus 'Manual Verification Required' `
                -Detected ("$($evts.Count) events in last $EventLookbackDays days. Top: " + ($top -join '; ')) `
                -Expected 'No unexplained recurring Critical/Error events' `
                -Detail 'Review each recurring Event ID for a known/benign cause before acceptance.'
        }
    } catch {
        if ($_.Exception.Message -match 'No events were found') {
            Add-Row -Category $Cat -Item 'Event Viewer critical/error review' -NormStatus 'Compliant' `
                -Detected "0 Critical/Error events in last $EventLookbackDays days" -Expected 'No unexplained recurring Critical/Error events'
        } else {
            Add-Row -Category $Cat -Item 'Event Viewer critical/error review' -NormStatus 'Unable to Check' `
                -Detected "Get-WinEvent failed: $($_.Exception.Message)" -Expected 'No unexplained recurring Critical/Error events'
        }
    }

    try {
        $shut = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 6008, 41; StartTime = (Get-Date).AddDays(-$EventLookbackDays) } -ErrorAction Stop)
        $status = if ($shut.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Unexpected shutdown/reboot history' -NormStatus $status `
            -Detected "$($shut.Count) event(s) (6008/41) in last $EventLookbackDays days" -Expected 'No unplanned shutdown/Kernel-Power events'
    } catch {
        if ($_.Exception.Message -match 'No events were found') {
            Add-Row -Category $Cat -Item 'Unexpected shutdown/reboot history' -NormStatus 'Compliant' `
                -Detected "0 events (6008/41) in last $EventLookbackDays days" -Expected 'No unplanned shutdown/Kernel-Power events'
        } else {
            Add-Row -Category $Cat -Item 'Unexpected shutdown/reboot history' -NormStatus 'Unable to Check' `
                -Detected "Get-WinEvent failed: $($_.Exception.Message)" -Expected 'No unplanned shutdown/Kernel-Power events'
        }
    }

    try {
        $bad = @(Get-CimInstance -ClassName Win32_PnPEntity -ErrorAction Stop | Where-Object { $_.ConfigManagerErrorCode -ne 0 })
        $status = if ($bad.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Device/driver issues' -NormStatus $status `
            -Detected $(if ($bad.Count -gt 0) { ($bad | ForEach-Object { "$($_.Name) (code $($_.ConfigManagerErrorCode))" }) -join '; ' } else { 'No devices in error state' }) `
            -Expected 'No devices report a non-zero ConfigManagerErrorCode'
    } catch {
        Add-Row -Category $Cat -Item 'Device/driver issues' -NormStatus 'Unable to Check' `
            -Detected "Win32_PnPEntity query failed: $($_.Exception.Message)" -Expected 'No devices in error state'
    }

    $timeSourceForFleet = ''
    try {
        $w32 = (& w32tm.exe /query /status) -join "`n"
        $srcMatch = [regex]::Match($w32, 'Source:\s*(.+)')
        $offMatch = [regex]::Match($w32, 'Phase Offset:\s*([\-0-9\.]+)s')
        $src = if ($srcMatch.Success) { $srcMatch.Groups[1].Value.Trim() } else { 'Unknown' }
        $timeSourceForFleet = $src
        if ($offMatch.Success) {
            $offsetSec = [double]$offMatch.Groups[1].Value
            $status = if ([math]::Abs($offsetSec) -gt 300) { 'Non-Compliant' } else { 'Compliant' }
            Add-Row -Category $Cat -Item 'In-guest time synchronization' -NormStatus $status `
                -Detected "Source=$src PhaseOffset=${offsetSec}s" -Expected 'Correct source, offset within seconds (< 5 min)'
        } else {
            Add-Row -Category $Cat -Item 'In-guest time synchronization' -NormStatus 'Manual Verification Required' `
                -Detected "Source=$src (offset not parsed)" -Expected 'Correct source, small offset, no errors'
        }
    } catch {
        Add-Row -Category $Cat -Item 'In-guest time synchronization' -NormStatus 'Unable to Check' `
            -Detected "w32tm /query /status failed: $($_.Exception.Message)" -Expected 'Correct source, small offset, no errors'
    }

    $pending = Test-PendingRebootFlags
    $pendingStatus = if ($pending.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
    Add-Row -Category $Cat -Item 'Pending reboot status' -NormStatus $pendingStatus `
        -Detected $(if ($pending.Count -gt 0) { $pending -join ', ' } else { 'No pending-reboot flags set' }) `
        -Expected 'No pending reboot flag outstanding at acceptance time'

    # ============================ 2. NETWORK & DNS (IN-GUEST) ============================
    $Cat = '2. Network & DNS (in-guest)'

    try {
        $ipc = Get-NetIPConfiguration -ErrorAction Stop | Where-Object { $_.IPv4Address }
        $summary = $ipc | ForEach-Object {
            $gw = if ($_.IPv4DefaultGateway) { $_.IPv4DefaultGateway.NextHop } else { '(none)' }
            "$($_.InterfaceAlias): $($_.IPv4Address.IPAddress)/$($_.IPv4Address.PrefixLength) GW=$gw"
        }
        Add-Row -Category $Cat -Item 'IP/subnet/gateway configuration' -NormStatus 'Manual Verification Required' `
            -Detected ($summary -join '; ') -Expected 'Matches the IP allocation sheet/IPAM exactly' `
            -Detail 'Compare each interface against IPAM/the network design document.'
    } catch {
        Add-Row -Category $Cat -Item 'IP/subnet/gateway configuration' -NormStatus 'Unable to Check' `
            -Detected "Get-NetIPConfiguration failed: $($_.Exception.Message)" -Expected 'Matches IP allocation sheet/IPAM'
    }

    $dnsServersForFleet = ''
    try {
        $dnsc = Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction Stop | Where-Object { $_.ServerAddresses }
        $allServers = @($dnsc.ServerAddresses) | Select-Object -Unique | Sort-Object
        $dnsServersForFleet = $allServers -join ','
        Add-Row -Category $Cat -Item 'DNS server configuration' -NormStatus 'Manual Verification Required' `
            -Detected ($allServers -join ', ') -Expected 'Matches the approved DNS server list for the site/domain' `
            -Detail 'Compared fleet-wide in Category 7.'
    } catch {
        Add-Row -Category $Cat -Item 'DNS server configuration' -NormStatus 'Unable to Check' `
            -Detected "Get-DnsClientServerAddress failed: $($_.Exception.Message)" -Expected 'Matches approved DNS server list'
    }

    $testNames = @($DnsTestNames)
    if ($testNames.Count -eq 0 -and $CsInfo -and $CsInfo.PartOfDomain -and $CsInfo.Domain) { $testNames = @($CsInfo.Domain) }
    if ($testNames.Count -eq 0) {
        Add-Row -Category $Cat -Item 'DNS resolution test' -NormStatus 'Manual Verification Required' `
            -Detected 'No DNS test target supplied and VM is not domain-joined' -Expected 'Successful resolution of required internal/external names' `
            -Detail 'Re-run with -DnsTestNames to automate this check.'
    } else {
        $fail = New-Object System.Collections.Generic.List[string]
        $ok   = New-Object System.Collections.Generic.List[string]
        foreach ($n in $testNames) {
            try { Resolve-DnsName -Name $n -ErrorAction Stop | Out-Null; $ok.Add($n) }
            catch { $fail.Add("$n ($($_.Exception.Message))") }
        }
        $status = if ($fail.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'DNS resolution test' -NormStatus $status `
            -Detected "Resolved: $($ok -join ', ') | Failed: $($fail -join '; ')" -Expected 'Successful resolution, no timeouts'
    }

    try {
        $dup = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Id = 4198 } -MaxEvents 20 -ErrorAction Stop)
        $status = if ($dup.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Duplicate/incorrect IP detection' -NormStatus $status `
            -Detected "$($dup.Count) duplicate-IP event(s) (4198) found" -Expected 'No duplicate IP address events logged'
    } catch {
        if ($_.Exception.Message -match 'No events were found') {
            Add-Row -Category $Cat -Item 'Duplicate/incorrect IP detection' -NormStatus 'Compliant' `
                -Detected '0 duplicate-IP events (4198)' -Expected 'No duplicate IP address events logged'
        } else {
            Add-Row -Category $Cat -Item 'Duplicate/incorrect IP detection' -NormStatus 'Unable to Check' `
                -Detected "Get-WinEvent failed: $($_.Exception.Message)" -Expected 'No duplicate IP address events logged'
        }
    }

    try {
        $nics = Get-NetAdapter -ErrorAction Stop
        $down = @($nics | Where-Object { $_.Status -eq 'Disconnected' })
        $status = if ($down.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'NIC link status and speed' -NormStatus $status `
            -Detected (($nics | ForEach-Object { "$($_.Name):$($_.Status)@$($_.LinkSpeed)" }) -join '; ') `
            -Expected 'Status Up at expected link speed'
    } catch {
        Add-Row -Category $Cat -Item 'NIC link status and speed' -NormStatus 'Unable to Check' `
            -Detected "Get-NetAdapter failed: $($_.Exception.Message)" -Expected 'Status Up at expected link speed'
    }

    if (@($RequiredConnectivity).Count -eq 0) {
        Add-Row -Category $Cat -Item 'Required connectivity paths' -NormStatus 'Manual Verification Required' `
            -Detected 'No target list supplied' -Expected 'All required paths per architecture diagram succeed' `
            -Detail 'Re-run with -RequiredConnectivity host:port,host:port,... to automate this check.'
    } else {
        $fail = New-Object System.Collections.Generic.List[string]
        $ok   = New-Object System.Collections.Generic.List[string]
        foreach ($target in $RequiredConnectivity) {
            $parts = $target -split ':', 2
            $h = $parts[0]; $p = if ($parts.Count -gt 1) { [int]$parts[1] } else { 443 }
            try {
                $t = Test-NetConnection -ComputerName $h -Port $p -WarningAction SilentlyContinue -ErrorAction Stop
                if ($t.TcpTestSucceeded) { $ok.Add($target) } else { $fail.Add($target) }
            } catch { $fail.Add("$target (error)") }
        }
        $status = if ($fail.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Required connectivity paths' -NormStatus $status `
            -Detected "OK: $($ok -join ', ') | Failed: $($fail -join ', ')" -Expected 'All required paths succeed'
    }

    # ============================ 3. WINDOWS ROLES & FEATURES ============================
    $Cat = '3. Windows Roles & Features'

    if ($SkipRolesInventory) {
        Add-Row -Category $Cat -Item 'Installed roles/features inventory' -NormStatus 'Not Applicable' `
            -Detected 'Skipped by operator (-SkipRolesInventory)' -Expected 'Documented list matches VM purpose'
        Add-Row -Category $Cat -Item 'Unexpected/unnecessary roles' -NormStatus 'Not Applicable' `
            -Detected 'Skipped by operator' -Expected 'Only required roles present'
        Add-Row -Category $Cat -Item 'Infrastructure role presence (IIS/DNS/DHCP/File Server)' -NormStatus 'Not Applicable' `
            -Detected 'Skipped by operator' -Expected 'Presence is deliberate and documented'
    } else {
        try {
            $feat = Get-WindowsFeature -ErrorAction Stop | Where-Object { $_.Installed }
            $names = ($feat | ForEach-Object { $_.Name }) -join ', '
            Add-Row -Category $Cat -Item 'Installed roles/features inventory' -NormStatus 'Manual Verification Required' `
                -Detected $names -Expected 'Documented list matches the VM intended function' `
                -Detail 'Compare against the VM build/purpose document.'
            Add-Row -Category $Cat -Item 'Unexpected/unnecessary roles' -NormStatus 'Manual Verification Required' `
                -Detected '(see inventory above)' -Expected 'Only roles required for the VM function are present'

            $infra = Get-WindowsFeature -Name DNS, DHCP, Web-Server, FS-FileServer -ErrorAction Stop | Where-Object { $_.Installed }
            if (@($infra).Count -eq 0) {
                Add-Row -Category $Cat -Item 'Infrastructure role presence (IIS/DNS/DHCP/File Server)' -NormStatus 'Compliant' `
                    -Detected 'None of DNS/DHCP/Web-Server/FS-FileServer installed' -Expected 'Presence is deliberate and coordinated with infra team'
            } else {
                Add-Row -Category $Cat -Item 'Infrastructure role presence (IIS/DNS/DHCP/File Server)' -NormStatus 'Manual Verification Required' `
                    -Detected (($infra | ForEach-Object { $_.Name }) -join ', ') -Expected 'Presence is deliberate and coordinated with infra team'
            }
        } catch {
            Add-Row -Category $Cat -Item 'Installed roles/features inventory' -NormStatus 'Unable to Check' `
                -Detected "Get-WindowsFeature failed (client OS or ServerManager unavailable): $($_.Exception.Message)" -Expected 'Documented list matches VM purpose'
            Add-Row -Category $Cat -Item 'Unexpected/unnecessary roles' -NormStatus 'Unable to Check' `
                -Detected 'Get-WindowsFeature unavailable' -Expected 'Only required roles present'
            Add-Row -Category $Cat -Item 'Infrastructure role presence (IIS/DNS/DHCP/File Server)' -NormStatus 'Unable to Check' `
                -Detected 'Get-WindowsFeature unavailable' -Expected 'Presence is deliberate and documented'
        }
    }

    # ============================ 4. SECURITY & HARDENING ============================
    $Cat = '4. Security & Hardening'

    $adminMembersForFleet = ''
    try {
        $admins = Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop
        $names = ($admins | ForEach-Object { $_.Name }) | Sort-Object
        $adminMembersForFleet = $names -join ','
        Add-Row -Category $Cat -Item 'Local users & Administrators group membership' -NormStatus 'Manual Verification Required' `
            -Detected ($names -join ', ') -Expected 'Limited to approved break-glass/service accounts per policy' `
            -Detail 'Compared fleet-wide in Category 7.'
    } catch {
        Add-Row -Category $Cat -Item 'Local users & Administrators group membership' -NormStatus 'Unable to Check' `
            -Detected "Get-LocalGroupMember failed: $($_.Exception.Message)" -Expected 'Limited to approved accounts'
    }

    try {
        $guest = Get-LocalUser -Name 'Guest' -ErrorAction Stop
        $status = if ($guest.Enabled) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Disabled/default/unused accounts' -NormStatus $status `
            -Detected "Guest.Enabled=$($guest.Enabled)" -Expected 'Guest disabled; no unused enabled accounts' `
            -Detail 'Guest account state only - review other local accounts for staleness manually.'
    } catch {
        Add-Row -Category $Cat -Item 'Disabled/default/unused accounts' -NormStatus 'Compliant' `
            -Detected 'Guest account not found (already removed/renamed)' -Expected 'Guest disabled; no unused enabled accounts'
    }

    try {
        $na = (& net.exe accounts) -join "`n"
        $pwLenMatch = [regex]::Match($na, 'password length:\s*(\d+)')
        $lockMatch  = [regex]::Match($na, 'lockout threshold:\s*(\d+|Never)', 'IgnoreCase')
        $ok = $true
        if ($pwLenMatch.Success -and [int]$pwLenMatch.Groups[1].Value -lt $MinPasswordLength) { $ok = $false }
        if ($lockMatch.Success -and $lockMatch.Groups[1].Value -ne 'Never' -and [int]$lockMatch.Groups[1].Value -gt $LockoutThreshold) { $ok = $false }
        Add-Row -Category $Cat -Item 'Password/security policy' -NormStatus $(if ($ok) { 'Compliant' } else { 'Non-Compliant' }) `
            -Detected ($na -replace '\s+', ' ').Trim() `
            -Expected "Min password length >= $MinPasswordLength, lockout threshold <= $LockoutThreshold"
    } catch {
        Add-Row -Category $Cat -Item 'Password/security policy' -NormStatus 'Unable to Check' `
            -Detected "net accounts failed: $($_.Exception.Message)" -Expected 'Meets org password/lockout baseline'
    }

    $firewallAllEnabledForFleet = $null
    try {
        $fw = Get-NetFirewallProfile -ErrorAction Stop
        $off = @($fw | Where-Object { -not $_.Enabled })
        $firewallAllEnabledForFleet = ($off.Count -eq 0)
        $status = if ($off.Count -gt 0) { 'Non-Compliant' } else { 'Compliant' }
        Add-Row -Category $Cat -Item 'Windows Firewall status' -NormStatus $status `
            -Detected (($fw | ForEach-Object { "$($_.Name)=$($_.Enabled)" }) -join ', ') -Expected 'All profiles Enabled'
    } catch {
        Add-Row -Category $Cat -Item 'Windows Firewall status' -NormStatus 'Unable to Check' `
            -Detected "Get-NetFirewallProfile failed: $($_.Exception.Message)" -Expected 'All profiles Enabled'
    }

    try {
        $mp = Get-MpComputerStatus -ErrorAction Stop
        $ageOk = $mp.AntivirusSignatureAge -le 2
        $status = if ($mp.RealTimeProtectionEnabled -and $ageOk) { 'Compliant' } else { 'Non-Compliant' }
        Add-Row -Category $Cat -Item 'Defender/EDR status' -NormStatus $status `
            -Detected "RealTimeProtectionEnabled=$($mp.RealTimeProtectionEnabled) SignatureAgeDays=$($mp.AntivirusSignatureAge)" `
            -Expected 'RealTimeProtectionEnabled=True, signature age <= 2 days'
    } catch {
        Add-Row -Category $Cat -Item 'Defender/EDR status' -NormStatus 'Unable to Check' `
            -Detected "Get-MpComputerStatus failed: $($_.Exception.Message)" -Expected 'RealTimeProtectionEnabled=True, signatures current'
    }

    try {
        $smb = Get-SmbServerConfiguration -ErrorAction Stop
        $smb1Ok = -not $smb.EnableSMB1Protocol
        $sigOk  = $smb.RequireSecuritySignature
        $status = if ($smb1Ok -and $sigOk) { 'Compliant' } else { 'Non-Compliant' }
        Add-Row -Category $Cat -Item 'SMB configuration (SMBv1, signing)' -NormStatus $status `
            -Detected "EnableSMB1Protocol=$($smb.EnableSMB1Protocol) RequireSecuritySignature=$($smb.RequireSecuritySignature)" `
            -Expected 'SMBv1 disabled; signing required'
    } catch {
        Add-Row -Category $Cat -Item 'SMB configuration (SMBv1, signing)' -NormStatus 'Unable to Check' `
            -Detected "Get-SmbServerConfiguration failed: $($_.Exception.Message)" -Expected 'SMBv1 disabled; signing required'
    }

    try {
        $ts = Get-CimInstance -Namespace root/cimv2/terminalservices -ClassName Win32_TerminalServiceSetting -ErrorAction Stop
        $nlaOn = [bool]$ts.UserAuthenticationRequired
        if (@($AllowedRdpSources).Count -eq 0) {
            Add-Row -Category $Cat -Item 'RDP configuration' -NormStatus $(if ($nlaOn) { 'Manual Verification Required' } else { 'Non-Compliant' }) `
                -Detected "NLA(UserAuthenticationRequired)=$nlaOn; no allowed-source list supplied" `
                -Expected 'NLA required; direct RDP restricted to the PAM/jump path' `
                -Detail 'Re-run with -AllowedRdpSources to also auto-check firewall scoping.'
        } else {
            $rules = @(Get-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue | Where-Object { $_.Enabled -eq 'True' })
            $addr = @($rules | Get-NetFirewallAddressFilter -ErrorAction SilentlyContinue | Select-Object -ExpandProperty RemoteAddress -Unique)
            $scopedOk = ($addr.Count -gt 0) -and ($addr -notcontains 'Any')
            $status = if ($nlaOn -and $scopedOk) { 'Compliant' } else { 'Non-Compliant' }
            Add-Row -Category $Cat -Item 'RDP configuration' -NormStatus $status `
                -Detected "NLA=$nlaOn RemoteAddressScope=$($addr -join ', ')" `
                -Expected "NLA required; scoped to $($AllowedRdpSources -join ', ')"
        }
    } catch {
        Add-Row -Category $Cat -Item 'RDP configuration' -NormStatus 'Unable to Check' `
            -Detected "Win32_TerminalServiceSetting query failed: $($_.Exception.Message)" -Expected 'NLA required; RDP scoped to PAM/jump path'
    }

    try {
        $legacy = 'SSL 2.0', 'SSL 3.0', 'TLS 1.0', 'TLS 1.1'
        $findings = New-Object System.Collections.Generic.List[string]
        $explicit = $false
        foreach ($proto in $legacy) {
            $v = Get-RegVal "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server" 'Enabled'
            if ($v.Found) {
                $explicit = $true
                if ($v.Value -ne 0) { $findings.Add("$proto Enabled=$($v.Value)") }
            }
        }
        if ($findings.Count -gt 0) {
            Add-Row -Category $Cat -Item 'TLS/security protocol configuration' -NormStatus 'Non-Compliant' `
                -Detected ($findings -join '; ') -Expected 'Only TLS 1.2/1.3 enabled'
        } elseif ($explicit) {
            Add-Row -Category $Cat -Item 'TLS/security protocol configuration' -NormStatus 'Compliant' `
                -Detected 'All legacy protocols explicitly disabled' -Expected 'Only TLS 1.2/1.3 enabled'
        } else {
            Add-Row -Category $Cat -Item 'TLS/security protocol configuration' -NormStatus 'Manual Verification Required' `
                -Detected 'No explicit SCHANNEL protocol overrides found' -Expected 'Only TLS 1.2/1.3 enabled' `
                -Detail 'Legacy protocols are not explicitly disabled by registry policy - verify against the OS build default and the org crypto baseline.'
        }
    } catch {
        Add-Row -Category $Cat -Item 'TLS/security protocol configuration' -NormStatus 'Unable to Check' `
            -Detected "SCHANNEL registry query failed: $($_.Exception.Message)" -Expected 'Only TLS 1.2/1.3 enabled'
    }

    Add-Row -Category $Cat -Item 'Security baseline/hardening compliance' -NormStatus 'Manual Verification Required' `
        -Detected 'Not scriptable from inside the guest' -Expected 'No open baseline compliance findings, or accepted-risk documented' `
        -Detail 'Confirm with the Security team against the CIS/Microsoft baseline GPO report.'

    try {
        $ap = (& auditpol.exe /get /category:*) -join "`n"
        $lines = ($ap -split "`n" | Where-Object { $_ -match '\S' } | Select-Object -First 40) -join ' | '
        Add-Row -Category $Cat -Item 'Audit policy configuration' -NormStatus 'Manual Verification Required' `
            -Detected $lines -Expected 'Matches org audit policy baseline (success/failure per category)'
    } catch {
        Add-Row -Category $Cat -Item 'Audit policy configuration' -NormStatus 'Unable to Check' `
            -Detected "auditpol failed: $($_.Exception.Message)" -Expected 'Matches org audit policy baseline'
    }

    # ============================ 5. INSTALLED AGENTS ============================
    $Cat = '5. Installed Agents'
    Test-AgentService -Name 'QualysAgent'     -Label 'Qualys agent'                          -Category $Cat
    try {
        $bin = Get-Service -ErrorAction Stop | Where-Object { $_.DisplayName -like '*Binalyze*' } | Select-Object -First 1
        if ($bin) {
            $status = if ($bin.Status -eq 'Running') { 'Compliant' } else { 'Non-Compliant' }
            Add-Row -Category $Cat -Item 'Binalyze AIR agent' -NormStatus $status -Detected "$($bin.DisplayName) Status=$($bin.Status)" -Expected 'Installed and Running'
        } else {
            Add-Row -Category $Cat -Item 'Binalyze AIR agent' -NormStatus 'Non-Compliant' -Detected 'No Binalyze service found' -Expected 'Installed and Running'
        }
    } catch {
        Add-Row -Category $Cat -Item 'Binalyze AIR agent' -NormStatus 'Unable to Check' -Detected "Get-Service failed: $($_.Exception.Message)" -Expected 'Installed and Running'
    }
    Test-AgentService -Name 'SplunkForwarder' -Label 'Splunk Universal Forwarder'            -Category $Cat
    Test-AgentService -Name 'Sense'           -Label 'Microsoft Defender for Endpoint (Sense)' -Category $Cat

    $arcExe = 'C:\Program Files\AzureConnectedMachineAgent\azcmagent.exe'
    $arcInstalled = Test-Path -LiteralPath $arcExe
    if ($arcInstalled) {
        try {
            $show = (& $arcExe show) -join "`n"
            $status = if ($show -match 'Connected') { 'Compliant' } else { 'Non-Compliant' }
            Add-Row -Category $Cat -Item 'Azure Arc agent' -NormStatus $status -Detected ($show -replace '\s+', ' ').Trim() -Expected 'Status: Connected'
        } catch {
            Add-Row -Category $Cat -Item 'Azure Arc agent' -NormStatus 'Unable to Check' -Detected "azcmagent show failed: $($_.Exception.Message)" -Expected 'Status: Connected'
        }
    } else {
        Add-Row -Category $Cat -Item 'Azure Arc agent' -NormStatus 'Non-Compliant' -Detected 'azcmagent.exe not found' -Expected 'Installed and Connected' `
            -Detail 'If Arc/AUM management is not intended for this VM, re-classify as Not Applicable manually.'
    }

    $agentSetForFleet = New-Object System.Collections.Generic.List[string]
    foreach ($n in @('QualysAgent', 'SplunkForwarder', 'Sense')) {
        if (Get-Service -Name $n -ErrorAction SilentlyContinue) { $agentSetForFleet.Add($n) }
    }
    if ($arcInstalled) { $agentSetForFleet.Add('AzureArc') }

    if (@($ExpectedAgentServices).Count -eq 0) {
        Add-Row -Category $Cat -Item 'Other monitoring/management agents' -NormStatus 'Not Applicable' `
            -Detected 'No additional org-mandated services supplied' -Expected 'n/a' `
            -Detail 'Re-run with -ExpectedAgentServices to check additional mandated agents.'
    } else {
        foreach ($svcName in $ExpectedAgentServices) {
            Test-AgentService -Name $svcName -Label "Mandated agent: $svcName" -Category $Cat
            if (Get-Service -Name $svcName -ErrorAction SilentlyContinue) { $agentSetForFleet.Add($svcName) }
        }
    }

    # ============================ 6. PATCHING & VULNERABILITY MGMT ============================
    $Cat = '6. Patching & Vulnerability Mgmt'

    $latestPatchDate = $null
    try {
        $hf = Get-HotFix -ErrorAction Stop | Sort-Object InstalledOn -Descending
        $top = $hf | Select-Object -First 5
        if (@($top).Count -gt 0 -and $top[0].InstalledOn) { $latestPatchDate = $top[0].InstalledOn }
        Add-Row -Category $Cat -Item 'Current patch level' -NormStatus 'Manual Verification Required' `
            -Detected (($top | ForEach-Object { "$($_.HotFixID)($($_.InstalledOn))" }) -join ', ') `
            -Expected 'Latest cumulative update within 1-2 patch cycles of current'
    } catch {
        Add-Row -Category $Cat -Item 'Current patch level' -NormStatus 'Unable to Check' `
            -Detected "Get-HotFix failed: $($_.Exception.Message)" -Expected 'Latest cumulative update within 1-2 patch cycles'
    }

    try {
        $qfe = Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction Stop
        $dates = $qfe | Where-Object { $_.InstalledOn } | ForEach-Object { try { [datetime]$_.InstalledOn } catch { $null } } | Where-Object { $_ }
        if ($dates -and -not $latestPatchDate) { $latestPatchDate = ($dates | Sort-Object -Descending | Select-Object -First 1) }
        if ($latestPatchDate) {
            $ageDays = [math]::Round(((Get-Date) - $latestPatchDate).TotalDays, 0)
            $status = if ($ageDays -gt $PatchAgeWarnDays) { 'Non-Compliant' } else { 'Compliant' }
            Add-Row -Category $Cat -Item 'Last successful update run' -NormStatus $status `
                -Detected "Most recent installed update: $latestPatchDate ($ageDays days ago)" -Expected "Successful update within the last $PatchAgeWarnDays days"
        } else {
            Add-Row -Category $Cat -Item 'Last successful update run' -NormStatus 'Manual Verification Required' `
                -Detected 'No parseable install dates returned' -Expected "Successful update within the last $PatchAgeWarnDays days"
        }
    } catch {
        Add-Row -Category $Cat -Item 'Last successful update run' -NormStatus 'Unable to Check' `
            -Detected "Win32_QuickFixEngineering query failed: $($_.Exception.Message)" -Expected "Successful update within the last $PatchAgeWarnDays days"
    }

    Add-Row -Category $Cat -Item 'Missing/pending updates' -NormStatus 'Manual Verification Required' `
        -Detected 'Not determinable from inside the guest alone' -Expected 'No missing Critical/Important updates outstanding' `
        -Detail 'Cross-reference the latest Qualys or Azure Update Manager assessment for this VM.'

    Add-Row -Category $Cat -Item 'Pending reboot (patch context)' -NormStatus $pendingStatus `
        -Detected $(if ($pending.Count -gt 0) { $pending -join ', ' } else { 'No pending-reboot flags set' }) `
        -Expected 'No pending reboot at time of acceptance' -Detail 'Reused from category 1 (Pending reboot status).'

    if ($arcInstalled) {
        try {
            $chk = (& $arcExe check) -join "`n"
            $status = if ($chk -match '(?i)fail|blocked|error') { 'Non-Compliant' } else { 'Compliant' }
            Add-Row -Category $Cat -Item 'Azure Arc connectivity for patching' -NormStatus $status `
                -Detected ($chk -replace '\s+', ' ').Trim() -Expected 'All required endpoints reachable, no proxy/firewall blocks'
        } catch {
            Add-Row -Category $Cat -Item 'Azure Arc connectivity for patching' -NormStatus 'Unable to Check' `
                -Detected "azcmagent check failed: $($_.Exception.Message)" -Expected 'All required endpoints reachable'
        }
    } else {
        Add-Row -Category $Cat -Item 'Azure Arc connectivity for patching' -NormStatus 'Non-Compliant' `
            -Detected 'azcmagent.exe not found' -Expected 'All required endpoints reachable'
    }

    try {
        $auPolicy = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' -ErrorAction Stop
        $wsus = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' 'WUServer'
        $detail = "AU policy present: NoAutoUpdate=$($auPolicy.NoAutoUpdate) AUOptions=$($auPolicy.AUOptions)"
        if ($wsus.Found) { $detail += " WUServer=$($wsus.Value)" }
        Add-Row -Category $Cat -Item 'Windows Update configuration conflicts' -NormStatus 'Non-Compliant' `
            -Detected $detail -Expected 'No conflicting policy source; the designated tool (e.g. AUM) is authoritative'
    } catch {
        Add-Row -Category $Cat -Item 'Windows Update configuration conflicts' -NormStatus 'Compliant' `
            -Detected 'No local WindowsUpdate\AU policy override present' -Expected 'No conflicting policy source'
    }

    Add-Row -Category $Cat -Item 'Azure Update Manager readiness' -NormStatus 'Manual Verification Required' `
        -Detected 'Not scriptable from inside the guest' -Expected 'Patch orchestration set correctly, recent assessment' `
        -Detail 'Confirm patch orchestration and last assessment date in the Azure portal.'
    Add-Row -Category $Cat -Item 'Qualys vulnerability posture' -NormStatus 'Manual Verification Required' `
        -Detected 'Not scriptable from inside the guest' -Expected 'No open Critical findings; High findings have a remediation plan' `
        -Detail 'Confirm in the Qualys console.'

    $Fleet.DnsServers          = $dnsServersForFleet
    $Fleet.TimeSource          = $timeSourceForFleet
    $Fleet.AdminMembers        = $adminMembersForFleet
    $Fleet.FirewallAllEnabled  = $firewallAllEnabledForFleet
    $Fleet.AgentSet            = (($agentSetForFleet | Sort-Object -Unique) -join ',')

} catch {
    $fatal = $_.Exception.Message
}

$meta = [PSCustomObject]@{ Hostname = $ComputerName; IPAddress = $IPAddresses; OS = $OSCaption; FatalError = $fatal }
# .ToArray() rather than @(...) - @() around a generic List instance is unreliable
# on some PowerShell builds; .ToArray() is well-defined on 5.1 and 7.x alike.
$envelope = [PSCustomObject]@{ Schema = 'osaccept-1'; Meta = $meta; Fleet = $Fleet; Rows = $Rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<OSACCEPT-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-OSACCEPT-ENVELOPE>>>'
'@

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
    $b64 = $sb.ToString()
    if ([string]::IsNullOrWhiteSpace($b64)) { return $null }
    try { return ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json) }
    catch { return $null }
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
            Reason = "VM '$Name' was not found in any connected vCenter ($serverList)."; Detected = '0 matches'; VM = $null; Server = $null }
    }
    if ($hits.Count -gt 1) {
        $where = ($hits | ForEach-Object { "$($_.Server.Name):$($_.VM.Id)" }) -join '; '
        return [pscustomobject]@{ Ok = $false; Stage = 'VM lookup'; ServerName = $hits[0].Server.Name
            Reason = "VM name '$Name' is ambiguous - $($hits.Count) matching VMs ($where). Skipped for safety."; Detected = "$($hits.Count) matches"; VM = $null; Server = $null }
    }
    $vm = $hits[0].VM; $srv = $hits[0].Server
    if ($vm.PowerState -ne 'PoweredOn') {
        return [pscustomobject]@{ Ok = $false; Stage = 'Power state'; ServerName = $srv.Name
            Reason = "VM is not powered on (PowerState=$($vm.PowerState))."; Detected = "PowerState=$($vm.PowerState)"; VM = $vm; Server = $srv }
    }
    $g = $vm.ExtensionData.Guest
    $toolsOk = ($g.ToolsRunningStatus -eq 'guestToolsRunning') -or ($g.ToolsStatus -in 'toolsOk', 'toolsOld')
    if (-not $toolsOk) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VMware Tools'; ServerName = $srv.Name
            Reason = "VMware Tools not installed/running (ToolsStatus=$($g.ToolsStatus), ToolsRunningStatus=$($g.ToolsRunningStatus))."; Detected = "ToolsStatus=$($g.ToolsStatus)"; VM = $vm; Server = $srv }
    }
    $isWin = ($g.GuestFamily -eq 'windowsGuest') -or ($g.GuestId -match 'windows') -or ($vm.Guest.OSFullName -match 'Windows')
    if (-not $isWin) {
        return [pscustomobject]@{ Ok = $false; Stage = 'Guest OS'; ServerName = $srv.Name
            Reason = "Guest OS is not Windows (GuestFamily=$($g.GuestFamily), OS=$($vm.Guest.OSFullName))."; Detected = "$($vm.Guest.OSFullName)"; VM = $vm; Server = $srv }
    }
    return [pscustomobject]@{ Ok = $true; Stage = 'OK'; ServerName = $srv.Name; Reason = ''; Detected = ''; VM = $vm; Server = $srv }
}

# VMware Tools status is the one item on the checklist that comes from the VM
# object in vCenter rather than from inside the guest - every other check
# depends on Invoke-VMScript, which itself depends on Tools being installed
# and running, so this is recorded for every VM that could be uniquely
# resolved, even when the VM is later skipped for being off or Tools-less.
function Split-CamelWords {
    # 'toolsOk' -> 'tools Ok', 'guestToolsNotRunning' -> 'guest Tools Not Running' - readability only,
    # never applied to the raw enum values used for comparisons.
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    return ($Text -creplace '([a-z])([A-Z])', '$1 $2')
}

function Get-ToolsStatusFields {
    param($VM)
    $g = $VM.ExtensionData.Guest
    $toolsStatus  = [string]$g.ToolsStatus
    $toolsRunning = [string]$g.ToolsRunningStatus
    $toolsVersion = [string]$VM.Guest.ToolsVersion
    $detected = "ToolsStatus= $(Split-CamelWords $toolsStatus) ToolsRunningStatus= $(Split-CamelWords $toolsRunning) ToolsVersion= $toolsVersion"
    $status =
        if ($toolsStatus -eq 'toolsOk' -and $toolsRunning -eq 'guestToolsRunning') { 'Compliant' }
        elseif ($toolsStatus -eq 'toolsOld' -and $toolsRunning -eq 'guestToolsRunning') { 'Manual Verification Required' }
        else { 'Non-Compliant' }
    $reason =
        if ($status -eq 'Manual Verification Required') { 'VMware Tools is running but out of date - schedule an upgrade.' }
        elseif ($status -eq 'Non-Compliant') { 'VMware Tools is not installed/running - guest heartbeat, quiesced snapshots/backups, and every other guest-ops check on this VM are unavailable until this is fixed.' }
        else { '' }
    return [pscustomobject]@{ Detected = $detected; Status = $status; Reason = $reason }
}

function Test-GuestPowerShell {
    param($VM, $Server, [pscredential]$Credential, [int]$ToolsWaitSecs)
    $probe = @'
try {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    $o = [pscustomobject]@{ PS = $PSVersionTable.PSVersion.ToString(); User = $id.Name; Elevated = [bool]$pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator); Host = $env:COMPUTERNAME }
    Write-Output ('<<<PROBE>>>' + ($o | ConvertTo-Json -Compress) + '<<<ENDPROBE>>>')
} catch {
    Write-Output ('<<<PROBE-ERR>>>' + $_.Exception.Message + '<<<ENDPROBE-ERR>>>')
}
'@
    try {
        $r = Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $probe `
                             -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop
    } catch {
        $m = $_.Exception.Message
        $reason = switch -Regex ($m) {
            'InvalidGuestLogin|authentication fail|incorrect user name or password|InvalidLogin|guest permissions|permission to perform this operation' {
                "Guest authentication failed - VMware Tools rejected the supplied credential ('$($Credential.UserName)'). Underlying: $m"; break }
            'Tools are not running|GuestOperationsUnavailable|VMware Tools is not|not installed|VIX' { "VMware Tools guest operations unavailable: $m"; break }
            'timed out|timeout' { "Timed out waiting for VMware Tools / guest response: $m"; break }
            default { "Guest PowerShell probe failed: $m" }
        }
        return [pscustomobject]@{ Ok = $false; Reason = $reason; PSVersion = ''; Elevated = $false; Hostname = '' }
    }
    $out = [string]$r.ScriptOutput
    if ($out -match '<<<PROBE>>>(.*?)<<<ENDPROBE>>>') {
        try { $p = $Matches[1] | ConvertFrom-Json }
        catch { return [pscustomobject]@{ Ok = $false; Reason = "Guest probe returned unreadable JSON: $($Matches[1])"; PSVersion = ''; Elevated = $false; Hostname = '' } }
        try { $ver = [version]$p.PS } catch { $ver = [version]'0.0' }
        if ($ver -lt [version]'5.1') {
            return [pscustomobject]@{ Ok = $false; Reason = "Guest PowerShell is $($p.PS); 5.1+ required."; PSVersion = $p.PS; Elevated = [bool]$p.Elevated; Hostname = $p.Host }
        }
        return [pscustomobject]@{ Ok = $true; Reason = ''; PSVersion = $p.PS; Elevated = [bool]$p.Elevated; Hostname = $p.Host }
    }
    if ($out -match '<<<PROBE-ERR>>>(.*?)<<<ENDPROBE-ERR>>>') {
        return [pscustomobject]@{ Ok = $false; Reason = "Guest probe raised: $($Matches[1])"; PSVersion = ''; Elevated = $false; Hostname = '' }
    }
    return [pscustomobject]@{ Ok = $false; Reason = "Guest probe produced no recognisable output. Raw: $((($out -replace '\s+', ' ').Trim()))"; PSVersion = ''; Elevated = $false; Hostname = '' }
}

# A single Invoke-VMScript carrying the whole payload can fail in some
# environments with a misleading "Could not locate Powershell script
# interpreter" error (the script overruns a guest-ops command-line length
# limit). Primary method here: Copy-VMGuestFile (one transfer, no limit);
# fallback: write Base64 into a guest file in small pieces, then decode
# in-guest. Guest operations only - no WinRM/PsExec/SMB/network path. The
# staged file holds no credential and is always deleted.
function Invoke-LargeGuestPayload {
    param(
        $VM, $Server, [pscredential]$Credential,
        [string]$PayloadText, [int]$ToolsWaitSecs,
        [int]$ChunkSize = 1500
    )
    $tag    = [guid]::NewGuid().ToString('N').Substring(0, 12)
    $gB64   = "C:\Windows\Temp\osacc$tag.b64"
    $gPs1   = "C:\Windows\Temp\osacc$tag.ps1"
    $b64    = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($PayloadText))
    $staged = $false
    Write-Log "  [stage] payload is $($PayloadText.Length) chars; delivering to guest as a file." 'INFO'

    try {
        if (Get-Command Copy-VMGuestFile -ErrorAction SilentlyContinue) {
            $local = Join-Path $env:TEMP "osacc$tag.ps1"
            try {
                [System.IO.File]::WriteAllText($local, $PayloadText, (New-Object System.Text.UTF8Encoding($false)))
                Copy-VMGuestFile -Source $local -Destination $gPs1 -VM $VM -Server $Server `
                                 -GuestCredential $Credential -LocalToGuest -Force -ErrorAction Stop | Out-Null
                $staged = $true
                Write-Log "  staged payload into guest via Copy-VMGuestFile." 'INFO'
            } catch {
                Write-Log "  Copy-VMGuestFile not usable ($(($_.Exception.Message -split "`n")[0].Trim())); using chunked staging." 'INFO'
            } finally {
                Remove-Item -LiteralPath $local -Force -ErrorAction SilentlyContinue
            }
        }

        if (-not $staged) {
            $total = [Math]::Ceiling($b64.Length / $ChunkSize)
            Write-Log "  staging payload into guest in $total chunk(s) of <= $ChunkSize chars." 'INFO'
            for ($i = 0; $i -lt $total; $i++) {
                $part = $b64.Substring($i * $ChunkSize, [Math]::Min($ChunkSize, $b64.Length - ($i * $ChunkSize)))
                $verb = if ($i -eq 0) { 'Set-Content' } else { 'Add-Content' }
                $st = "$verb -LiteralPath '$gB64' -Value '$part'"
                Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $st `
                                -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop | Out-Null
            }
            $decoder = @"
`$raw = [System.IO.File]::ReadAllText('$gB64')
`$bytes = [Convert]::FromBase64String(( `$raw -replace '\s','' ))
[System.IO.File]::WriteAllText('$gPs1', [System.Text.Encoding]::UTF8.GetString(`$bytes))
"@
            Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $decoder `
                            -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop | Out-Null
        }

        $runner = "& powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File '$gPs1' 2>&1"
        $res = Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $runner `
                               -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop
        return [string]$res.ScriptOutput
    }
    finally {
        $cleanup = "Remove-Item -LiteralPath '$gB64','$gPs1' -Force -ErrorAction SilentlyContinue"
        try {
            Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $cleanup `
                            -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}

function New-CentralRow {
    param($vCenter, $VMName, $GuestHostname, $IPAddress, $OS, $Category, $Item, $DetectedValue, $ExpectedValue, $Status, $Reason)
    [PSCustomObject]([ordered]@{
        vCenter = $vCenter; VMName = $VMName; GuestHostname = $GuestHostname; IPAddress = $IPAddress; OS = $OS
        Category = $Category; Item = $Item; DetectedValue = $DetectedValue; ExpectedValue = $ExpectedValue
        Status = $Status; 'Error/Reason' = $Reason; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

function Show-VmResults {
    param([string]$VmBanner, $Rows)
    Write-Host ""
    Write-Host ("##### $VmBanner #####") -ForegroundColor Cyan
    $lastCat = $null
    foreach ($row in $Rows) {
        if ($row.Category -ne $lastCat) {
            Write-Host ""
            Write-Host ("=== {0} ===" -f $row.Category) -ForegroundColor Cyan
            $lastCat = $row.Category
        }
        $color = switch ($row.Status) {
            'Compliant'                     { 'Green' }
            'Manual Verification Required'  { 'DarkYellow' }
            'Not Applicable'                { 'Gray' }
            'Unable to Check'               { 'Yellow' }
            default                         { 'Red' }
        }
        Write-Host ("  [{0}] " -f $row.Status) -ForegroundColor $color -NoNewline
        Write-Host $row.Item -ForegroundColor White
        if ($row.DetectedValue) { Write-Host ("      Detected: {0}" -f (([string]$row.DetectedValue) -replace '\s+', ' ').Trim()) -ForegroundColor Gray }
    }
}

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows = New-Object System.Collections.Generic.List[object]
$fleetData   = New-Object System.Collections.Generic.List[object]
$summary = [ordered]@{ Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0; Manual = 0; Unable = 0; SkippedFailed = 0 }
$inventory = Get-VMInventory -Servers $connectedServers

$dnsTestLiteral      = ConvertTo-PSArrayLiteral -Items $DnsTestNames
$connLiteral         = ConvertTo-PSArrayLiteral -Items $RequiredConnectivity
$agentsLiteral       = ConvertTo-PSArrayLiteral -Items $ExpectedAgentServices
$rdpSourcesLiteral   = ConvertTo-PSArrayLiteral -Items $AllowedRdpSources

$payloadForGuest = $PayloadTemplate.
    Replace('{{EVENT_LOOKBACK_DAYS}}', ([string]$EventLookbackDays)).
    Replace('{{DISK_FREE_PCT}}', ([string]$DiskFreePercentThreshold)).
    Replace('{{UPTIME_WARN_DAYS}}', ([string]$UptimeWarnDays)).
    Replace('{{PATCH_AGE_DAYS}}', ([string]$PatchAgeWarnDays)).
    Replace('{{MIN_PW_LENGTH}}', ([string]$MinPasswordLength)).
    Replace('{{LOCKOUT_THRESHOLD}}', ([string]$LockoutThreshold)).
    Replace('{{SKIP_ROLES}}', $(if ($SkipRolesInventory) { '$true' } else { '$false' })).
    Replace('{{DNS_TEST_NAMES}}', $dnsTestLiteral).
    Replace('{{REQUIRED_CONNECTIVITY}}', $connLiteral).
    Replace('{{EXPECTED_AGENT_SERVICES}}', $agentsLiteral).
    Replace('{{ALLOWED_RDP_SOURCES}}', $rdpSourcesLiteral)

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Assessing VM '$vmName' ===" 'INFO'
    try {
        $val = Resolve-AndValidateVM -Name $vmName -Inventory $inventory

        if ($val.VM) {
            $toolsFields = Get-ToolsStatusFields -VM $val.VM
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -GuestHostname $val.VM.Guest.HostName -IPAddress ($val.VM.Guest.IPAddress -join ',') -OS $val.VM.Guest.OSFullName `
                -Category '1. Windows OS Health' -Item 'VMware Tools status' -DetectedValue $toolsFields.Detected `
                -ExpectedValue 'ToolsStatus= tools Ok, ToolsRunningStatus= guest Tools Running' -Status $toolsFields.Status -Reason $toolsFields.Reason))
        }

        if (-not $val.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -Item $val.Stage -DetectedValue $val.Detected -ExpectedValue '' `
                -Status 'Unable to Check' -Reason $val.Reason))
            Write-Log "SKIP '$vmName' [$($val.Stage)]: $($val.Reason)" 'WARN'
            continue
        }
        $vm = $val.VM; $srv = $val.Server

        $probe = Test-GuestPowerShell -VM $vm -Server $srv -Credential $GuestCredential -ToolsWaitSecs $ToolsWaitSecs
        if (-not $probe.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $vm.Guest.HostName -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -Item 'Guest authentication / PowerShell' -DetectedValue '' -ExpectedValue 'PowerShell 5.1+ reachable via VMware Tools' `
                -Status 'Unable to Check' -Reason $probe.Reason))
            Write-Log "SKIP '$vmName' [probe]: $($probe.Reason)" 'WARN'
            continue
        }

        Write-Log "'$vmName': staging + running read-only assessment payload." 'INFO'
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $payloadForGuest -ToolsWaitSecs $ToolsWaitSecs

        $envelope = Read-EnvelopeFromScriptOutput -Output $scriptOutput -StartMarker '<<<OSACCEPT-ENVELOPE-B64>>>' -EndMarker '<<<END-OSACCEPT-ENVELOPE>>>'
        if ($null -eq $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($scriptOutput) { ($scriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -Item 'Guest payload result' -DetectedValue '' -ExpectedValue 'Base64 result envelope' `
                -Status 'Unable to Check' -Reason "No parseable result envelope. Output start: $snippet"))
            Write-Log "'$vmName': no parseable envelope returned." 'ERROR'
            continue
        }

        $meta  = $envelope.Meta
        $ghost = if ($meta.Hostname)  { $meta.Hostname }  else { $vm.Guest.HostName }
        $gip   = if ($meta.IPAddress) { $meta.IPAddress } else { ($vm.Guest.IPAddress -join ',') }
        $gos   = if ($meta.OS)        { $meta.OS }        else { $vm.Guest.OSFullName }
        if ($meta.FatalError) { Write-Log "'$vmName': guest payload reported a fatal error: $($meta.FatalError)" 'ERROR' }

        $summary.Checked++
        $f = @{ NonCompliant = $false; Manual = $false; Unable = $false }
        foreach ($r in @($envelope.Rows)) {
            switch ($r.Status) {
                'Non-Compliant'                { $f.NonCompliant = $true }
                'Manual Verification Required' { $f.Manual = $true }
                'Unable to Check'              { $f.Unable = $true }
            }
            $reason = if ($r.Status -in 'Manual Verification Required', 'Unable to Check', 'Non-Compliant') { $r.Detail } else { '' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category $r.Category -Item $r.Item -DetectedValue $r.DetectedValue -ExpectedValue $r.ExpectedValue `
                -Status $r.Status -Reason $reason))
        }
        if ($f.NonCompliant) { $summary.NonCompliant++ } else { $summary.Compliant++ }
        if ($f.Manual)       { $summary.Manual++ }
        if ($f.Unable)       { $summary.Unable++ }

        if ($envelope.Fleet) {
            $fleetData.Add([pscustomobject]@{
                VMName = $vmName; vCenter = $srv.Name
                DnsServers = [string]$envelope.Fleet.DnsServers
                TimeSource = [string]$envelope.Fleet.TimeSource
                AdminMembers = [string]$envelope.Fleet.AdminMembers
                FirewallAllEnabled = $envelope.Fleet.FirewallAllEnabled
                AgentSet = [string]$envelope.Fleet.AgentSet
            }) | Out-Null
        }

        Show-VmResults -VmBanner "$vmName  @ $($srv.Name)   ($ghost / $gip)" -Rows @($envelope.Rows | ForEach-Object {
            [pscustomobject]@{ Category = $_.Category; Item = $_.Item; Status = $_.Status; DetectedValue = $_.DetectedValue }
        })

        Write-Log "'$vmName': assessed. nonCompliant=$($f.NonCompliant) manual=$($f.Manual) unable=$($f.Unable)" 'INFO'
    }
    catch {
        $summary.SkippedFailed++
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        try {
            $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -Item 'Processing' -DetectedValue '' -ExpectedValue '' `
                -Status 'Unable to Check' -Reason $_.Exception.Message))
        } catch { }
        continue
    }
}

# =====================================================================================
# 5. Category 7 - Configuration Consistency (host-side fleet comparison)
# =====================================================================================
function Get-Mode {
    param([string[]]$Values)
    $nonEmpty = @($Values | Where-Object { $_ -ne $null -and $_ -ne '' })
    if ($nonEmpty.Count -eq 0) { return '' }
    return ($nonEmpty | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
}

if ($fleetData.Count -ge 2) {
    $Cat7 = '7. Configuration Consistency'
    $modeDns    = Get-Mode -Values @($fleetData | ForEach-Object { $_.DnsServers })
    $modeTime   = Get-Mode -Values @($fleetData | ForEach-Object { $_.TimeSource })
    $modeAdmins = Get-Mode -Values @($fleetData | ForEach-Object { $_.AdminMembers })
    $modeAgents = Get-Mode -Values @($fleetData | ForEach-Object { $_.AgentSet })
    $fwEnabledCount = @($fleetData | Where-Object { $_.FirewallAllEnabled -eq $true }).Count
    $modeFw = $fwEnabledCount -ge ([math]::Ceiling($fleetData.Count / 2.0))

    foreach ($fd in $fleetData) {
        $centralRows.Add((New-CentralRow -vCenter $fd.vCenter -VMName $fd.VMName -GuestHostname '' -IPAddress '' -OS '' `
            -Category $Cat7 -Item 'DNS configuration consistency' -DetectedValue $fd.DnsServers -ExpectedValue $modeDns `
            -Status $(if ($fd.DnsServers -eq $modeDns) { 'Compliant' } else { 'Non-Compliant' }) `
            -Reason $(if ($fd.DnsServers -ne $modeDns) { 'Deviates from the fleet-common DNS server set.' } else { '' })))

        $centralRows.Add((New-CentralRow -vCenter $fd.vCenter -VMName $fd.VMName -GuestHostname '' -IPAddress '' -OS '' `
            -Category $Cat7 -Item 'Time sync configuration consistency' -DetectedValue $fd.TimeSource -ExpectedValue $modeTime `
            -Status $(if ($fd.TimeSource -eq $modeTime) { 'Compliant' } else { 'Non-Compliant' }) `
            -Reason $(if ($fd.TimeSource -ne $modeTime) { 'Deviates from the fleet-common time source.' } else { '' })))

        $centralRows.Add((New-CentralRow -vCenter $fd.vCenter -VMName $fd.VMName -GuestHostname '' -IPAddress '' -OS '' `
            -Category $Cat7 -Item 'Security agent version/coverage consistency' -DetectedValue $fd.AgentSet -ExpectedValue $modeAgents `
            -Status $(if ($fd.AgentSet -eq $modeAgents) { 'Compliant' } else { 'Non-Compliant' }) `
            -Reason $(if ($fd.AgentSet -ne $modeAgents) { 'Deviates from the fleet-common agent set.' } else { '' })))

        $centralRows.Add((New-CentralRow -vCenter $fd.vCenter -VMName $fd.VMName -GuestHostname '' -IPAddress '' -OS '' `
            -Category $Cat7 -Item 'Local administrator configuration consistency' -DetectedValue $fd.AdminMembers -ExpectedValue $modeAdmins `
            -Status $(if ($fd.AdminMembers -eq $modeAdmins) { 'Compliant' } else { 'Non-Compliant' }) `
            -Reason $(if ($fd.AdminMembers -ne $modeAdmins) { 'Deviates from the fleet-common local admin membership model.' } else { '' })))

        $centralRows.Add((New-CentralRow -vCenter $fd.vCenter -VMName $fd.VMName -GuestHostname '' -IPAddress '' -OS '' `
            -Category $Cat7 -Item 'Firewall configuration consistency' -DetectedValue $fd.FirewallAllEnabled -ExpectedValue $modeFw `
            -Status $(if ($fd.FirewallAllEnabled -eq $modeFw) { 'Compliant' } else { 'Non-Compliant' }) `
            -Reason $(if ($fd.FirewallAllEnabled -ne $modeFw) { 'Deviates from the fleet-common firewall posture.' } else { '' })))
    }
    $centralRows.Add((New-CentralRow -vCenter '' -VMName '(fleet)' -GuestHostname '' -IPAddress '' -OS '' `
        -Category $Cat7 -Item 'Patch configuration consistency' -DetectedValue '' -ExpectedValue 'Single authoritative patch source fleet-wide' `
        -Status 'Manual Verification Required' -Reason 'Requires the Azure/AUM console to confirm every VM uses the same patch source - not derivable from the guest alone.'))
} else {
    Write-Log "Fewer than 2 VMs produced usable fleet data - Category 7 consistency pass skipped." 'WARN'
}

# =====================================================================================
# =====================================================================================
# 6. HTML dashboard template
# =====================================================================================
# Single-quoted here-string (no PS variable expansion) - {{TOKENS}} are swapped for real
# JSON further down. Category and item order is reconstructed from first-seen order in
# ROWS (already emitted in checklist order per VM), so there is no separate canonical
# order list to keep in sync here either.
$DashboardTemplate = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>OS Acceptance Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@600;700;800&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#F3F6F5; --surface:#FFFFFF; --surface-2:#EBF1EF; --border:#D7E0DD;
    --ink:#16211F; --ink-dim:#57655F; --ink-faint:#8A968F;
    --accent:#147D77; --accent-soft:rgba(20,125,119,.12);
    --good:#1F8A4C; --good-soft:rgba(31,138,76,.11);
    --bad:#C23B34; --bad-soft:rgba(194,59,52,.10);
    --warn:#A6720C; --warn-soft:rgba(166,114,12,.12);
    --mute:#68757E; --mute-soft:rgba(104,117,126,.12);
    --shadow:0 1px 2px rgba(20,33,31,.06), 0 10px 28px -16px rgba(20,33,31,.16);
    --radius:10px;
  }
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]){
      --bg:#0A0F0E; --surface:#111917; --surface-2:#16201D; --border:#223029;
      --ink:#E7EEEB; --ink-dim:#9CACA5; --ink-faint:#647570;
      --accent:#5BC7BF; --accent-soft:rgba(91,199,191,.16);
      --good:#4CC786; --good-soft:rgba(76,199,134,.14);
      --bad:#F0685F; --bad-soft:rgba(240,104,95,.15);
      --warn:#E5B84B; --warn-soft:rgba(229,184,75,.15);
      --mute:#8FA09A; --mute-soft:rgba(143,160,154,.14);
      --shadow:0 1px 2px rgba(0,0,0,.4), 0 14px 34px -18px rgba(0,0,0,.6);
    }
  }
  :root[data-theme="dark"]{
    --bg:#0A0F0E; --surface:#111917; --surface-2:#16201D; --border:#223029;
    --ink:#E7EEEB; --ink-dim:#9CACA5; --ink-faint:#647570;
    --accent:#5BC7BF; --accent-soft:rgba(91,199,191,.16);
    --good:#4CC786; --good-soft:rgba(76,199,134,.14);
    --bad:#F0685F; --bad-soft:rgba(240,104,95,.15);
    --warn:#E5B84B; --warn-soft:rgba(229,184,75,.15);
    --mute:#8FA09A; --mute-soft:rgba(143,160,154,.14);
    --shadow:0 1px 2px rgba(0,0,0,.4), 0 14px 34px -18px rgba(0,0,0,.6);
  }
  *{box-sizing:border-box;}
  html,body{margin:0; padding:0;}
  body{
    background:var(--bg); color:var(--ink);
    font-family:'IBM Plex Sans', -apple-system, Segoe UI, Roboto, Arial, sans-serif;
    font-size:14px; line-height:1.5;
  }
  ::selection{background:var(--accent-soft);}
  .wrap{max-width:1280px; margin:0 auto; padding:22px 24px 90px;}
  code, .mono{font-family:'IBM Plex Mono', Consolas, Menlo, monospace;}
  header{
    display:flex; justify-content:space-between; align-items:flex-end; gap:20px;
    padding-bottom:18px; margin-bottom:20px; border-bottom:1px solid var(--border); flex-wrap:wrap;
  }
  .h-eyebrow{font-family:'IBM Plex Mono'; font-size:11px; color:var(--accent); text-transform:uppercase; letter-spacing:.12em; margin-bottom:6px; font-weight:500;}
  h1{font-family:'Archivo'; font-weight:800; font-size:28px; margin:0; letter-spacing:-.01em; text-wrap:balance;}
  .h-meta{display:flex; gap:18px; flex-wrap:wrap; font-size:12.5px; color:var(--ink-dim); margin-top:4px;}
  .h-meta b{color:var(--ink); font-weight:600;}
  .h-right{text-align:right; font-size:12px; color:var(--ink-faint); font-family:'IBM Plex Mono';}
  .kpis{display:grid; grid-template-columns:repeat(5,1fr); gap:12px; margin-bottom:18px;}
  .kpi{background:var(--surface); border:1px solid var(--border); border-radius:var(--radius); padding:14px 16px; box-shadow:var(--shadow); position:relative; overflow:hidden;}
  .kpi::before{content:''; position:absolute; left:0; top:0; bottom:0; width:3px;}
  .kpi.total::before{background:var(--accent);} .kpi.good::before{background:var(--good);}
  .kpi.bad::before{background:var(--bad);} .kpi.warn::before{background:var(--warn);} .kpi.mute::before{background:var(--mute);}
  .kpi-label{font-size:10.5px; text-transform:uppercase; letter-spacing:.07em; color:var(--ink-faint); font-weight:600;}
  .kpi-value{font-family:'Archivo'; font-weight:800; font-size:30px; margin-top:5px; font-variant-numeric:tabular-nums;}
  .kpi.good .kpi-value{color:var(--good);} .kpi.bad .kpi-value{color:var(--bad);}
  .kpi.warn .kpi-value{color:var(--warn);} .kpi.mute .kpi-value{color:var(--mute);}
  .kpi-sub{font-size:11px; color:var(--ink-faint); margin-top:3px;}
  .panel{background:var(--surface); border:1px solid var(--border); border-radius:var(--radius); box-shadow:var(--shadow); padding:16px 18px; margin-bottom:16px;}
  .panel h2{font-family:'Archivo'; font-size:14.5px; font-weight:700; margin:0 0 4px;}
  .panel .panel-sub{font-size:11.5px; color:var(--ink-faint); margin-bottom:14px;}
  .exec-grid{display:grid; grid-template-columns:220px 1fr; gap:30px; align-items:center;}
  .exec-donut{display:flex; justify-content:center;}
  .exec-legend{display:grid; grid-template-columns:repeat(2,1fr); gap:10px 22px;}
  .leg-row{display:flex; align-items:center; justify-content:space-between; padding:8px 0; border-bottom:1px solid var(--border);}
  .leg-row .lbl{display:flex; align-items:center; gap:8px; font-size:12.5px; font-weight:600; color:var(--ink-dim);}
  .leg-row .val{font-family:'Archivo'; font-size:19px; font-weight:800; font-variant-numeric:tabular-nums;}
  .leg-row.good .val{color:var(--good);} .leg-row.bad .val{color:var(--bad);}
  .leg-row.warn .val{color:var(--warn);} .leg-row.mute .val{color:var(--mute);}
  .verdict{grid-column:1/-1; margin-top:4px; padding-top:14px; border-top:1px solid var(--border); display:flex; align-items:baseline; gap:10px; flex-wrap:wrap;}
  .verdict .vlabel{font-size:10.5px; text-transform:uppercase; letter-spacing:.08em; color:var(--ink-faint); font-weight:700;}
  .verdict .vtext{font-family:'Archivo'; font-size:22px; font-weight:800;}
  .verdict .vtext.good{color:var(--good);} .verdict .vtext.bad{color:var(--bad);} .verdict .vtext.warn{color:var(--warn);}
  .verdict .vnote{font-size:11.5px; color:var(--ink-faint);}
  .donut-grid{display:grid; grid-template-columns:repeat(auto-fill, minmax(150px,1fr)); gap:6px;}
  .donut-tile{display:flex; flex-direction:column; align-items:center; text-align:center; padding:12px 8px; border-radius:8px; cursor:pointer; transition:background .12s;}
  .donut-tile:hover{background:var(--surface-2);}
  .donut-tile .dname{font-size:11px; font-weight:600; margin-top:6px; line-height:1.3;}
  .donut-tile .dcounts{font-family:'IBM Plex Mono'; font-size:9.5px; color:var(--ink-faint); margin-top:2px;}
  @media(max-width:900px){ .exec-grid{grid-template-columns:1fr;} .exec-donut{padding-bottom:6px;} }
  .roster{display:flex; gap:8px; flex-wrap:wrap; margin-bottom:18px;}
  .roster-card{
    background:var(--surface); border:1px solid var(--border); border-radius:8px; padding:8px 12px;
    display:flex; align-items:center; gap:9px; box-shadow:var(--shadow); min-width:0;
  }
  .roster-dot{width:8px; height:8px; border-radius:50%; flex-shrink:0;}
  .roster-name{font-family:'IBM Plex Mono'; font-size:12px; font-weight:500;}
  .roster-sub{font-size:10.5px; color:var(--ink-faint);}
  .roster-card.skip{opacity:.65; border-style:dashed;}
  .controls{display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin-bottom:18px;}
  .chip{font-size:11.5px; font-weight:600; padding:6px 12px; border-radius:100px; border:1px solid var(--border); background:var(--surface); color:var(--ink-dim); cursor:pointer; user-select:none; display:flex; align-items:center; gap:6px; transition:border-color .12s, color .12s;}
  .chip:hover{border-color:var(--accent);}
  .chip .n{font-family:'IBM Plex Mono'; opacity:.75;}
  .chip.active{border-color:var(--accent); color:var(--accent); background:var(--accent-soft);}
  .swatch{width:7px; height:7px; border-radius:50%;}
  .swatch.good{background:var(--good);} .swatch.bad{background:var(--bad);} .swatch.warn{background:var(--warn);} .swatch.mute{background:var(--mute);}
  #search{margin-left:auto; background:var(--surface); border:1px solid var(--border); border-radius:8px; padding:7px 12px; font-size:12.5px; color:var(--ink); width:250px;}
  #search:focus{outline:2px solid var(--accent-soft); border-color:var(--accent);}
  #search::placeholder{color:var(--ink-faint);}
  .not-assessed{background:var(--mute-soft); border:1px solid var(--border); border-radius:var(--radius); padding:10px 16px; margin-bottom:18px; font-size:12px; color:var(--ink-dim);}
  .not-assessed b{color:var(--ink);}
  .not-assessed .row{padding:3px 0;}
  .not-assessed .mono{color:var(--ink-faint);}
  .layout{display:grid; grid-template-columns:220px 1fr; gap:26px; align-items:start;}
  .sidenav{position:sticky; top:16px; background:var(--surface); border:1px solid var(--border); border-radius:var(--radius); box-shadow:var(--shadow); padding:10px; max-height:calc(100vh - 32px); overflow-y:auto;}
  .sidenav a{
    display:flex; justify-content:space-between; align-items:center; gap:8px; padding:8px 10px; border-radius:7px;
    font-size:12px; color:var(--ink-dim); text-decoration:none; margin-bottom:2px; font-weight:500;
  }
  .sidenav a:hover{background:var(--surface-2); color:var(--ink);}
  .sidenav .cnt{font-family:'IBM Plex Mono'; font-size:10px; color:var(--ink-faint);}
  .sidenav .issue{background:var(--bad-soft); color:var(--bad); font-family:'IBM Plex Mono'; font-size:10px; font-weight:700; padding:1px 6px; border-radius:100px;}
  .cat-section{margin-bottom:8px; scroll-margin-top:16px;}
  .cat-heading{
    font-family:'Archivo'; font-weight:800; font-size:17px; margin:26px 0 12px; padding-bottom:8px;
    border-bottom:2px solid var(--border); display:flex; align-items:center; gap:10px;
  }
  .cat-heading:first-child{margin-top:0;}
  .cat-heading .idx{color:var(--accent); font-family:'IBM Plex Mono'; font-size:14px;}
  .item-card{background:var(--surface); border:1px solid var(--border); border-radius:var(--radius); box-shadow:var(--shadow); padding:14px 16px; margin-bottom:12px; scroll-margin-top:16px;}
  .item-top{display:flex; justify-content:space-between; align-items:flex-start; gap:14px; margin-bottom:10px;}
  .item-title{font-family:'Archivo'; font-weight:700; font-size:14px;}
  .item-exp{font-size:11px; color:var(--ink-faint); margin-top:3px;}
  .item-exp .mono{color:var(--ink-dim);}
  .item-tally{display:flex; gap:5px; flex-shrink:0;}
  .tally{font-family:'IBM Plex Mono'; font-size:10.5px; font-weight:700; padding:2px 8px; border-radius:100px;}
  .tally.good{background:var(--good-soft); color:var(--good);} .tally.bad{background:var(--bad-soft); color:var(--bad);}
  .tally.warn{background:var(--warn-soft); color:var(--warn);} .tally.mute{background:var(--mute-soft); color:var(--mute);}
  .vm-scroll{max-height:230px; overflow-y:auto; border-top:1px solid var(--border);}
  .vm-row{display:grid; grid-template-columns:130px 1fr; gap:12px; padding:8px 2px; border-bottom:1px solid var(--border); align-items:start;}
  .vm-row:last-child{border-bottom:none;}
  .vm-who{display:flex; align-items:center; gap:7px; padding-top:1px;}
  .vm-who .name{font-family:'IBM Plex Mono'; font-size:11.5px; font-weight:500;}
  .vm-detail .det{font-family:'IBM Plex Mono'; font-size:11.5px; color:var(--ink-dim); word-break:break-word;}
  .vm-detail .status-txt{font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:.04em; margin-right:6px;}
  .vm-detail .status-txt.good{color:var(--good);} .vm-detail .status-txt.bad{color:var(--bad);}
  .vm-detail .status-txt.warn{color:var(--warn);} .vm-detail .status-txt.mute{color:var(--mute);}
  .vm-detail .reason{font-size:11px; color:var(--warn); margin-top:3px;}
  .vm-detail .reason.crit{color:var(--bad);}
  .empty-state{text-align:center; padding:30px 0; color:var(--ink-faint); font-size:13px;}
  footer{margin-top:26px; padding-top:16px; border-top:1px solid var(--border); font-size:11.5px; color:var(--ink-faint);}
  @media(max-width:900px){
    .kpis{grid-template-columns:repeat(2,1fr);}
    .layout{grid-template-columns:1fr;}
    .sidenav{position:static; max-height:none;}
    #search{width:100%; margin-left:0;}
    .vm-row{grid-template-columns:1fr;}
  }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <div class="h-eyebrow">Windows OS Acceptance &mdash; Assessment Results</div>
      <h1>OS Acceptance Results</h1>
      <div class="h-meta">
        <span><b id="mTotal">-</b> VMs targeted</span>
        <span><b id="mAssessed">-</b> assessed &middot; <b id="mSkipped">-</b> not assessed</span>
        <span>vCenter: <b id="vcenterList">-</b></span>
      </div>
    </div>
    <div class="h-right" id="runMetaRight"></div>
  </header>

  <div class="panel">
    <h2>Fleet compliance &mdash; executive summary</h2>
    <div class="panel-sub">Every checklist item &times; every assessed VM, rolled up into one verdict</div>
    <div class="exec-grid">
      <div class="exec-donut" id="execDonut"></div>
      <div>
        <div class="exec-legend" id="execLegend"></div>
        <div class="verdict"><span class="vlabel">Overall assessment</span><span class="vtext" id="verdictText"></span><span class="vnote" id="verdictNote"></span></div>
      </div>
    </div>
  </div>

  <div class="kpis" id="kpiStrip"></div>

  <div class="panel">
    <h2>Compliance by category</h2>
    <div class="panel-sub">Item-level results aggregated across assessed VMs &mdash; click a donut to jump to that category</div>
    <div class="donut-grid" id="catChart"></div>
  </div>

  <div class="roster" id="roster"></div>
  <div id="notAssessed"></div>

  <div class="controls">
    <div id="statusChips" style="display:flex; gap:10px; flex-wrap:wrap;"></div>
    <input id="search" type="text" placeholder="Filter by VM, category, or item...">
  </div>

  <div class="layout">
    <nav class="sidenav" id="sidenav"></nav>
    <main id="main"></main>
  </div>

  <footer>
    Read-only assessment &mdash; <code class="mono">Invoke-OSAcceptanceCheck.ps1</code> via VMware Tools guest ops, plus a VMware Tools status read from the VM object in vCenter (the one item checked before any in-guest command is attempted).
  </footer>
</div>

<script>
const ROWS = {{ROWS_JSON}};
const RUN_META = {{RUN_META_JSON}};

function statusClass(status){
  if (status === 'Compliant') return 'good';
  if (status === 'Non-Compliant') return 'bad';
  if (status === 'Manual Verification Required') return 'warn';
  return 'mute';
}
const STATUS_LABEL = { good:'Compliant', bad:'Non-Compliant', warn:'Manual Review', mute:'Unable to Check' };

const VM_META = {};
const catOrder = [];
const itemsByCat = {};
const dataByKey = {};

ROWS.forEach(row=>{
  if (!VM_META[row.vm]) VM_META[row.vm] = { os:'', ip:'', vCenter:row.vCenter||'', cats:new Set(), skipReason:'' };
  const vm = VM_META[row.vm];
  if (row.os) vm.os = row.os;
  if (row.ip) vm.ip = row.ip;
  if (row.vCenter) vm.vCenter = row.vCenter;
  vm.cats.add(row.cat);

  if (row.cat === 'Validation') { vm.skipReason = row.reason || row.det || 'Not assessed.'; return; }

  if (!catOrder.includes(row.cat)) { catOrder.push(row.cat); itemsByCat[row.cat] = []; }
  if (!itemsByCat[row.cat].includes(row.item)) itemsByCat[row.cat].push(row.item);

  const key = row.cat + '||' + row.item;
  if (!dataByKey[key]) dataByKey[key] = { exp: row.exp, vms:{} };
  dataByKey[key].vms[row.vm] = { s: statusClass(row.status), label: row.status, d: row.det, r: row.reason };
});

Object.keys(VM_META).forEach(v=>{
  const cats = VM_META[v].cats;
  if (cats.size >= 3) VM_META[v].state = 'assessed';
  else if (cats.size === 1 && cats.has('Validation')) VM_META[v].state = 'notfound';
  else VM_META[v].state = 'skipped';
});
const ASSESSED = Object.keys(VM_META).filter(v => VM_META[v].state === 'assessed');

let activeStatuses = new Set(['good','bad','warn','mute']);
let searchTerm = '';

function computeFlags(){
  const flags = {}; ASSESSED.forEach(v => flags[v] = { good:0, bad:0, warn:0, mute:0 });
  Object.values(dataByKey).forEach(entry=>{
    ASSESSED.forEach(v=>{ const r = entry.vms[v]; if (r) flags[v][r.s]++; });
  });
  return flags;
}
const FLAGS = computeFlags();

function donutSvg(counts, total, size){
  const order = ['good','bad','warn','mute'];
  const r = 62, cx = 80, cy = 80, sw = 22;
  const circ = 2 * Math.PI * r;
  let acc = 0, arcs = '';
  order.forEach(k=>{
    const v = counts[k] || 0;
    if (!v) return;
    const len = circ * (v / total);
    arcs += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" style="stroke:var(--${k})" stroke-width="${sw}" stroke-dasharray="${len} ${circ-len}" stroke-dashoffset="${-acc}" transform="rotate(-90 ${cx} ${cy})"/>`;
    acc += len;
  });
  if (!total) arcs = `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" style="stroke:var(--border)" stroke-width="${sw}"/>`;
  return `<svg viewBox="0 0 160 160" width="${size}" height="${size}" role="img" aria-label="Compliance donut">
    ${arcs}
    <text x="80" y="76" text-anchor="middle" font-family="Archivo" font-size="${size>=180?28:16}" font-weight="800" style="fill:var(--ink)">${total}</text>
    <text x="80" y="${size>=180?96:92}" text-anchor="middle" font-family="IBM Plex Mono" font-size="${size>=180?11:8}" style="fill:var(--ink-faint)">points</text>
  </svg>`;
}

function computeItemTotals(){
  const t = { good:0, bad:0, warn:0, mute:0 };
  Object.values(dataByKey).forEach(entry=>{
    ASSESSED.forEach(v=>{ const r = entry.vms[v]; if (r) t[r.s]++; });
  });
  t.total = t.good+t.bad+t.warn+t.mute;
  return t;
}

function renderExecSummary(){
  const t = computeItemTotals();
  document.getElementById('execDonut').innerHTML = donutSvg(t, t.total, 200);
  document.getElementById('execLegend').innerHTML = `
    <div class="leg-row good"><span class="lbl"><span class="swatch good"></span>Compliant</span><span class="val">${t.good}</span></div>
    <div class="leg-row bad"><span class="lbl"><span class="swatch bad"></span>Non-Compliant</span><span class="val">${t.bad}</span></div>
    <div class="leg-row warn"><span class="lbl"><span class="swatch warn"></span>Manual Review</span><span class="val">${t.warn}</span></div>
    <div class="leg-row mute"><span class="lbl"><span class="swatch mute"></span>Unable to Check</span><span class="val">${t.mute}</span></div>`;

  const vEl = document.getElementById('verdictText'); const nEl = document.getElementById('verdictNote');
  if (t.total === 0) {
    vEl.textContent = 'No Data'; vEl.className = 'vtext warn';
    nEl.textContent = 'No VM in this run produced any checklist results.';
  } else if (t.bad > 0) {
    vEl.textContent = 'Not Ready for Acceptance'; vEl.className = 'vtext bad';
    nEl.textContent = `${t.bad} non-compliant finding(s) block acceptance until remediated or formally risk-accepted.`;
  } else if (t.warn > 0 || t.mute > 0) {
    vEl.textContent = 'Accepted with Conditions'; vEl.className = 'vtext warn';
    nEl.textContent = `${t.warn + t.mute} item(s) need a human sign-off before this fleet is fully closed out.`;
  } else {
    vEl.textContent = 'Accepted'; vEl.className = 'vtext good';
    nEl.textContent = 'No open findings across any assessed VM.';
  }
}

function renderKpis(){
  const skippedCount = Object.values(VM_META).filter(v => v.state !== 'assessed').length;
  const compliant = ASSESSED.filter(v => FLAGS[v].bad === 0).length;
  const nonCompliant = ASSESSED.filter(v => FLAGS[v].bad > 0).length;
  const manual = ASSESSED.filter(v => FLAGS[v].warn > 0).length;
  const unable = ASSESSED.filter(v => FLAGS[v].mute > 0).length;

  document.getElementById('mTotal').textContent = Object.keys(VM_META).length;
  document.getElementById('mAssessed').textContent = ASSESSED.length;
  document.getElementById('mSkipped').textContent = skippedCount;
  document.getElementById('vcenterList').textContent = RUN_META.vCenters || '-';
  document.getElementById('runMetaRight').innerHTML = `Run ${RUN_META.generatedAt}<br>${RUN_META.csvFile}`;

  document.getElementById('kpiStrip').innerHTML = `
    <div class="kpi total"><div class="kpi-label">VMs Assessed</div><div class="kpi-value">${ASSESSED.length}</div><div class="kpi-sub">of ${Object.keys(VM_META).length} targeted</div></div>
    <div class="kpi good"><div class="kpi-label">Compliant</div><div class="kpi-value">${compliant}</div><div class="kpi-sub">no gaps found</div></div>
    <div class="kpi bad"><div class="kpi-label">Non-Compliant</div><div class="kpi-value">${nonCompliant}</div><div class="kpi-sub">gaps found</div></div>
    <div class="kpi warn"><div class="kpi-label">Manual Review</div><div class="kpi-value">${manual}</div><div class="kpi-sub">need a human check</div></div>
    <div class="kpi mute"><div class="kpi-label">Not Assessed</div><div class="kpi-value">${skippedCount}</div><div class="kpi-sub">Tools/power/lookup failure</div></div>`;
}

function renderRoster(){
  const el = document.getElementById('roster');
  let html = '';
  ASSESSED.forEach(v=>{
    const bad = FLAGS[v].bad > 0;
    html += `<div class="roster-card"><span class="roster-dot" style="background:var(--${bad?'bad':'good'})"></span>
      <div><div class="roster-name">${v}</div><div class="roster-sub">${bad ? FLAGS[v].bad+' non-compliant' : 'compliant'} &middot; ${VM_META[v].ip||VM_META[v].os||''}</div></div></div>`;
  });
  Object.keys(VM_META).filter(v=>VM_META[v].state!=='assessed').forEach(v=>{
    html += `<div class="roster-card skip"><span class="roster-dot" style="background:var(--mute)"></span>
      <div><div class="roster-name">${v}</div><div class="roster-sub">not assessed</div></div></div>`;
  });
  el.innerHTML = html;
}

function renderNotAssessed(){
  const list = Object.entries(VM_META).filter(([,m]) => m.state !== 'assessed');
  if (list.length === 0) { document.getElementById('notAssessed').innerHTML = ''; return; }
  document.getElementById('notAssessed').innerHTML = `<div class="not-assessed">
    <b>${list.length} VM(s) not fully assessed:</b>
    ${list.map(([name,m])=>`<div class="row">${name} &mdash; <span class="mono">${m.skipReason}</span>${m.state==='skipped' ? ' (VMware Tools status recorded below, if reachable)' : ''}</div>`).join('')}
  </div>`;
}

function renderCatChart(){
  const el = document.getElementById('catChart');
  el.innerHTML = catOrder.map(cat=>{
    const c = { good:0, bad:0, warn:0, mute:0 }; let total = 0;
    itemsByCat[cat].forEach(item=>{
      const entry = dataByKey[cat+'||'+item];
      ASSESSED.forEach(v=>{ const r = entry.vms[v]; if (r) { c[r.s]++; total++; } });
    });
    const slug = cat.replace(/[^A-Za-z0-9]+/g,'-');
    return `
    <div class="donut-tile" onclick="document.getElementById('cat-${slug}').scrollIntoView({behavior:'smooth', block:'start'})">
      ${donutSvg(c, total, 108)}
      <div class="dname">${cat.replace(/^\d+\.\s*/,'')}</div>
      <div class="dcounts">${c.good}&#10003; ${c.bad}&#10005; ${c.warn}? ${c.mute}&#8213;</div>
    </div>`;
  }).join('');
}

function renderStatusChips(){
  const totals = { good:0, bad:0, warn:0, mute:0 };
  Object.values(dataByKey).forEach(entry=>{
    ASSESSED.forEach(v=>{ const r = entry.vms[v]; if (r) totals[r.s]++; });
  });
  const order = ['good','bad','warn','mute'];
  document.getElementById('statusChips').innerHTML = order.map(k => `
    <div class="chip ${activeStatuses.has(k)?'active':''}" onclick="toggleStatus('${k}')">
      <span class="swatch ${k}"></span>${STATUS_LABEL[k]} <span class="n">${totals[k]}</span>
    </div>`).join('');
}
function toggleStatus(k){
  if (activeStatuses.has(k)) { if (activeStatuses.size>1) activeStatuses.delete(k); } else activeStatuses.add(k);
  renderStatusChips(); renderMain();
}

function matchesSearch(hayParts){
  if (!searchTerm) return true;
  return hayParts.join(' ').toLowerCase().includes(searchTerm);
}

function renderSidenav(visibleCounts, issueCounts){
  document.getElementById('sidenav').innerHTML = catOrder.map(cat=>{
    const slug = cat.replace(/[^A-Za-z0-9]+/g,'-');
    return `<a href="#cat-${slug}">
      <span>${cat}</span>
      <span style="display:flex; gap:6px; align-items:center;">
        ${issueCounts[cat] ? `<span class="issue">${issueCounts[cat]}</span>` : ''}
        <span class="cnt">${visibleCounts[cat]||0}</span>
      </span>
    </a>`;
  }).join('');
}

function renderMain(){
  const visibleCounts = {}; const issueCounts = {};
  let bodyHtml = '';

  catOrder.forEach((cat, ci)=>{
    const slug = cat.replace(/[^A-Za-z0-9]+/g,'-');
    let sectionHtml = '';
    let sectionVisibleCount = 0;
    let sectionIssues = 0;

    itemsByCat[cat].forEach(item=>{
      const entry = dataByKey[cat+'||'+item];
      const rowsForItem = Object.entries(entry.vms).filter(([vm, r]) =>
        activeStatuses.has(r.s) && matchesSearch([cat, item, vm, r.d, r.r||'']));
      if (rowsForItem.length === 0) return;
      sectionVisibleCount++;

      const tally = { good:0, bad:0, warn:0, mute:0 };
      Object.values(entry.vms).forEach(r=>tally[r.s]++);
      if (tally.bad > 0) sectionIssues += tally.bad;

      const rowsHtml = rowsForItem.map(([vm, r])=>`
        <div class="vm-row">
          <div class="vm-who"><span class="status-dot" style="width:7px;height:7px;border-radius:50%;background:var(--${r.s})"></span><span class="name">${vm}</span></div>
          <div class="vm-detail">
            <span class="status-txt ${r.s}">${r.label}</span><span class="det">${r.d}</span>
            ${r.r ? `<div class="reason ${r.s==='bad'?'crit':''}">${r.r}</div>` : ''}
          </div>
        </div>`).join('');

      sectionHtml += `
      <div class="item-card">
        <div class="item-top">
          <div><div class="item-title">${item}</div>
            <div class="item-exp">Expected: <span class="mono">${entry.exp}</span></div></div>
          <div class="item-tally">
            ${tally.good?`<span class="tally good">${tally.good}</span>`:''}
            ${tally.bad?`<span class="tally bad">${tally.bad}</span>`:''}
            ${tally.warn?`<span class="tally warn">${tally.warn}</span>`:''}
            ${tally.mute?`<span class="tally mute">${tally.mute}</span>`:''}
          </div>
        </div>
        <div class="vm-scroll">${rowsHtml}</div>
      </div>`;
    });

    visibleCounts[cat] = sectionVisibleCount;
    issueCounts[cat] = sectionIssues;
    if (sectionVisibleCount === 0) return;

    bodyHtml += `
    <section class="cat-section" id="cat-${slug}">
      <div class="cat-heading"><span class="idx">${String(ci+1).padStart(2,'0')}</span>${cat}</div>
      ${sectionHtml}
    </section>`;
  });

  document.getElementById('main').innerHTML = bodyHtml || '<div class="empty-state">No items match the current filter, or no checklist rows were produced by this run.</div>';
  renderSidenav(visibleCounts, issueCounts);
}

document.getElementById('search').addEventListener('input', e=>{
  searchTerm = e.target.value.trim().toLowerCase();
  renderMain();
});

renderExecSummary();
renderKpis();
renderRoster();
renderNotAssessed();
renderCatChart();
renderStatusChips();
renderMain();
</script>
</body>
</html>
'@

$HtmlPath = Join-Path $OutputPath "OSAcceptance_$RunStamp.html"
$rowsForJson = @($centralRows | ForEach-Object {
    [ordered]@{
        vm = $_.VMName; vCenter = $_.vCenter; os = $_.OS; ip = $_.IPAddress
        cat = $_.Category; item = $_.Item; status = $_.Status
        det = $_.DetectedValue; exp = $_.ExpectedValue; reason = $_.'Error/Reason'
    }
})
if ($rowsForJson.Count -eq 0) {
    $rowsJson = '[]'
} elseif ($rowsForJson.Count -eq 1) {
    # ConvertTo-Json on a single-item array collapses to a bare object on PS 5.1 (no -AsArray there) - wrap by hand.
    $rowsJson = '[' + ($rowsForJson | ConvertTo-Json -Depth 6 -Compress) + ']'
} else {
    $rowsJson = $rowsForJson | ConvertTo-Json -Depth 6 -Compress
}
$rowsJson = $rowsJson.Replace('</script', '<\/script')

$runMetaObj = [ordered]@{
    generatedAt   = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    csvFile       = (Split-Path -Leaf $CsvPath)
    vCenters      = (($connectedServers | ForEach-Object { $_.Name }) -join ', ')
    totalTargeted = $vmNames.Count
}
$runMetaJson = ($runMetaObj | ConvertTo-Json -Compress).Replace('</script', '<\/script')

$htmlOut = $DashboardTemplate.Replace('{{ROWS_JSON}}', $rowsJson).Replace('{{RUN_META_JSON}}', $runMetaJson)
Set-Content -LiteralPath $HtmlPath -Value $htmlOut -Encoding UTF8
Write-Log "HTML dashboard written: $HtmlPath"

# 7. Compliance matrix workbook (VM rows x checklist-item columns)
# =====================================================================================
# One row per VM, one column per Category/Item the fleet actually produced results for -
# column order is first-seen order in $centralRows (already emitted in checklist order per
# VM), so there is no separate canonical column list to keep in sync. Cell = Status, color
# coded; the Detected/Expected/Reason detail that would otherwise need ~3x the columns is
# attached as a cell comment instead, so the sheet stays scannable at compliance-matrix width.
function New-ExcelColor {
    # Excel COM Interior.Color is an OLE_COLOR (0x00BBGGRR) - NOT plain RGB packing.
    param([int]$R, [int]$G, [int]$B)
    return ($B * 65536) + ($G * 256) + $R
}
$ColorGood = New-ExcelColor 198 239 206   # Excel's built-in "Good" green
$ColorBad  = New-ExcelColor 255 199 206   # Excel's built-in "Bad" red
$ColorWarn = New-ExcelColor 255 235 156   # Excel's built-in "Neutral" amber
$ColorMute = New-ExcelColor 217 217 217   # grey - Unable to Check / Not Applicable
$ColorHead = New-ExcelColor 232 236 236   # header fill

$xlsxOk   = $false
$XlsxPath = Join-Path $OutputPath "OSAcceptance_$RunStamp.xlsx"
try {
    $catOrder   = New-Object System.Collections.Generic.List[string]         # categories, in first-seen order
    $itemsByCat = @{}                                                        # category -> List[string] of items, in first-seen order within that category
    $cellData   = @{}                                                        # "VM||Category||Item" -> @{ Status; Detected; Expected; Reason }
    $vmOrder    = New-Object System.Collections.Generic.List[string]
    $vmMetaRow  = @{}

    foreach ($r in $centralRows) {
        if (-not $vmOrder.Contains($r.VMName)) {
            $vmOrder.Add($r.VMName) | Out-Null
            $vmMetaRow[$r.VMName] = @{ vCenter = $r.vCenter; Host = $r.GuestHostname; IP = $r.IPAddress; OS = $r.OS }
        } else {
            if (-not $vmMetaRow[$r.VMName].Host -and $r.GuestHostname) { $vmMetaRow[$r.VMName].Host = $r.GuestHostname }
            if (-not $vmMetaRow[$r.VMName].IP   -and $r.IPAddress)     { $vmMetaRow[$r.VMName].IP   = $r.IPAddress }
            if (-not $vmMetaRow[$r.VMName].OS   -and $r.OS)            { $vmMetaRow[$r.VMName].OS   = $r.OS }
        }
        if ($r.Category -eq 'Validation') { continue }
        # Two-level ordering (category first, then item within it) - not a single flat first-seen pass -
        # so a category's columns stay contiguous even when a later VM is the first to hit an item in a
        # category that an earlier VM had already partially populated (e.g. one VM alone triggers CPU
        # counters failing, or -SkipRolesInventory differs between runs).
        if (-not $catOrder.Contains($r.Category)) {
            $catOrder.Add($r.Category) | Out-Null
            $itemsByCat[$r.Category] = New-Object System.Collections.Generic.List[string]
        }
        if (-not $itemsByCat[$r.Category].Contains($r.Item)) { $itemsByCat[$r.Category].Add($r.Item) | Out-Null }
        $cellData["$($r.VMName)||$($r.Category)||$($r.Item)"] = @{ Status = $r.Status; Detected = $r.DetectedValue; Expected = $r.ExpectedValue; Reason = $r.'Error/Reason' }
    }

    $itemColumns     = New-Object System.Collections.Generic.List[string]    # "Category||Item" keys, in final column order
    $itemColumnLabel = @{}
    $itemColumnCat   = @{}
    foreach ($cat in $catOrder) {
        foreach ($item in $itemsByCat[$cat]) {
            $key = "$cat||$item"
            $itemColumns.Add($key) | Out-Null
            $itemColumnCat[$key]   = $cat
            $itemColumnLabel[$key] = $item
        }
    }

    $IdentityCols = @('VM Name', 'vCenter', 'Guest Hostname', 'IP Address', 'OS')
    $XlsxPath     = Join-Path $OutputPath "OSAcceptance_$RunStamp.xlsx"
    $xlsxOk       = $false
    $excel = $null; $wb = $null; $ws = $null
    try {
        $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'OS Acceptance'

        for ($i = 0; $i -lt $IdentityCols.Count; $i++) {
            $c = $i + 1
            $ws.Cells.Item(2, $c) = $IdentityCols[$i]
            $ws.Range($ws.Cells.Item(1, $c), $ws.Cells.Item(2, $c)).Merge() | Out-Null
        }
        $firstItemCol = $IdentityCols.Count + 1

        $col = $firstItemCol
        $prevCat = $null
        $catStartCol = $col
        foreach ($key in $itemColumns) {
            $cat = $itemColumnCat[$key]
            if ($cat -ne $prevCat) {
                if ($prevCat -and ($col - 1) -ge $catStartCol) {
                    $ws.Range($ws.Cells.Item(1, $catStartCol), $ws.Cells.Item(1, $col - 1)).Merge() | Out-Null
                }
                $ws.Cells.Item(1, $col) = $cat
                $catStartCol = $col
                $prevCat = $cat
            }
            $ws.Cells.Item(2, $col) = $itemColumnLabel[$key]
            $col++
        }
        if ($prevCat -and ($col - 1) -ge $catStartCol) {
            $ws.Range($ws.Cells.Item(1, $catStartCol), $ws.Cells.Item(1, $col - 1)).Merge() | Out-Null
        }
        $lastCol = $col - 1

        $headerRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(2, $lastCol))
        $headerRange.Font.Bold = $true
        $headerRange.Interior.Color = $ColorHead
        $headerRange.WrapText = $true
        $headerRange.VerticalAlignment = -4108   # xlVAlignCenter / xlHAlignCenter share this value
        $headerRange.HorizontalAlignment = -4108
        $ws.Rows.Item(1).RowHeight = 20
        $ws.Rows.Item(2).RowHeight = 46

        $rowIdx = 3
        foreach ($vm in $vmOrder) {
            $meta = $vmMetaRow[$vm]
            $ws.Cells.Item($rowIdx, 1) = $vm
            $ws.Cells.Item($rowIdx, 2) = $meta.vCenter
            $ws.Cells.Item($rowIdx, 3) = $meta.Host
            $ws.Cells.Item($rowIdx, 4) = $meta.IP
            $ws.Cells.Item($rowIdx, 5) = $meta.OS

            $c = $firstItemCol
            foreach ($key in $itemColumns) {
                $cd = $cellData["$vm||$key"]
                if ($cd) {
                    $cell = $ws.Cells.Item($rowIdx, $c)
                    # Compliant/Non-Compliant: color + the pass/fail word is enough at a glance.
                    # Manual Verification Required/Unable to Check: the whole point is that a human
                    # has to read the actual value to make the call, so put it in the cell itself
                    # instead of behind a hover comment.
                    $cell.Value2 = if ($cd.Status -in 'Manual Verification Required', 'Unable to Check') {
                        if ($cd.Detected) { $cd.Detected } else { $cd.Status }
                    } else { $cd.Status }
                    $cell.Interior.Color = switch ($cd.Status) {
                        'Compliant' { $ColorGood }
                        'Non-Compliant' { $ColorBad }
                        'Manual Verification Required' { $ColorWarn }
                        default { $ColorMute }
                    }
                    $note = "Status: $($cd.Status)`nDetected: $($cd.Detected)`nExpected: $($cd.Expected)"
                    if ($cd.Reason) { $note += "`nReason: $($cd.Reason)" }
                    [void]$cell.AddComment($note)
                    $cell.Comment.Shape.TextFrame.AutoSize = $true
                }
                $c++
            }
            $rowIdx++
        }

        $ws.Columns.Item(1).ColumnWidth = 22
        $ws.Columns.Item(2).ColumnWidth = 20
        $ws.Columns.Item(3).ColumnWidth = 24
        $ws.Columns.Item(4).ColumnWidth = 16
        $ws.Columns.Item(5).ColumnWidth = 26
        for ($c = $firstItemCol; $c -le $lastCol; $c++) { $ws.Columns.Item($c).ColumnWidth = 15 }

        $ws.Activate()
        $ws.Cells.Item(3, $firstItemCol).Select() | Out-Null
        $excel.ActiveWindow.FreezePanes = $true

        $wb.SaveAs($XlsxPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        $excel.Quit()
        $xlsxOk = $true
        Write-Log "Compliance matrix workbook written: $XlsxPath ($($vmOrder.Count) VM rows x $($itemColumns.Count) item columns)."
    } catch {
        Write-Log "Excel COM automation failed - falling back to a wide-format CSV instead of .xlsx: $($_.Exception.Message)" 'WARN'
        try { if ($wb) { $wb.Close($false) } } catch { }
        try { if ($excel) { $excel.Quit() } } catch { }
    } finally {
        foreach ($comObj in @($ws, $wb, $excel)) {
            if ($comObj) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObj) }
        }
        Remove-Variable excel, wb, ws -ErrorAction SilentlyContinue
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if (-not $xlsxOk) {
        $XlsxPath = Join-Path $OutputPath "OSAcceptance_Matrix_$RunStamp.csv"
        $wideRows = foreach ($vm in $vmOrder) {
            $meta = $vmMetaRow[$vm]
            $obj = [ordered]@{ 'VM Name' = $vm; vCenter = $meta.vCenter; 'Guest Hostname' = $meta.Host; 'IP Address' = $meta.IP; OS = $meta.OS }
            foreach ($key in $itemColumns) {
                $cd = $cellData["$vm||$key"]
                $obj["$($itemColumnCat[$key]) - $($itemColumnLabel[$key])"] =
                    if (-not $cd) { '' }
                    elseif ($cd.Status -in 'Manual Verification Required', 'Unable to Check') { if ($cd.Detected) { "$($cd.Status): $($cd.Detected)" } else { $cd.Status } }
                    else { $cd.Status }
            }
            [PSCustomObject]$obj
        }
        $wideRows | Export-Csv -LiteralPath $XlsxPath -NoTypeInformation -Encoding UTF8
        Write-Log "Wide-format compliance matrix CSV written instead: $XlsxPath"
    }
} catch {
    # Guarantees the run still reaches the CSV export below even if compliance-matrix
    # generation (Excel COM or its CSV fallback) fails for a reason the inner try/catch
    # did not anticipate - a single bad row should never cost the whole report.
    Write-Log "Compliance matrix generation failed entirely - neither .xlsx nor a fallback CSV was written: $($_.Exception.Message)" 'ERROR'
}

# =====================================================================================
# 8. Central report + console summary
# =====================================================================================
$centralRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV written: $CsvPath ($($centralRows.Count) rows)."

Write-Host ""
Write-Host "============= OS ACCEPTANCE SUMMARY =============" -ForegroundColor Green
Write-Host ("  Total VMs                     : {0}" -f $summary.Total)
Write-Host ("  Successfully Checked          : {0}" -f $summary.Checked)
Write-Host ("  Compliant (no gaps found)     : {0}" -f $summary.Compliant)
Write-Host ("  Non-Compliant (gaps found)    : {0}" -f $summary.NonCompliant) -ForegroundColor $(if ($summary.NonCompliant) { 'Yellow' } else { 'Gray' })
Write-Host ("  Manual Verification Required  : {0}" -f $summary.Manual) -ForegroundColor DarkYellow
Write-Host ("  Unable to Check (some items)  : {0}" -f $summary.Unable) -ForegroundColor $(if ($summary.Unable) { 'Yellow' } else { 'Gray' })
Write-Host ("  Skipped / Failed to process   : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "===================================================" -ForegroundColor Green
Write-Host ""
Write-Host "CSV       : $CsvPath"
Write-Host "Dashboard : $HtmlPath"
Write-Host "$(if ($xlsxOk) { 'Workbook ' } else { 'Matrix   ' }) : $XlsxPath"
Write-Host "Log       : $LogPath"
Write-Log "OS Acceptance run complete."
