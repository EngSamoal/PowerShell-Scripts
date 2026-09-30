#Requires -Version 5.1
<#
.SYNOPSIS
    Centralised (PowerCLI) front-end for post-remediation-check.ps1.

    Strictly READ-ONLY verification of the same controls the local
    post-remediation-check.ps1 verifies, run against every VM in a list from a
    single admin workstation over VMware Tools Guest Operations. No WinRM / PsExec /
    SMB / RDP / guest network path is used.

.DESCRIPTION
    The in-guest payload contains ONLY Get-* / registry reads. It never calls
    Set-*, New-*, Remove-*, Stop-*, Disable-* or Add-*. It changes nothing.

    Per VM it returns, for every checked item, one of:
        Compliant
        Non-Compliant
        Manual Verification Required
        Unable to Check
    together with the actual detected value and the expected value, so the CSV is
    auditable.

    Same per-VM validation battery as the remediation wrapper (existence,
    non-ambiguity across connected vCenters, powered on, Windows, VMware Tools
    running, guest auth, PowerShell 5.1+). A validation failure skips only that VM
    and is recorded with an exact reason; one bad VM never stops the batch.

    Notes on two checks that were weaker in the original local script (behaviour
    preserved, but the detected value now states the gap):
      * "ASR Rules Configured" only confirms that *some* ASR rules exist, not that
        the full recommended set is at the intended action. The detected value
        reports the rule count so a reviewer can see it.
      * "Legacy SSL/TLS Disabled" only inspects TLS 1.0 (Server). The detected value
        says so; a full check is out of scope for a read-only verifier and is an
        opt-in remediation item anyway.

.PARAMETER VMListPath      Text file of VM names (default C:\temp\vmlist.txt). '#' lines ignored.
.PARAMETER CredentialPath  Export-Clixml PSCredential for the guest admin (default C:\temp\wincred.xml). In-memory only.
.PARAMETER OutputPath      Folder on the admin machine for the CSV + log.
.PARAMETER ToolsWaitSecs   Invoke-VMScript VMware Tools wait, seconds (default 180).

.EXAMPLE
    .\Invoke-RemotePostRemediationCheck.ps1
    .\Invoke-RemotePostRemediationCheck.ps1 -VMListPath C:\temp\vmlist.txt -OutputPath D:\Reports

.NOTES
    Prerequisites: PowerCLI imported; already Connect-VIServer'd to the target
    vCenter(s); vCenter account has the three "Guest Operation ..." privileges.
    This script issues no vSphere write of any kind.
#>

[CmdletBinding()]
param(
    [string] $VMListPath     = 'C:\temp\vmlist.txt',
    [string] $CredentialPath = 'C:\temp\wincred.xml',
    [string] $OutputPath     = (Join-Path $PSScriptRoot 'SecurityGap_Reports'),

    # Force every target to be treated as a workgroup / non-domain server. Controls
    # that only apply to domain members (Machine Identity Isolation, Kerberos
    # encryption types) are then reported "Not Applicable" with a workgroup hint,
    # without relying on in-guest domain auto-detection.
    [switch] $TreatAllAsWorkgroup,

    # Category 7 (Extended Baseline): value at or below which Winlogon\CachedLogonsCount
    # is treated Compliant. Default 4 (CIS). Qualys QID 90007 only fully clears at 0 on
    # a workgroup server - match whatever -CachedLogonsCount you passed to the
    # remediation script.
    [ValidateRange(0, 50)]
    [int]    $CachedLogonsTarget = 4,

    # Skip the whole "7. Extended Baseline (Cybersecurity Review)" verification block.
    [switch] $SkipExtendedBaseline,
    # --- ISO / Qualys Policy Compliance (iso.xlsx) verification - category 8 ---
    [switch] $SkipIsoPolicy,
    [switch] $SkipIeHardening,

    [ValidateRange(30, 3600)]
    [int]    $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'

Write-Host "Invoke-RemotePostRemediationCheck.ps1 - READ-ONLY verification  [build 2026-09-04b-postcheck-iso-pc]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
Write-Host ("If the build above is not '2026-09-04b-postcheck-iso-pc' you are running an OLD copy - update it.") -ForegroundColor DarkGray

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
$CsvPath  = Join-Path $OutputPath "SecurityGap_PostCheck_$RunStamp.csv"
$LogPath  = Join-Path $OutputPath "SecurityGap_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Post-check run started. VMs=$($vmNames.Count)."

# =====================================================================================
# 2. In-guest READ-ONLY payload (mirrors post-remediation-check.ps1)
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'
$CachedLogonsTarget   = {{CACHED_TARGET}}
$SkipExtendedBaseline = {{SKIP_EXTENDED_BASELINE}}
$SkipIsoPolicy        = {{SKIP_ISO}}
$SkipIeHardening      = {{SKIP_IE}}

# Each row carries BOTH representations:
#   Header / Label / LocalStatus / Detail  -> reproduce the local post-remediation-check.ps1
#                                             on-screen output exactly (Show-Result style)
#   Category / Item / Status / Detected / Expected -> the auditable central CSV
$Rows = New-Object System.Collections.Generic.List[object]
function Add-Row {
    param(
        [string]$Header, [string]$Label, [string]$LocalStatus, [string]$Detail = '',
        [string]$Category, [string]$Item, [string]$NormStatus = '',
        [string]$Detected = '', [string]$Expected = ''
    )
    if (-not $Item) { $Item = $Label }
    if (-not $NormStatus) {
        if     ($LocalStatus -like 'OK*')     { $NormStatus = 'Compliant' }
        elseif ($LocalStatus -like 'MANUAL*') { $NormStatus = 'Manual Verification Required' }
        elseif ($LocalStatus -like 'N/A*')    { $NormStatus = 'Not Applicable' }
        else                                  { $NormStatus = 'Non-Compliant' }
    }
    $Rows.Add([PSCustomObject]@{
        Header = $Header; Label = $Label; LocalStatus = $LocalStatus; Detail = $Detail
        Category = $Category; Item = $Item; Status = $NormStatus
        DetectedValue = $Detected; ExpectedValue = $Expected
    }) | Out-Null
}
# Read a single registry value; returns [pscustomobject]@{ Found; Value } - never throws.
function Get-RegVal {
    param([string]$Path, [string]$Name)
    try {
        $ip = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return [pscustomobject]@{ Found = $true; Value = $ip.$Name }
    } catch {
        return [pscustomobject]@{ Found = $false; Value = $null }
    }
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

# --- Workgroup / non-domain detection ---------------------------------------
# Used to classify domain-only controls (Machine Identity Isolation, Kerberos
# encryption types) as "Not Applicable" on workgroup servers. Two independent
# signals: PartOfDomain=$false OR DomainRole in {0,2} (Standalone Workstation /
# Standalone Server). The admin -TreatAllAsWorkgroup switch forces it on.
$TreatAllAsWorkgroup = {{TREAT_ALL_AS_WORKGROUP}}
$IsWorkgroup   = $false
$WorkgroupWhy  = ''
if ($TreatAllAsWorkgroup) {
    $IsWorkgroup  = $true
    $WorkgroupWhy = 'operator passed -TreatAllAsWorkgroup'
} else {
    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        $partOfDomain = [bool]$cs.PartOfDomain
        $domainRole   = [int]$cs.DomainRole
        if ((-not $partOfDomain) -or ($domainRole -in 0, 2)) {
            $IsWorkgroup  = $true
            $WorkgroupWhy = "PartOfDomain=$partOfDomain; DomainRole=$domainRole"
        }
    } catch {
        # Could not determine - do NOT assume workgroup (would hide real gaps on a
        # domain box). Operator can force with -TreatAllAsWorkgroup.
        $IsWorkgroup  = $false
        $WorkgroupWhy = "domain membership query failed: $($_.Exception.Message)"
    }
}

$fatal = $null
try {
    # ==========================================================================
    $H = 'VBS & Credential Protection - Enable Credential Guard, HVCI/VBS, Machine Identity Isolation'
    $v = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' 'LsaCfgFlags'
    Add-Row -Header $H -Label 'Credential Guard' `
        -LocalStatus $(if ($v.Found -and $v.Value -in @(1, 2)) { 'OK' } else { 'NOT SET' }) `
        -Category 'VBS & Credential Protection' -Item 'Credential Guard (LsaCfgFlags)' `
        -Detected ("LsaCfgFlags={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '1 or 2'

    $vbs  = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity'
    $hvci = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' 'Enabled'
    $hvciVbsOk = ($vbs.Found -and $vbs.Value -eq 1) -and ($hvci.Found -and $hvci.Value -eq 1)
    Add-Row -Header $H -Label 'HVCI/VBS' `
        -LocalStatus $(if ($hvciVbsOk) { 'OK' } else { 'NOT SET' }) `
        -Category 'VBS & Credential Protection' -Item 'HVCI / VBS' `
        -Detected ("EnableVirtualizationBasedSecurity={0}; HVCI.Enabled={1}" -f $(if ($vbs.Found) { $vbs.Value } else { '(not set)' }), $(if ($hvci.Found) { $hvci.Value } else { '(not set)' })) `
        -Expected 'VBS=1 and HVCI.Enabled=1'

    $v = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' 'MachineIdentityIsolation'
    if ($IsWorkgroup) {
        Add-Row -Header $H -Label 'Machine Identity Isolation' -LocalStatus 'N/A' `
            -Detail "NOT APPLICABLE - this server is in a workgroup (not domain-joined). Machine Identity Isolation protects the Active Directory computer-account credential; a workgroup server has none, so the policy is a no-op that does not persist. [$WorkgroupWhy]" `
            -Category 'VBS & Credential Protection' -Item 'Machine Identity Isolation' -NormStatus 'Not Applicable' `
            -Detected ("Workgroup ($WorkgroupWhy); MachineIdentityIsolation={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected 'N/A unless domain-joined'
    } else {
        Add-Row -Header $H -Label 'Machine Identity Isolation' `
            -LocalStatus $(if ($v.Found -and $v.Value -eq 2) { 'OK' } else { "NOT SET (found: $(if ($v.Found) { $v.Value } else { '' }))" }) `
            -Category 'VBS & Credential Protection' -Item 'Machine Identity Isolation' `
            -Detected ("MachineIdentityIsolation={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '2 (Enforcement Mode)'
    }

    # ==========================================================================
    $H = 'Microsoft Defender - Enable ASR enforcement and Defender reporting controls'
    try {
        $status = Get-MpComputerStatus -ErrorAction Stop
        Add-Row -Header $H -Label 'Real-Time Protection' `
            -LocalStatus $(if ($status.RealTimeProtectionEnabled) { 'OK' } else { 'NOT SET' }) `
            -Category 'Microsoft Defender' -Item 'Real-Time Protection' `
            -Detected ("RealTimeProtectionEnabled={0}" -f $status.RealTimeProtectionEnabled) -Expected 'True'
    } catch {
        Add-Row -Header $H -Label 'Real-Time Protection' -LocalStatus 'NOT SET' -Detail 'Defender not available' `
            -Category 'Microsoft Defender' -Item 'Real-Time Protection' -NormStatus 'Unable to Check' `
            -Detected 'Get-MpComputerStatus failed' -Expected 'True'
    }
    try {
        $pref = Get-MpPreference -ErrorAction Stop
        Add-Row -Header $H -Label 'Cloud/MAPS Reporting' `
            -LocalStatus $(if ($pref.MAPSReporting -eq 2) { 'OK' } else { 'NOT SET' }) `
            -Category 'Microsoft Defender' -Item 'Cloud / MAPS Reporting' `
            -Detected ("MAPSReporting={0}" -f $pref.MAPSReporting) -Expected '2 (Advanced)'
        $asrCount = @($pref.AttackSurfaceReductionRules_Ids).Count
        Add-Row -Header $H -Label 'ASR Rules Configured' `
            -LocalStatus $(if ($asrCount -gt 0) { "OK ($asrCount rules)" } else { 'NOT SET' }) `
            -Category 'Microsoft Defender' -Item 'ASR Rules Configured' `
            -Detected ("$asrCount rule(s) configured") `
            -Expected '>= 1 rule (original check); recommended set = 16 rules at Enabled'
    } catch {
        Add-Row -Header $H -Label 'Defender Preferences' -LocalStatus 'NOT SET' -Detail 'Defender not available' `
            -Category 'Microsoft Defender' -Item 'Defender Preferences (MAPS / ASR)' -NormStatus 'Unable to Check' `
            -Detected 'Get-MpPreference failed' -Expected ''
    }
    Add-Row -Header $H -Label 'Tamper Protection' -LocalStatus 'MANUAL' `
        -Detail 'Cannot be checked/set by script - verify in Windows Security app or Intune' `
        -Category 'Microsoft Defender' -Item 'Tamper Protection' `
        -Detected 'Not scriptable' -Expected 'Enabled'

    # ==========================================================================
    $H = 'Authentication Security - Disable WDigest; enable Kerberos SHA256/SHA384/SHA512'
    $v = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential'
    Add-Row -Header $H -Label 'WDigest Disabled' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
        -Category 'Authentication Security' -Item 'WDigest Disabled' `
        -Detected ("UseLogonCredential={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '0'

    $v = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' 'SupportedEncryptionTypes'
    if ($IsWorkgroup) {
        Add-Row -Header $H -Label 'Kerberos AES128/256 Supported' -LocalStatus 'N/A' `
            -Detail "NOT APPLICABLE - this server is in a workgroup. There is no Active Directory domain or KDC; authentication is NTLM / local accounts only, so Kerberos is not used. [$WorkgroupWhy]" `
            -Category 'Authentication Security' -Item 'Kerberos AES128/256 Supported' -NormStatus 'Not Applicable' `
            -Detected ("Workgroup ($WorkgroupWhy); SupportedEncryptionTypes={0}" -f $(if ($v.Found) { ('0x{0:X}' -f [int]$v.Value) } else { '(not set)' })) `
            -Expected 'N/A unless domain-joined'
        Add-Row -Header $H -Label 'Kerberos SHA256/SHA384/SHA512' -LocalStatus 'N/A' `
            -Detail "NOT APPLICABLE - this server is in a workgroup; Kerberos (and its AES-SHA2 cipher suites) is only used by domain members. [$WorkgroupWhy]" `
            -Category 'Authentication Security' -Item 'Kerberos SHA256/SHA384/SHA512 (AES-SHA2 / RFC 8009)' -NormStatus 'Not Applicable' `
            -Detected "Workgroup ($WorkgroupWhy)" -Expected 'N/A unless domain-joined'
    } else {
        $aesOk = $v.Found -and (([int]$v.Value) -band 0x18) -eq 0x18
        Add-Row -Header $H -Label 'Kerberos AES128/256 Supported' `
            -LocalStatus $(if ($aesOk) { 'OK' } else { 'NOT SET' }) `
            -Category 'Authentication Security' -Item 'Kerberos AES128/256 Supported' `
            -Detected ("SupportedEncryptionTypes={0}" -f $(if ($v.Found) { ('0x{0:X} ({0})' -f [int]$v.Value) } else { '(not set)' })) `
            -Expected 'bits 0x18 (AES128+AES256) set'

        Add-Row -Header $H -Label 'Kerberos SHA256/SHA384/SHA512' -LocalStatus 'MANUAL' `
            -Detail 'Not verifiable by script - requires manual check/configuration (no confirmed Windows mechanism yet, see Microsoft guidance)' `
            -Category 'Authentication Security' -Item 'Kerberos SHA256/SHA384/SHA512 (AES-SHA2 / RFC 8009)' `
            -Detected 'Not verifiable by registry check' -Expected 'AES-SHA2 suites (etype 19/20) negotiated'
    }

    # ==========================================================================
    $H = 'Remote Management - Harden or disable WinRM; restrict listener filters'
    $winrm = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service'
    $u = Get-RegVal $winrm 'AllowUnencrypted'
    Add-Row -Header $H -Label 'WinRM Unencrypted Blocked' `
        -LocalStatus $(if ($u.Found -and $u.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
        -Category 'Remote Management' -Item 'WinRM Unencrypted Traffic Blocked' `
        -Detected ("AllowUnencrypted={0}" -f $(if ($u.Found) { $u.Value } else { '(not set)' })) -Expected '0'
    $b = Get-RegVal $winrm 'auth_Basic'
    Add-Row -Header $H -Label 'WinRM Basic Auth Blocked' `
        -LocalStatus $(if ($b.Found -and $b.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
        -Category 'Remote Management' -Item 'WinRM Basic Auth Blocked' `
        -Detected ("auth_Basic={0}" -f $(if ($b.Found) { $b.Value } else { '(not set)' })) -Expected '0'
    try {
        $rules = Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' }
        $addr = @($rules | Get-NetFirewallAddressFilter -ErrorAction Stop | Select-Object -ExpandProperty RemoteAddress -Unique)
        if ($addr.Count -eq 0) {
            Add-Row -Header $H -Label 'WinRM Listener Restricted' -LocalStatus 'MANUAL' -Detail 'No enabled WinRM firewall rules found - verify manually' `
                -Category 'Remote Management' -Item 'WinRM Listener Restricted to Management Sources' `
                -Detected 'No enabled WinRM firewall rules' -Expected 'Scoped to management subnets/hosts'
        } elseif ($addr.Count -eq 1 -and $addr[0] -eq 'Any') {
            Add-Row -Header $H -Label 'WinRM Listener Restricted' -LocalStatus 'MANUAL' -Detail 'Currently open to Any source - supply -WinRmAllowedSourceRange and re-apply' `
                -Category 'Remote Management' -Item 'WinRM Listener Restricted to Management Sources' `
                -Detected 'RemoteAddress=Any' -Expected 'Scoped to management subnets/hosts'
        } else {
            Add-Row -Header $H -Label 'WinRM Listener Restricted' -LocalStatus 'OK' -Detail ("Restricted to: {0}" -f ($addr -join ', ')) `
                -Category 'Remote Management' -Item 'WinRM Listener Restricted to Management Sources' `
                -Detected ("RemoteAddress={0}" -f ($addr -join ', ')) -Expected 'Scoped to management subnets/hosts'
        }
    } catch {
        Add-Row -Header $H -Label 'WinRM Listener Restricted' -LocalStatus 'MANUAL' -Detail 'Could not verify - check firewall rules manually' `
            -Category 'Remote Management' -Item 'WinRM Listener Restricted to Management Sources' -NormStatus 'Unable to Check' `
            -Detected 'Get-NetFirewallRule failed' -Expected 'Scoped to management subnets/hosts'
    }

    # ==========================================================================
    $H = 'SMB & RPC Security - Enable SMB auditing; configure secure Printer RPC settings'
    try {
        $smb = Get-SmbServerConfiguration -ErrorAction Stop
        if ($smb.PSObject.Properties.Name -notcontains 'AuditSmb1Access') {
            Add-Row -Header $H -Label 'SMB1 Access Auditing' -LocalStatus 'N/A' `
                -Detail 'NOT APPLICABLE - the AuditSmb1Access setting requires Windows Server 2022 or later; this OS build does not expose it.' `
                -Category 'SMB & RPC Security' -Item 'SMB1 Access Auditing' -NormStatus 'Not Applicable' `
                -Detected 'AuditSmb1Access property not present on this OS build' -Expected 'N/A on builds older than Server 2022'
        } else {
            Add-Row -Header $H -Label 'SMB1 Access Auditing' `
                -LocalStatus $(if ($smb.AuditSmb1Access) { 'OK' } else { 'NOT SET' }) `
                -Category 'SMB & RPC Security' -Item 'SMB1 Access Auditing' `
                -Detected ("AuditSmb1Access={0}" -f $smb.AuditSmb1Access) -Expected 'True'
        }
    } catch {
        Add-Row -Header $H -Label 'SMB1 Access Auditing' -LocalStatus 'NOT SET' -Detail 'Could not read SMB config' `
            -Category 'SMB & RPC Security' -Item 'SMB1 Access Auditing' -NormStatus 'Unable to Check' `
            -Detected 'Get-SmbServerConfiguration failed' -Expected 'True'
    }
    $v = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' 'RpcAuthnLevelPrivacyEnabled'
    Add-Row -Header $H -Label 'Printer RPC Privacy' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
        -Category 'SMB & RPC Security' -Item 'Printer RPC Packet Privacy' `
        -Detected ("RpcAuthnLevelPrivacyEnabled={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '1'
    $v = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' 'RestrictDriverInstallationToAdministrators'
    Add-Row -Header $H -Label 'Point-and-Print Restricted' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
        -Category 'SMB & RPC Security' -Item 'Point and Print Restricted' `
        -Detected ("RestrictDriverInstallationToAdministrators={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '1'

    # ==========================================================================
    $H = 'Security Baseline - Apply Microsoft Windows Server 2025 Security Baseline GPO (IE, SmartScreen, ActiveX, SSL/TLS, scripting, Protected Mode)'
    $v = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' 'EnableSmartScreen'
    Add-Row -Header $H -Label 'SmartScreen Enabled' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
        -Category '2025 Security Baseline' -Item 'SmartScreen Enabled' `
        -Detected ("EnableSmartScreen={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '1'

    $z = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'
    $z1200 = Get-RegVal $z '1200'; $z1400 = Get-RegVal $z '1400'; $z2500 = Get-RegVal $z '2500'
    $ieOk = ($z1200.Found -and $z1200.Value -eq 3) -and ($z1400.Found -and $z1400.Value -eq 3) -and ($z2500.Found -and $z2500.Value -eq 0)
    Add-Row -Header $H -Label 'IE Zone (ActiveX/Scripting/Protected Mode)' `
        -LocalStatus $(if ($ieOk) { 'OK' } else { 'NOT SET' }) `
        -Category '2025 Security Baseline' -Item 'IE Zone (ActiveX / Scripting / Protected Mode)' `
        -Detected ("1200={0}; 1400={1}; 2500={2}" -f $(if ($z1200.Found) { $z1200.Value } else { '(ns)' }), $(if ($z1400.Found) { $z1400.Value } else { '(ns)' }), $(if ($z2500.Found) { $z2500.Value } else { '(ns)' })) `
        -Expected '1200=3, 1400=3, 2500=0'

    try {
        $feat = Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction Stop
        if (-not $feat) {
            Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'N/A' -Detail 'NOT APPLICABLE - the Internet Explorer optional feature does not exist on this Windows Server build; there is nothing to remove.' `
                -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -NormStatus 'Not Applicable' -Detected 'Feature not present on this build' -Expected 'N/A - feature absent'
        } elseif ($feat.State -eq 'Disabled') {
            Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'OK' `
                -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -Detected 'State=Disabled' -Expected 'Disabled / not present'
        } else {
            Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'NOT SET' `
                -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -Detected ("State={0}" -f $feat.State) -Expected 'Disabled / not present'
        }
    } catch {
        Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'N/A' -Detail 'NOT APPLICABLE - the Internet Explorer optional feature does not exist on this Windows Server build; there is nothing to remove.' `
            -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -NormStatus 'Not Applicable' -Detected 'Feature does not exist on this build' -Expected 'N/A - feature absent'
    }

    $v = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' 'Enabled'
    Add-Row -Header $H -Label 'Legacy SSL/TLS Disabled' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 0) { 'OK' } else { 'MANUAL' }) `
        -Detail 'Opt-in only (-DisableLegacyTls) - confirm intentionally' `
        -Category '2025 Security Baseline' -Item 'Legacy SSL/TLS Disabled (spot check: TLS 1.0 Server)' `
        -Detected ("TLS 1.0\Server\Enabled={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected 'Enabled=0 (opt-in remediation)'

    Add-Row -Header $H -Label 'Full Baseline GPO Applied' -LocalStatus 'MANUAL' `
        -Detail 'Not verifiable by registry check - confirm via gpresult /h or re-run LGPO.exe /g' `
        -Category '2025 Security Baseline' -Item 'Full Windows Server 2025 Security Baseline GPO Applied' `
        -Detected 'Not verifiable by registry check' -Expected 'Baseline GPO linked / applied'

    # ==========================================================================
    # 7. Extended Baseline (Cybersecurity Review) - verifies the same items the
    #    remediation script's Category 7 remediates. READ-ONLY: Get-* / registry
    #    reads / "auditpol /get" / Get-LocalUser only.
    # ==========================================================================
    if (-not $SkipExtendedBaseline) {
        $CC = '7. Extended Baseline (Cybersecurity Review)'

        $H = '7. Extended Baseline (Cybersecurity Review) - VBS / LSA / logging'
        # HVCI - Enabled with UEFI lock
        $p = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'
        $en = Get-RegVal $p 'Enabled'; $lk = Get-RegVal $p 'Locked'
        Add-Row -Header $H -Label 'HVCI - Enabled with UEFI lock' `
            -LocalStatus $(if ($en.Found -and $en.Value -eq 1 -and $lk.Found -and $lk.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
            -Category $CC -Item 'HVCI - UEFI lock' `
            -Detected ("Enabled={0}; Locked={1}" -f $(if ($en.Found) { $en.Value } else { '(ns)' }), $(if ($lk.Found) { $lk.Value } else { '(ns)' })) `
            -Expected 'Enabled=1 and Locked=1'

        # Kernel-mode HW-enforced Stack Protection - enforcement (not Audit)
        $p = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\KernelShadowStacks'
        $en = Get-RegVal $p 'Enabled'; $am = Get-RegVal $p 'AuditMode'
        Add-Row -Header $H -Label 'Kernel-mode HW Stack Protection - enforcement' `
            -LocalStatus $(if ($en.Found -and $en.Value -eq 1 -and $am.Found -and $am.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
            -Detail $(if ($en.Found -and $en.Value -eq 1 -and $am.Found -and $am.Value -eq 1) { 'currently in Audit mode - control expects Enforcement' } else { '' }) `
            -Category $CC -Item 'Kernel-mode HW-enforced Stack Protection' `
            -Detected ("Enabled={0}; AuditMode={1}" -f $(if ($en.Found) { $en.Value } else { '(ns)' }), $(if ($am.Found) { $am.Value } else { '(ns)' })) `
            -Expected 'Enabled=1 and AuditMode=0'

        # LSASS RunAsPPL - UEFI lock (value 1, not 2)
        $v = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' 'RunAsPPL'
        Add-Row -Header $H -Label 'LSASS RunAsPPL - UEFI lock' `
            -LocalStatus $(if ($v.Found -and $v.Value -eq 1) { 'OK' } else { "NOT SET (found: $(if ($v.Found) { $v.Value } else { '' }))" }) `
            -Detail $(if ($v.Found -and $v.Value -eq 2) { 'RunAsPPL=2 = enabled WITHOUT UEFI lock; control expects 1 (with lock)' } else { '' }) `
            -Category $CC -Item 'LSASS RunAsPPL - UEFI lock' `
            -Detected ("RunAsPPL={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected 'RunAsPPL=1'

        # PowerShell Script Block (+ Invocation) Logging
        $p = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'
        $sb = Get-RegVal $p 'EnableScriptBlockLogging'; $si = Get-RegVal $p 'EnableScriptBlockInvocationLogging'
        Add-Row -Header $H -Label 'Script Block + Invocation Logging' `
            -LocalStatus $(if ($sb.Found -and $sb.Value -eq 1 -and $si.Found -and $si.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
            -Category $CC -Item 'PowerShell Script Block (and Invocation) Logging' `
            -Detected ("EnableScriptBlockLogging={0}; EnableScriptBlockInvocationLogging={1}" -f $(if ($sb.Found) { $sb.Value } else { '(ns)' }), $(if ($si.Found) { $si.Value } else { '(ns)' })) `
            -Expected 'both = 1'

        # Credential Guard - policy value / GPO override
        $dg = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' 'LsaCfgFlags'
        $lc = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' 'LsaCfgFlags'
        if ($dg.Found -and $dg.Value -eq 0) {
            Add-Row -Header $H -Label 'Credential Guard - policy value' -LocalStatus 'NOT SET' `
                -Detail 'A GPO is DISABLING Credential Guard (Policies\...\DeviceGuard\LsaCfgFlags=0). Set "Turn On VBS > Credential Guard Configuration" to "Enabled with UEFI lock".' `
                -Category $CC -Item 'Credential Guard - policy value' `
                -Detected ("Policies\...\DeviceGuard\LsaCfgFlags=0; Control\LSA\LsaCfgFlags={0}" -f $(if ($lc.Found) { $lc.Value } else { '(ns)' })) `
                -Expected 'LsaCfgFlags=1 (Enabled with UEFI lock)'
        } else {
            Add-Row -Header $H -Label 'Credential Guard - policy value' `
                -LocalStatus $(if ($dg.Found -and $dg.Value -eq 1 -and $lc.Found -and $lc.Value -in @(1, 2)) { 'OK' } else { 'NOT SET' }) `
                -Category $CC -Item 'Credential Guard - policy value' `
                -Detected ("Policies\...\DeviceGuard\LsaCfgFlags={0}; Control\LSA\LsaCfgFlags={1}" -f $(if ($dg.Found) { $dg.Value } else { '(ns)' }), $(if ($lc.Found) { $lc.Value } else { '(ns)' })) `
                -Expected 'Policies LsaCfgFlags=1 and Control\LSA LsaCfgFlags in {1,2}'
        }

        $H = '7. Extended Baseline (Cybersecurity Review) - WinRM / Kerberos PKINIT / Ink / Audit / Cached logons'
        # WinRM auto-configuration disabled
        $wp = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service'
        $ac = Get-RegVal $wp 'AllowAutoConfig'
        Add-Row -Header $H -Label 'WinRM auto-configuration disabled' `
            -LocalStatus $(if ((-not $ac.Found) -or $ac.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
            -Category $CC -Item 'WinRM auto-configuration disabled' `
            -Detected ("AllowAutoConfig={0}" -f $(if ($ac.Found) { $ac.Value } else { '(not configured)' })) -Expected '0 or not configured'

        # WinRM listener IPv4/IPv6 filter
        $f4 = Get-RegVal $wp 'IPv4Filter'; $f6 = Get-RegVal $wp 'IPv6Filter'
        $f4v = if ($f4.Found) { [string]$f4.Value } else { '' }
        $f6v = if ($f6.Found) { [string]$f6.Value } else { '' }
        $filterOpen = ($f4v -eq '*' -or $f4v -eq '') -and ($f6v -eq '*' -or $f6v -eq '')
        Add-Row -Header $H -Label 'WinRM listener IPv4/IPv6 filter' `
            -LocalStatus $(if ($filterOpen) { 'MANUAL' } else { 'OK' }) `
            -Detail $(if ($filterOpen) { "IPv4Filter='$f4v', IPv6Filter='$f6v' - open to all. Scope to a management range (-WinRmFilterRange on the remediation script)." } else { "IPv4Filter='$f4v'; IPv6Filter='$f6v'" }) `
            -Category $CC -Item 'WinRM listener IPv4/IPv6 filter' `
            -Detected ("IPv4Filter='{0}'; IPv6Filter='{1}'" -f $f4v, $f6v) -Expected 'Scoped to a management IP range (not * / blank)'

        # PKINIT hash SHA256 / SHA384 / SHA512 = Supported (3)
        foreach ($alg in 'SHA256', 'SHA384', 'SHA512') {
            $v = Get-RegVal ("HKLM:\SOFTWARE\Policies\Microsoft\Windows\PKINITHashAlgorithms\{0}" -f $alg) 'Support'
            Add-Row -Header $H -Label ("PKINIT hash {0} = Supported" -f $alg) `
                -LocalStatus $(if ($v.Found -and $v.Value -eq 3) { 'OK' } else { 'NOT SET' }) `
                -Detail 'Benchmark value only - no AD / PKINIT smart-card logon in a workgroup.' `
                -Category $CC -Item ("PKINIT hash {0}" -f $alg) `
                -Detected ("Support={0}" -f $(if ($v.Found) { $v.Value } else { '(not set / default)' })) -Expected 'Support=3 (Supported)'
        }

        # Windows Ink Workspace disabled
        $v = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsInkWorkspace' 'AllowWindowsInkWorkspace'
        Add-Row -Header $H -Label 'Windows Ink Workspace disabled' `
            -LocalStatus $(if ($v.Found -and $v.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
            -Category $CC -Item 'Windows Ink Workspace disabled' `
            -Detected ("AllowWindowsInkWorkspace={0}" -f $(if ($v.Found) { $v.Value } else { '(not configured)' })) -Expected 'AllowWindowsInkWorkspace=0'

        # Audit "Audit Policy Change" = Success and Failure
        try {
            $csv = (& auditpol /get /subcategory:"Audit Policy Change" /r 2>$null | ConvertFrom-Csv)
            $row = $csv | Where-Object { $_.Subcategory -eq 'Audit Policy Change' } | Select-Object -First 1
            $inc = if ($row) { [string]$row.'Inclusion Setting' } else { '(unknown)' }
            Add-Row -Header $H -Label 'Audit "Audit Policy Change"' `
                -LocalStatus $(if ($inc -eq 'Success and Failure') { 'OK' } else { 'NOT SET' }) `
                -Category $CC -Item 'Audit "Audit Policy Change"' `
                -Detected ("Inclusion Setting={0}" -f $inc) -Expected 'Success and Failure'
        } catch {
            Add-Row -Header $H -Label 'Audit "Audit Policy Change"' -LocalStatus 'NOT SET' -Detail 'auditpol read failed' `
                -Category $CC -Item 'Audit "Audit Policy Change"' -NormStatus 'Unable to Check' `
                -Detected 'auditpol /get failed' -Expected 'Success and Failure'
        }

        # Cached logons count
        $v = Get-RegVal 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' 'CachedLogonsCount'
        $clc = if ($v.Found) { [int]$v.Value } else { $null }
        Add-Row -Header $H -Label 'Cached logons count' `
            -LocalStatus $(if ($null -ne $clc -and $clc -le $CachedLogonsTarget) { 'OK' } else { 'NOT SET' }) `
            -Detail $(if ($null -ne $clc -and $clc -gt 0) { "value $clc - Qualys QID 90007 only clears at 0 on a workgroup server" } else { '' }) `
            -Category $CC -Item 'Cached logons count' `
            -Detected ("CachedLogonsCount={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected ("<= {0}" -f $CachedLogonsTarget)

        $H = '7. Extended Baseline (Cybersecurity Review) - Local Security Policy / patch (verify with evidence)'
        # Built-in Guest account renamed (checkable via well-known SID -501)
        try {
            $guest = Get-LocalUser -ErrorAction Stop | Where-Object { $_.SID.Value -like '*-501' } | Select-Object -First 1
            if ($guest) {
                Add-Row -Header $H -Label 'Built-in Guest account renamed (QID 105228)' `
                    -LocalStatus $(if ($guest.Name -ne 'Guest') { 'OK' } else { 'NOT SET' }) `
                    -Detail ("Guest account name = '{0}', Enabled = {1}" -f $guest.Name, $guest.Enabled) `
                    -Category $CC -Item 'Built-in Guest account renamed' `
                    -Detected ("Name='{0}'; Enabled={1}" -f $guest.Name, $guest.Enabled) -Expected "Renamed (not 'Guest') and Disabled"
            } else {
                Add-Row -Header $H -Label 'Built-in Guest account renamed (QID 105228)' -LocalStatus 'MANUAL' -Detail 'Built-in Guest (SID -501) not found - verify manually' `
                    -Category $CC -Item 'Built-in Guest account renamed' -Detected 'SID -501 account not found' -Expected 'Renamed and Disabled'
            }
        } catch {
            Add-Row -Header $H -Label 'Built-in Guest account renamed (QID 105228)' -LocalStatus 'MANUAL' -Detail 'Get-LocalUser not available - verify with secedit export' `
                -Category $CC -Item 'Built-in Guest account renamed' -NormStatus 'Manual Verification Required' -Detected 'unable to read local users' -Expected 'Renamed and Disabled'
        }

        # Deny access to this computer from the network - read-only verify via a secedit export.
        try {
            $cfg = Join-Path $env:TEMP ("pcur_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
            & secedit /export /cfg $cfg /areas USER_RIGHTS /quiet 2>&1 | Out-Null
            $dl = (Select-String -LiteralPath $cfg -Pattern '^\s*SeDenyNetworkLogonRight\s*=' -ErrorAction SilentlyContinue | Select-Object -First 1).Line
            Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue
            $dvals = if ($dl) { ($dl -split '=', 2)[1].Trim() } else { '(not set)' }
            $dok = $dl -and @('S-1-5-32-546', 'S-1-5-113', 'S-1-5-114' | Where-Object { $dl -notmatch [regex]::Escape($_) }).Count -eq 0
            Add-Row -Header $H -Label 'Deny access to this computer from the network' `
                -LocalStatus $(if ($dok) { 'OK' } else { 'NOT SET' }) `
                -Category $CC -Item 'Deny access from network (user right)' `
                -Detected ("SeDenyNetworkLogonRight = {0}" -f $dvals) -Expected 'contains S-1-5-32-546, S-1-5-113, S-1-5-114'
        } catch {
            Add-Row -Header $H -Label 'Deny access to this computer from the network' -LocalStatus 'MANUAL' -Detail 'secedit export failed - verify manually' `
                -Category $CC -Item 'Deny access from network (user right)' -NormStatus 'Manual Verification Required' -Detected 'secedit /export failed' -Expected 'contains S-1-5-32-546, S-1-5-113, S-1-5-114'
        }

        # Deny logon through RDS - excluded from automatic remediation (workgroup admin RDP); verify manually.
        Add-Row -Header $H -Label 'Deny logon through Remote Desktop Services' -LocalStatus 'MANUAL' `
            -Detail 'Excluded from automatic remediation by design (workgroup local-admin RDP). Verify SeDenyRemoteInteractiveLogonRight manually if applied (secedit /export).' `
            -Category $CC -Item 'Deny logon through RDS (user right)' -Detected 'Not remediated by script' -Expected 'Guests (+ Local account if RDP not needed)'

        # Allow Administrator account lockout - read-only verify via a secedit export.
        try {
            $cfg = Join-Path $env:TEMP ("pcsa_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
            & secedit /export /cfg $cfg /areas SECURITYPOLICY /quiet 2>&1 | Out-Null
            $al = (Select-String -LiteralPath $cfg -Pattern '^\s*AllowAdministratorLockout\s*=' -ErrorAction SilentlyContinue | Select-Object -First 1).Line
            Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue
            $aval = if ($al) { ($al -split '=', 2)[1].Trim() } else { '(not set)' }
            Add-Row -Header $H -Label 'Allow Administrator account lockout' `
                -LocalStatus $(if ($aval -eq '1') { 'OK' } else { 'NOT SET' }) `
                -Category $CC -Item 'Allow Administrator account lockout' `
                -Detected ("AllowAdministratorLockout = {0}" -f $aval) -Expected 'AllowAdministratorLockout = 1'
        } catch {
            Add-Row -Header $H -Label 'Allow Administrator account lockout' -LocalStatus 'MANUAL' -Detail 'secedit export failed - verify manually' `
                -Category $CC -Item 'Allow Administrator account lockout' -NormStatus 'Manual Verification Required' -Detected 'secedit /export failed' -Expected 'AllowAdministratorLockout = 1'
        }

        # August 2026 Windows update - checkable via build revision (UBR)
        try {
            $cv  = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
            $bld = [int]$cv.CurrentBuildNumber
            $ubr = [int]$cv.UBR
            $fullBuild = "$bld.$ubr"
            $patched = ($bld -gt 26100) -or ($bld -eq 26100 -and $ubr -ge 33222)
            Add-Row -Header $H -Label 'August 2026 Windows Update (QID 92439)' `
                -LocalStatus $(if ($patched) { 'OK' } else { 'NOT SET' }) `
                -Detail $(if ($patched) { '' } else { 'install KB5120228 / KB5120233 via Windows Update, then re-scan' }) `
                -Category $CC -Item 'August 2026 Windows Update' `
                -Detected ("OS build {0}" -f $fullBuild) -Expected 'build >= 26100.33222'
        } catch {
            Add-Row -Header $H -Label 'August 2026 Windows Update (QID 92439)' -LocalStatus 'MANUAL' -Detail 'could not read OS build revision' `
                -Category $CC -Item 'August 2026 Windows Update' -NormStatus 'Manual Verification Required' -Detected 'UBR unreadable' -Expected 'build >= 26100.33222'
        }

        Add-Row -Header $H -Label 'QID 92446 - Defender EoP zero-day' -LocalStatus 'MANUAL' `
            -Detail 'No Microsoft patch at assessment time - track MSRC advisory, apply when released; keep Defender platform/signatures current.' `
            -Category $CC -Item 'QID 92446 - Defender EoP zero-day' -Detected 'No patch at assessment time' -Expected 'Patched when available'
        Add-Row -Header $H -Label 'Third-party agent vulnerabilities' -LocalStatus 'MANUAL' `
            -Detail 'Splunk UF / Azure Arc agent / ServiceNow agent / Binalyze AIR - upgrade with the respective owners (Go / gRPC / .NET / Ruby vulns).' `
            -Category $CC -Item 'Third-party agent vulnerabilities' -Detected 'Vulnerable agent builds present' -Expected 'Agents on fixed versions'
    } else {
        Add-Row -Header '7. Extended Baseline (Cybersecurity Review)' -Label 'Extended Baseline verification' -LocalStatus 'N/A' `
            -Detail '-SkipExtendedBaseline was passed - Category 7 was not verified this run.' `
            -Category '7. Extended Baseline (Cybersecurity Review)' -Item 'All items' -NormStatus 'Not Applicable' `
            -Detected 'Skipped by operator' -Expected 'n/a'
    }

    # ==========================================================================
    # 8. ISO Policy Compliance (Qualys PC) - READ-ONLY verification of the
    #    Control IDs from iso.xlsx. Uses the existing Add-Row / Get-RegVal
    #    helpers only. Header/Label/LocalStatus + Category/Item/Status/Detected/
    #    Expected exactly as categories 1-7. Get-* / registry / auditpol / secedit
    #    /export only - never writes.
    # ==========================================================================
    if (-not $SkipIsoPolicy) {
        $IH = '8. ISO Policy Compliance (Qualys PC)'
        function Test-IsoReg {
            param([int]$Cid,[string]$Label,[string]$Path,[string]$Name,[string]$Type,$Data,[bool]$IE=$false)
            if ($IE -and $SkipIeHardening) { return }
            $v = Get-RegVal $Path $Name
            $ok = $false
            if ($v.Found) { if ($Type -eq 'String') { $ok = ([string]$v.Value -eq [string]$Data) } else { $ok = ([int]$v.Value -eq [int]$Data) } }
            Add-Row -Header $IH -Label ("CID $Cid - $Label") `
                -LocalStatus $(if ($ok) { 'OK' } else { 'NOT SET' }) `
                -Category $IH -Item ("CID $Cid - $Label") `
                -Detected ("{0} = {1}" -f $Name, $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected ("{0} = {1}" -f $Name, $Data)
        }
        $IsoRegV = @(
        @{ Cid=1366; Title='''Accounts: Limit local account use of blank passwords to console logon only'''; Path='HKLM:\System\CurrentControlSet\Control\Lsa'; Name='LimitBlankPasswordUse'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=2584; Title='''UAC: Only elevate UIAccess applications installed in secure locations'''; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System'; Name='EnableSecureUIAPaths'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=2586; Title='''UAC: Admin Approval Mode for the Built-in Administrator account'''; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System'; Name='FilterAdministratorToken'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate='$EnableUacHardening'; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=2587; Title='''UAC: Behavior of the elevation prompt for administrators in Admin Approval Mode'''; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System'; Name='ConsentPromptBehaviorAdmin'; Type='DWord'; Data=2; Reboot=$false; IE=$false; Gate='$EnableUacHardening'; Risk='Medium - interactive admins get a consent prompt (minor). Real risk is non-interactive automation that elevates via runas/Task Scheduler ''Run with highest privileges'' while a user is NOT logged on - test the maintenance jobs. Consider value 5 (prompt for non-Windows binaries) as a lighter option if 2 breaks tooling.' }
        @{ Cid=3940; Title='''UAC: Virtualize file and registry write failures to per-user locations'''; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System'; Name='EnableVirtualization'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate='$EnableUacHardening'; Risk='Low/Medium - this is the Windows default (1). If it was deliberately set to 0 for a specific 64-bit-only application requirement, confirm before reverting (64-bit apps are never virtualized regardless).' }
        @{ Cid=8274; Title='''Configure Windows Defender SmartScreen'' (Explorer)'; Path='HKLM:\Software\Policies\Microsoft\Windows\System'; Name='EnableSmartScreen'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=9024; Title='''Apply UAC restrictions to local accounts on network logons'' (LocalAccountTokenFilterPolicy)'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System'; Name='LocalAccountTokenFilterPolicy'; Type='DWord'; Data=0; Reboot=$false; IE=$false; Gate='$RestrictLocalAcctNetworkLogon'; Risk='Medium - any tool that manages this host remotely by authenticating with a LOCAL account (not a domain account, not the VM guest-ops channel) will lose admin capability. VMware guest operations, console and RDP are unaffected. Verify remote-admin tooling does not depend on local-account network logon before applying.' }
        @{ Cid=10068; Title='IE - Restricted Sites Zone - ''Access data sources across domains'' (Zones\4\1406)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1406'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=10472; Title='''Turn On Virtualization Based Security'''; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Name='EnableVirtualizationBasedSecurity'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=10473; Title='''Turn On VBS (Credential Guard Configuration)'''; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Name='LsaCfgFlags'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=10474; Title='''Turn On VBS (Enable Virtualization Based Protection of Code Integrity)'' - HVCI'; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Name='HypervisorEnforcedCodeIntegrity'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=10475; Title='''Turn On VBS (Select Platform Security Level)'''; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Name='RequirePlatformSecurityFeatures'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=10970; Title='''Script Block Invocation Logging'''; Path='HKLM:\Software\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'; Name='EnableScriptBlockInvocationLogging'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=11573; Title='IE - Security Zones: Do not allow users to add/delete sites'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings'; Name='Security_zones_map_edit'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11574; Title='IE - Security Zones: Do not allow users to change policies'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings'; Name='Security_options_edit'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11615; Title='IE - Allow software to run or install even if the signature is invalid'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Download'; Name='RunInvalidSignatures'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11619; Title='IE - Turn off Encryption Support'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings'; Name='SecureProtocols'; Type='DWord'; Data=2560; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11621; Title='IE - Check for signatures on downloaded programs'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Download'; Name='CheckExeSignatures'; Type='String'; Data='yes'; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11742; Title='IE - Restricted Sites Zone - ''Download signed ActiveX controls'' (Zones\4\1001)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1001'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11743; Title='IE - Restricted Sites Zone - ''Download unsigned ActiveX controls'' (Zones\4\1004)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1004'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11744; Title='IE - Restricted Sites Zone - ''Run ActiveX controls and plugins'' (Zones\4\1200)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1200'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11745; Title='IE - Restricted Sites Zone - ''Script ActiveX controls marked safe for scripting'' (Zones\4\1405)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1405'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11746; Title='IE - Restricted Sites Zone - ''Initialize and script ActiveX controls not marked as safe'' (Zones\4\1201)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1201'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11747; Title='IE - Restricted Sites Zone - ''Allow file downloads'' (Zones\4\1803)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1803'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11750; Title='IE - Restricted Sites Zone - ''Allow META REFRESH'' (Zones\4\1608)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1608'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11753; Title='IE - Restricted Sites Zone - ''Allow drag and drop or copy and paste files'' (Zones\4\1802)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1802'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11754; Title='IE - Restricted Sites Zone - ''Launching applications and files in an IFRAME'' (Zones\4\1804)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1804'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11755; Title='IE - Restricted Sites Zone - ''Navigate windows and frames across different domains'' (Zones\4\1607)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1607'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11757; Title='IE - Restricted Sites Zone - ''Userdata persistence'' (Zones\4\1606)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1606'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11758; Title='IE - Restricted Sites Zone - ''Allow Active scripting'' (Zones\4\1400)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1400'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11759; Title='IE - Restricted Sites Zone - ''Scripting of Java applets'' (Zones\4\1402)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1402'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11760; Title='IE - Restricted Sites Zone - ''Allow updates to status bar via script'' (Zones\4\2103)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2103'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11761; Title='IE - Restricted Sites Zone - ''Logon options (Anonymous logon)'' (Zones\4\1A00)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1A00'; Type='DWord'; Data=196608; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11763; Title='IE - Internet Zone - ''Download signed ActiveX controls'' (Zones\3\1001)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1001'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11764; Title='IE - Internet Zone - ''Download unsigned ActiveX controls'' (Zones\3\1004)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1004'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11767; Title='IE - Internet Zone - ''Initialize and script ActiveX controls not marked as safe'' (Zones\3\1201)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1201'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11770; Title='IE - Internet Zone - ''Access data sources across domains'' (Zones\3\1406)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1406'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11774; Title='IE - Internet Zone - ''Allow drag and drop or copy and paste files'' (Zones\3\1802)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1802'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11775; Title='IE - Internet Zone - ''Launching applications and files in an IFRAME'' (Zones\3\1804)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1804'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11776; Title='IE - Internet Zone - ''Navigate windows and frames across different domains'' (Zones\3\1607)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1607'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11778; Title='IE - Internet Zone - ''Userdata persistence'' (Zones\3\1606)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1606'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11781; Title='IE - Internet Zone - ''Allow updates to status bar via script'' (Zones\3\2103)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2103'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11782; Title='IE - Internet Zone - ''Logon options (Prompt for user name and password)'' (Zones\3\1A00)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1A00'; Type='DWord'; Data=65536; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11787; Title='IE - Local Intranet Zone - ''Initialize and script ActiveX controls not marked as safe'' (Zones\1\1201)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1'; Name='1201'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11807; Title='IE - Trusted Sites Zone - ''Initialize and script ActiveX controls not marked as safe'' (Zones\2\1201)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2'; Name='1201'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11930; Title='IE - Prevent Bypassing SmartScreen Filter Warnings'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\PhishingFilter'; Name='PreventOverride'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11947; Title='IE - Check for server certificate revocation'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings'; Name='CertificateRevocation'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11948; Title='IE - Turn on certificate address mismatch warning'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings'; Name='WarnOnBadCertRecving'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=11949; Title='IE - Prevent ignoring certificate errors'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings'; Name='PreventIgnoreCertErrors'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12011; Title='IE - Internet Zone - ''Allow paste operations via scripts'' (Zones\3\1407)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1407'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12027; Title='IE - Internet Zone - ''Turn on Cross Site Scripting Filter (Enable)'' (Zones\3\1409)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1409'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12028; Title='IE - Internet Zone - ''Run .NET Framework-reliant components signed with Authenticode'' (Zones\3\2001)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2001'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12029; Title='IE - Internet Zone - ''Use pop-up blocker (Enable)'' (Zones\3\1809)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1809'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12030; Title='IE - Internet Zone - ''Allow scriptlets'' (Zones\3\1209)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1209'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12032; Title='IE - Internet Zone - ''Run .NET Framework-reliant components not signed with Authenticode'' (Zones\3\2004)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2004'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12033; Title='IE - Internet Zone - ''Allow scripting of Internet Explorer WebBrowser control'' (Zones\3\1206)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1206'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12037; Title='IE - Internet Zone - ''Launching programs and unsafe files (Prompt)'' (Zones\3\1806)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='1806'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12038; Title='IE - Internet Zone - ''Automatic prompting for file downloads'' (Zones\3\2200)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2200'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12048; Title='IE - Internet Zone - ''Allow loading of XAML files'' (Zones\3\2402)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2402'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12050; Title='IE - Internet Zone - ''Include local directory path when uploading files to a server'' (Zones\3\160A)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='160A'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12051; Title='IE - Internet Zone - ''Enable dragging of content from different domains within a window'' (Zones\3\2708)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2708'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12052; Title='IE - Internet Zone - ''Enable dragging of content from different domains across windows'' (Zones\3\2709)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2709'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12053; Title='IE - Internet Zone - ''Allow script-initiated windows without size or position constraints'' (Zones\3\2102)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2102'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12055; Title='IE - Internet Zone - ''Web sites in less privileged zones can navigate into this zone'' (Zones\3\2101)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2101'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12057; Title='IE - Intranet Sites: Include all network paths (UNCs)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap'; Name='UNCAsIntranet'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12059; Title='IE - Restricted Sites Zone - ''Turn on Cross-Site Scripting Filter (Enable)'' (Zones\4\1409)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1409'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12060; Title='IE - Restricted Sites Zone - ''Run .NET Framework-reliant components signed with Authenticode'' (Zones\4\2001)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2001'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12061; Title='IE - Restricted Sites Zone - ''Allow paste operations via script'' (Zones\4\1407)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1407'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12064; Title='IE - Restricted Sites Zone - ''Launching programs and unsafe files'' (Zones\4\1806)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1806'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12065; Title='IE - Restricted Sites Zone - ''Automatic prompting for file downloads'' (Zones\4\2200)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2200'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12066; Title='IE - Restricted Sites Zone - ''Allow loading of XAML files'' (Zones\4\2402)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2402'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12068; Title='IE - Restricted Sites Zone - ''Allow scripting of Internet Explorer WebBrowser controls'' (Zones\4\1206)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1206'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12069; Title='IE - Restricted Sites Zone - ''Allow Binary and Script Behaviors'' (Zones\4\2000)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2000'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12070; Title='IE - Restricted Sites Zone - ''Use Pop-up Blocker (Enable)'' (Zones\4\1809)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1809'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12071; Title='IE - Restricted Sites Zone - ''Allow Scriptlets'' (Zones\4\1209)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='1209'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12074; Title='IE - Restricted Sites Zone - ''Use SmartScreen Filter (Enable)'' (Zones\4\2301)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2301'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12075; Title='IE - Restricted Sites Zone - ''Run .NET Framework-reliant components not signed with Authenticode'' (Zones\4\2004)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2004'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12076; Title='IE - Restricted Sites Zone - ''Allow script-initiated windows without size or position constraints'' (Zones\4\2102)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2102'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12077; Title='IE - Restricted Sites Zone - ''Include local directory path when uploading files to a server'' (Zones\4\160A)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='160A'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12078; Title='IE - Restricted Sites Zone - ''Enable dragging of content from different domains within a window'' (Zones\4\2708)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2708'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12079; Title='IE - Restricted Sites Zone - ''Web sites in less privileged zones can navigate into this zone'' (Zones\4\2101)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2101'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12081; Title='IE - Restricted Sites Zone - ''Enable dragging of content from different domains across windows'' (Zones\4\2709)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='2709'; Type='DWord'; Data=3; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12101; Title='IE - Turn off the Security Settings Check feature'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Security'; Name='DisableSecuritySettingsCheck'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12105; Title='IE - Restrict ActiveX Install - IE Processes ((Reserved))'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_ACTIVEXINSTALL'; Name='(Reserved)'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12106; Title='IE - Restrict ActiveX Install - IE Processes (explorer.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_ACTIVEXINSTALL'; Name='explorer.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12107; Title='IE - Restrict ActiveX Install - IE Processes (iexplore.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_ACTIVEXINSTALL'; Name='iexplore.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12120; Title='IE - Consistent Mime Handling - IE Processes ((Reserved))'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_HANDLING'; Name='(Reserved)'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12121; Title='IE - Consistent Mime Handling - IE Processes (explorer.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_HANDLING'; Name='explorer.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12122; Title='IE - Consistent Mime Handling - IE Processes (iexplore.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_MIME_HANDLING'; Name='iexplore.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12123; Title='IE - Restrict File Download - IE Processes ((Reserved))'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD'; Name='(Reserved)'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12124; Title='IE - Restrict File Download - IE Processes (explorer.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD'; Name='explorer.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12125; Title='IE - Restrict File Download - IE Processes (iexplore.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_RESTRICT_FILEDOWNLOAD'; Name='iexplore.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12126; Title='IE - Protection From Zone Elevation - IE Processes ((Reserved))'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ZONE_ELEVATION'; Name='(Reserved)'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12127; Title='IE - Protection From Zone Elevation - IE Processes (explorer.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ZONE_ELEVATION'; Name='explorer.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12128; Title='IE - Protection From Zone Elevation - IE Processes (iexplore.exe)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_ZONE_ELEVATION'; Name='iexplore.exe'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12131; Title='IE - Internet Zone - ''Don''t run antimalware programs against ActiveX controls (Disable)'' (Zones\3\270C)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='270C'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12132; Title='IE - Local Intranet Zone - ''Don''t run antimalware programs against ActiveX controls (Disable)'' (Zones\1\270C)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1'; Name='270C'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12133; Title='IE - Restricted Sites Zone - ''Don''t run antimalware programs against ActiveX controls (Disable)'' (Zones\4\270C)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\4'; Name='270C'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12135; Title='IE - Trusted Sites Zone - ''Don''t run antimalware programs against ActiveX controls (Disable)'' (Zones\2\270C)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\2'; Name='270C'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12161; Title='IE - Prevent Managing SmartScreen Filter (mode On)'; Path='HKLM:\Software\Policies\Microsoft\Internet Explorer\PhishingFilter'; Name='EnabledV9'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12164; Title='IE - Internet Zone - ''Turn on SmartScreen Filter scan (Enable)'' (Zones\3\2301)'; Path='HKLM:\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'; Name='2301'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12165; Title='IE - Remove the "Run this time" button for outdated ActiveX controls'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Ext'; Name='RunThisTimeEnabled'; Type='DWord'; Data=0; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=12166; Title='IE - Turn off blocking of outdated ActiveX controls (Disabled)'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\Ext'; Name='VersionCheckEnabled'; Type='DWord'; Data=1; Reboot=$false; IE=$true; Gate=''; Risk='' }
        @{ Cid=13918; Title='''Turn On VBS: Require UEFI Memory Attributes Table'' (HVCIMATRequired)'; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Name='HVCIMATRequired'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=16104; Title='''Turn On VBS (Secure Launch Configuration)'' (ConfigureSystemGuardLaunch)'; Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Name='ConfigureSystemGuardLaunch'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=25349; Title='''Kernel-mode Hardware-enforced Stack Protection'' (enforcement)'; Path='HKLM:\Software\Policies\Microsoft\Windows\DeviceGuard'; Name='ConfigureKernelShadowStacksLaunch'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=25358; Title='''Configure NetBIOS settings'' (EnableNetbios) - disable NetBIOS name resolution'; Path='HKLM:\Software\Policies\Microsoft\Windows NT\DNSClient'; Name='EnableNetbios'; Type='DWord'; Data=0; Reboot=$false; IE=$false; Gate='$DisableNetbios'; Risk='Medium - if any legacy application, script or share access relies on a NetBIOS (short) name that is not resolvable via DNS, it will break. Confirm DNS has records for everything these servers talk to by short name; test, then roll out. Value 2 is a lower-risk interim step.' }
        @{ Cid=25359; Title='''Configure RPC listener settings: Authentication protocol to use for incoming RPC connections'' (ForceKerberosForRpc)'; Path='HKLM:\Software\Policies\Microsoft\Windows NT\Printers\RPC'; Name='ForceKerberosForRpc'; Type='DWord'; Data=0; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=25361; Title='''Configure RPC listener settings: Protocols to allow for incoming RPC connections'' (RpcProtocols)'; Path='HKLM:\Software\Policies\Microsoft\Windows NT\Printers\RPC'; Name='RpcProtocols'; Type='DWord'; Data=5; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=25938; Title='''Configures LSASS to run as a protected process (Enabled with UEFI Lock)'' (RunAsPPL)'; Path='HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'; Name='RunAsPPL'; Type='DWord'; Data=1; Reboot=$true; IE=$false; Gate=''; Risk='Registry value set; takes effect only after a reboot and (for VBS/Device Guard items) once the VM exposes Secure Boot + virtualization. This tool never reboots.' }
        @{ Cid=27615; Title='''Control whether exclusions are visible to Local Admins'' (HideExclusionsFromLocalAdmins)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender'; Name='HideExclusionsFromLocalAdmins'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29570; Title='''Network\LanmanServer: AuditClientDoesNotSupportEncryption'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer'; Name='AuditClientDoesNotSupportEncryption'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29571; Title='''Network\LanmanServer: AuditClientDoesNotSupportSigning'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer'; Name='AuditClientDoesNotSupportSigning'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29572; Title='''Network\LanmanServer: AuditInsecureGuestLogon'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer'; Name='AuditInsecureGuestLogon'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29573; Title='''Network\LanmanServer: EnableAuthRateLimiter'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer'; Name='EnableAuthRateLimiter'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29574; Title='''Lanman Server: Enable remote mailslots'''; Path='HKLM:\Software\Policies\Microsoft\Windows\Bowser'; Name='EnableMailslots'; Type='DWord'; Data=0; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29577; Title='''Network\LanmanServer: InvalidAuthenticationDelayTimeInMs'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer'; Name='InvalidAuthenticationDelayTimeInMs'; Type='DWord'; Data=2000; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29578; Title='''Network\LanmanWorkstation: AuditInsecureGuestLogon'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanWorkstation'; Name='AuditInsecureGuestLogon'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29579; Title='''Network\LanmanWorkstation: AuditServerDoesNotSupportEncryption'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanWorkstation'; Name='AuditServerDoesNotSupportEncryption'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29580; Title='''Network\LanmanWorkstation: AuditServerDoesNotSupportSigning'''; Path='HKLM:\Software\Policies\Microsoft\Windows\LanmanWorkstation'; Name='AuditServerDoesNotSupportSigning'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29581; Title='''Lanman Workstation: Enable remote mailslots'''; Path='HKLM:\Software\Policies\Microsoft\Windows\NetworkProvider'; Name='EnableMailslots'; Type='DWord'; Data=0; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29592; Title='''Control whether exclusions are visible to local users'' (HideExclusionsFromLocalUsers)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender'; Name='HideExclusionsFromLocalUsers'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29594; Title='''Configure real-time protection and Security Intelligence Updates during OOBE'' (OobeEnableRtpAndSigUpdate)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender\Real-Time Protection'; Name='OobeEnableRtpAndSigUpdate'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29595; Title='''Configure whether to report Dynamic Signature dropped events'' (EnableDynamicSignatureDroppedEventReporting)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender\Reporting'; Name='EnableDynamicSignatureDroppedEventReporting'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29596; Title='''Scan excluded files and directories during quick scans'' (QuickScanIncludeExclusions)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender\Scan'; Name='QuickScanIncludeExclusions'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29741; Title='''Kerberos - Configure hash algorithms for certificate logon (PKInitHashAlgorithmConfigurationEnabled)'' - client'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters'; Name='PKInitHashAlgorithmConfigurationEnabled'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29742; Title='''Kerberos - Configure hash algorithms for certificate logon (PKInitSHA1)'' - client'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters'; Name='PKINITSHA1'; Type='DWord'; Data=1; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29743; Title='''Kerberos - Configure hash algorithms for certificate logon (PKInitSHA256)'' - client'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters'; Name='PKINITSHA256'; Type='DWord'; Data=3; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29744; Title='''Kerberos - Configure hash algorithms for certificate logon (PKInitSHA384)'' - client'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters'; Name='PKINITSHA384'; Type='DWord'; Data=3; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=29745; Title='''Kerberos - Configure hash algorithms for certificate logon (PKInitSHA512)'' - client'; Path='HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System\Kerberos\Parameters'; Name='PKInitSHA512'; Type='DWord'; Data=3; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=30458; Title='''Select the channel for Microsoft Defender daily security intelligence updates'' (SignaturesRing)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender'; Name='SignaturesRing'; Type='String'; Data='5'; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=30459; Title='''Select the channel for Microsoft Defender monthly engine updates'' (EngineRing)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender'; Name='EngineRing'; Type='String'; Data='5'; Reboot=$false; IE=$false; Gate=''; Risk='' }
        @{ Cid=30460; Title='''Select the channel for Microsoft Defender monthly platform updates'' (PlatformRing)'; Path='HKLM:\Software\Policies\Microsoft\Windows Defender'; Name='PlatformRing'; Type='String'; Data='5'; Reboot=$false; IE=$false; Gate=''; Risk='' }
        )
        foreach ($it in $IsoRegV) { Test-IsoReg -Cid $it.Cid -Label $it.Title -Path $it.Path -Name $it.Name -Type $it.Type -Data $it.Data -IE $it.IE }

        # CID 4501 - auditpol
        try {
            $csv = (& auditpol /get /subcategory:"Audit Policy Change" /r 2>$null | ConvertFrom-Csv)
            $row = $csv | Where-Object { $_.Subcategory -eq 'Audit Policy Change' } | Select-Object -First 1
            $inc = if ($row) { [string]$row.'Inclusion Setting' } else { '(unknown)' }
            Add-Row -Header $IH -Label 'CID 4501 - Audit "Audit Policy Change"' -LocalStatus $(if ($inc -eq 'Success and Failure') { 'OK' } else { 'NOT SET' }) `
                -Category $IH -Item 'CID 4501 - Audit "Audit Policy Change"' -Detected ("Inclusion Setting=$inc") -Expected 'Success and Failure'
        } catch {
            Add-Row -Header $IH -Label 'CID 4501 - Audit "Audit Policy Change"' -LocalStatus 'NOT SET' -Detail 'auditpol read failed' `
                -Category $IH -Item 'CID 4501 - Audit "Audit Policy Change"' -NormStatus 'Unable to Check' -Detected 'auditpol /get failed' -Expected 'Success and Failure'
        }

        # CID 3950/3951/3952 - Windows Firewall profile state
        foreach ($fp in @('Domain','Private','Public')) {
            $cid = switch ($fp) { 'Domain' {3952} 'Private' {3951} 'Public' {3950} }
            try {
                $en = (Get-NetFirewallProfile -Profile $fp -ErrorAction Stop).Enabled
                Add-Row -Header $IH -Label ("CID {0} - Windows Firewall state ({1})" -f $cid, $fp) -LocalStatus $(if ($en -eq $true) { 'OK' } else { 'NOT SET' }) `
                    -Category $IH -Item ("CID {0} - Windows Firewall state ({1})" -f $cid, $fp) -Detected ("Enabled=$en") -Expected 'Enabled=True (On)'
            } catch {
                $v = Get-RegVal ("HKLM:\Software\Policies\Microsoft\WindowsFirewall\{0}Profile" -f $fp) 'EnableFirewall'
                Add-Row -Header $IH -Label ("CID {0} - Windows Firewall state ({1})" -f $cid, $fp) -LocalStatus $(if ($v.Found -and $v.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
                    -Category $IH -Item ("CID {0} - Windows Firewall state ({1})" -f $cid, $fp) -Detected ("EnableFirewall={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected 'On (1)'
            }
        }

        # CID 29576/29583 - minimum SMB dialect
        foreach ($sd in @(
            @{ Cid=29576; Key='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer' },
            @{ Cid=29583; Key='HKLM:\Software\Policies\Microsoft\Windows\LanmanWorkstation' } )) {
            $v = Get-RegVal $sd.Key 'MinSmb2Dialect'
            Add-Row -Header $IH -Label ("CID {0} - Mandate minimum SMB version" -f $sd.Cid) -LocalStatus $(if ($v.Found -and [int]$v.Value -eq 768) { 'OK' } else { 'NOT SET' }) `
                -Category $IH -Item ("CID {0} - Mandate minimum SMB version" -f $sd.Cid) -Detected ("MinSmb2Dialect={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '768 (SMB 3.0.0)'
        }

        # CID 1071 / 2342 - account policy (secedit export, read-only)
        try {
            $cfg = Join-Path $env:TEMP ("iso_pc_sa_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
            & secedit /export /cfg $cfg /areas SECURITYPOLICY /quiet 2>&1 | Out-Null
            $mpl = (Select-String -LiteralPath $cfg -Pattern '^\s*MinimumPasswordLength\s*=\s*(\d+)' | Select-Object -First 1)
            $lbc = (Select-String -LiteralPath $cfg -Pattern '^\s*LockoutBadCount\s*=\s*(\d+)' | Select-Object -First 1)
            Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue
            $mplV = if ($mpl) { [int]$mpl.Matches[0].Groups[1].Value } else { $null }
            $lbcV = if ($lbc) { [int]$lbc.Matches[0].Groups[1].Value } else { $null }
            Add-Row -Header $IH -Label 'CID 1071 - Minimum Password Length' -LocalStatus $(if ($null -ne $mplV -and $mplV -ge 14) { 'OK' } else { 'NOT SET' }) `
                -Category $IH -Item 'CID 1071 - Minimum Password Length' -Detected ("MinimumPasswordLength={0}" -f $(if ($null -ne $mplV) { $mplV } else { '(not set)' })) -Expected '>= 14'
            Add-Row -Header $IH -Label 'CID 2342 - Account Lockout Threshold' -LocalStatus $(if ($null -ne $lbcV -and $lbcV -ge 1 -and $lbcV -le 3) { 'OK' } else { 'NOT SET' }) `
                -Category $IH -Item 'CID 2342 - Account Lockout Threshold' -Detected ("LockoutBadCount={0}" -f $(if ($null -ne $lbcV) { $lbcV } else { '(not set)' })) -Expected 'in range 1-3'
        } catch {
            Add-Row -Header $IH -Label 'CID 1071 / 2342 - account policy' -LocalStatus 'MANUAL' -Detail 'secedit export failed - check with net accounts' `
                -Category $IH -Item 'CID 1071 / 2342 - account policy' -NormStatus 'Unable to Check' -Detected 'secedit /export failed' -Expected 'MinPwLen>=14; LockoutThreshold 1-3'
        }

        # CID 2196 - Deny access from the network (user right; secedit export)
        try {
            $cfg = Join-Path $env:TEMP ("iso_pc_ura_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
            & secedit /export /cfg $cfg /areas USER_RIGHTS /quiet 2>&1 | Out-Null
            $l = (Select-String -LiteralPath $cfg -Pattern '^\s*SeDenyNetworkLogonRight\s*=' | Select-Object -First 1).Line
            Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue
            $dok = $l -and @('S-1-5-32-546','S-1-5-113','S-1-5-114' | Where-Object { $l -notmatch [regex]::Escape($_) }).Count -eq 0
            Add-Row -Header $IH -Label 'CID 2196 - Deny access from network' -LocalStatus $(if ($dok) { 'OK' } else { 'NOT SET' }) `
                -Category $IH -Item 'CID 2196 - Deny access from network (user right)' -Detected ("SeDenyNetworkLogonRight = {0}" -f $(if ($l) { ($l -split '=',2)[1].Trim() } else { '(not set)' })) -Expected 'contains S-1-5-32-546, S-1-5-113, S-1-5-114'
        } catch {
            Add-Row -Header $IH -Label 'CID 2196 - Deny access from network' -LocalStatus 'MANUAL' -Detail 'secedit export failed' `
                -Category $IH -Item 'CID 2196 - Deny access from network (user right)' -NormStatus 'Manual Verification Required' -Detected 'secedit /export failed' -Expected 'contains S-1-5-32-546, S-1-5-113, S-1-5-114'
        }

        # CID 25357 / 30456 - ASR rules
        try {
            $pref = Get-MpPreference -ErrorAction Stop
            foreach ($ar in @(
                @{ Cid=25357; Guid='56a863a9-875e-4185-98a7-b882c64b5ce5' },
                @{ Cid=30456; Guid='a8f5898e-1dc8-49a9-9878-85004b8a61e6' } )) {
                $i = [array]::IndexOf(@($pref.AttackSurfaceReductionRules_Ids), $ar.Guid)
                $act = if ($i -ge 0) { [int](@($pref.AttackSurfaceReductionRules_Actions)[$i]) } else { -1 }
                Add-Row -Header $IH -Label ("CID {0} - ASR rule Block" -f $ar.Cid) -LocalStatus $(if ($act -eq 1) { 'OK' } else { 'NOT SET' }) `
                    -Category $IH -Item ("CID {0} - ASR rule {1}" -f $ar.Cid, $ar.Guid) -Detected ("action={0}" -f $(if ($act -ge 0) { $act } else { 'not configured' })) -Expected 'Block (1)'
            }
        } catch {
            Add-Row -Header $IH -Label 'CID 25357 / 30456 - ASR rules' -LocalStatus 'NOT SET' -Detail 'Get-MpPreference failed' `
                -Category $IH -Item 'CID 25357 / 30456 - ASR rules' -NormStatus 'Unable to Check' -Detected 'Get-MpPreference failed' -Expected 'Block (1)'
        }

        # Not Applicable (workgroup / role) - CID 29740, 29586-29590
        foreach ($na in @(
            @{ Cid=29740; T='Machine Identity Isolation'; R='workgroup - no AD machine-account secret to isolate' },
            @{ Cid=29586; T='KDC PKINIT hash config enabled'; R='KDC role not installed (not a Domain Controller)' },
            @{ Cid=29587; T='KDC PKINIT SHA1'; R='KDC role not installed (not a Domain Controller)' },
            @{ Cid=29588; T='KDC PKINIT SHA256'; R='KDC role not installed (not a Domain Controller)' },
            @{ Cid=29589; T='KDC PKINIT SHA384'; R='KDC role not installed (not a Domain Controller)' },
            @{ Cid=29590; T='KDC PKINIT SHA512'; R='KDC role not installed (not a Domain Controller)' } )) {
            Add-Row -Header $IH -Label ("CID {0} - {1}" -f $na.Cid, $na.T) -LocalStatus 'N/A' -Detail ("NOT APPLICABLE - $($na.R). [$WorkgroupWhy]") `
                -Category $IH -Item ("CID {0} - {1}" -f $na.Cid, $na.T) -NormStatus 'Not Applicable' -Detected 'Not applicable on a workgroup member server' -Expected 'N/A'
        }

        # Manual remediation recommended - CID 2186/2195/2200/2392/2642/25350/30461 and per-user IE 10821/17333/30467
        foreach ($mv in @(
            @{ Cid=2186;  T="'Back up files and directories' right"; E='Administrators only (Splunk grant needs exception or Event Log Readers)' },
            @{ Cid=2195;  T="'Debug Programs' right"; E='Administrators only (remove ServiceNow Users unless justified)' },
            @{ Cid=2200;  T="'Deny log on through Remote Desktop Services' right"; E='Guests (+ Local account once VM-console access confirmed)' },
            @{ Cid=2392;  T="'Manage auditing and security log' right"; E='Administrators only (move Splunk to Event Log Readers)' },
            @{ Cid=2642;  T="'Impersonate a client after authentication' right"; E='Administrators, LOCAL/NETWORK SERVICE, SERVICE (IIS_IUSRS = exception if IIS installed)' },
            @{ Cid=25350; T="'Allow Custom SSPs and APs to be loaded into LSASS'"; E='Decision required (Qualys wants 1; hardened stance is 0) - see report' },
            @{ Cid=30461; T="'Allow network protection on Windows Server'"; E='Decision required (Qualys wants 0) - confirm Defender standard' },
            @{ Cid=10821; T='IE per-user autocomplete (FormSuggest Passwords)'; E='no  (per HKU hive + .DEFAULT)' },
            @{ Cid=17333; T='IE per-user autocomplete (FormSuggest PW Ask)'; E='no  (per HKU hive + .DEFAULT)' },
            @{ Cid=30467; T='IE per-user autocomplete (Control Panel\FormSuggest Passwords)'; E='1  (per HKU hive + .DEFAULT)' } )) {
            Add-Row -Header $IH -Label ("CID {0} - {1}" -f $mv.Cid, $mv.T) -LocalStatus 'MANUAL' -Detail 'See consolidated remediation report - manual / change-review item.' `
                -Category $IH -Item ("CID {0} - {1}" -f $mv.Cid, $mv.T) -NormStatus 'Manual Verification Required' -Detected 'Not remediated by script' -Expected $mv.E
        }
    } else {
        Add-Row -Header '8. ISO Policy Compliance (Qualys PC)' -Label 'ISO Policy Compliance verification' -LocalStatus 'N/A' `
            -Detail '-SkipIsoPolicy was passed - Category 8 was not verified this run.' `
            -Category '8. ISO Policy Compliance (Qualys PC)' -Item 'All items' -NormStatus 'Not Applicable' -Detected 'Skipped by operator' -Expected 'n/a'
    }

} catch {
    $fatal = $_.Exception.Message
}

$meta = [PSCustomObject]@{ Hostname = $ComputerName; IPAddress = $IPAddresses; OS = $OSCaption; FatalError = $fatal }
# .ToArray() rather than @(...) - @() around a generic List instance is unreliable
# on some PowerShell builds; .ToArray() is well-defined on 5.1 and 7.x alike.
$envelope = [PSCustomObject]@{ Schema = 'secgap-postcheck-2'; Meta = $meta; Rows = $Rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<SECGAP-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-SECGAP-ENVELOPE>>>'
'@

# =====================================================================================
# 3. Admin-side helpers (identical validation to the remediation wrapper)
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
            Reason = "VM is not powered on (PowerState=$($vm.PowerState))."; Detected = "PowerState=$($vm.PowerState)"; VM = $null; Server = $null }
    }
    $g = $vm.ExtensionData.Guest
    $toolsOk = ($g.ToolsRunningStatus -eq 'guestToolsRunning') -or ($g.ToolsStatus -in 'toolsOk', 'toolsOld')
    if (-not $toolsOk) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VMware Tools'; ServerName = $srv.Name
            Reason = "VMware Tools not installed/running (ToolsStatus=$($g.ToolsStatus), ToolsRunningStatus=$($g.ToolsRunningStatus))."; Detected = "ToolsStatus=$($g.ToolsStatus)"; VM = $null; Server = $null }
    }
    $isWin = ($g.GuestFamily -eq 'windowsGuest') -or ($g.GuestId -match 'windows') -or ($vm.Guest.OSFullName -match 'Windows')
    if (-not $isWin) {
        return [pscustomobject]@{ Ok = $false; Stage = 'Guest OS'; ServerName = $srv.Name
            Reason = "Guest OS is not Windows (GuestFamily=$($g.GuestFamily), OS=$($vm.Guest.OSFullName))."; Detected = "$($vm.Guest.OSFullName)"; VM = $null; Server = $null }
    }
    return [pscustomobject]@{ Ok = $true; Stage = 'OK'; ServerName = $srv.Name; Reason = ''; Detected = ''; VM = $vm; Server = $srv }
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

# See Invoke-RemoteSecurityGapRemediation.ps1 for the full rationale. A single
# Invoke-VMScript carrying the whole payload fails in some environments with the
# misleading "Could not locate Powershell script interpreter" error (the script is
# placed on a command line that overruns a guest-ops length limit). Primary method
# here: Copy-VMGuestFile (one transfer, no limit); fallback: write Base64 into a
# guest file in small pieces, then decode in-guest. Guest operations only - no
# WinRM/PsExec/SMB/ESXi-or-guest network path. Staged file holds no credential and
# is always deleted.
function Invoke-LargeGuestPayload {
    param(
        $VM, $Server, [pscredential]$Credential,
        [string]$PayloadText, [int]$ToolsWaitSecs,
        [int]$ChunkSize = 1500
    )
    $tag    = [guid]::NewGuid().ToString('N').Substring(0, 12)
    $gB64   = "C:\Windows\Temp\sg$tag.b64"
    $gPs1   = "C:\Windows\Temp\sg$tag.ps1"
    $b64    = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($PayloadText))
    $staged = $false
    Write-Log "  [stage] payload is $($PayloadText.Length) chars; delivering to guest as a file." 'INFO'

    try {
        if (Get-Command Copy-VMGuestFile -ErrorAction SilentlyContinue) {
            $local = Join-Path $env:TEMP "sg$tag.ps1"
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
    param($vCenter, $VMName, $GuestHostname, $IPAddress, $OS, $Category, $SecurityCheck, $DetectedValue, $ExpectedValue, $Status, $Action, $RequiresReboot, $RiskNote, $Reason)
    [PSCustomObject]([ordered]@{
        vCenter = $vCenter; VMName = $VMName; GuestHostname = $GuestHostname; IPAddress = $IPAddress; OS = $OS
        Category = $Category; SecurityCheck = $SecurityCheck; DetectedValue = $DetectedValue; ExpectedValue = $ExpectedValue
        Status = $Status; Action = $Action; RequiresReboot = $RequiresReboot; RiskNote = $RiskNote
        'Error/Reason' = $Reason; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

# On-screen output, modelled on the local post-remediation-check.ps1:
#   "  [<STATUS>] <Label>"   (OK*=Green, MANUAL*=DarkYellow, N/A*=Gray, else Red;
#                             label White)
# Alignment fixes vs. the local script:
#   * the status token is right-padded to the WIDEST token in this VM's result set
#     (min 8), so a long token like "[OK (19 rules)]" or "[NOT SET (found: 3)]"
#     no longer pushes its label out of column;
#   * a SHORT detail (<= $InlineDetailMax chars) is still printed inline " - ..."
#     like the local script; a LONG justification (Machine Identity Isolation,
#     Kerberos, etc.) is word-wrapped onto its own indented line(s) directly under
#     the label, so it never disturbs the aligned [status]/label columns.
function Show-VmResults {
    param([string]$VmBanner, $Rows, [int]$InlineDetailMax = 70, [int]$WrapWidth = 100)
    $HeaderColor   = 'Cyan'
    $SubPointColor = 'White'

    $statusWidth = 8
    foreach ($row in $Rows) {
        $l = ([string]$row.LocalStatus).Length
        if ($l -gt $statusWidth) { $statusWidth = $l }
    }
    $contIndent = ' ' * ($statusWidth + 5)   # 2 leading spaces + '[' + status + '] '

    Write-Host ""
    Write-Host ("##### $VmBanner #####") -ForegroundColor Cyan
    $lastHeader = $null
    foreach ($row in $Rows) {
        if ($row.Header -and $row.Header -ne $lastHeader) {
            Write-Host ""
            Write-Host ("=== {0} ===" -f $row.Header) -ForegroundColor $HeaderColor
            $lastHeader = $row.Header
        }
        $s = [string]$row.LocalStatus
        $statusColor = switch -Wildcard ($s) {
            'OK*'     { 'Green' }
            'MANUAL*' { 'DarkYellow' }
            'N/A*'    { 'Gray' }
            default   { 'Red' }
        }
        Write-Host ("  [{0}] " -f $s.PadRight($statusWidth)) -ForegroundColor $statusColor -NoNewline
        Write-Host $row.Label -ForegroundColor $SubPointColor -NoNewline

        $text = if ($row.Detail) { (([string]$row.Detail) -replace '\s+', ' ').Trim() } else { '' }
        if (-not $text) {
            Write-Host ""
        } elseif ($text.Length -le $InlineDetailMax) {
            Write-Host (" - {0}" -f $text) -ForegroundColor Gray            # inline, like the local script
        } else {
            Write-Host ""                                                  # finish the label line
            while ($text.Length -gt $WrapWidth) {
                $cut = $text.LastIndexOf(' ', [Math]::Min($WrapWidth, $text.Length - 1))
                if ($cut -le 0) { $cut = $WrapWidth }
                Write-Host ($contIndent + $text.Substring(0, $cut).TrimEnd()) -ForegroundColor Gray
                $text = $text.Substring($cut).TrimStart()
            }
            Write-Host ($contIndent + $text) -ForegroundColor Gray
        }
    }
}

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows = New-Object System.Collections.Generic.List[object]
$summary = [ordered]@{ Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0; Manual = 0; NotApplicable = 0; Unable = 0; SkippedFailed = 0 }
$inventory = Get-VMInventory -Servers $connectedServers

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Verifying VM '$vmName' ===" 'INFO'
    try {
        $val = Resolve-AndValidateVM -Name $vmName -Inventory $inventory
        if (-not $val.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -SecurityCheck $val.Stage -DetectedValue $val.Detected -ExpectedValue '' `
                -Status 'Unable to Check' -Action 'Skipped - VM not processed' -RequiresReboot $false -RiskNote '' -Reason $val.Reason))
            Write-Log "SKIP '$vmName' [$($val.Stage)]: $($val.Reason)" 'WARN'
            Show-VmResults -VmBanner "$vmName  [SKIPPED - not processed]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = $val.Stage; LocalStatus = 'SKIP'; Detail = $val.Reason })
            continue
        }
        $vm = $val.VM; $srv = $val.Server

        $probe = Test-GuestPowerShell -VM $vm -Server $srv -Credential $GuestCredential -ToolsWaitSecs $ToolsWaitSecs
        if (-not $probe.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $vm.Guest.HostName -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest authentication / PowerShell' -DetectedValue '' -ExpectedValue 'PowerShell 5.1+ reachable via VMware Tools' `
                -Status 'Unable to Check' -Action 'Skipped - VM not processed' -RequiresReboot $false -RiskNote '' -Reason $probe.Reason))
            Write-Log "SKIP '$vmName' [probe]: $($probe.Reason)" 'WARN'
            Show-VmResults -VmBanner "$vmName  [SKIPPED - not processed]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = 'Guest authentication / PowerShell'; LocalStatus = 'SKIP'; Detail = $probe.Reason })
            continue
        }

        Write-Log "'$vmName': staging + running read-only verification payload." 'INFO'
        $payloadForGuest = $PayloadTemplate.Replace('{{TREAT_ALL_AS_WORKGROUP}}', $(if ($TreatAllAsWorkgroup) { '$true' } else { '$false' }))
        $payloadForGuest = $payloadForGuest.Replace('{{CACHED_TARGET}}', ([string][int]$CachedLogonsTarget))
        $payloadForGuest = $payloadForGuest.Replace('{{SKIP_EXTENDED_BASELINE}}', $(if ($SkipExtendedBaseline) { '$true' } else { '$false' }))
        $payloadForGuest = $payloadForGuest.Replace('{{SKIP_ISO}}', $(if ($SkipIsoPolicy) { '$true' } else { '$false' }))
        $payloadForGuest = $payloadForGuest.Replace('{{SKIP_IE}}', $(if ($SkipIeHardening) { '$true' } else { '$false' }))
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $payloadForGuest -ToolsWaitSecs $ToolsWaitSecs

        $envelope = Read-EnvelopeFromScriptOutput -Output $scriptOutput -StartMarker '<<<SECGAP-ENVELOPE-B64>>>' -EndMarker '<<<END-SECGAP-ENVELOPE>>>'
        if ($null -eq $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($scriptOutput) { ($scriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest payload result' -DetectedValue '' -ExpectedValue 'Base64 result envelope' `
                -Status 'Unable to Check' -Action 'Skipped - unparseable result' -RequiresReboot $false -RiskNote '' -Reason "No parseable result envelope. Output start: $snippet"))
            Write-Log "'$vmName': no parseable envelope returned." 'ERROR'
            Show-VmResults -VmBanner "$vmName  [SKIPPED - not processed]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = 'Guest payload result'; LocalStatus = 'SKIP'; Detail = "No parseable result envelope. Output start: $snippet" })
            continue
        }

        $meta  = $envelope.Meta
        $ghost = if ($meta.Hostname)  { $meta.Hostname }  else { $vm.Guest.HostName }
        $gip   = if ($meta.IPAddress) { $meta.IPAddress } else { ($vm.Guest.IPAddress -join ',') }
        $gos   = if ($meta.OS)        { $meta.OS }        else { $vm.Guest.OSFullName }
        if ($meta.FatalError) { Write-Log "'$vmName': guest payload reported a fatal error: $($meta.FatalError)" 'ERROR' }

        $summary.Checked++
        $f = @{ NonCompliant = $false; Manual = $false; Unable = $false; NA = $false }
        foreach ($r in @($envelope.Rows)) {
            switch ($r.Status) {
                'Non-Compliant'                { $f.NonCompliant = $true }
                'Manual Verification Required' { $f.Manual = $true }
                'Unable to Check'              { $f.Unable = $true }
                'Not Applicable'               { $f.NA = $true }
            }
            $reason = if ($r.Status -in 'Manual Verification Required', 'Unable to Check', 'Non-Compliant', 'Not Applicable') { $r.Detail } else { '' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category $r.Category -SecurityCheck $r.Item -DetectedValue $r.DetectedValue -ExpectedValue $r.ExpectedValue `
                -Status $r.Status -Action 'Verification only (read-only)' -RequiresReboot $false -RiskNote '' -Reason $reason))
        }
        if ($f.NonCompliant) { $summary.NonCompliant++ }
        if ($f.Manual)       { $summary.Manual++ }
        if ($f.NA)           { $summary.NotApplicable++ }
        if ($f.Unable)       { $summary.Unable++ }
        if (-not $f.NonCompliant) { $summary.Compliant++ }

        # On-screen result for this VM, formatted exactly like the local
        # post-remediation-check.ps1 (in addition to the central CSV).
        Show-VmResults -VmBanner "$vmName  @ $($srv.Name)   ($ghost / $gip)" -Rows @($envelope.Rows)

        Write-Log "'$vmName': verified. nonCompliant=$($f.NonCompliant) manual=$($f.Manual) unable=$($f.Unable)" 'INFO'
    }
    catch {
        $summary.SkippedFailed++
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        try {
            $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -SecurityCheck 'Processing' -DetectedValue '' -ExpectedValue '' `
                -Status 'Unable to Check' -Action 'Skipped - exception' -RequiresReboot $false -RiskNote '' -Reason $_.Exception.Message))
            Show-VmResults -VmBanner "$vmName  [SKIPPED - error]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = 'Processing'; LocalStatus = 'SKIP'; Detail = $_.Exception.Message })
        } catch { }
        continue
    }
}

# =====================================================================================
# 5. Central report + console summary
# =====================================================================================
$centralRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV written: $CsvPath ($($centralRows.Count) rows)."

Write-Host ""
Write-Host "============= POST-CHECK SUMMARY =============" -ForegroundColor Green
Write-Host ("  Total VMs                     : {0}" -f $summary.Total)
Write-Host ("  Successfully Checked          : {0}" -f $summary.Checked)
Write-Host ("  Compliant (no gaps found)     : {0}" -f $summary.Compliant)
Write-Host ("  Non-Compliant (gaps found)    : {0}" -f $summary.NonCompliant) -ForegroundColor $(if ($summary.NonCompliant) { 'Yellow' } else { 'Gray' })
Write-Host ("  Manual Verification Required  : {0}" -f $summary.Manual) -ForegroundColor DarkYellow
Write-Host ("  Not Applicable (workgroup etc): {0}" -f $summary.NotApplicable) -ForegroundColor Gray
Write-Host ("  Unable to Check (some items)  : {0}" -f $summary.Unable) -ForegroundColor $(if ($summary.Unable) { 'Yellow' } else { 'Gray' })
Write-Host ("  Skipped / Failed to process   : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Report : $CsvPath"
Write-Host "Log    : $LogPath"
Write-Log "Post-check run complete."
