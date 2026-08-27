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

    [ValidateRange(30, 3600)]
    [int]    $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'

Write-Host "Invoke-RemotePostRemediationCheck.ps1 - READ-ONLY verification  [build 2026-08-28f-local-format]" -ForegroundColor Magenta
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
    Add-Row -Header $H -Label 'Machine Identity Isolation' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 2) { 'OK' } else { "NOT SET (found: $(if ($v.Found) { $v.Value } else { '' }))" }) `
        -Category 'VBS & Credential Protection' -Item 'Machine Identity Isolation' `
        -Detected ("MachineIdentityIsolation={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '2 (Enforcement Mode)'

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
        Add-Row -Header $H -Label 'SMB1 Access Auditing' `
            -LocalStatus $(if ($smb.AuditSmb1Access) { 'OK' } else { 'NOT SET' }) `
            -Category 'SMB & RPC Security' -Item 'SMB1 Access Auditing' `
            -Detected ("AuditSmb1Access={0}" -f $smb.AuditSmb1Access) -Expected 'True'
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
            Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'OK' -Detail 'N/A - this feature does not exist on this Windows Server build, nothing to remove' `
                -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -Detected 'Feature not present on this build' -Expected 'Disabled / not present'
        } elseif ($feat.State -eq 'Disabled') {
            Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'OK' `
                -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -Detected 'State=Disabled' -Expected 'Disabled / not present'
        } else {
            Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'NOT SET' `
                -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -Detected ("State={0}" -f $feat.State) -Expected 'Disabled / not present'
        }
    } catch {
        Add-Row -Header $H -Label 'Legacy IE11 Removed' -LocalStatus 'OK' -Detail 'N/A - this feature does not exist on this Windows Server build, nothing to remove' `
            -Category '2025 Security Baseline' -Item 'Legacy Internet Explorer 11 Removed' -NormStatus 'Compliant' -Detected 'Feature does not exist on this build' -Expected 'Disabled / not present'
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

# Reproduces the local post-remediation-check.ps1 on-screen output EXACTLY:
# a Cyan "=== <category header> ===" for each group, then one Show-Result line per
# item -> "  [<STATUS padded to 8>] <Label>[ - <Detail>]" coloured OK*=Green,
# MANUAL*=DarkYellow, everything else=Red, label in White, detail in Gray.
# The only addition is the "##### <VM> #####" banner so multi-VM runs stay readable.
function Show-VmResults {
    param([string]$VmBanner, $Rows)
    $HeaderColor   = 'Cyan'
    $SubPointColor = 'White'
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
            default   { 'Red' }
        }
        Write-Host ("  [{0}]" -f $s.PadRight(8)) -ForegroundColor $statusColor -NoNewline
        Write-Host (" {0}" -f $row.Label) -ForegroundColor $SubPointColor -NoNewline
        if ($row.Detail) { Write-Host (" - {0}" -f $row.Detail) -ForegroundColor Gray } else { Write-Host "" }
    }
}

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows = New-Object System.Collections.Generic.List[object]
$summary = [ordered]@{ Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0; Manual = 0; Unable = 0; SkippedFailed = 0 }
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
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $PayloadTemplate -ToolsWaitSecs $ToolsWaitSecs

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
        $f = @{ NonCompliant = $false; Manual = $false; Unable = $false }
        foreach ($r in @($envelope.Rows)) {
            switch ($r.Status) {
                'Non-Compliant'                { $f.NonCompliant = $true }
                'Manual Verification Required' { $f.Manual = $true }
                'Unable to Check'              { $f.Unable = $true }
            }
            $reason = if ($r.Status -in 'Manual Verification Required', 'Unable to Check', 'Non-Compliant') { $r.Detail } else { '' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category $r.Category -SecurityCheck $r.Item -DetectedValue $r.DetectedValue -ExpectedValue $r.ExpectedValue `
                -Status $r.Status -Action 'Verification only (read-only)' -RequiresReboot $false -RiskNote '' -Reason $reason))
        }
        if ($f.NonCompliant) { $summary.NonCompliant++ }
        if ($f.Manual)       { $summary.Manual++ }
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
Write-Host ("  Unable to Check (some items)  : {0}" -f $summary.Unable) -ForegroundColor $(if ($summary.Unable) { 'Yellow' } else { 'Gray' })
Write-Host ("  Skipped / Failed to process   : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Report : $CsvPath"
Write-Host "Log    : $LogPath"
Write-Log "Post-check run complete."
