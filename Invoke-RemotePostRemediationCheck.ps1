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

Write-Host "Invoke-RemotePostRemediationCheck.ps1 - READ-ONLY verification" -ForegroundColor Magenta

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

$Rows = New-Object System.Collections.Generic.List[object]
function Add-Row {
    param([string]$Category, [string]$Item, [string]$Status, [string]$Detected = '', [string]$Expected = '', [string]$Details = '')
    $Rows.Add([PSCustomObject]@{
        Category = $Category; Item = $Item; Status = $Status
        DetectedValue = $Detected; ExpectedValue = $Expected; Details = $Details
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
    # === VBS & Credential Protection ===
    $r = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' 'LsaCfgFlags'
    Add-Row 'VBS & Credential Protection' 'Credential Guard (LsaCfgFlags)' `
        $(if ($r.Found -and $r.Value -in @(1, 2)) { 'Compliant' } else { 'Non-Compliant' }) `
        "LsaCfgFlags=$(if ($r.Found) { $r.Value } else { '(not set)' })" '1 or 2'

    $vbs  = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity'
    $hvci = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' 'Enabled'
    $hvciVbsOk = ($vbs.Found -and $vbs.Value -eq 1) -and ($hvci.Found -and $hvci.Value -eq 1)
    Add-Row 'VBS & Credential Protection' 'HVCI / VBS' `
        $(if ($hvciVbsOk) { 'Compliant' } else { 'Non-Compliant' }) `
        "EnableVirtualizationBasedSecurity=$(if ($vbs.Found) { $vbs.Value } else { '(not set)' }); HVCI.Enabled=$(if ($hvci.Found) { $hvci.Value } else { '(not set)' })" `
        'VBS=1 and HVCI.Enabled=1'

    $r = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' 'MachineIdentityIsolation'
    Add-Row 'VBS & Credential Protection' 'Machine Identity Isolation' `
        $(if ($r.Found -and $r.Value -eq 2) { 'Compliant' } else { 'Non-Compliant' }) `
        "MachineIdentityIsolation=$(if ($r.Found) { $r.Value } else { '(not set)' })" '2 (Enforcement Mode)'

    # === Microsoft Defender ===
    try {
        $status = Get-MpComputerStatus -ErrorAction Stop
        Add-Row 'Microsoft Defender' 'Real-Time Protection' `
            $(if ($status.RealTimeProtectionEnabled) { 'Compliant' } else { 'Non-Compliant' }) `
            "RealTimeProtectionEnabled=$($status.RealTimeProtectionEnabled)" 'True'
    } catch {
        Add-Row 'Microsoft Defender' 'Real-Time Protection' 'Unable to Check' 'Get-MpComputerStatus failed' 'True' $_.Exception.Message
    }
    try {
        $pref = Get-MpPreference -ErrorAction Stop
        Add-Row 'Microsoft Defender' 'Cloud / MAPS Reporting' `
            $(if ($pref.MAPSReporting -eq 2) { 'Compliant' } else { 'Non-Compliant' }) `
            "MAPSReporting=$($pref.MAPSReporting)" '2 (Advanced)'
        $asrCount = @($pref.AttackSurfaceReductionRules_Ids).Count
        Add-Row 'Microsoft Defender' 'ASR Rules Configured' `
            $(if ($asrCount -gt 0) { 'Compliant' } else { 'Non-Compliant' }) `
            "$asrCount rule(s) configured" '>= 1 rule (original check); recommended set = 16 rules at Enabled' `
            'Original local check only verifies that at least one ASR rule exists, not that the full recommended set is at the intended action.'
    } catch {
        Add-Row 'Microsoft Defender' 'Defender Preferences (MAPS / ASR)' 'Unable to Check' 'Get-MpPreference failed' '' $_.Exception.Message
    }
    Add-Row 'Microsoft Defender' 'Tamper Protection' 'Manual Verification Required' `
        'Not scriptable' 'Enabled' 'Cannot be read/set by script - verify in the Windows Security app or Intune / Defender for Endpoint.'

    # === Authentication Security ===
    $r = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential'
    Add-Row 'Authentication Security' 'WDigest Disabled' `
        $(if ($r.Found -and $r.Value -eq 0) { 'Compliant' } else { 'Non-Compliant' }) `
        "UseLogonCredential=$(if ($r.Found) { $r.Value } else { '(not set)' })" '0'

    $r = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' 'SupportedEncryptionTypes'
    $aesOk = $r.Found -and (([int]$r.Value) -band 0x18) -eq 0x18
    Add-Row 'Authentication Security' 'Kerberos AES128/256 Supported' `
        $(if ($aesOk) { 'Compliant' } else { 'Non-Compliant' }) `
        "SupportedEncryptionTypes=$(if ($r.Found) { ('0x{0:X} ({0})' -f [int]$r.Value) } else { '(not set)' })" `
        'bits 0x18 (AES128+AES256) set'

    Add-Row 'Authentication Security' 'Kerberos SHA256/SHA384/SHA512 (AES-SHA2 / RFC 8009)' 'Manual Verification Required' `
        'Not verifiable by registry check' 'AES-SHA2 suites (etype 19/20) negotiated' `
        'No confirmed local registry mechanism. Use the remote Kerberos check (Invoke-RemoteKerberosCheck.ps1) against your Domain Controllers for event-based evidence (Event ID 4768 TicketEncryptionType 0x13/0x14).'

    # === Remote Management ===
    $winrm = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service'
    $u = Get-RegVal $winrm 'AllowUnencrypted'
    $b = Get-RegVal $winrm 'auth_Basic'
    Add-Row 'Remote Management' 'WinRM Unencrypted Traffic Blocked' `
        $(if ($u.Found -and $u.Value -eq 0) { 'Compliant' } else { 'Non-Compliant' }) `
        "AllowUnencrypted=$(if ($u.Found) { $u.Value } else { '(not set)' })" '0'
    Add-Row 'Remote Management' 'WinRM Basic Auth Blocked' `
        $(if ($b.Found -and $b.Value -eq 0) { 'Compliant' } else { 'Non-Compliant' }) `
        "auth_Basic=$(if ($b.Found) { $b.Value } else { '(not set)' })" '0'
    try {
        $rules = Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' }
        $addr = @($rules | Get-NetFirewallAddressFilter -ErrorAction Stop | Select-Object -ExpandProperty RemoteAddress -Unique)
        if ($addr.Count -eq 0) {
            Add-Row 'Remote Management' 'WinRM Listener Restricted to Management Sources' 'Manual Verification Required' 'No enabled WinRM firewall rules found' 'Scoped to management subnets/hosts' 'Verify firewall rules manually.'
        } elseif ($addr.Count -eq 1 -and $addr[0] -eq 'Any') {
            Add-Row 'Remote Management' 'WinRM Listener Restricted to Management Sources' 'Manual Verification Required' 'RemoteAddress=Any' 'Scoped to management subnets/hosts' 'Currently open to any source - supply -WinRmAllowedSourceRange to the remediation script and re-apply.'
        } else {
            Add-Row 'Remote Management' 'WinRM Listener Restricted to Management Sources' 'Compliant' "RemoteAddress=$($addr -join ', ')" 'Scoped to management subnets/hosts'
        }
    } catch {
        Add-Row 'Remote Management' 'WinRM Listener Restricted to Management Sources' 'Unable to Check' 'Get-NetFirewallRule failed' 'Scoped to management subnets/hosts' $_.Exception.Message
    }

    # === SMB & RPC Security ===
    try {
        $smb = Get-SmbServerConfiguration -ErrorAction Stop
        Add-Row 'SMB & RPC Security' 'SMB1 Access Auditing' `
            $(if ($smb.AuditSmb1Access) { 'Compliant' } else { 'Non-Compliant' }) `
            "AuditSmb1Access=$($smb.AuditSmb1Access)" 'True'
    } catch {
        Add-Row 'SMB & RPC Security' 'SMB1 Access Auditing' 'Unable to Check' 'Get-SmbServerConfiguration failed' 'True' $_.Exception.Message
    }
    $r = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' 'RpcAuthnLevelPrivacyEnabled'
    Add-Row 'SMB & RPC Security' 'Printer RPC Packet Privacy' `
        $(if ($r.Found -and $r.Value -eq 1) { 'Compliant' } else { 'Non-Compliant' }) `
        "RpcAuthnLevelPrivacyEnabled=$(if ($r.Found) { $r.Value } else { '(not set)' })" '1'
    $r = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' 'RestrictDriverInstallationToAdministrators'
    Add-Row 'SMB & RPC Security' 'Point and Print Restricted' `
        $(if ($r.Found -and $r.Value -eq 1) { 'Compliant' } else { 'Non-Compliant' }) `
        "RestrictDriverInstallationToAdministrators=$(if ($r.Found) { $r.Value } else { '(not set)' })" '1'

    # === Security Baseline ===
    $r = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' 'EnableSmartScreen'
    Add-Row '2025 Security Baseline' 'SmartScreen Enabled' `
        $(if ($r.Found -and $r.Value -eq 1) { 'Compliant' } else { 'Non-Compliant' }) `
        "EnableSmartScreen=$(if ($r.Found) { $r.Value } else { '(not set)' })" '1'

    $z = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'
    $z1200 = Get-RegVal $z '1200'; $z1400 = Get-RegVal $z '1400'; $z2500 = Get-RegVal $z '2500'
    $ieOk = ($z1200.Found -and $z1200.Value -eq 3) -and ($z1400.Found -and $z1400.Value -eq 3) -and ($z2500.Found -and $z2500.Value -eq 0)
    Add-Row '2025 Security Baseline' 'IE Zone (ActiveX / Scripting / Protected Mode)' `
        $(if ($ieOk) { 'Compliant' } else { 'Non-Compliant' }) `
        "1200=$(if ($z1200.Found) { $z1200.Value } else { '(ns)' }); 1400=$(if ($z1400.Found) { $z1400.Value } else { '(ns)' }); 2500=$(if ($z2500.Found) { $z2500.Value } else { '(ns)' })" `
        '1200=3, 1400=3, 2500=0'

    try {
        $feat = Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction Stop
        if (-not $feat) {
            Add-Row '2025 Security Baseline' 'Legacy Internet Explorer 11 Removed' 'Compliant' 'Feature not present on this build' 'Disabled / not present'
        } elseif ($feat.State -eq 'Disabled') {
            Add-Row '2025 Security Baseline' 'Legacy Internet Explorer 11 Removed' 'Compliant' 'State=Disabled' 'Disabled / not present'
        } else {
            Add-Row '2025 Security Baseline' 'Legacy Internet Explorer 11 Removed' 'Non-Compliant' "State=$($feat.State)" 'Disabled / not present'
        }
    } catch {
        Add-Row '2025 Security Baseline' 'Legacy Internet Explorer 11 Removed' 'Compliant' 'Feature does not exist on this build' 'Disabled / not present' $_.Exception.Message
    }

    $r = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' 'Enabled'
    Add-Row '2025 Security Baseline' 'Legacy SSL/TLS Disabled (spot check: TLS 1.0 Server)' `
        $(if ($r.Found -and $r.Value -eq 0) { 'Compliant' } else { 'Manual Verification Required' }) `
        "TLS 1.0\Server\Enabled=$(if ($r.Found) { $r.Value } else { '(not set)' })" 'Enabled=0 (opt-in remediation)' `
        'Original local check inspects only TLS 1.0 (Server). SSL 2.0/3.0, TLS 1.1 and the client side are not covered here; this is an opt-in remediation item (-DisableLegacyTls).'

    Add-Row '2025 Security Baseline' 'Full Windows Server 2025 Security Baseline GPO Applied' 'Manual Verification Required' `
        'Not verifiable by registry check' 'Baseline GPO linked / applied' 'Confirm via gpresult /h or by re-running LGPO.exe /g against the official baseline backup.'

} catch {
    $fatal = $_.Exception.Message
}

$meta = [PSCustomObject]@{ Hostname = $ComputerName; IPAddress = $IPAddresses; OS = $OSCaption; FatalError = $fatal }
# .ToArray() rather than @(...) - @() around a generic List instance is unreliable
# on some PowerShell builds; .ToArray() is well-defined on 5.1 and 7.x alike.
$envelope = [PSCustomObject]@{ Schema = 'secgap-postcheck-1'; Meta = $meta; Rows = $Rows.ToArray() }
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

function New-CentralRow {
    param($vCenter, $VMName, $GuestHostname, $IPAddress, $OS, $Category, $SecurityCheck, $DetectedValue, $ExpectedValue, $Status, $Action, $RequiresReboot, $RiskNote, $Reason)
    [PSCustomObject]([ordered]@{
        vCenter = $vCenter; VMName = $VMName; GuestHostname = $GuestHostname; IPAddress = $IPAddress; OS = $OS
        Category = $Category; SecurityCheck = $SecurityCheck; DetectedValue = $DetectedValue; ExpectedValue = $ExpectedValue
        Status = $Status; Action = $Action; RequiresReboot = $RequiresReboot; RiskNote = $RiskNote
        'Error/Reason' = $Reason; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
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
            continue
        }

        Write-Log "'$vmName': running read-only verification payload." 'INFO'
        $run = Invoke-VMScript -VM $vm -Server $srv -GuestCredential $GuestCredential -ScriptText $PayloadTemplate `
                               -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop

        $envelope = Read-EnvelopeFromScriptOutput -Output $run.ScriptOutput -StartMarker '<<<SECGAP-ENVELOPE-B64>>>' -EndMarker '<<<END-SECGAP-ENVELOPE>>>'
        if ($null -eq $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($run.ScriptOutput) { ($run.ScriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest payload result' -DetectedValue '' -ExpectedValue 'Base64 result envelope' `
                -Status 'Unable to Check' -Action 'Skipped - unparseable result' -RequiresReboot $false -RiskNote '' -Reason "No parseable result envelope. Output start: $snippet"))
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
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category $r.Category -SecurityCheck $r.Item -DetectedValue $r.DetectedValue -ExpectedValue $r.ExpectedValue `
                -Status $r.Status -Action 'Verification only (read-only)' -RequiresReboot $false -RiskNote '' -Reason $r.Details))
        }
        if ($f.NonCompliant) { $summary.NonCompliant++ }
        if ($f.Manual)       { $summary.Manual++ }
        if ($f.Unable)       { $summary.Unable++ }
        if (-not $f.NonCompliant) { $summary.Compliant++ }
        Write-Log "'$vmName': verified. nonCompliant=$($f.NonCompliant) manual=$($f.Manual) unable=$($f.Unable)" 'INFO'
    }
    catch {
        $summary.SkippedFailed++
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        try {
            $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -SecurityCheck 'Processing' -DetectedValue '' -ExpectedValue '' `
                -Status 'Unable to Check' -Action 'Skipped - exception' -RequiresReboot $false -RiskNote '' -Reason $_.Exception.Message))
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
