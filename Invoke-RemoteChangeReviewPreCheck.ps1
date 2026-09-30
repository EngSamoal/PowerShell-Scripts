#Requires -Version 5.1
<#
.SYNOPSIS
    Centralised (PowerCLI) front-end for the change-review pre-check.

    Strictly READ-ONLY. Runs the same 6 checks as the local
    Invoke-ChangeReviewPreCheck.ps1 against every VM in a list, from a single
    admin workstation, over VMware Tools Guest Operations. No WinRM / PsExec /
    SMB / RDP / guest network path is used.

.DESCRIPTION
    Covers the 6 "Script - Review" items from the remediation action plan:
    Windows Firewall, SMB minimum version, NetBIOS, Account Lockout Threshold,
    UAC hardening, and VBS & Credential Guard prerequisites. For each item it
    reports one of:
        OK      - safe to enable the matching -Apply switch
        MANUAL  - review with the relevant owner before enabling it
    together with the detected detail, so the CSV is auditable. This script
    never calls Set-*, New-*, Remove-*, Stop-*, Disable-* or Add-* - it only
    reads.

    Same per-VM validation battery as the remediation/post-check wrappers
    (existence, non-ambiguity across connected vCenters, powered on, Windows,
    VMware Tools running, guest auth, PowerShell 5.1+). A validation failure
    skips only that VM and is recorded with an exact reason - one bad VM never
    stops the batch.

.PARAMETER VMListPath      Text file of VM names (default C:\temp\vmlist.txt). '#' lines ignored.
.PARAMETER CredentialPath  Export-Clixml PSCredential for the guest admin (default C:\temp\wincred.xml). In-memory only.
.PARAMETER OutputPath      Folder on the admin machine for the CSV + log.
.PARAMETER ToolsWaitSecs   Invoke-VMScript VMware Tools wait, seconds (default 180).

.EXAMPLE
    .\Invoke-RemoteChangeReviewPreCheck.ps1
    .\Invoke-RemoteChangeReviewPreCheck.ps1 -VMListPath C:\temp\vmlist.txt -OutputPath D:\Reports

.NOTES
    Prerequisites: PowerCLI imported; already Connect-VIServer'd to the target
    vCenter(s); vCenter account has the three "Guest Operation ..." privileges.
    This script issues no vSphere write of any kind, and makes no change
    inside any guest either. Run it BEFORE enabling -EnableFirewallProfiles,
    -SetSmbMinSmb3, -DisableNetbios, -ApplyAccountPolicy, -EnableUacHardening
    or the VBS/Credential Guard items in Invoke-RemoteSecurityGapRemediation.ps1.
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

Write-Host "Invoke-RemoteChangeReviewPreCheck.ps1 - READ-ONLY verification  [build 2026-10-01a-precheck-remote]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
Write-Host ("If the build above is not '2026-10-01a-precheck-remote' you are running an OLD copy - update it.") -ForegroundColor DarkGray

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
$CsvPath  = Join-Path $OutputPath "SecurityGap_PreCheckReview_$RunStamp.csv"
$LogPath  = Join-Path $OutputPath "SecurityGap_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Change-review pre-check run started. VMs=$($vmNames.Count)."

# =====================================================================================
# 2. In-guest READ-ONLY payload - the 6 "Script - Review" checks, elevated recommended.
#    Never calls Set-*/New-*/Remove-*/Stop-*/Disable-*/Add-*.
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'

$Rows = New-Object System.Collections.Generic.List[object]
function Add-Row {
    param(
        [string]$Header, [string]$Label, [string]$LocalStatus, [string]$Detail = '',
        [string]$Category, [string]$Item, [string]$NormStatus = '',
        [string]$Detected = '', [string]$Expected = ''
    )
    if (-not $Item) { $Item = $Label }
    if (-not $NormStatus) {
        if     ($LocalStatus -like 'OK*')     { $NormStatus = 'OK - safe to proceed' }
        elseif ($LocalStatus -like 'MANUAL*') { $NormStatus = 'Review Required' }
        else                                  { $NormStatus = 'Unable to Check' }
    }
    $Rows.Add([PSCustomObject]@{
        Header = $Header; Label = $Label; LocalStatus = $LocalStatus; Detail = $Detail
        Category = $Category; Item = $Item; Status = $NormStatus
        DetectedValue = $Detected; ExpectedValue = $Expected
    }) | Out-Null
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

$IsElevated = $false
try { $IsElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) } catch { }

$fatal = $null
try {
    if (-not $IsElevated) {
        Add-Row -Header 'Session' -Label 'Elevation' -LocalStatus 'MANUAL' `
            -Detail 'Guest session is not elevated - the firewall and SMB session checks below may show incomplete results. Re-run with an administrator credential for complete results.' `
            -Category 'Session' -Item 'Elevation' -Detected "Elevated=$IsElevated" -Expected 'Elevated=True'
    }

    # ==========================================================================
    $H = '1. Windows Firewall - inbound allow rules (before enabling -EnableFirewallProfiles)'
    try {
        $rules = @(Get-NetFirewallRule -Enabled True -Direction Inbound -ErrorAction Stop | Where-Object { $_.Action -eq 'Allow' })
        $rdp   = $rules | Where-Object { $_.DisplayName -match 'Remote Desktop' }
        $ports = @($rules | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue | Where-Object { $_.LocalPort -and $_.LocalPort -ne 'Any' } | Select-Object -ExpandProperty LocalPort -Unique | Select-Object -First 10)
        Add-Row -Header $H -Label 'RDP (3389) allow rule present' `
            -LocalStatus $(if ($rdp) { 'OK' } else { 'MANUAL' }) `
            -Detail $(if ($rdp) { "$($rules.Count) enabled inbound allow rule(s) total." } else { "No enabled RDP rule found among $($rules.Count) allow rule(s) - add one before enabling the firewall or you may lose remote access." }) `
            -Category '1. Windows Firewall' -Item 'RDP allow rule present' `
            -Detected ("{0} inbound allow rule(s); sample ports: {1}" -f $rules.Count, ($ports -join ', ')) -Expected 'RDP (3389) covered by an enabled allow rule'
    } catch {
        Add-Row -Header $H -Label 'Firewall rule check' -LocalStatus 'MANUAL' -Detail "Could not query firewall rules: $($_.Exception.Message)" `
            -Category '1. Windows Firewall' -Item 'Firewall rule check' -NormStatus 'Unable to Check' -Detected 'Get-NetFirewallRule failed' -Expected 'Enumerable inbound allow rules'
    }

    # ==========================================================================
    $H = '2. SMB minimum version - active sessions (before enabling -SetSmbMinSmb3)'
    try {
        $sessions = @(Get-SmbSession -ErrorAction Stop)
        if ($sessions.Count -eq 0) {
            Add-Row -Header $H -Label 'Clients on SMB 2.0.2/2.1' -LocalStatus 'OK' -Detail 'No active SMB sessions right now - also check scheduled/off-hours jobs (backup, NAS) separately.' `
                -Category '2. SMB Minimum Version' -Item 'Active SMB session dialects' -Detected '0 active sessions' -Expected 'No client on dialect 2.0.2/2.1'
        } else {
            $old = @($sessions | Where-Object { $_.Dialect -in '2.0.2','2.1' })
            $summary = ($sessions | Group-Object Dialect | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ', '
            Add-Row -Header $H -Label 'Clients on SMB 2.0.2/2.1' `
                -LocalStatus $(if ($old) { 'MANUAL' } else { 'OK' }) `
                -Detail $(if ($old) { "$($old.Count) of $($sessions.Count) session(s) on an older dialect - they will lose SMB access if the minimum is raised to 3.0.0: $(($old.ClientComputerName | Select-Object -Unique) -join ', ')" } else { "$($sessions.Count) active session(s), all SMB 3.0+." }) `
                -Category '2. SMB Minimum Version' -Item 'Active SMB session dialects' -Detected $summary -Expected 'No client on dialect 2.0.2/2.1'
        }
    } catch {
        Add-Row -Header $H -Label 'SMB session check' -LocalStatus 'MANUAL' -Detail "Could not query SMB sessions (requires elevation): $($_.Exception.Message)" `
            -Category '2. SMB Minimum Version' -Item 'Active SMB session dialects' -NormStatus 'Unable to Check' -Detected 'Get-SmbSession failed' -Expected 'No client on dialect 2.0.2/2.1'
    }

    # ==========================================================================
    $H = '3. NetBIOS - DNS resolvability (before enabling -DisableNetbios)'
    try {
        $dns = Resolve-DnsName -Name $ComputerName -ErrorAction Stop
        Add-Row -Header $H -Label 'DNS resolves this hostname' -LocalStatus 'OK' `
            -Detail ("Resolved to: {0}" -f (($dns | Select-Object -ExpandProperty IPAddress) -join ', ')) `
            -Category '3. NetBIOS' -Item 'DNS resolves hostname' -Detected 'DNS lookup succeeded' -Expected 'DNS lookup succeeds'
    } catch {
        Add-Row -Header $H -Label 'DNS resolves this hostname' -LocalStatus 'MANUAL' `
            -Detail 'DNS lookup failed - if anything reaches this server by NetBIOS/short name only, disabling NetBIOS will break it. Verify DNS records exist first.' `
            -Category '3. NetBIOS' -Item 'DNS resolves hostname' -Detected 'DNS lookup failed' -Expected 'DNS lookup succeeds'
    }

    # ==========================================================================
    $H = '4. Account Lockout Threshold - current policy (before enabling -ApplyAccountPolicy)'
    try {
        $na  = & net accounts 2>$null
        $lt  = ($na | Select-String 'Lockout threshold').ToString().Trim()
        $ld  = ($na | Select-String 'Lockout duration').ToString().Trim()
        Add-Row -Header $H -Label 'Current lockout policy captured' -LocalStatus 'MANUAL' `
            -Detail 'Manually confirm with the app/service owners which local or service accounts (e.g. Splunk, ServiceNow, backup agents) log in frequently enough to risk hitting a lower threshold before lowering it.' `
            -Category '4. Account Lockout Threshold' -Item 'Current lockout policy' -Detected ("$lt; $ld") -Expected 'Reviewed with account owners'
    } catch {
        Add-Row -Header $H -Label 'Account policy check' -LocalStatus 'MANUAL' -Detail "Could not read account policy: $($_.Exception.Message)" `
            -Category '4. Account Lockout Threshold' -Item 'Current lockout policy' -NormStatus 'Unable to Check' -Detected 'net accounts failed' -Expected 'Reviewed with account owners'
    }

    # ==========================================================================
    $H = '5. UAC Hardening - elevated scheduled tasks (before enabling -EnableUacHardening)'
    try {
        $tasks = @(Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.Principal.UserId -match 'Administrator|SYSTEM' -or $_.Principal.RunLevel -eq 'Highest' })
        $sample = ($tasks | Select-Object -First 5 -ExpandProperty TaskName) -join '; '
        Add-Row -Header $H -Label 'Elevated scheduled tasks' `
            -LocalStatus $(if ($tasks.Count -gt 0) { 'MANUAL' } else { 'OK' }) `
            -Detail $(if ($tasks.Count -gt 0) { "$($tasks.Count) task(s) run elevated/as Administrator - confirm none depend on silent (no-prompt) elevation. Examples: $sample" } else { 'No elevated scheduled tasks found.' }) `
            -Category '5. UAC Hardening' -Item 'Elevated scheduled tasks' -Detected ("$($tasks.Count) elevated task(s)") -Expected 'No task relies on silent elevation'
    } catch {
        Add-Row -Header $H -Label 'Scheduled task check' -LocalStatus 'MANUAL' -Detail "Could not query scheduled tasks: $($_.Exception.Message)" `
            -Category '5. UAC Hardening' -Item 'Elevated scheduled tasks' -NormStatus 'Unable to Check' -Detected 'Get-ScheduledTask failed' -Expected 'No task relies on silent elevation'
    }

    # ==========================================================================
    $H = '6. VBS and Credential Guard - platform prerequisites (before this reboot-required item)'
    try {
        $sb = Confirm-SecureBootUEFI -ErrorAction Stop
        Add-Row -Header $H -Label 'Secure Boot enabled' -LocalStatus $(if ($sb) { 'OK' } else { 'MANUAL' }) `
            -Detail $(if (-not $sb) { 'Secure Boot is OFF - enable it on the VM in vCenter before running the VBS/Credential Guard items, or they will stay "configured, not running".' }) `
            -Category '6. VBS and Credential Guard' -Item 'Secure Boot enabled' -Detected "SecureBoot=$sb" -Expected 'SecureBoot=True'
    } catch {
        Add-Row -Header $H -Label 'Secure Boot enabled' -LocalStatus 'MANUAL' `
            -Detail 'Could not determine Secure Boot state (not UEFI, or access denied) - verify the VM firmware is set to EFI with Secure Boot in vCenter.' `
            -Category '6. VBS and Credential Guard' -Item 'Secure Boot enabled' -NormStatus 'Unable to Check' -Detected 'Confirm-SecureBootUEFI failed' -Expected 'SecureBoot=True'
    }
    try {
        $ci = Get-ComputerInfo -ErrorAction Stop
        $flags = [ordered]@{
            'Firmware virtualization enabled'  = [bool]$ci.HyperVRequirementVirtualizationFirmwareEnabled
            'Second Level Address Translation'  = [bool]$ci.HyperVRequirementSecondLevelAddressTranslation
            'Data Execution Prevention available' = [bool]$ci.HyperVRequirementDataExecutionPreventionAvailable
        }
        $failing = @($flags.GetEnumerator() | Where-Object { -not $_.Value } | ForEach-Object { $_.Key })
        Add-Row -Header $H -Label 'Hardware virtualization exposed to guest' `
            -LocalStatus $(if ($failing.Count -eq 0) { 'OK' } else { 'MANUAL' }) `
            -Detail $(if ($failing.Count -gt 0) { "Not met: $($failing -join ', '). Expose Secure Boot + hardware virtualization to this VM in vCenter first - VBS cannot start without them." }) `
            -Category '6. VBS and Credential Guard' -Item 'Hardware virtualization exposed to guest' `
            -Detected (($flags.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ') -Expected 'All three = True'
    } catch {
        Add-Row -Header $H -Label 'Hardware virtualization exposed to guest' -LocalStatus 'MANUAL' -Detail "Could not query platform requirements: $($_.Exception.Message)" `
            -Category '6. VBS and Credential Guard' -Item 'Hardware virtualization exposed to guest' -NormStatus 'Unable to Check' -Detected 'Get-ComputerInfo failed' -Expected 'All three = True'
    }

} catch {
    $fatal = $_.Exception.Message
}

$meta = [PSCustomObject]@{ Hostname = $ComputerName; IPAddress = $IPAddresses; OS = $OSCaption; Elevated = $IsElevated; FatalError = $fatal }
# .ToArray() rather than @(...) - @() around a generic List instance is unreliable
# on some PowerShell builds; .ToArray() is well-defined on 5.1 and 7.x alike.
$envelope = [PSCustomObject]@{ Schema = 'secgap-precheck-1'; Meta = $meta; Rows = $Rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<SECGAP-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-SECGAP-ENVELOPE>>>'
'@

# =====================================================================================
# 3. Admin-side helpers (identical to the remediation/post-check wrappers)
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

# Same rationale as the remediation/post-check wrappers: a single Invoke-VMScript
# carrying the whole payload can fail with a misleading "Could not locate
# Powershell script interpreter" error because the script is placed on a command
# line that overruns a guest-ops length limit. Primary method: Copy-VMGuestFile
# (one transfer, no limit). Fallback: write Base64 into a guest file in small
# pieces, then decode in-guest. Guest operations only - no WinRM/PsExec/SMB/
# ESXi-or-guest network path. The staged file holds no credential and is always
# deleted.
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
    param($vCenter, $VMName, $GuestHostname, $IPAddress, $OS, $Category, $SecurityCheck, $DetectedValue, $ExpectedValue, $Status, $Detail)
    [PSCustomObject]([ordered]@{
        vCenter = $vCenter; VMName = $VMName; GuestHostname = $GuestHostname; IPAddress = $IPAddress; OS = $OS
        Category = $Category; SecurityCheck = $SecurityCheck; DetectedValue = $DetectedValue; ExpectedValue = $ExpectedValue
        Status = $Status; Detail = $Detail; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

# On-screen output - same look as Invoke-RemotePostRemediationCheck.ps1's
# Show-VmResults: OK=Green, MANUAL=DarkYellow, else Red. No N/A track here.
function Show-VmResults {
    param([string]$VmBanner, $Rows, [int]$InlineDetailMax = 70, [int]$WrapWidth = 100)
    $HeaderColor   = 'Cyan'
    $SubPointColor = 'White'

    $statusWidth = 6
    foreach ($row in $Rows) {
        $l = ([string]$row.LocalStatus).Length
        if ($l -gt $statusWidth) { $statusWidth = $l }
    }
    $contIndent = ' ' * ($statusWidth + 5)

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
        Write-Host ("  [{0}] " -f $s.PadRight($statusWidth)) -ForegroundColor $statusColor -NoNewline
        Write-Host $row.Label -ForegroundColor $SubPointColor -NoNewline

        $text = if ($row.Detail) { (([string]$row.Detail) -replace '\s+', ' ').Trim() } else { '' }
        if (-not $text) {
            Write-Host ""
        } elseif ($text.Length -le $InlineDetailMax) {
            Write-Host (" - {0}" -f $text) -ForegroundColor Gray
        } else {
            Write-Host ""
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
$summary = [ordered]@{ Total = 0; Checked = 0; AllPass = 0; NeedsReview = 0; Unable = 0; SkippedFailed = 0 }
$inventory = Get-VMInventory -Servers $connectedServers

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Pre-checking VM '$vmName' ===" 'INFO'
    try {
        $val = Resolve-AndValidateVM -Name $vmName -Inventory $inventory
        if (-not $val.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -SecurityCheck $val.Stage -DetectedValue $val.Detected -ExpectedValue '' `
                -Status 'Unable to Check' -Detail $val.Reason))
            Write-Log "SKIP '$vmName' [$($val.Stage)]: $($val.Reason)" 'WARN'
            Show-VmResults -VmBanner "$vmName  [SKIPPED - not processed]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = $val.Stage; LocalStatus = 'MANUAL'; Detail = $val.Reason })
            continue
        }
        $vm = $val.VM; $srv = $val.Server

        $probe = Test-GuestPowerShell -VM $vm -Server $srv -Credential $GuestCredential -ToolsWaitSecs $ToolsWaitSecs
        if (-not $probe.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $vm.Guest.HostName -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest authentication / PowerShell' -DetectedValue '' -ExpectedValue 'PowerShell 5.1+ reachable via VMware Tools' `
                -Status 'Unable to Check' -Detail $probe.Reason))
            Write-Log "SKIP '$vmName' [probe]: $($probe.Reason)" 'WARN'
            Show-VmResults -VmBanner "$vmName  [SKIPPED - not processed]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = 'Guest authentication / PowerShell'; LocalStatus = 'MANUAL'; Detail = $probe.Reason })
            continue
        }

        Write-Log "'$vmName': staging + running read-only pre-check payload." 'INFO'
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $PayloadTemplate -ToolsWaitSecs $ToolsWaitSecs

        $envelope = Read-EnvelopeFromScriptOutput -Output $scriptOutput -StartMarker '<<<SECGAP-ENVELOPE-B64>>>' -EndMarker '<<<END-SECGAP-ENVELOPE>>>'
        if ($null -eq $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($scriptOutput) { ($scriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest payload result' -DetectedValue '' -ExpectedValue 'Base64 result envelope' `
                -Status 'Unable to Check' -Detail "No parseable result envelope. Output start: $snippet"))
            Write-Log "'$vmName': no parseable envelope returned." 'ERROR'
            Show-VmResults -VmBanner "$vmName  [SKIPPED - not processed]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = 'Guest payload result'; LocalStatus = 'MANUAL'; Detail = "No parseable result envelope. Output start: $snippet" })
            continue
        }

        $meta  = $envelope.Meta
        $ghost = if ($meta.Hostname)  { $meta.Hostname }  else { $vm.Guest.HostName }
        $gip   = if ($meta.IPAddress) { $meta.IPAddress } else { ($vm.Guest.IPAddress -join ',') }
        $gos   = if ($meta.OS)        { $meta.OS }        else { $vm.Guest.OSFullName }
        if ($meta.FatalError) { Write-Log "'$vmName': guest payload reported a fatal error: $($meta.FatalError)" 'ERROR' }

        $summary.Checked++
        $f = @{ Review = $false; Unable = $false }
        foreach ($r in @($envelope.Rows)) {
            if ($r.Status -eq 'Review Required') { $f.Review = $true }
            if ($r.Status -eq 'Unable to Check')  { $f.Unable = $true }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category $r.Category -SecurityCheck $r.Item -DetectedValue $r.DetectedValue -ExpectedValue $r.ExpectedValue `
                -Status $r.Status -Detail $r.Detail))
        }
        if ($f.Review) { $summary.NeedsReview++ } else { $summary.AllPass++ }
        if ($f.Unable) { $summary.Unable++ }

        Show-VmResults -VmBanner "$vmName  @ $($srv.Name)   ($ghost / $gip)" -Rows @($envelope.Rows)
        Write-Log "'$vmName': pre-checked. needsReview=$($f.Review) unable=$($f.Unable)" 'INFO'
    }
    catch {
        $summary.SkippedFailed++
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        try {
            $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -SecurityCheck 'Processing' -DetectedValue '' -ExpectedValue '' `
                -Status 'Unable to Check' -Detail $_.Exception.Message))
            Show-VmResults -VmBanner "$vmName  [SKIPPED - error]" -Rows @([pscustomobject]@{ Header = 'Validation'; Label = 'Processing'; LocalStatus = 'MANUAL'; Detail = $_.Exception.Message })
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
Write-Host "============= CHANGE-REVIEW PRE-CHECK SUMMARY =============" -ForegroundColor Green
Write-Host ("  Total VMs                     : {0}" -f $summary.Total)
Write-Host ("  Successfully Checked          : {0}" -f $summary.Checked)
Write-Host ("  All Items Pass (safe to apply): {0}" -f $summary.AllPass) -ForegroundColor $(if ($summary.AllPass) { 'Green' } else { 'Gray' })
Write-Host ("  Items Needing Review          : {0}" -f $summary.NeedsReview) -ForegroundColor $(if ($summary.NeedsReview) { 'DarkYellow' } else { 'Gray' })
Write-Host ("  Unable to Check (some items)  : {0}" -f $summary.Unable) -ForegroundColor $(if ($summary.Unable) { 'Yellow' } else { 'Gray' })
Write-Host ("  Skipped / Failed to process   : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "=============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Report : $CsvPath"
Write-Host "Log    : $LogPath"
Write-Log "Change-review pre-check run complete."
