#Requires -Version 5.1
<#
.SYNOPSIS
    Centralised (PowerCLI) front-end for the change-review POST-check.

    Strictly READ-ONLY. Verifies that the 6 "Script - Review" remediation
    items actually took effect, against every VM in a list, from a single
    admin workstation, over VMware Tools Guest Operations. Run this AFTER
    enabling -EnableFirewallProfiles / -SetSmbMinSmb3 / -DisableNetbios /
    -ApplyAccountPolicy / -EnableUacHardening / the VBS &Credential Guard
    items in Invoke-RemoteSecurityGapRemediation.ps1.

.DESCRIPTION
    Mirrors Invoke-RemoteChangeReviewPreCheck.ps1's 6 categories, but checks
    the END STATE instead of the pre-conditions:
        1. Windows Firewall           - profiles actually enabled
        2. SMB Minimum Version        - MinSmb2Dialect actually 768 (3.0.0)
        3. NetBIOS                    - EnableNetbios actually 0
        4. Account Lockout Threshold  - policy actually in range / raised
        5. UAC Hardening              - registry values actually set
        6. VBS & Credential Guard     - registry set AND actually running
    Reports OK / NOT SET / MANUAL per item. Never calls Set-*, New-*,
    Remove-*, Stop-*, Disable-* or Add-* - it only reads. Same per-VM
    validation battery and same console style as the other wrappers.

.PARAMETER VMListPath      Text file of VM names (default C:\temp\vmlist.txt). '#' lines ignored.
.PARAMETER CredentialPath  Export-Clixml PSCredential for the guest admin (default C:\temp\wincred.xml). In-memory only.
.PARAMETER OutputPath      Folder on the admin machine for the CSV + log.
.PARAMETER ToolsWaitSecs   Invoke-VMScript VMware Tools wait, seconds (default 180).
.PARAMETER MinPasswordLength       Target used to judge CID 1071 (default 14 - match what you passed to -MinPasswordLength).
.PARAMETER AccountLockoutThreshold Target used to judge CID 2342 (default 3 - match what you passed to -AccountLockoutThreshold).

.EXAMPLE
    .\Invoke-RemoteChangeReviewPostCheck.ps1

.NOTES
    Prerequisites: PowerCLI imported; already Connect-VIServer'd to the target
    vCenter(s); vCenter account has the three "Guest Operation ..." privileges.
    This script issues no vSphere write of any kind, and makes no change
    inside any guest either.
#>

[CmdletBinding()]
param(
    [string] $VMListPath     = 'C:\temp\vmlist.txt',
    [string] $CredentialPath = 'C:\temp\wincred.xml',
    [string] $OutputPath     = (Join-Path $PSScriptRoot 'SecurityGap_Reports'),

    [ValidateRange(8, 256)]
    [int]    $MinPasswordLength = 14,
    [ValidateRange(1, 20)]
    [int]    $AccountLockoutThreshold = 3,

    [ValidateRange(30, 3600)]
    [int]    $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'

Write-Host "Invoke-RemoteChangeReviewPostCheck.ps1 - READ-ONLY verification  [build 2026-10-01a-precheck-postvalidate]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
Write-Host ("If the build above is not '2026-10-01a-precheck-postvalidate' you are running an OLD copy - update it.") -ForegroundColor DarkGray

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
$CsvPath  = Join-Path $OutputPath "SecurityGap_PostCheckReview_$RunStamp.csv"
$LogPath  = Join-Path $OutputPath "SecurityGap_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Change-review post-check run started. VMs=$($vmNames.Count)."

# =====================================================================================
# 2. In-guest READ-ONLY payload - verifies the END STATE of the 6 "Script - Review"
#    items. Never calls Set-*/New-*/Remove-*/Stop-*/Disable-*/Add-*.
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'
$MinPasswordLength       = {{MIN_PW_LEN}}
$AccountLockoutThreshold = {{LOCKOUT_THRESHOLD}}

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
    $H = '1. Windows Firewall - profile state (CID 3950/3951/3952)'
    foreach ($fp in @('Domain','Private','Public')) {
        try {
            $en = (Get-NetFirewallProfile -Profile $fp -ErrorAction Stop).Enabled
            Add-Row -Header $H -Label "Firewall state ($fp)" -LocalStatus $(if ($en -eq $true) { 'OK' } else { 'NOT SET' }) `
                -Category '1. Windows Firewall' -Item "Firewall state ($fp)" -Detected "Enabled=$en" -Expected 'Enabled=True'
        } catch {
            $v = Get-RegVal ("HKLM:\Software\Policies\Microsoft\WindowsFirewall\{0}Profile" -f $fp) 'EnableFirewall'
            Add-Row -Header $H -Label "Firewall state ($fp)" -LocalStatus $(if ($v.Found -and $v.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
                -Category '1. Windows Firewall' -Item "Firewall state ($fp)" -Detected ("EnableFirewall={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected 'On (1)'
        }
    }

    # ==========================================================================
    $H = '2. SMB Minimum Version (CID 29576/29583)'
    foreach ($sd in @(
        @{ Side='Server';  Key='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer' },
        @{ Side='Client'; Key='HKLM:\Software\Policies\Microsoft\Windows\LanmanWorkstation' } )) {
        $v = Get-RegVal $sd.Key 'MinSmb2Dialect'
        Add-Row -Header $H -Label ("Mandate minimum SMB version ({0})" -f $sd.Side) `
            -LocalStatus $(if ($v.Found -and [int]$v.Value -eq 768) { 'OK' } else { 'NOT SET' }) `
            -Category '2. SMB Minimum Version' -Item ("Mandate minimum SMB version ({0})" -f $sd.Side) `
            -Detected ("MinSmb2Dialect={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '768 (SMB 3.0.0)'
    }

    # ==========================================================================
    $H = '3. NetBIOS (CID 25358)'
    $v = Get-RegVal 'HKLM:\Software\Policies\Microsoft\Windows NT\DNSClient' 'EnableNetbios'
    Add-Row -Header $H -Label 'NetBIOS name resolution disabled' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
        -Detail $(if (-not ($v.Found -and $v.Value -eq 0)) { 'Full effect requires a reboot or adapter re-init if just applied.' }) `
        -Category '3. NetBIOS' -Item 'NetBIOS name resolution disabled' `
        -Detected ("EnableNetbios={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected '0 (Disabled)'

    # ==========================================================================
    $H = '4. Account Lockout Threshold (CID 1071/2342)'
    try {
        $cfg = Join-Path $env:TEMP ("pcr_sa_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
        & secedit /export /cfg $cfg /areas SECURITYPOLICY /quiet 2>&1 | Out-Null
        $mpl = (Select-String -LiteralPath $cfg -Pattern '^\s*MinimumPasswordLength\s*=\s*(\d+)' | Select-Object -First 1)
        $lbc = (Select-String -LiteralPath $cfg -Pattern '^\s*LockoutBadCount\s*=\s*(\d+)' | Select-Object -First 1)
        Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue
        $mplV = if ($mpl) { [int]$mpl.Matches[0].Groups[1].Value } else { $null }
        $lbcV = if ($lbc) { [int]$lbc.Matches[0].Groups[1].Value } else { $null }
        Add-Row -Header $H -Label 'Minimum Password Length' `
            -LocalStatus $(if ($null -ne $mplV -and $mplV -ge $MinPasswordLength) { 'OK' } else { 'NOT SET' }) `
            -Category '4. Account Lockout Threshold' -Item 'Minimum Password Length' `
            -Detected ("MinimumPasswordLength={0}" -f $(if ($null -ne $mplV) { $mplV } else { '(not set)' })) -Expected (">= {0}" -f $MinPasswordLength)
        Add-Row -Header $H -Label 'Account Lockout Threshold' `
            -LocalStatus $(if ($null -ne $lbcV -and $lbcV -ge 1 -and $lbcV -le $AccountLockoutThreshold) { 'OK' } else { 'NOT SET' }) `
            -Category '4. Account Lockout Threshold' -Item 'Account Lockout Threshold' `
            -Detected ("LockoutBadCount={0}" -f $(if ($null -ne $lbcV) { $lbcV } else { '(not set)' })) -Expected ("1 to {0}" -f $AccountLockoutThreshold)
    } catch {
        Add-Row -Header $H -Label 'Account policy' -LocalStatus 'MANUAL' -Detail "secedit export failed: $($_.Exception.Message)" `
            -Category '4. Account Lockout Threshold' -Item 'Account policy' -NormStatus 'Unable to Check' -Detected 'secedit /export failed' -Expected 'see -MinPasswordLength / -AccountLockoutThreshold'
    }

    # ==========================================================================
    $H = '5. UAC Hardening (CID 2586/2587/3940)'
    $p = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System'
    $fat = Get-RegVal $p 'FilterAdministratorToken'
    Add-Row -Header $H -Label 'Admin Approval Mode for built-in Administrator' `
        -LocalStatus $(if ($fat.Found -and $fat.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
        -Category '5. UAC Hardening' -Item 'Admin Approval Mode for built-in Administrator' `
        -Detected ("FilterAdministratorToken={0}" -f $(if ($fat.Found) { $fat.Value } else { '(not set)' })) -Expected '1 (Enabled)'
    $cpb = Get-RegVal $p 'ConsentPromptBehaviorAdmin'
    Add-Row -Header $H -Label 'Elevation prompt behavior' `
        -LocalStatus $(if ($cpb.Found -and $cpb.Value -eq 2) { 'OK' } else { 'NOT SET' }) `
        -Category '5. UAC Hardening' -Item 'Elevation prompt behavior' `
        -Detected ("ConsentPromptBehaviorAdmin={0}" -f $(if ($cpb.Found) { $cpb.Value } else { '(not set)' })) -Expected '2 (Prompt for consent on the secure desktop)'
    $ev = Get-RegVal $p 'EnableVirtualization'
    Add-Row -Header $H -Label 'File/registry write virtualization' `
        -LocalStatus $(if ($ev.Found -and $ev.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
        -Category '5. UAC Hardening' -Item 'File/registry write virtualization' `
        -Detected ("EnableVirtualization={0}" -f $(if ($ev.Found) { $ev.Value } else { '(not set)' })) -Expected '1 (Enabled)'

    # ==========================================================================
    $H = '6. VBS and Credential Guard (CID 10473)'
    $lcf = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' 'LsaCfgFlags'
    $configured = $lcf.Found -and $lcf.Value -eq 1
    $running = $false
    try {
        $dg = Get-CimInstance -Namespace root\Microsoft\Windows\DeviceGuard -ClassName Win32_DeviceGuard -ErrorAction Stop
        $running = @($dg.SecurityServicesRunning) -contains 1
    } catch { }
    Add-Row -Header $H -Label 'Credential Guard configuration' `
        -LocalStatus $(if ($configured -and $running) { 'OK' } elseif ($configured -and -not $running) { 'MANUAL' } else { 'NOT SET' }) `
        -Detail $(if ($configured -and -not $running) { 'Registry is set (LsaCfgFlags=1) but Credential Guard is not yet reported as running - a reboot is pending, or the VM does not yet expose Secure Boot + virtualization in vCenter.' }) `
        -Category '6. VBS and Credential Guard' -Item 'Credential Guard configuration' `
        -Detected ("LsaCfgFlags={0}; SecurityServicesRunning contains CredGuard={1}" -f $(if ($lcf.Found) { $lcf.Value } else { '(not set)' }), $running) -Expected 'LsaCfgFlags=1 and actually running'

} catch {
    $fatal = $_.Exception.Message
}

$meta = [PSCustomObject]@{ Hostname = $ComputerName; IPAddress = $IPAddresses; OS = $OSCaption; FatalError = $fatal }
# .ToArray() rather than @(...) - @() around a generic List instance is unreliable
# on some PowerShell builds; .ToArray() is well-defined on 5.1 and 7.x alike.
$envelope = [PSCustomObject]@{ Schema = 'secgap-postcheck-review-1'; Meta = $meta; Rows = $Rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<SECGAP-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-SECGAP-ENVELOPE>>>'
'@

# =====================================================================================
# 3. Admin-side helpers (identical to the pre-check/post-check wrappers)
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

# Same rationale as the other wrappers: deliver the payload as a staged guest
# file (Copy-VMGuestFile, or chunked Base64 fallback) rather than one long
# Invoke-VMScript command line. Guest operations only.
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

# Same console style as the other wrappers: OK=Green, MANUAL=DarkYellow, else Red.
function Show-VmResults {
    param([string]$VmBanner, $Rows, [int]$InlineDetailMax = 70, [int]$WrapWidth = 100)
    $HeaderColor   = 'Cyan'
    $SubPointColor = 'White'

    $statusWidth = 7
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
$summary = [ordered]@{ Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0; Manual = 0; SkippedFailed = 0 }
$inventory = Get-VMInventory -Servers $connectedServers

$payloadForGuest = $PayloadTemplate.Replace('{{MIN_PW_LEN}}', ([string][int]$MinPasswordLength)).Replace('{{LOCKOUT_THRESHOLD}}', ([string][int]$AccountLockoutThreshold))

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Post-check-review for VM '$vmName' ===" 'INFO'
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

        Write-Log "'$vmName': staging + running read-only post-check-review payload." 'INFO'
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $payloadForGuest -ToolsWaitSecs $ToolsWaitSecs

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
        $f = @{ NonCompliant = $false; Manual = $false }
        foreach ($r in @($envelope.Rows)) {
            if ($r.Status -eq 'Non-Compliant')                { $f.NonCompliant = $true }
            if ($r.Status -eq 'Manual Verification Required')  { $f.Manual = $true }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category $r.Category -SecurityCheck $r.Item -DetectedValue $r.DetectedValue -ExpectedValue $r.ExpectedValue `
                -Status $r.Status -Detail $r.Detail))
        }
        if ($f.NonCompliant) { $summary.NonCompliant++ }
        if ($f.Manual)       { $summary.Manual++ }
        if (-not $f.NonCompliant) { $summary.Compliant++ }

        Show-VmResults -VmBanner "$vmName  @ $($srv.Name)   ($ghost / $gip)" -Rows @($envelope.Rows)
        Write-Log "'$vmName': post-checked. nonCompliant=$($f.NonCompliant) manual=$($f.Manual)" 'INFO'
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
Write-Host "======= CHANGE-REVIEW POST-CHECK SUMMARY (did it take effect?) =======" -ForegroundColor Green
Write-Host ("  Total VMs                     : {0}" -f $summary.Total)
Write-Host ("  Successfully Checked          : {0}" -f $summary.Checked)
Write-Host ("  Fully Compliant               : {0}" -f $summary.Compliant) -ForegroundColor $(if ($summary.Compliant) { 'Green' } else { 'Gray' })
Write-Host ("  Non-Compliant (gaps remain)   : {0}" -f $summary.NonCompliant) -ForegroundColor $(if ($summary.NonCompliant) { 'Red' } else { 'Gray' })
Write-Host ("  Manual Verification Required  : {0}" -f $summary.Manual) -ForegroundColor $(if ($summary.Manual) { 'DarkYellow' } else { 'Gray' })
Write-Host ("  Skipped / Failed to process   : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "========================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Report : $CsvPath"
Write-Host "Log    : $LogPath"
Write-Log "Change-review post-check run complete."
