#Requires -Version 5.1
<#
.SYNOPSIS
    Read-only verification for the extra Microsoft Security Baseline controls raised
    by the Cybersecurity team's Qualys review of SF-UFM-APP01.

    Companion to Invoke-RemoteExtendedBaselineRemediation.ps1 - verifies exactly the
    same items, changes nothing, and uses the same on-screen format as
    Invoke-RemotePostRemediationCheck.ps1:
        Cyan "=== <group header> ==="  then one line per item
        "  [<STATUS>] <Label> [- <detail>]"   (OK*=Green, MANUAL*=DarkYellow,
        N/A*=Gray, else Red)
    plus a central CSV with the auditable columns.

.PARAMETER VMListPath / CredentialPath / OutputPath / ToolsWaitSecs
    Same as the other remote scripts. Defaults: C:\temp\vmlist.txt,
    C:\temp\wincred.xml, .\SecurityGap_Reports, 180.

.PARAMETER CachedLogonsTarget
    Value at or below which Winlogon\CachedLogonsCount is treated Compliant.
    Default 4 (CIS). Qualys QID 90007 only fully clears at 0.

.EXAMPLE
    .\Invoke-RemoteExtendedBaselinePostCheck.ps1
    .\Invoke-RemoteExtendedBaselinePostCheck.ps1 -CachedLogonsTarget 0

.NOTES
    In-guest payload contains ONLY Get-* / registry reads / auditpol /get / net user.
    No Set-*, no vSphere write. Prerequisites: PowerCLI imported and Connect-VIServer'd;
    vCenter account has the "Guest Operation ..." privileges.
#>

[CmdletBinding()]
param(
    [string] $VMListPath     = 'C:\temp\vmlist.txt',
    [string] $CredentialPath = 'C:\temp\wincred.xml',
    [string] $OutputPath     = (Join-Path $PSScriptRoot 'SecurityGap_Reports'),

    [ValidateRange(0, 50)]
    [int]    $CachedLogonsTarget = 4,

    [ValidateRange(30, 3600)]
    [int]    $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'

Write-Host "Invoke-RemoteExtendedBaselinePostCheck.ps1 - READ-ONLY verification  [build 2026-08-31a-extbaseline-postcheck]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray

# =====================================================================================
# 1. Pre-flight
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
$CsvPath  = Join-Path $OutputPath "SecurityGap_ExtBaseline_PostCheck_$RunStamp.csv"
$LogPath  = Join-Path $OutputPath "SecurityGap_ExtBaseline_PostCheck_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Extended-baseline post-check run started. VMs=$($vmNames.Count)."

# =====================================================================================
# 2. In-guest READ-ONLY payload ({{CACHED_TARGET}} replaced with an integer literal)
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'
$CachedLogonsTarget = {{CACHED_TARGET}}

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
function Get-RegVal {
    param([string]$Path, [string]$Name)
    try { $ip = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop; return [pscustomobject]@{ Found = $true; Value = $ip.$Name } }
    catch { return [pscustomobject]@{ Found = $false; Value = $null } }
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
    # =====================================================================
    $H = 'Extended Baseline (Cybersecurity Review) - VBS / LSA / logging'
    # 1. HVCI - UEFI lock
    $p = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'
    $en = Get-RegVal $p 'Enabled'; $lk = Get-RegVal $p 'Locked'
    Add-Row -Header $H -Label 'HVCI - Enabled with UEFI lock' `
        -LocalStatus $(if ($en.Found -and $en.Value -eq 1 -and $lk.Found -and $lk.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
        -Category 'Extended Baseline' -Item 'HVCI - UEFI lock' `
        -Detected ("Enabled={0}; Locked={1}" -f $(if ($en.Found) { $en.Value } else { '(ns)' }), $(if ($lk.Found) { $lk.Value } else { '(ns)' })) `
        -Expected 'Enabled=1 and Locked=1'

    # 2. Kernel-mode HW-enforced Stack Protection - enforcement
    $p = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\KernelShadowStacks'
    $en = Get-RegVal $p 'Enabled'; $am = Get-RegVal $p 'AuditMode'
    Add-Row -Header $H -Label 'Kernel-mode HW Stack Protection - enforcement' `
        -LocalStatus $(if ($en.Found -and $en.Value -eq 1 -and $am.Found -and $am.Value -eq 0) { 'OK' } elseif ($en.Found -and $en.Value -eq 1) { 'NOT SET' } else { 'NOT SET' }) `
        -Detail $(if ($en.Found -and $en.Value -eq 1 -and $am.Found -and $am.Value -eq 1) { 'currently in Audit mode - control expects Enforcement' } else { '' }) `
        -Category 'Extended Baseline' -Item 'Kernel-mode HW-enforced Stack Protection' `
        -Detected ("Enabled={0}; AuditMode={1}" -f $(if ($en.Found) { $en.Value } else { '(ns)' }), $(if ($am.Found) { $am.Value } else { '(ns)' })) `
        -Expected 'Enabled=1 and AuditMode=0'

    # 3. LSASS RunAsPPL - UEFI lock
    $v = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' 'RunAsPPL'
    Add-Row -Header $H -Label 'LSASS RunAsPPL - UEFI lock' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 1) { 'OK' } else { "NOT SET (found: $(if ($v.Found) { $v.Value } else { '' }))" }) `
        -Detail $(if ($v.Found -and $v.Value -eq 2) { 'RunAsPPL=2 = enabled WITHOUT UEFI lock; control expects 1 (with lock)' } else { '' }) `
        -Category 'Extended Baseline' -Item 'LSASS RunAsPPL - UEFI lock' `
        -Detected ("RunAsPPL={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected 'RunAsPPL=1'

    # 4. PowerShell Script Block (+ Invocation) Logging
    $p = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'
    $sb = Get-RegVal $p 'EnableScriptBlockLogging'; $si = Get-RegVal $p 'EnableScriptBlockInvocationLogging'
    Add-Row -Header $H -Label 'Script Block + Invocation Logging' `
        -LocalStatus $(if ($sb.Found -and $sb.Value -eq 1 -and $si.Found -and $si.Value -eq 1) { 'OK' } else { 'NOT SET' }) `
        -Category 'Extended Baseline' -Item 'PowerShell Script Block (and Invocation) Logging' `
        -Detected ("EnableScriptBlockLogging={0}; EnableScriptBlockInvocationLogging={1}" -f $(if ($sb.Found) { $sb.Value } else { '(ns)' }), $(if ($si.Found) { $si.Value } else { '(ns)' })) `
        -Expected 'both = 1'

    # 11. Credential Guard policy value / GPO override
    $dg = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' 'LsaCfgFlags'
    $lc = Get-RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' 'LsaCfgFlags'
    if ($dg.Found -and $dg.Value -eq 0) {
        Add-Row -Header $H -Label 'Credential Guard - policy value' -LocalStatus 'NOT SET' `
            -Detail 'A GPO is DISABLING Credential Guard (Policies\...\DeviceGuard\LsaCfgFlags=0). Set "Turn On VBS > Credential Guard Configuration" to "Enabled with UEFI lock".' `
            -Category 'Extended Baseline' -Item 'Credential Guard - policy value' `
            -Detected ("Policies\...\DeviceGuard\LsaCfgFlags=0; Control\LSA\LsaCfgFlags={0}" -f $(if ($lc.Found) { $lc.Value } else { '(ns)' })) `
            -Expected 'LsaCfgFlags=1 (Enabled with UEFI lock)'
    } else {
        Add-Row -Header $H -Label 'Credential Guard - policy value' `
            -LocalStatus $(if ($dg.Found -and $dg.Value -eq 1 -and $lc.Found -and $lc.Value -in @(1, 2)) { 'OK' } else { 'NOT SET' }) `
            -Category 'Extended Baseline' -Item 'Credential Guard - policy value' `
            -Detected ("Policies\...\DeviceGuard\LsaCfgFlags={0}; Control\LSA\LsaCfgFlags={1}" -f $(if ($dg.Found) { $dg.Value } else { '(ns)' }), $(if ($lc.Found) { $lc.Value } else { '(ns)' })) `
            -Expected 'Policies LsaCfgFlags=1 and Control\LSA LsaCfgFlags in {1,2}'
    }

    # =====================================================================
    $H = 'Extended Baseline (Cybersecurity Review) - WinRM / Kerberos PKINIT / Ink / Audit / Cached logons'
    # 5. WinRM AllowAutoConfig
    $wp = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service'
    $ac = Get-RegVal $wp 'AllowAutoConfig'
    Add-Row -Header $H -Label 'WinRM auto-configuration disabled' `
        -LocalStatus $(if ((-not $ac.Found) -or $ac.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
        -Category 'Extended Baseline' -Item 'WinRM auto-configuration disabled' `
        -Detected ("AllowAutoConfig={0}" -f $(if ($ac.Found) { $ac.Value } else { '(not configured)' })) -Expected '0 or not configured'

    # 6. WinRM IPv4/IPv6 filter
    $f4 = Get-RegVal $wp 'IPv4Filter'; $f6 = Get-RegVal $wp 'IPv6Filter'
    $f4v = if ($f4.Found) { [string]$f4.Value } else { '' }
    $f6v = if ($f6.Found) { [string]$f6.Value } else { '' }
    $filterOpen = ($f4v -eq '*' -or $f4v -eq '' ) -and ($f6v -eq '*' -or $f6v -eq '')
    Add-Row -Header $H -Label 'WinRM listener IPv4/IPv6 filter' `
        -LocalStatus $(if ($filterOpen) { 'MANUAL' } else { 'OK' }) `
        -Detail $(if ($filterOpen) { "IPv4Filter='$f4v', IPv6Filter='$f6v' - open to all. Scope to a management range." } else { "IPv4Filter='$f4v'; IPv6Filter='$f6v'" }) `
        -Category 'Extended Baseline' -Item 'WinRM listener IPv4/IPv6 filter' `
        -Detected ("IPv4Filter='{0}'; IPv6Filter='{1}'" -f $f4v, $f6v) -Expected 'Scoped to a management IP range (not * / blank)'

    # 7. PKINIT hash SHA256 / SHA384 / SHA512 = Supported (3)
    foreach ($alg in 'SHA256', 'SHA384', 'SHA512') {
        $v = Get-RegVal ("HKLM:\SOFTWARE\Policies\Microsoft\Windows\PKINITHashAlgorithms\{0}" -f $alg) 'Support'
        Add-Row -Header $H -Label ("PKINIT hash {0} = Supported" -f $alg) `
            -LocalStatus $(if ($v.Found -and $v.Value -eq 3) { 'OK' } else { 'NOT SET' }) `
            -Detail 'Benchmark value only - no AD / PKINIT smart-card logon in a workgroup.' `
            -Category 'Extended Baseline' -Item ("PKINIT hash {0}" -f $alg) `
            -Detected ("Support={0}" -f $(if ($v.Found) { $v.Value } else { '(not set / default)' })) -Expected 'Support=3 (Supported)'
    }

    # 8. Windows Ink Workspace
    $v = Get-RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsInkWorkspace' 'AllowWindowsInkWorkspace'
    Add-Row -Header $H -Label 'Windows Ink Workspace disabled' `
        -LocalStatus $(if ($v.Found -and $v.Value -eq 0) { 'OK' } else { 'NOT SET' }) `
        -Category 'Extended Baseline' -Item 'Windows Ink Workspace disabled' `
        -Detected ("AllowWindowsInkWorkspace={0}" -f $(if ($v.Found) { $v.Value } else { '(not configured)' })) -Expected 'AllowWindowsInkWorkspace=0'

    # 9. Audit "Audit Policy Change" = Success and Failure
    try {
        $csv = (& auditpol /get /subcategory:"Audit Policy Change" /r 2>$null | ConvertFrom-Csv)
        $row = $csv | Where-Object { $_.Subcategory -eq 'Audit Policy Change' } | Select-Object -First 1
        $inc = if ($row) { [string]$row.'Inclusion Setting' } else { '(unknown)' }
        Add-Row -Header $H -Label 'Audit "Audit Policy Change"' `
            -LocalStatus $(if ($inc -eq 'Success and Failure') { 'OK' } else { 'NOT SET' }) `
            -Category 'Extended Baseline' -Item 'Audit "Audit Policy Change"' `
            -Detected ("Inclusion Setting={0}" -f $inc) -Expected 'Success and Failure'
    } catch {
        Add-Row -Header $H -Label 'Audit "Audit Policy Change"' -LocalStatus 'NOT SET' -Detail 'auditpol read failed' `
            -Category 'Extended Baseline' -Item 'Audit "Audit Policy Change"' -NormStatus 'Unable to Check' `
            -Detected 'auditpol /get failed' -Expected 'Success and Failure'
    }

    # 10. Cached logons count
    $v = Get-RegVal 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' 'CachedLogonsCount'
    $clc = if ($v.Found) { [int]$v.Value } else { $null }
    Add-Row -Header $H -Label 'Cached logons count' `
        -LocalStatus $(if ($null -ne $clc -and $clc -le $CachedLogonsTarget) { 'OK' } else { 'NOT SET' }) `
        -Detail $(if ($null -ne $clc -and $clc -gt 0) { "value $clc - Qualys QID 90007 only clears at 0 on a workgroup server" } else { '' }) `
        -Category 'Extended Baseline' -Item 'Cached logons count' `
        -Detected ("CachedLogonsCount={0}" -f $(if ($v.Found) { $v.Value } else { '(not set)' })) -Expected ("<= {0}" -f $CachedLogonsTarget)

    # =====================================================================
    $H = 'Extended Baseline (Cybersecurity Review) - Local Security Policy / patch (verify with evidence)'
    # 15. Guest account renamed (checkable via well-known SID -501)
    try {
        $guest = Get-LocalUser -ErrorAction Stop | Where-Object { $_.SID.Value -like '*-501' } | Select-Object -First 1
        if ($guest) {
            Add-Row -Header $H -Label 'Built-in Guest account renamed (QID 105228)' `
                -LocalStatus $(if ($guest.Name -ne 'Guest') { 'OK' } else { 'NOT SET' }) `
                -Detail ("Guest account name = '{0}', Enabled = {1}" -f $guest.Name, $guest.Enabled) `
                -Category 'Extended Baseline' -Item 'Built-in Guest account renamed' `
                -Detected ("Name='{0}'; Enabled={1}" -f $guest.Name, $guest.Enabled) -Expected "Renamed (not 'Guest') and Disabled"
        } else {
            Add-Row -Header $H -Label 'Built-in Guest account renamed (QID 105228)' -LocalStatus 'MANUAL' -Detail 'Built-in Guest (SID -501) not found - verify manually' `
                -Category 'Extended Baseline' -Item 'Built-in Guest account renamed' -Detected 'SID -501 account not found' -Expected 'Renamed and Disabled'
        }
    } catch {
        Add-Row -Header $H -Label 'Built-in Guest account renamed (QID 105228)' -LocalStatus 'MANUAL' -Detail 'Get-LocalUser not available - verify with secedit export' `
            -Category 'Extended Baseline' -Item 'Built-in Guest account renamed' -NormStatus 'Manual Verification Required' -Detected 'unable to read local users' -Expected 'Renamed and Disabled'
    }

    # Unused local accounts flagged by Qualys
    try {
        $flagged = @('s.l.jganzon.m', 's.l.samohamed.m')
        $present = Get-LocalUser -ErrorAction Stop | Where-Object { $flagged -contains $_.Name } |
                   ForEach-Object { "$($_.Name)(Enabled=$($_.Enabled))" }
        Add-Row -Header $H -Label 'Unused local accounts (s.l.jganzon.m / s.l.samohamed.m)' `
            -LocalStatus $(if (-not $present) { 'OK' } else { 'MANUAL' }) `
            -Detail $(if ($present) { "present: $($present -join ', ') - disable / remove after owner confirmation" } else { 'not present on this host' }) `
            -Category 'Extended Baseline' -Item 'Unused local accounts' `
            -Detected $(if ($present) { ($present -join ', ') } else { 'none of the flagged accounts present' }) -Expected 'No stale local accounts'
    } catch {
        Add-Row -Header $H -Label 'Unused local accounts (s.l.jganzon.m / s.l.samohamed.m)' -LocalStatus 'MANUAL' -Detail 'unable to read local users' `
            -Category 'Extended Baseline' -Item 'Unused local accounts' -NormStatus 'Manual Verification Required' -Detected 'unable to read local users' -Expected 'No stale local accounts'
    }

    # 12/13/14 - user rights + admin lockout: not readable without secedit export
    Add-Row -Header $H -Label 'Deny access to this computer from the network' -LocalStatus 'MANUAL' `
        -Detail 'Verify SeDenyNetworkLogonRight includes Guests + Local account + Local account & Administrators (secedit /export /cfg secpol.txt).' `
        -Category 'Extended Baseline' -Item 'Deny access from network (user right)' -Detected 'Not readable without secedit export' -Expected '*S-1-5-32-546,*S-1-5-113,*S-1-5-114'
    Add-Row -Header $H -Label 'Deny logon through Remote Desktop Services' -LocalStatus 'MANUAL' `
        -Detail 'Verify SeDenyRemoteInteractiveLogonRight includes Guests + Local account (secedit /export).' `
        -Category 'Extended Baseline' -Item 'Deny logon through RDS (user right)' -Detected 'Not readable without secedit export' -Expected '*S-1-5-32-546,*S-1-5-113'
    Add-Row -Header $H -Label 'Allow Administrator account lockout' -LocalStatus 'MANUAL' `
        -Detail 'Verify [System Access] AllowAdministratorLockout = 1 and an account lockout threshold is set (secedit /export / "net accounts").' `
        -Category 'Extended Baseline' -Item 'Allow Administrator account lockout' -Detected 'Not readable without secedit export' -Expected 'Enabled (1)'

    # 17. August 2026 Windows update - checkable via UBR
    try {
        $cv  = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $bld = [int]$cv.CurrentBuildNumber
        $ubr = [int]$cv.UBR
        $fullBuild = "$bld.$ubr"
        $patched = ($bld -gt 26100) -or ($bld -eq 26100 -and $ubr -ge 33222)
        Add-Row -Header $H -Label 'August 2026 Windows Update (QID 92439)' `
            -LocalStatus $(if ($patched) { 'OK' } else { 'NOT SET' }) `
            -Detail $(if ($patched) { '' } else { 'install KB5120228 / KB5120233 via Windows Update, then re-scan' }) `
            -Category 'Extended Baseline' -Item 'August 2026 Windows Update' `
            -Detected ("OS build {0}" -f $fullBuild) -Expected 'build >= 26100.33222'
    } catch {
        Add-Row -Header $H -Label 'August 2026 Windows Update (QID 92439)' -LocalStatus 'MANUAL' -Detail 'could not read OS build revision' `
            -Category 'Extended Baseline' -Item 'August 2026 Windows Update' -NormStatus 'Manual Verification Required' -Detected 'UBR unreadable' -Expected 'build >= 26100.33222'
    }

    Add-Row -Header $H -Label 'QID 92446 - Defender EoP zero-day' -LocalStatus 'MANUAL' `
        -Detail 'No Microsoft patch at assessment time - track MSRC advisory, apply when released; keep Defender platform/signatures current.' `
        -Category 'Extended Baseline' -Item 'QID 92446 - Defender EoP zero-day' -Detected 'No patch at assessment time' -Expected 'Patched when available'
    Add-Row -Header $H -Label 'Third-party agent vulnerabilities' -LocalStatus 'MANUAL' `
        -Detail 'Splunk UF / Azure Arc agent / ServiceNow agent / Binalyze AIR - upgrade with the respective owners (Go / gRPC / .NET / Ruby vulns).' `
        -Category 'Extended Baseline' -Item 'Third-party agent vulnerabilities' -Detected 'Vulnerable agent builds present' -Expected 'Agents on fixed versions'

} catch {
    $fatal = $_.Exception.Message
}

$meta = [PSCustomObject]@{ Hostname = $ComputerName; IPAddress = $IPAddresses; OS = $OSCaption; FatalError = $fatal }
$envelope = [PSCustomObject]@{ Schema = 'secgap-extbaseline-postcheck-1'; Meta = $meta; Rows = $Rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<SECGAP-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-SECGAP-ENVELOPE>>>'
'@

# =====================================================================================
# 3. Admin-side helpers (same as the other remote scripts)
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

function Invoke-LargeGuestPayload {
    param($VM, $Server, [pscredential]$Credential, [string]$PayloadText, [int]$ToolsWaitSecs, [int]$ChunkSize = 1500)
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

# Aligned local-format renderer (status column padded to widest token; long detail wraps).
function Show-VmResults {
    param([string]$VmBanner, $Rows, [int]$InlineDetailMax = 70, [int]$WrapWidth = 100)
    $HeaderColor   = 'Cyan'
    $SubPointColor = 'White'
    $statusWidth = 8
    foreach ($row in $Rows) { $l = ([string]$row.LocalStatus).Length; if ($l -gt $statusWidth) { $statusWidth = $l } }
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
            'N/A*'    { 'Gray' }
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
$summary = [ordered]@{ Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0; Manual = 0; NotApplicable = 0; Unable = 0; SkippedFailed = 0 }
$inventory = Get-VMInventory -Servers $connectedServers
$guestPayload = $PayloadTemplate.Replace('{{CACHED_TARGET}}', [string][int]$CachedLogonsTarget)

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
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $guestPayload -ToolsWaitSecs $ToolsWaitSecs

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
        $f = @{ NonCompliant = $false; Manual = $false; NA = $false; Unable = $false }
        foreach ($r in @($envelope.Rows)) {
            switch ($r.Status) {
                'Non-Compliant'                { $f.NonCompliant = $true }
                'Manual Verification Required' { $f.Manual = $true }
                'Not Applicable'               { $f.NA = $true }
                'Unable to Check'              { $f.Unable = $true }
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

        Show-VmResults -VmBanner "$vmName  @ $($srv.Name)   ($ghost / $gip)" -Rows @($envelope.Rows)
        Write-Log "'$vmName': verified. nonCompliant=$($f.NonCompliant) manual=$($f.Manual) na=$($f.NA) unable=$($f.Unable)" 'INFO'
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
Write-Host "===== EXTENDED-BASELINE POST-CHECK SUMMARY =====" -ForegroundColor Green
Write-Host ("  Total VMs                     : {0}" -f $summary.Total)
Write-Host ("  Successfully Checked          : {0}" -f $summary.Checked)
Write-Host ("  Compliant (no gaps found)     : {0}" -f $summary.Compliant)
Write-Host ("  Non-Compliant (gaps found)    : {0}" -f $summary.NonCompliant) -ForegroundColor $(if ($summary.NonCompliant) { 'Yellow' } else { 'Gray' })
Write-Host ("  Manual Verification Required  : {0}" -f $summary.Manual) -ForegroundColor DarkYellow
Write-Host ("  Not Applicable (workgroup etc): {0}" -f $summary.NotApplicable) -ForegroundColor Gray
Write-Host ("  Unable to Check (some items)  : {0}" -f $summary.Unable) -ForegroundColor $(if ($summary.Unable) { 'Yellow' } else { 'Gray' })
Write-Host ("  Skipped / Failed to process   : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "===============================================" -ForegroundColor Green
Write-Host ""
Write-Host "Report : $CsvPath"
Write-Host "Log    : $LogPath"
Write-Log "Extended-baseline post-check run complete."
