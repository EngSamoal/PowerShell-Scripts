#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Server Security Remediation - runs entirely locally on the server being fixed.
    Companion to Get-WindowsServerSecurityAudit-Local.ps1. UNLIKE the audit script, this ONE
    DOES change configuration when run with -Apply.

.DESCRIPTION
    Addresses the gaps found by the audit script, covering the same 6 areas:
      1. VBS & Credential Protection (Credential Guard, HVCI/Memory Integrity)
      2. Microsoft Defender (real-time protection, ASR rules, cloud-delivered protection)
      3. Authentication Security (WDigest, Kerberos supported encryption types)
      4. Remote Management (WinRM hardening or full disable)
      5. SMB & RPC Security (SMB1 auditing, Printer RPC / PrintNightmare, Point and Print)
      6. Windows Server 2025 Security Baseline (narrow SmartScreen/IE-zone/legacy-IE items
         directly, PLUS the option to invoke Microsoft's own LGPO.exe against an official
         baseline GPO backup for genuine full-baseline compliance - see .NOTES)

    SAFETY MODEL - dry-run by default:
      Without -Apply, this script changes NOTHING. It evaluates each item, decides whether it
      would need to change anything, and reports that as "DryRun - would apply" (or
      "AlreadyCompliant" if nothing would change). Nothing is written to the registry, no
      service is touched, no cmdlet that changes state is ever invoked.

      With -Apply, changes are made through PowerShell's own confirmation pipeline
      ($PSCmdlet.ShouldProcess) - so -Apply can be combined with -WhatIf to preview exactly what
      would run without executing it, or with -Confirm to be prompted per item. Elevation
      (Run as Administrator) is REQUIRED when -Apply is used; the script refuses to proceed
      without it rather than silently failing half the items.

    REBOOTS - this script never reboots or schedules a reboot. Credential Guard/HVCI and the
    legacy-IE-feature removal only take effect after a reboot; every item that needs one is
    still applied (the registry/feature-state change is made immediately) but is called out
    individually in its own result row AND summarized in one list at the end of the run, so you
    can schedule reboots deliberately rather than have this script decide that for you.

    HIGH-RISK ITEMS ARE OPT-IN, NOT DEFAULT - even under -Apply:
      -DisableLegacyTls   Disables SSL 2.0/3.0 and TLS 1.0/1.1 via Schannel. Off by default:
                          this can break any client or integration that hasn't moved to TLS 1.2+.
      -EnforceAesOnlyKerberos
                          Sets Kerberos SupportedEncryptionTypes to AES-only, REMOVING RC4/DES
                          support. Off by default: without this switch, the script only ADDS
                          AES128/AES256 support to whatever is already configured - it never
                          removes an encryption type, so it cannot break existing Kerberos auth.
                          Only pass this once you've confirmed no legacy service/device in your
                          environment still requires RC4.
      -DisableWinRM       Stops and disables the WinRM service entirely. Off by default: the
                          default behavior instead just hardens WinRM (disables unencrypted
                          traffic and Basic auth) without removing remote-management capability,
                          since some environments have other tooling that depends on WinRM being
                          up. Pass this explicitly if you want it fully disabled.

    Tamper Protection is NOT remediated by this script - Microsoft deliberately blocks changing
    it via PowerShell/script (by design, to stop malware from disabling it the same way). It's
    reported as "Manual/External Required - enable via Windows Security app or Intune."

.PARAMETER Apply
    Master switch. Without it, the script only reports what it WOULD do and changes nothing.
    Required (along with an elevated session) before any change is made.

.PARAMETER OutputPath
    Folder for the CSV/log. Created if it doesn't exist. Defaults to .\SecurityRemediation_Reports
    next to this script.

.PARAMETER SkipCredentialGuard
.PARAMETER SkipDefenderASR
.PARAMETER SkipAuthHardening
.PARAMETER SkipWinRM
.PARAMETER SkipSmbRpc
.PARAMETER SkipBaselineIndicators
    Each skips its whole category. All default to off (i.e. every category runs).

.PARAMETER DisableWinRM
    Opt-in: fully stop and disable the WinRM service instead of the default (harden only).

.PARAMETER WinRmAllowedSourceRange
    Optional list of IP addresses/CIDR ranges (e.g. '10.1.1.0/24','10.1.2.5') that should be the
    ONLY sources allowed to reach WinRM. When supplied (and -DisableWinRM is not used), restricts
    the "Windows Remote Management" firewall rules' RemoteAddress to exactly this list - this is
    the actual "restrict listener filters" hardening step. Not guessed at or defaulted: without
    this parameter the step is reported Manual/External Required rather than assume what your
    management subnet is, since a wrong guess here could either lock out legitimate admin access
    or fail to restrict anything.

.PARAMETER EnforceAesOnlyKerberos
    Opt-in: set Kerberos SupportedEncryptionTypes to AES-only (removes RC4/DES). Default
    behavior only adds AES support without removing anything already configured.

.PARAMETER DisableLegacyTls
    Opt-in: disable SSL 2.0/3.0 and TLS 1.0/1.1 via Schannel. High app-compat risk - off by
    default even under -Apply.

.PARAMETER AsrRuleAction
    'Enabled' (block - default, matches "enable ASR enforcement") or 'AuditMode' (log only, for
    a safer pilot rollout before enforcing). Applies to every ASR rule this script configures.

.PARAMETER LgpoPath
.PARAMETER BaselineGpoBackupPath
    Optional. If both point to a real LGPO.exe and a real extracted Windows Server 2025 Security
    Baseline GPO backup folder (from Microsoft's Security Compliance Toolkit), passing -Apply
    invokes `LGPO.exe /g <BaselineGpoBackupPath>` as one atomic step - the officially supported
    way to apply the full baseline, instead of hand-replicating hundreds of settings here. If
    either path is omitted or doesn't exist, this step is reported Manual/External Required with
    the exact command to run once you have the tool and backup staged.

.EXAMPLE
    # See what would change, on this server, with no risk of touching anything:
    .\Repair-WindowsServerSecurityGaps-Local.ps1

.EXAMPLE
    # Preview exactly what -Apply would run without running it:
    .\Repair-WindowsServerSecurityGaps-Local.ps1 -Apply -WhatIf

.EXAMPLE
    # Apply the safe defaults (elevated session required):
    .\Repair-WindowsServerSecurityGaps-Local.ps1 -Apply

.EXAMPLE
    # Apply everything including the high-risk opt-ins, and fully disable WinRM:
    .\Repair-WindowsServerSecurityGaps-Local.ps1 -Apply -DisableWinRM -EnforceAesOnlyKerberos -DisableLegacyTls

.EXAMPLE
    # Apply, and also push the real Server 2025 baseline GPO via LGPO.exe:
    .\Repair-WindowsServerSecurityGaps-Local.ps1 -Apply -LgpoPath 'C:\Tools\LGPO\LGPO.exe' -BaselineGpoBackupPath 'C:\Tools\Baseline\WS2025-Baseline-GPO'

.NOTES
    Companion to Get-WindowsServerSecurityAudit-Local.ps1 - run the audit script after this one
    (and after any scheduled reboot) to confirm the gaps are actually closed, rather than trusting
    this script's own "Applied" status as the final word: a registry write can succeed while the
    underlying feature still fails to activate (e.g. Credential Guard/HVCI need the VM's virtual
    hardware to expose Secure Boot and hardware-assisted virtualization - a vSphere-side VM
    setting this script cannot see or change from inside the guest).

    Category 6 (Security Baseline) is intentionally narrow by default: SmartScreen, IE zone
    ActiveX/Scripting/Protected Mode, and the legacy IE11 optional feature are remediated
    directly, because they're individually low-risk and were specifically called out. The real
    "Windows Server 2025 Security Baseline GPO" is hundreds of settings; reproducing all of them
    as ad-hoc registry writes here would risk both silently missing some and introducing
    undocumented compatibility breaks. Use -LgpoPath/-BaselineGpoBackupPath (see above) to apply
    the genuine baseline via Microsoft's own tool instead.

    KNOWN GAPS - deliberately not remediated, not an oversight:
      Kerberos SHA256/SHA384/SHA512 (RFC 8009 AES-SHA2 suites)
                                  Only the classic SupportedEncryptionTypes bitmask is remediated
                                  here (AES128/AES256-SHA1, the long-standing Windows Kerberos
                                  encryption types). Whether/how Windows Server 2025 exposes the
                                  newer SHA-2-based Kerberos suites via a local registry value was
                                  not confirmed as of this script's authoring - do not treat the
                                  Kerberos remediation above as covering this specific requirement.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [switch]$Apply,
    [string]$OutputPath = (Join-Path $PSScriptRoot 'SecurityRemediation_Reports'),

    [switch]$SkipCredentialGuard,
    [switch]$SkipDefenderASR,
    [switch]$SkipAuthHardening,
    [switch]$SkipWinRM,
    [switch]$SkipSmbRpc,
    [switch]$SkipBaselineIndicators,

    [switch]$DisableWinRM,
    [string[]]$WinRmAllowedSourceRange,
    [switch]$EnforceAesOnlyKerberos,
    [switch]$DisableLegacyTls,

    [ValidateSet('Enabled', 'AuditMode')]
    [string]$AsrRuleAction = 'Enabled',

    [string]$LgpoPath = 'C:\Tools\LGPO\LGPO.exe',
    [string]$BaselineGpoBackupPath = 'C:\Tools\Baseline\GPOs'
)

$ProgressPreference = 'SilentlyContinue'
$ScriptBuild = '2026-08-17-02-winrm-listener-filter'
Write-Host "Repair-WindowsServerSecurityGaps-Local.ps1 - build $ScriptBuild" -ForegroundColor Magenta
if ($Apply) {
    Write-Host "*** -Apply IS SET: this run WILL change configuration on this server. ***" -ForegroundColor Red
} else {
    Write-Host "DRY-RUN MODE (default) - nothing will be changed. Pass -Apply to actually make changes." -ForegroundColor Yellow
}

try {
    $IsElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    $IsElevated = $null
}
if ($Apply -and $IsElevated -ne $true) {
    throw "Refusing to -Apply: this session is not confirmed to be running elevated (Run as Administrator). Re-run from an elevated PowerShell/ISE window. (Dry-run without -Apply does not require elevation.)"
}
if (-not $Apply -and $IsElevated -eq $false) {
    Write-Host "NOTE: this session is not elevated. Dry-run results below still show what WOULD be done, but re-run elevated with -Apply when you're ready to make changes." -ForegroundColor Yellow
}

$ComputerName = $env:COMPUTERNAME
try {
    $IPAddresses = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop | Where-Object { $_.InterfaceAlias -notmatch 'Loopback' -and $_.IPAddress -notmatch '^169\.254\.' }).IPAddress -join ', '
} catch {
    $IPAddresses = $null
}
if ([string]::IsNullOrWhiteSpace($IPAddresses)) {
    try {
        $IPAddresses = ([System.Net.Dns]::GetHostAddresses($ComputerName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString
    } catch {
        $IPAddresses = 'Unknown'
    }
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogPath  = Join-Path $OutputPath "SecurityRemediation_${ComputerName}_$RunStamp.log"
$CsvPath  = Join-Path $OutputPath "SecurityRemediation_${ComputerName}_$RunStamp.csv"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

$Results = New-Object System.Collections.Generic.List[object]
$RebootRequiredItems = New-Object System.Collections.Generic.List[string]

function Add-ResultRow {
    param([string]$Category, [string]$Item, [string]$Status, [string]$Details = '', [bool]$RequiresReboot = $false, [string]$RiskNote = '')
    $Results.Add([PSCustomObject]@{
        ComputerName    = $ComputerName
        IPAddress       = $IPAddresses
        Category        = $Category
        Item            = $Item
        Status          = $Status
        Details         = $Details
        RequiresReboot  = $RequiresReboot
        RiskNote        = $RiskNote
        Timestamp       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }) | Out-Null
    if ($RequiresReboot -and $Status -eq 'Applied') { $RebootRequiredItems.Add($Item) | Out-Null }
}

# Runs $CheckBlock to see if the item already matches the desired state; if it does, records
# AlreadyCompliant without touching anything. Otherwise, either records DryRun (no -Apply) or
# goes through ShouldProcess and $ApplyBlock (with -Apply), catching and recording any failure.
function Invoke-Remediation {
    param(
        [string]$Category,
        [string]$Item,
        [string]$Description,
        [scriptblock]$CheckBlock,
        [scriptblock]$ApplyBlock,
        [bool]$RequiresReboot = $false,
        [string]$RiskNote = ''
    )
    try {
        $alreadyCompliant = & $CheckBlock
    } catch {
        Add-ResultRow -Category $Category -Item $Item -Status 'Unable to Verify' -Details "Pre-check failed: $($_.Exception.Message)" -RiskNote $RiskNote
        return
    }
    if ($alreadyCompliant) {
        Add-ResultRow -Category $Category -Item $Item -Status 'AlreadyCompliant' -Details $Description -RiskNote $RiskNote
        return
    }
    if (-not $Apply) {
        Add-ResultRow -Category $Category -Item $Item -Status 'DryRun - would apply' -Details $Description -RequiresReboot $RequiresReboot -RiskNote $RiskNote
        return
    }
    if ($PSCmdlet.ShouldProcess("$ComputerName - $Item", $Description)) {
        try {
            & $ApplyBlock
            Add-ResultRow -Category $Category -Item $Item -Status 'Applied' -Details $Description -RequiresReboot $RequiresReboot -RiskNote $RiskNote
        } catch {
            Add-ResultRow -Category $Category -Item $Item -Status 'Failed' -Details "$Description - ERROR: $($_.Exception.Message)" -RequiresReboot $RequiresReboot -RiskNote $RiskNote
        }
    } else {
        Add-ResultRow -Category $Category -Item $Item -Status 'Skipped (declined at confirm prompt)' -Details $Description -RequiresReboot $RequiresReboot -RiskNote $RiskNote
    }
}

Write-Log "Run started. Build $ScriptBuild. Computer: $ComputerName ($IPAddresses). Apply: $Apply. Elevated: $IsElevated"

# ===========================================================================
# 1. VBS & Credential Protection
# ===========================================================================
if (-not $SkipCredentialGuard) {
    $vbsCaveat = "Requires a reboot AND requires the VM's virtual hardware to expose Secure Boot + hardware-assisted virtualization (a vSphere-side VM setting this script cannot see/change from inside the guest) - if that's not enabled at the VM level, this will still show 'Configured but NOT Running' after reboot."

    $p = @{
        Category = 'VBS & Credential Protection'; Item = 'Enable Virtualization Based Security'
        Description = 'Set HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard EnableVirtualizationBasedSecurity=1, RequirePlatformSecurityFeatures=1 (Secure Boot)'
        RequiresReboot = $true; RiskNote = $vbsCaveat
        CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Name EnableVirtualizationBasedSecurity -ErrorAction SilentlyContinue; $v -and $v.EnableVirtualizationBasedSecurity -eq 1 }
        ApplyBlock = {
            New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Name EnableVirtualizationBasedSecurity -Value 1 -Type DWord -Force
            Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Name RequirePlatformSecurityFeatures -Value 1 -Type DWord -Force
        }
    }
    Invoke-Remediation @p

    $p = @{
        Category = 'VBS & Credential Protection'; Item = 'Enable HVCI / Memory Integrity'
        Description = 'Set HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity Enabled=1'
        RequiresReboot = $true; RiskNote = $vbsCaveat
        CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' -Name Enabled -ErrorAction SilentlyContinue; $v -and $v.Enabled -eq 1 }
        ApplyBlock = {
            New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' -Name Enabled -Value 1 -Type DWord -Force
        }
    }
    Invoke-Remediation @p

    $p = @{
        Category = 'VBS & Credential Protection'; Item = 'Enable Credential Guard'
        Description = 'Set HKLM:\SYSTEM\CurrentControlSet\Control\LSA LsaCfgFlags=1 (Enabled with UEFI lock - Microsoft-recommended default; harder to disable later without physical/firmware access, which is intentional)'
        RequiresReboot = $true; RiskNote = $vbsCaveat
        CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -ErrorAction SilentlyContinue; $v -and $v.LsaCfgFlags -in @(1, 2) }
        ApplyBlock = { Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -Value 1 -Type DWord -Force }
    }
    Invoke-Remediation @p

    $p = @{
        Category = 'VBS & Credential Protection'; Item = 'Enable Machine Identity Isolation'
        Description = 'Set HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard MachineIdentityIsolation=2 (Enabled - Enforcement Mode; machine password becomes IUM-bound only, isolated from LSASS). This is the same registry value the "Machine Identity Isolation Configuration" option under the "Turn On Virtualization Based Security" GPO writes.'
        RequiresReboot = $true
        RiskNote = "$vbsCaveat Also: this policy is confirmed in Microsoft's DeviceGuard Policy CSP reference but is not confirmed generally available on every Windows Server 2025 build - if unsupported on this build, the write is a harmless no-op (verify the 'Machine Identity Isolation Configuration' dropdown is present under gpedit.msc if in doubt)."
        CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' -Name MachineIdentityIsolation -ErrorAction SilentlyContinue; $v -and $v.MachineIdentityIsolation -eq 2 }
        ApplyBlock = {
            New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' -Name MachineIdentityIsolation -Value 2 -Type DWord -Force
        }
    }
    Invoke-Remediation @p
} else {
    Add-ResultRow -Category 'VBS & Credential Protection' -Status 'Skipped (category)' -Item 'All items' -Details '-SkipCredentialGuard was passed'
}

# ===========================================================================
# 2. Microsoft Defender
# ===========================================================================
if (-not $SkipDefenderASR) {
    $mpAvailable = [bool](Get-Command Set-MpPreference -ErrorAction SilentlyContinue)
    if (-not $mpAvailable) {
        Add-ResultRow -Category 'Microsoft Defender' -Item 'All items' -Status 'Manual/External Required' -Details 'Set-MpPreference/Add-MpPreference not found - Defender module absent (disabled, uninstalled, or replaced by third-party AV). Remediate via that product''s own console.'
    } else {
        $p = @{
            Category = 'Microsoft Defender'; Item = 'Real-Time Protection'
            Description = 'Set-MpPreference -DisableRealtimeMonitoring $false'
            CheckBlock = { $s = Get-MpComputerStatus -ErrorAction Stop; $s.RealTimeProtectionEnabled -eq $true }
            ApplyBlock = { Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop }
        }
        Invoke-Remediation @p

        $p = @{
            Category = 'Microsoft Defender'; Item = 'Cloud-Delivered Protection (MAPS Reporting)'
            Description = "Set-MpPreference -MAPSReporting Advanced -SubmitSamplesConsent SendSafeSamples"
            CheckBlock = { $pref = Get-MpPreference -ErrorAction Stop; $pref.MAPSReporting -eq 2 }
            ApplyBlock = { Set-MpPreference -MAPSReporting Advanced -SubmitSamplesConsent SendSafeSamples -ErrorAction Stop }
        }
        Invoke-Remediation @p

        # Standard Microsoft-recommended ASR rule set (GUID -> friendly name for logging only).
        $AsrRules = [ordered]@{
            '56a863a9-875e-4185-98a7-b882c64b5ce5' = 'Block abuse of exploited vulnerable signed drivers'
            '7674ba52-37eb-4a4f-a9a1-f0f9a1619a2c' = 'Block Adobe Reader from creating child processes'
            'd4f940ab-401b-4efc-aadc-ad5f3c50688a' = 'Block all Office applications from creating child processes'
            '9e6c4e1f-7d60-472f-ba1a-a39ef669e4b2' = 'Block credential stealing from LSASS'
            'be9ba2d9-53ea-4cdc-84e5-9b1eeee46550' = 'Block executable content from email client/webmail'
            '01443614-cd74-433a-b99e-2ecdc07bfc25' = 'Block executable files from running unless prevalence/age/trusted-list criteria met'
            '5beb7efe-fd9a-4556-801d-275e5ffc04cc' = 'Block execution of potentially obfuscated scripts'
            'd3e037e1-3eb8-44c8-a917-57927947596d' = 'Block JavaScript/VBScript from launching downloaded executable content'
            '3b576869-a4ec-4529-8536-b80a7769e899' = 'Block Office applications from creating executable content'
            '75668c1f-73b5-4cf0-bb93-3ecf5cb7cc84' = 'Block Office applications from injecting code into other processes'
            '26190899-1602-49e8-8b27-eb1d0a1ce869' = 'Block Office communication application from creating child processes'
            'e6db77e5-3df2-4cf1-b95a-636979351e5b' = 'Block persistence through WMI event subscription'
            'd1e49aac-8f56-4280-b9ba-993a6d77406c' = 'Block process creations from PSExec and WMI commands'
            '33ddedf1-c6e0-47cb-833e-de6133960387' = 'Block rebooting machine in Safe Mode'
            'b2b3f03d-6a65-4f7b-a9c7-1c7ef74a9ba4' = 'Block untrusted/unsigned processes running from USB'
            '92e97fa1-2edf-4476-bdd6-9dd0b4dddc7b' = 'Block Win32 API calls from Office macros'
        }
        $ruleIdsCsv = ($AsrRules.Keys -join ',')
        $p = @{
            Category = 'Microsoft Defender'; Item = 'Attack Surface Reduction (ASR) Rules'
            Description = "Add-MpPreference -AttackSurfaceReductionRules_Ids <16 standard rules> -AttackSurfaceReductionRules_Actions $AsrRuleAction (rule GUIDs: $ruleIdsCsv)"
            RiskNote = 'ASR rules can break legitimate line-of-business behavior (Office macros, script interpreters, PSExec-based tooling). Consider -AsrRuleAction AuditMode for a pilot period before switching to Enabled (block) fleet-wide.'
            CheckBlock = {
                $pref = Get-MpPreference -ErrorAction Stop
                $ids = @($pref.AttackSurfaceReductionRules_Ids)
                $actions = @($pref.AttackSurfaceReductionRules_Actions)
                $desiredActionCode = if ($AsrRuleAction -eq 'Enabled') { 1 } else { 2 }
                $allSet = $true
                foreach ($ruleId in $AsrRules.Keys) {
                    $idx = [array]::IndexOf($ids, $ruleId)
                    if ($idx -lt 0 -or [int]$actions[$idx] -ne $desiredActionCode) { $allSet = $false; break }
                }
                $allSet
            }
            ApplyBlock = { Add-MpPreference -AttackSurfaceReductionRules_Ids @($AsrRules.Keys) -AttackSurfaceReductionRules_Actions $AsrRuleAction -ErrorAction Stop }
        }
        Invoke-Remediation @p

        Add-ResultRow -Category 'Microsoft Defender' -Item 'Tamper Protection' -Status 'Manual/External Required' -Details 'Microsoft blocks changing Tamper Protection via PowerShell/script by design (prevents malware from disabling it the same way). Enable via the Windows Security app on this server, or centrally via Intune/Microsoft Defender for Endpoint policy.'
    }
} else {
    Add-ResultRow -Category 'Microsoft Defender' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipDefenderASR was passed'
}

# ===========================================================================
# 3. Authentication Security
# ===========================================================================
if (-not $SkipAuthHardening) {
    $p = @{
        Category = 'Authentication Security'; Item = 'Disable WDigest (UseLogonCredential)'
        Description = 'Set HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest UseLogonCredential=0'
        CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' -Name UseLogonCredential -ErrorAction SilentlyContinue; $v -and $v.UseLogonCredential -eq 0 }
        ApplyBlock = {
            New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' -Name UseLogonCredential -Value 0 -Type DWord -Force
        }
    }
    Invoke-Remediation @p

    if ($EnforceAesOnlyKerberos) {
        $p = @{
            Category = 'Authentication Security'; Item = 'Kerberos Encryption Types - Enforce AES-only'
            Description = 'Set HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters SupportedEncryptionTypes=0x18 (AES128+AES256 ONLY - removes RC4/DES support)'
            RiskNote = 'REMOVES RC4/DES support. Will break Kerberos auth for any client, service account, or trust that has not been confirmed to support AES. Only run this after confirming no legacy dependency needs RC4.'
            CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' -Name SupportedEncryptionTypes -ErrorAction SilentlyContinue; $v -and $v.SupportedEncryptionTypes -eq 0x18 }
            ApplyBlock = {
                New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' -Force -ErrorAction SilentlyContinue | Out-Null
                Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' -Name SupportedEncryptionTypes -Value 0x18 -Type DWord -Force
            }
        }
        Invoke-Remediation @p
    } else {
        $p = @{
            Category = 'Authentication Security'; Item = 'Kerberos Encryption Types - Ensure AES Supported'
            Description = 'Add AES128+AES256 (0x18) to whatever SupportedEncryptionTypes bits are already set under HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters - never removes an existing type, so this cannot break current auth. Pass -EnforceAesOnlyKerberos to additionally remove RC4/DES.'
            CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' -Name SupportedEncryptionTypes -ErrorAction SilentlyContinue; $v -and (([int]$v.SupportedEncryptionTypes) -band 0x18) -eq 0x18 }
            ApplyBlock = {
                $path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters'
                New-Item -Path $path -Force -ErrorAction SilentlyContinue | Out-Null
                $existing = (Get-ItemProperty -Path $path -Name SupportedEncryptionTypes -ErrorAction SilentlyContinue).SupportedEncryptionTypes
                if ($null -eq $existing) { $existing = 0 }
                $newValue = ([int]$existing) -bor 0x18
                Set-ItemProperty -Path $path -Name SupportedEncryptionTypes -Value $newValue -Type DWord -Force
            }
        }
        Invoke-Remediation @p
    }
} else {
    Add-ResultRow -Category 'Authentication Security' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipAuthHardening was passed'
}

# ===========================================================================
# 4. Remote Management (WinRM)
# ===========================================================================
if (-not $SkipWinRM) {
    if ($DisableWinRM) {
        $p = @{
            Category = 'Remote Management'; Item = 'Disable WinRM Service'
            Description = 'Stop-Service WinRM; Set-Service WinRM -StartupType Disabled'
            RiskNote = 'Fully removes remote-management capability via WinRM on this server. Confirm no other tooling (monitoring agents, other automation) depends on it before applying.'
            CheckBlock = { $svc = Get-Service -Name WinRM -ErrorAction Stop; $svc.Status -eq 'Stopped' -and (Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'").StartMode -eq 'Disabled' }
            ApplyBlock = { Stop-Service -Name WinRM -Force -ErrorAction Stop; Set-Service -Name WinRM -StartupType Disabled -ErrorAction Stop }
        }
        Invoke-Remediation @p
    } else {
        $winrmSvcPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service'
        if (-not (Test-Path $winrmSvcPath)) {
            Add-ResultRow -Category 'Remote Management' -Item 'Harden WinRM (disable unencrypted traffic + Basic auth)' -Status 'Not Applicable' -Details "$winrmSvcPath does not exist - WinRM has never been configured on this host, so there is nothing to harden. Run 'winrm quickconfig' first if WinRM management access is actually needed here."
        } else {
            $p = @{
                Category = 'Remote Management'; Item = 'Harden WinRM (disable unencrypted traffic + Basic auth)'
                Description = "Set $winrmSvcPath AllowUnencrypted=0, auth_Basic=0. Pass -DisableWinRM instead to fully stop/disable the service."
                CheckBlock = { $v = Get-ItemProperty -Path $winrmSvcPath -ErrorAction SilentlyContinue; $v -and $v.AllowUnencrypted -eq 0 -and $v.auth_Basic -eq 0 }
                ApplyBlock = {
                    Set-ItemProperty -Path $winrmSvcPath -Name AllowUnencrypted -Value 0 -Type DWord -Force
                    Set-ItemProperty -Path $winrmSvcPath -Name auth_Basic -Value 0 -Type DWord -Force
                }
            }
            Invoke-Remediation @p
        }

        if (-not $WinRmAllowedSourceRange -or $WinRmAllowedSourceRange.Count -eq 0) {
            Add-ResultRow -Category 'Remote Management' -Item 'Restrict WinRM Listener to Management Sources' -Status 'Manual/External Required' -Details "Not restricted - no -WinRmAllowedSourceRange supplied. Re-run with -WinRmAllowedSourceRange @('10.1.1.0/24') (your actual management subnet(s)/host IPs) to scope the 'Windows Remote Management' firewall rules to only those sources. Left open to any source by default rather than guessed, since a wrong assumption here could lock out legitimate admin access."
        } elseif (-not (Get-Command Get-NetFirewallRule -ErrorAction SilentlyContinue)) {
            Add-ResultRow -Category 'Remote Management' -Item 'Restrict WinRM Listener to Management Sources' -Status 'Unable to Verify' -Details 'Get-NetFirewallRule/Set-NetFirewallRule not available (NetSecurity module missing) - cannot inspect or restrict the WinRM firewall rules on this host.'
        } else {
            $p = @{
                Category = 'Remote Management'; Item = 'Restrict WinRM Listener to Management Sources'
                Description = "Set-NetFirewallRule on the 'Windows Remote Management' firewall rule group -RemoteAddress $($WinRmAllowedSourceRange -join ', ') - only these sources will be able to reach the WinRM listener."
                RiskNote = 'If this list omits a host/subnet that legitimately needs WinRM access (e.g. this very management workstation, or a monitoring server), that access will be cut off. Double-check the range before applying.'
                CheckBlock = {
                    $rules = @(Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' })
                    if ($rules.Count -eq 0) { throw "No enabled 'Windows Remote Management' firewall rules found - cannot verify/restrict scope." }
                    $allMatch = $true
                    foreach ($rule in $rules) {
                        $currentAddr = @($rule | Get-NetFirewallAddressFilter -ErrorAction Stop | Select-Object -ExpandProperty RemoteAddress)
                        $desired = @($WinRmAllowedSourceRange)
                        if (@(Compare-Object $currentAddr $desired -SyncWindow 0).Count -ne 0) { $allMatch = $false }
                    }
                    $allMatch
                }
                ApplyBlock = {
                    $rules = Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction Stop
                    foreach ($rule in $rules) { $rule | Set-NetFirewallRule -RemoteAddress $WinRmAllowedSourceRange -ErrorAction Stop }
                }
            }
            Invoke-Remediation @p
        }
    }
} else {
    Add-ResultRow -Category 'Remote Management' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipWinRM was passed'
}

# ===========================================================================
# 5. SMB & RPC Security
# ===========================================================================
if (-not $SkipSmbRpc) {
    if (Get-Command Set-SmbServerConfiguration -ErrorAction SilentlyContinue) {
        $p = @{
            Category = 'SMB & RPC Security'; Item = 'Enable SMB1 Access Auditing'
            Description = 'Set-SmbServerConfiguration -AuditSmb1Access $true'
            CheckBlock = {
                $smb = Get-SmbServerConfiguration -ErrorAction Stop
                if ($smb.PSObject.Properties.Name -notcontains 'AuditSmb1Access') { throw 'AuditSmb1Access property not present on this OS build (needs Windows Server 2022+) - Not Applicable, not a failure.' }
                $smb.AuditSmb1Access -eq $true
            }
            ApplyBlock = { Set-SmbServerConfiguration -AuditSmb1Access $true -Confirm:$false -ErrorAction Stop }
        }
        Invoke-Remediation @p
    } else {
        Add-ResultRow -Category 'SMB & RPC Security' -Item 'Enable SMB1 Access Auditing' -Status 'Not Applicable' -Details 'Get-SmbServerConfiguration/Set-SmbServerConfiguration not available (SMB module missing or OS build predates AuditSmb1Access, which needs Windows Server 2022+).'
    }

    $p = @{
        Category = 'SMB & RPC Security'; Item = 'Printer RPC Packet Privacy (PrintNightmare mitigation)'
        Description = 'Set HKLM:\SYSTEM\CurrentControlSet\Control\Print RpcAuthnLevelPrivacyEnabled=1'
        CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' -Name RpcAuthnLevelPrivacyEnabled -ErrorAction SilentlyContinue; $v -and $v.RpcAuthnLevelPrivacyEnabled -eq 1 }
        ApplyBlock = { Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' -Name RpcAuthnLevelPrivacyEnabled -Value 1 -Type DWord -Force }
    }
    Invoke-Remediation @p

    $p = @{
        Category = 'SMB & RPC Security'; Item = 'Point and Print Restrictions'
        Description = 'Set HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint RestrictDriverInstallationToAdministrators=1, NoWarningNoElevationOnInstall=0'
        CheckBlock = {
            $v = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' -ErrorAction SilentlyContinue
            $v -and $v.RestrictDriverInstallationToAdministrators -eq 1 -and $v.NoWarningNoElevationOnInstall -eq 0
        }
        ApplyBlock = {
            $path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint'
            New-Item -Path $path -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path $path -Name RestrictDriverInstallationToAdministrators -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $path -Name NoWarningNoElevationOnInstall -Value 0 -Type DWord -Force
        }
    }
    Invoke-Remediation @p
} else {
    Add-ResultRow -Category 'SMB & RPC Security' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipSmbRpc was passed'
}

# ===========================================================================
# 6. Windows Server 2025 Security Baseline
# ===========================================================================
if (-not $SkipBaselineIndicators) {
    $p = @{
        Category = '2025 Security Baseline'; Item = 'SmartScreen (Explorer policy)'
        Description = 'Set HKLM:\SOFTWARE\Policies\Microsoft\Windows\System EnableSmartScreen=1'
        CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name EnableSmartScreen -ErrorAction SilentlyContinue; $v -and $v.EnableSmartScreen -eq 1 }
        ApplyBlock = {
            New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name EnableSmartScreen -Value 1 -Type DWord -Force
        }
    }
    Invoke-Remediation @p

    $p = @{
        Category = '2025 Security Baseline'; Item = 'IE Zone (Internet) - Disable ActiveX + Scripting, Enable Protected Mode'
        Description = 'Set HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3 : 1200=3 (Disable ActiveX), 1400=3 (Disable Scripting), 2500=0 (Protected Mode ON)'
        CheckBlock = {
            $v = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3' -ErrorAction SilentlyContinue
            $v -and $v.'1200' -eq 3 -and $v.'1400' -eq 3 -and $v.'2500' -eq 0
        }
        ApplyBlock = {
            $path = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'
            New-Item -Path $path -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path $path -Name '1200' -Value 3 -Type DWord -Force
            Set-ItemProperty -Path $path -Name '1400' -Value 3 -Type DWord -Force
            Set-ItemProperty -Path $path -Name '2500' -Value 0 -Type DWord -Force
        }
    }
    Invoke-Remediation @p

    if (Get-Command Get-WindowsOptionalFeature -ErrorAction SilentlyContinue) {
        $p = @{
            Category = '2025 Security Baseline'; Item = 'Remove Legacy Internet Explorer 11 Feature'
            Description = 'Disable-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -NoRestart'
            RequiresReboot = $true
            CheckBlock = { (Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction Stop).State -eq 'Disabled' }
            ApplyBlock = { Disable-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -NoRestart -ErrorAction Stop | Out-Null }
        }
        Invoke-Remediation @p
    } else {
        Add-ResultRow -Category '2025 Security Baseline' -Item 'Remove Legacy Internet Explorer 11 Feature' -Status 'Not Applicable' -Details 'Get-WindowsOptionalFeature not available (e.g. Server Core or restricted session).'
    }

    if ($DisableLegacyTls) {
        $protoList = @('SSL 2.0', 'SSL 3.0', 'TLS 1.0', 'TLS 1.1')
        foreach ($proto in $protoList) {
            foreach ($side in @('Client', 'Server')) {
                $sidePath = Join-Path (Join-Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols' $proto) $side
                $p = @{
                    Category = '2025 Security Baseline'; Item = "Disable $proto ($side)"
                    Description = "Set $sidePath Enabled=0, DisabledByDefault=1"
                    RequiresReboot = $true
                    RiskNote = 'High app-compat risk: breaks any client/integration still requiring this protocol version. Confirmed opt-in via -DisableLegacyTls.'
                    CheckBlock = { $v = Get-ItemProperty -Path $sidePath -ErrorAction SilentlyContinue; $v -and $v.Enabled -eq 0 -and $v.DisabledByDefault -eq 1 }
                    ApplyBlock = {
                        New-Item -Path $sidePath -Force -ErrorAction SilentlyContinue | Out-Null
                        Set-ItemProperty -Path $sidePath -Name Enabled -Value 0 -Type DWord -Force
                        Set-ItemProperty -Path $sidePath -Name DisabledByDefault -Value 1 -Type DWord -Force
                    }
                }
                Invoke-Remediation @p
            }
        }
    } else {
        Add-ResultRow -Category '2025 Security Baseline' -Item 'Disable Legacy SSL/TLS (SSL 2.0/3.0, TLS 1.0/1.1)' -Status 'Manual/External Required' -Details 'Skipped by default - high app-compat risk. Re-run with -DisableLegacyTls once you have confirmed no client/integration still requires these protocol versions.'
    }

    $lgpoReady = $LgpoPath -and (Test-Path $LgpoPath) -and $BaselineGpoBackupPath -and (Test-Path $BaselineGpoBackupPath)
    if ($lgpoReady) {
        $p = @{
            Category = '2025 Security Baseline'; Item = 'Apply full Windows Server 2025 Security Baseline GPO via LGPO.exe'
            Description = "& '$LgpoPath' /g '$BaselineGpoBackupPath'"
            RequiresReboot = $true
            RiskNote = 'Applies hundreds of settings at once (the actual baseline, not just the narrow items above). Review the baseline GPO backup contents before running against production if you have not already.'
            CheckBlock = { $false }
            ApplyBlock = {
                $lgpoOutput = & $LgpoPath /g $BaselineGpoBackupPath 2>&1
                if ($LASTEXITCODE -ne 0) { throw "LGPO.exe exited with code $LASTEXITCODE. Output: $lgpoOutput" }
            }
        }
        Invoke-Remediation @p
    } else {
        Add-ResultRow -Category '2025 Security Baseline' -Item 'Apply full Windows Server 2025 Security Baseline GPO' -Status 'Manual/External Required' -Details "Not staged on this run (-LgpoPath/-BaselineGpoBackupPath not both supplied/valid). This script only remediates the narrow SmartScreen/IE-zone/legacy-IE items above, NOT the full official baseline. To apply the real baseline: download the Windows Server 2025 Security Baseline from Microsoft's Security Compliance Toolkit, extract it, then either re-run this script with -LgpoPath '<path to LGPO.exe>' -BaselineGpoBackupPath '<path to extracted baseline GPO backup>', or run 'LGPO.exe /g <path>' yourself, or link the baseline GPO in Active Directory for the OU containing these servers."
    }
} else {
    Add-ResultRow -Category '2025 Security Baseline' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipBaselineIndicators was passed'
}

# ---------------------------------------------------------------------------------------------
# On-screen summary + export
# ---------------------------------------------------------------------------------------------
Write-Host ""
Write-Host "=== $ComputerName ($IPAddresses) ===" -ForegroundColor Green
foreach ($row in $Results) {
    Write-Host ("  {0,-25} {1,-55} : {2}" -f $row.Category, $row.Item, $row.Status)
}

if ($RebootRequiredItems.Count -gt 0) {
    Write-Host ""
    Write-Host "REBOOT REQUIRED to complete these applied changes:" -ForegroundColor Red
    foreach ($item in ($RebootRequiredItems | Select-Object -Unique)) { Write-Host "  - $item" -ForegroundColor Red }
    Write-Host "This script does NOT reboot the server itself - schedule these during a maintenance window, then re-run Get-WindowsServerSecurityAudit-Local.ps1 to confirm they took effect." -ForegroundColor Red
}

$Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV exported: $CsvPath"

$failedCount = @($Results | Where-Object { $_.Status -eq 'Failed' }).Count
if ($failedCount -gt 0) { Write-Log "$failedCount item(s) FAILED to apply - see CSV/log for details." 'WARN' }

Write-Log "Run complete. $($Results.Count) total item rows for $ComputerName. Apply=$Apply."
Write-Host ""
Write-Host "Results: $CsvPath"
Write-Host "Log:     $LogPath"
