#Requires -Version 5.1
<#
.SYNOPSIS
    Centralised (PowerCLI) front-end for RepairWindowsServerSecurityGapsLocal.ps1.

    Runs the SAME Windows Server security assessment / remediation logic as the local
    script, but pushes it into multiple Windows Server guests from a single admin
    workstation using VMware Tools Guest Operations (Invoke-VMScript). No WinRM, no
    PsExec, no SMB admin shares, no RDP, and no IP connectivity from the admin box to
    the guests is required - everything goes admin -> vCenter -> ESXi -> VMware Tools.

.DESCRIPTION
    SAFETY MODEL (unchanged from the local script):
      * DRY-RUN BY DEFAULT. Without -Apply nothing is changed in any guest. Every item
        is evaluated and reported as "Compliant" / "Would remediate (dry-run)".
      * -Apply is the master switch. It is injected into the in-guest payload as
        $Apply = $true. The payload still refuses to change anything if the guest
        session is not elevated.
      * HIGH-RISK ITEMS STAY OPT-IN even under -Apply:
          -DisableWinRM              fully stop+disable WinRM (default: harden only)
          -EnforceAesOnlyKerberos    Kerberos AES-only, removes RC4/DES
                                     (default: only ADD AES, never remove a type)
          -DisableLegacyTls          disable SSL2/3 + TLS1.0/1.1 via Schannel
        None of these do anything unless you pass them explicitly here; the wrapper
        only ever injects $true for a switch you supplied.
      * NEVER reboots a guest. Items that need a reboot to take effect are still
        applied, flagged RequiresReboot=$true per row, and listed in the summary.
      * NEVER touches VMware / vSphere: no power operations, no snapshots, no hardware
        changes, no Secure Boot / vTPM / virtualization-extension changes, no
        networking changes. The only vCenter cmdlets used are Get-VM (read) and
        Invoke-VMScript (guest program execution).

    CONFIRMATION PIPELINE: the local script uses $PSCmdlet.ShouldProcess with
    ConfirmImpact 'High', which needs an interactive console to answer. That cannot
    work through Guest Operations, so the payload sets $ConfirmPreference='None' in
    the guest (only relevant when -Apply is set; the dry-run path never reaches
    ShouldProcess). The genuine safety gates - dry-run default, explicit -Apply,
    explicit high-risk switches, elevation check - are all still enforced. This
    wrapper is itself [CmdletBinding(SupportsShouldProcess)], so `-WhatIf` here
    downgrades every target to a dry-run and still produces the full report.

    PER-VM VALIDATION (each VM is validated independently; a failure skips only that
    VM and is recorded with an exact reason - one bad VM never stops the batch):
      1. VM exists in one of the currently connected vCenters
      2. VM name is not ambiguous / duplicated across the connected vCenters
      3. VM is powered on
      4. Guest OS is Windows
      5. VMware Tools is installed and running
      6. Guest authentication succeeds (probe via Invoke-VMScript)
      7. PowerShell 5.1+ is available in the guest (same probe)

.PARAMETER Apply
    Master switch. Without it: assessment only, nothing changes. With it: remediation
    runs in every validated guest, gated by the same rules as the local script.

.PARAMETER VMListPath
    Text file of VM names, one per line. Blank lines and lines starting with '#' are
    ignored. Default C:\temp\vmlist.txt.

.PARAMETER CredentialPath
    Export-Clixml PSCredential for the Windows guest admin account. Default
    C:\temp\wincred.xml. Decrypted in memory only; never written anywhere; never
    placed in script text, logs, CSVs or command lines.

.PARAMETER OutputPath
    Folder on the ADMIN machine for the central CSV + log. Created if missing.

.PARAMETER SkipCredentialGuard / SkipDefenderASR / SkipAuthHardening / SkipWinRM / SkipSmbRpc / SkipBaselineIndicators
    Skip a whole category (passed straight through to the in-guest script).

.PARAMETER SkipBaselineGpoApply
    Skip ONLY the "Apply full Windows Server 2025 Security Baseline GPO via LGPO.exe"
    item, leaving every other baseline check running. The local script's check for
    that item is deliberately always-not-compliant (it cannot verify hundreds of
    baseline settings), so without this switch it re-runs LGPO.exe and re-flags a
    reboot on EVERY run. Pass this once you have applied the baseline GPO and
    rebooted; re-run without it when you update the baseline backup.

.PARAMETER DisableWinRM
    Opt-in: fully stop and disable WinRM in the guest instead of hardening it.

.PARAMETER WinRmAllowedSourceRange
    Optional list of IPs / CIDR ranges to scope the "Windows Remote Management"
    firewall rules to. Each entry is validated to contain only [0-9A-Fa-f.:/] before
    it is injected. Without it, that step is reported "Manual Verification Required"
    (exactly like the local script).

.PARAMETER EnforceAesOnlyKerberos
    Opt-in: Kerberos SupportedEncryptionTypes = AES-only (removes RC4/DES).

.PARAMETER DisableLegacyTls
    Opt-in: disable SSL 2.0/3.0 and TLS 1.0/1.1 via Schannel. High app-compat risk.

.PARAMETER AsrRuleAction
    'Enabled' (block, default) or 'AuditMode' (log only) for the ASR rule set.

.PARAMETER LgpoPath / BaselineGpoBackupPath
    GUEST-SIDE paths. If both exist inside the guest, -Apply runs
    `LGPO.exe /g <BaselineGpoBackupPath>` there. If not, that step is reported
    "Manual Verification Required" with the exact command - same as the local script.
    You must stage those files in the guest yourself (this wrapper never copies
    files into a guest).

.PARAMETER ToolsWaitSecs
    Seconds Invoke-VMScript waits for VMware Tools to be responsive. Default 180.

.EXAMPLE
    # 1) Pre-check / dry run - changes nothing anywhere:
    .\Invoke-RemoteSecurityGapRemediation.ps1

.EXAMPLE
    # 2) Controlled remediation with the safe defaults:
    .\Invoke-RemoteSecurityGapRemediation.ps1 -Apply

.EXAMPLE
    # 2b) Remediation including the high-risk opt-ins:
    .\Invoke-RemoteSecurityGapRemediation.ps1 -Apply -DisableWinRM -EnforceAesOnlyKerberos -DisableLegacyTls

.EXAMPLE
    # Preview what -Apply would do, per VM, without executing in any guest:
    .\Invoke-RemoteSecurityGapRemediation.ps1 -Apply -WhatIf

.NOTES
    Prerequisites on the admin machine:
      * VMware PowerCLI installed and imported.
      * Already connected to the target vCenter(s): Connect-VIServer ... (this script
        does NOT connect or disconnect - it uses the sessions you already have).
      * The vCenter account needs the guest-operation privileges:
        "Guest Operation Program Execution", "Guest Operation Modifications",
        "Guest Operation Queries".
    This wrapper does not snapshot, reboot or modify any VM. If your change process
    requires a snapshot before -Apply, take it yourself first.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [switch]  $Apply,

    [string]  $VMListPath     = 'C:\temp\vmlist.txt',
    [string]  $CredentialPath = 'C:\temp\wincred.xml',
    [string]  $OutputPath     = (Join-Path $PSScriptRoot 'SecurityGap_Reports'),

    [switch]  $SkipCredentialGuard,
    [switch]  $SkipDefenderASR,
    [switch]  $SkipAuthHardening,
    [switch]  $SkipWinRM,
    [switch]  $SkipSmbRpc,
    [switch]  $SkipBaselineIndicators,
    [switch]  $SkipBaselineGpoApply,

    [switch]  $DisableWinRM,
    [string[]]$WinRmAllowedSourceRange,
    [switch]  $EnforceAesOnlyKerberos,
    [switch]  $DisableLegacyTls,

    [ValidateSet('Enabled', 'AuditMode')]
    [string]  $AsrRuleAction = 'Enabled',

    [string]  $LgpoPath              = 'C:\Tools\LGPO\LGPO.exe',
    [string]  $BaselineGpoBackupPath = 'C:\Tools\Baseline\GPOs',

    [ValidateRange(30, 3600)]
    [int]     $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'
$ScriptBuild = '2026-08-28g-baseline-gpo-optout'

# =====================================================================================
# 0. Banner
# =====================================================================================
Write-Host ("Invoke-RemoteSecurityGapRemediation.ps1  [build $ScriptBuild]") -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
Write-Host ("If the build above is not '2026-08-28g-baseline-gpo-optout' you are running an OLD copy - update it.") -ForegroundColor DarkGray
if ($Apply) {
    Write-Host "*** -Apply IS SET: validated guests WILL have configuration changed. ***" -ForegroundColor Red
    Write-Host "    This tool never snapshots, reboots, or alters any VM/vSphere setting." -ForegroundColor Yellow
    Write-Host "    Take snapshots via your normal change process first if policy requires it." -ForegroundColor Yellow
} else {
    Write-Host "DRY-RUN MODE (default) - no guest will be changed. Pass -Apply to remediate." -ForegroundColor Yellow
}

# =====================================================================================
# 1. Pre-flight on the admin machine
# =====================================================================================
if (-not (Get-Command Invoke-VMScript -ErrorAction SilentlyContinue)) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop }
    catch { throw "VMware PowerCLI (VMware.VimAutomation.Core) is not available. Install PowerCLI and retry." }
}

$connectedServers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($connectedServers.Count -eq 0) {
    throw "No connected vCenter session. Run Connect-VIServer <vcenter> first; this script uses the existing session(s)."
}
Write-Host ("Connected vCenter(s): {0}" -f (($connectedServers | ForEach-Object { $_.Name }) -join ', ')) -ForegroundColor Gray

if (-not (Test-Path -LiteralPath $VMListPath)) { throw "VM list not found: $VMListPath" }
$vmNames = @(Get-Content -LiteralPath $VMListPath |
             ForEach-Object { $_.Trim() } |
             Where-Object { $_ -and -not $_.StartsWith('#') } |
             Select-Object -Unique)
if ($vmNames.Count -eq 0) { throw "VM list '$VMListPath' contained no usable VM names." }
Write-Host ("Target VMs: {0}" -f $vmNames.Count) -ForegroundColor Gray

if (-not (Test-Path -LiteralPath $CredentialPath)) { throw "Guest credential file not found: $CredentialPath" }
try {
    $GuestCredential = Import-Clixml -LiteralPath $CredentialPath
} catch {
    throw ("Failed to import guest credential from '$CredentialPath'. Export-Clixml credentials are DPAPI-protected " +
           "and can only be read by the same Windows account on the same machine that created them. Recreate it in " +
           "this context with:  Get-Credential | Export-Clixml '$CredentialPath'.  Underlying error: $($_.Exception.Message)")
}
if (-not ($GuestCredential -is [pscredential])) {
    throw "'$CredentialPath' did not deserialise to a PSCredential object."
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$CsvName  = if ($Apply) { "SecurityGap_Remediation_$RunStamp.csv" } else { "SecurityGap_PreCheck_$RunStamp.csv" }
$CsvPath  = Join-Path $OutputPath $CsvName
$LogPath  = Join-Path $OutputPath "SecurityGap_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Run started. Build=$ScriptBuild. Script=$PSCommandPath. Apply=$Apply. VMs=$($vmNames.Count). vCenter(s)=$(($connectedServers | ForEach-Object { $_.Name }) -join ', ')."

# =====================================================================================
# 2. In-guest payload template (verbatim security logic from
#    RepairWindowsServerSecurityGapsLocal.ps1; only the param() surface and the
#    reporting tail are adapted for centralised collection). #__INJECT_PARAMS__ is
#    replaced with literal variable assignments before each run - it never carries
#    a secret.
# =====================================================================================
$PayloadTemplate = @'
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param()

#__INJECT_PARAMS__

$ProgressPreference = 'SilentlyContinue'
$ScriptBuild = '2026-08-17-02-winrm-listener-filter / remote-payload-r1'

$Results             = New-Object System.Collections.Generic.List[object]
$RebootRequiredItems = New-Object System.Collections.Generic.List[string]
$GuestLog            = New-Object System.Collections.Generic.List[string]
function Write-GuestLog { param([string]$Message, [string]$Level = 'INFO')
    $GuestLog.Add(("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)) | Out-Null
}

$ComputerName = $env:COMPUTERNAME
try {
    $IPAddresses = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
        Where-Object { $_.InterfaceAlias -notmatch 'Loopback' -and $_.IPAddress -notmatch '^169\.254\.' }).IPAddress -join ', '
} catch { $IPAddresses = $null }
if ([string]::IsNullOrWhiteSpace($IPAddresses)) {
    try {
        $IPAddresses = ([System.Net.Dns]::GetHostAddresses($ComputerName) |
            Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString
    } catch { $IPAddresses = 'Unknown' }
}
try { $OSCaption = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Caption } catch { $OSCaption = 'Unknown' }

function Add-ResultRow {
    param([string]$Category, [string]$Item, [string]$Status, [string]$Details = '',
          [bool]$RequiresReboot = $false, [string]$RiskNote = '',
          [string]$DetectedValue = '', [string]$ExpectedValue = '')
    $Results.Add([PSCustomObject]@{
        Category       = $Category
        Item           = $Item
        Status         = $Status
        Details        = $Details
        DetectedValue  = $DetectedValue
        ExpectedValue  = $ExpectedValue
        RequiresReboot = $RequiresReboot
        RiskNote       = $RiskNote
        Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }) | Out-Null
    if ($RequiresReboot -and $Status -eq 'Applied') { $RebootRequiredItems.Add($Item) | Out-Null }
}

# Same contract as the local script's Invoke-Remediation, plus an optional
# CurrentStateBlock/ExpectedValue used only to enrich the report (it never changes
# what is checked or applied).
function Invoke-Remediation {
    param(
        [string]$Category, [string]$Item, [string]$Description,
        [scriptblock]$CheckBlock, [scriptblock]$ApplyBlock,
        [bool]$RequiresReboot = $false, [string]$RiskNote = '',
        [scriptblock]$CurrentStateBlock, [string]$ExpectedValue = ''
    )
    $detected = ''
    if ($CurrentStateBlock) {
        try { $detected = [string](& $CurrentStateBlock) } catch { $detected = "(unreadable: $($_.Exception.Message))" }
    }
    try {
        $alreadyCompliant = & $CheckBlock
    } catch {
        Add-ResultRow -Category $Category -Item $Item -Status 'Unable to Verify' -Details "Pre-check failed: $($_.Exception.Message)" -RiskNote $RiskNote -DetectedValue $detected -ExpectedValue $ExpectedValue
        return
    }
    if ($alreadyCompliant) {
        Add-ResultRow -Category $Category -Item $Item -Status 'AlreadyCompliant' -Details $Description -RiskNote $RiskNote -DetectedValue $detected -ExpectedValue $ExpectedValue
        return
    }
    if (-not $Apply) {
        Add-ResultRow -Category $Category -Item $Item -Status 'DryRun - would apply' -Details $Description -RequiresReboot $RequiresReboot -RiskNote $RiskNote -DetectedValue $detected -ExpectedValue $ExpectedValue
        return
    }
    if ($PSCmdlet.ShouldProcess("$ComputerName - $Item", $Description)) {
        try {
            & $ApplyBlock
            $after = $detected
            if ($CurrentStateBlock) { try { $after = [string](& $CurrentStateBlock) } catch { } }
            Add-ResultRow -Category $Category -Item $Item -Status 'Applied' -Details $Description -RequiresReboot $RequiresReboot -RiskNote $RiskNote -DetectedValue $after -ExpectedValue $ExpectedValue
        } catch {
            Add-ResultRow -Category $Category -Item $Item -Status 'Failed' -Details "$Description - ERROR: $($_.Exception.Message)" -RequiresReboot $RequiresReboot -RiskNote $RiskNote -DetectedValue $detected -ExpectedValue $ExpectedValue
        }
    } else {
        Add-ResultRow -Category $Category -Item $Item -Status 'Skipped (declined at confirm prompt)' -Details $Description -RequiresReboot $RequiresReboot -RiskNote $RiskNote -DetectedValue $detected -ExpectedValue $ExpectedValue
    }
}

function Get-RegValueString { param([string]$Path, [string]$Name)
    try { $ip = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop; return ("{0}={1}" -f $Name, [string]$ip.$Name) }
    catch { return "$Name=(not set)" }
}

$fatal = $null
try {
    try {
        $IsElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $IsElevated = $null }
    if ($Apply -and $IsElevated -ne $true) {
        throw "Refusing to -Apply inside guest: session is not confirmed elevated (Administrator). No changes made."
    }
    Write-GuestLog "Start. Build $ScriptBuild. $ComputerName ($IPAddresses). Apply=$Apply. Elevated=$IsElevated"

    # ===========================================================================
    # 1. VBS & Credential Protection
    # ===========================================================================
    if (-not $SkipCredentialGuard) {
        $vbsCaveat = "Requires a reboot AND requires the VM's virtual hardware to expose Secure Boot + hardware-assisted virtualization (a vSphere-side VM setting this script cannot see/change from inside the guest) - if that's not enabled at the VM level, this will still show 'Configured but NOT Running' after reboot."

        $p = @{
            Category = 'VBS & Credential Protection'; Item = 'Enable Virtualization Based Security'
            Description = 'Set HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard EnableVirtualizationBasedSecurity=1, RequirePlatformSecurityFeatures=1 (Secure Boot)'
            RequiresReboot = $true; RiskNote = $vbsCaveat
            ExpectedValue = 'EnableVirtualizationBasedSecurity=1'
            CurrentStateBlock = { Get-RegValueString 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' 'EnableVirtualizationBasedSecurity' }
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
            ExpectedValue = 'Enabled=1'
            CurrentStateBlock = { Get-RegValueString 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' 'Enabled' }
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
            ExpectedValue = 'LsaCfgFlags=1 or 2'
            CurrentStateBlock = { Get-RegValueString 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' 'LsaCfgFlags' }
            CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -ErrorAction SilentlyContinue; $v -and $v.LsaCfgFlags -in @(1, 2) }
            ApplyBlock = { Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -Value 1 -Type DWord -Force }
        }
        Invoke-Remediation @p

        $p = @{
            Category = 'VBS & Credential Protection'; Item = 'Enable Machine Identity Isolation'
            Description = 'Set HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard MachineIdentityIsolation=2 (Enabled - Enforcement Mode; machine password becomes IUM-bound only, isolated from LSASS). This is the same registry value the "Machine Identity Isolation Configuration" option under the "Turn On Virtualization Based Security" GPO writes.'
            RequiresReboot = $true
            RiskNote = "$vbsCaveat Also: this policy is confirmed in Microsoft's DeviceGuard Policy CSP reference but is not confirmed generally available on every Windows Server 2025 build - if unsupported on this build, the write is a harmless no-op (verify the 'Machine Identity Isolation Configuration' dropdown is present under gpedit.msc if in doubt)."
            ExpectedValue = 'MachineIdentityIsolation=2'
            CurrentStateBlock = { Get-RegValueString 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' 'MachineIdentityIsolation' }
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
                ExpectedValue = 'RealTimeProtectionEnabled=True'
                CurrentStateBlock = { try { 'RealTimeProtectionEnabled=' + (Get-MpComputerStatus -ErrorAction Stop).RealTimeProtectionEnabled } catch { 'unknown' } }
                CheckBlock = { $s = Get-MpComputerStatus -ErrorAction Stop; $s.RealTimeProtectionEnabled -eq $true }
                ApplyBlock = { Set-MpPreference -DisableRealtimeMonitoring $false -ErrorAction Stop }
            }
            Invoke-Remediation @p

            $p = @{
                Category = 'Microsoft Defender'; Item = 'Cloud-Delivered Protection (MAPS Reporting)'
                Description = "Set-MpPreference -MAPSReporting Advanced -SubmitSamplesConsent SendSafeSamples"
                ExpectedValue = 'MAPSReporting=2 (Advanced)'
                CurrentStateBlock = { try { 'MAPSReporting=' + (Get-MpPreference -ErrorAction Stop).MAPSReporting } catch { 'unknown' } }
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
                ExpectedValue = "16 rules = $AsrRuleAction"
                CurrentStateBlock = {
                    try {
                        $pref = Get-MpPreference -ErrorAction Stop
                        $ids = @($pref.AttackSurfaceReductionRules_Ids)
                        $actions = @($pref.AttackSurfaceReductionRules_Actions)
                        $desiredActionCode = if ($AsrRuleAction -eq 'Enabled') { 1 } else { 2 }
                        $set = 0
                        foreach ($ruleId in $AsrRules.Keys) {
                            $idx = [array]::IndexOf($ids, $ruleId)
                            if ($idx -ge 0 -and [int]$actions[$idx] -eq $desiredActionCode) { $set++ }
                        }
                        "$set/$($AsrRules.Count) rules at desired action; $($ids.Count) rule(s) configured in total"
                    } catch { 'unknown' }
                }
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
            ExpectedValue = 'UseLogonCredential=0'
            CurrentStateBlock = { Get-RegValueString 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential' }
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
                ExpectedValue = 'SupportedEncryptionTypes=0x18 (24)'
                CurrentStateBlock = { Get-RegValueString 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' 'SupportedEncryptionTypes' }
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
                ExpectedValue = 'SupportedEncryptionTypes has bits 0x18 set'
                CurrentStateBlock = { Get-RegValueString 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' 'SupportedEncryptionTypes' }
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
                ExpectedValue = 'Status=Stopped, StartMode=Disabled'
                CurrentStateBlock = { try { $s = Get-Service -Name WinRM -ErrorAction Stop; "Status=$($s.Status), StartMode=$((Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'").StartMode)" } catch { 'unknown' } }
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
                    ExpectedValue = 'AllowUnencrypted=0, auth_Basic=0'
                    CurrentStateBlock = { "$(Get-RegValueString $winrmSvcPath 'AllowUnencrypted'); $(Get-RegValueString $winrmSvcPath 'auth_Basic')" }
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
                    ExpectedValue = ($WinRmAllowedSourceRange -join ', ')
                    CurrentStateBlock = { try { 'RemoteAddress=' + ((@(Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' } | Get-NetFirewallAddressFilter -ErrorAction Stop | Select-Object -ExpandProperty RemoteAddress -Unique)) -join ',') } catch { 'unknown' } }
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
                ExpectedValue = 'AuditSmb1Access=True'
                CurrentStateBlock = { try { 'AuditSmb1Access=' + (Get-SmbServerConfiguration -ErrorAction Stop).AuditSmb1Access } catch { 'unknown' } }
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
            ExpectedValue = 'RpcAuthnLevelPrivacyEnabled=1'
            CurrentStateBlock = { Get-RegValueString 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' 'RpcAuthnLevelPrivacyEnabled' }
            CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' -Name RpcAuthnLevelPrivacyEnabled -ErrorAction SilentlyContinue; $v -and $v.RpcAuthnLevelPrivacyEnabled -eq 1 }
            ApplyBlock = { Set-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' -Name RpcAuthnLevelPrivacyEnabled -Value 1 -Type DWord -Force }
        }
        Invoke-Remediation @p

        $p = @{
            Category = 'SMB & RPC Security'; Item = 'Point and Print Restrictions'
            Description = 'Set HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint RestrictDriverInstallationToAdministrators=1, NoWarningNoElevationOnInstall=0'
            ExpectedValue = 'RestrictDriverInstallationToAdministrators=1, NoWarningNoElevationOnInstall=0'
            CurrentStateBlock = { "$(Get-RegValueString 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' 'RestrictDriverInstallationToAdministrators'); $(Get-RegValueString 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' 'NoWarningNoElevationOnInstall')" }
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
            ExpectedValue = 'EnableSmartScreen=1'
            CurrentStateBlock = { Get-RegValueString 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' 'EnableSmartScreen' }
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
            ExpectedValue = '1200=3, 1400=3, 2500=0'
            CurrentStateBlock = {
                $z = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3'
                "$(Get-RegValueString $z '1200'); $(Get-RegValueString $z '1400'); $(Get-RegValueString $z '2500')"
            }
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
                ExpectedValue = 'State=Disabled'
                CurrentStateBlock = { try { 'State=' + (Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction Stop).State } catch { 'not present' } }
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
                        ExpectedValue = 'Enabled=0, DisabledByDefault=1'
                        CurrentStateBlock = { "$(Get-RegValueString $sidePath 'Enabled'); $(Get-RegValueString $sidePath 'DisabledByDefault')" }
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
        if ($SkipBaselineGpoApply) {
            # Operator has already pushed the baseline GPO once and does not want it
            # re-asserted every run. The local script's CheckBlock for this item is
            # { $false } (it cannot verify hundreds of baseline settings), so without
            # this switch it re-applies + re-flags a reboot on EVERY run.
            Add-ResultRow -Category '2025 Security Baseline' -Item 'Apply full Windows Server 2025 Security Baseline GPO via LGPO.exe' -Status 'Manual/External Required' -Details 'Skipped by -SkipBaselineGpoApply. The baseline GPO backup was assumed already applied; this item is not re-pushed and does not flag a reboot. Re-run WITHOUT -SkipBaselineGpoApply to re-assert the full baseline (e.g. after updating the baseline backup).'
        } elseif ($lgpoReady) {
            $p = @{
                Category = '2025 Security Baseline'; Item = 'Apply full Windows Server 2025 Security Baseline GPO via LGPO.exe'
                Description = "& '$LgpoPath' /g '$BaselineGpoBackupPath'"
                RequiresReboot = $true
                RiskNote = 'Applies hundreds of settings at once (the actual baseline, not just the narrow items above). Review the baseline GPO backup contents before running against production if you have not already. NOTE: the check for this item is deliberately always-not-compliant, so it re-applies and re-flags a reboot on every run - pass -SkipBaselineGpoApply once you have applied it and rebooted.'
                ExpectedValue = 'LGPO.exe exit code 0'
                CheckBlock = { $false }
                ApplyBlock = {
                    $lgpoOutput = & $LgpoPath /g $BaselineGpoBackupPath 2>&1
                    if ($LASTEXITCODE -ne 0) { throw "LGPO.exe exited with code $LASTEXITCODE. Output: $lgpoOutput" }
                }
            }
            Invoke-Remediation @p
        } else {
            Add-ResultRow -Category '2025 Security Baseline' -Item 'Apply full Windows Server 2025 Security Baseline GPO' -Status 'Manual/External Required' -Details "Not staged in this guest (-LgpoPath/-BaselineGpoBackupPath not both present inside the guest). This script only remediates the narrow SmartScreen/IE-zone/legacy-IE items above, NOT the full official baseline. To apply the real baseline: stage LGPO.exe and an extracted Windows Server 2025 Security Baseline GPO backup in the guest, then re-run with -LgpoPath/-BaselineGpoBackupPath pointing at them, or run 'LGPO.exe /g <path>' in the guest, or link the baseline GPO in Active Directory for the OU containing these servers."
        }
    } else {
        Add-ResultRow -Category '2025 Security Baseline' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipBaselineIndicators was passed'
    }
} catch {
    $fatal = $_.Exception.Message
    Write-GuestLog "FATAL: $fatal" 'ERROR'
}

# ---------------------------------------------------------------------------------------------
# Emit a single Base64(JSON) envelope between markers. Base64 keeps the payload
# immune to any newline / quoting handling in Invoke-VMScript ScriptOutput.
# ---------------------------------------------------------------------------------------------
$meta = [PSCustomObject]@{
    Hostname       = $ComputerName
    IPAddress      = $IPAddresses
    OS             = $OSCaption
    Apply          = [bool]$Apply
    Elevated       = $IsElevated
    Build          = $ScriptBuild
    RebootRequired = @(($RebootRequiredItems | Select-Object -Unique))
    GuestLog       = $GuestLog.ToArray()
    FatalError     = $fatal
}
# .ToArray() rather than @(...) - @() around a generic List instance is unreliable
# on some PowerShell builds; .ToArray() is well-defined on 5.1 and 7.x alike.
$envelope = [PSCustomObject]@{ Schema = 'secgap-remediation-1'; Meta = $meta; Rows = $Results.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<SECGAP-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-SECGAP-ENVELOPE>>>'
'@

# =====================================================================================
# 3. Helper functions on the admin machine
# =====================================================================================

function ConvertTo-PsBool { param([bool]$Value) if ($Value) { '$true' } else { '$false' } }

function ConvertTo-PsSingleQuoted {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return "''" }
    if ($Value -match "[`r`n]") { throw "Refusing to inject a value containing a newline: '$Value'" }
    "'" + ($Value -replace "'", "''") + "'"
}

function New-RemediationPayload {
    param([bool]$EffectiveApply)

    if ($WinRmAllowedSourceRange -and $WinRmAllowedSourceRange.Count -gt 0) {
        $parts = foreach ($entry in $WinRmAllowedSourceRange) {
            $clean = ($entry -replace '[^0-9A-Fa-f\.:/]', '')
            if ($clean -ne $entry -or [string]::IsNullOrWhiteSpace($clean)) {
                throw "WinRmAllowedSourceRange entry '$entry' contains characters outside [0-9A-Fa-f.:/] - refused."
            }
            ConvertTo-PsSingleQuoted $clean
        }
        $rangesLiteral = '@(' + ($parts -join ',') + ')'
    } else {
        $rangesLiteral = '@()'
    }

    $inject = @"
`$Apply                  = $(ConvertTo-PsBool $EffectiveApply)
`$SkipCredentialGuard    = $(ConvertTo-PsBool $SkipCredentialGuard)
`$SkipDefenderASR        = $(ConvertTo-PsBool $SkipDefenderASR)
`$SkipAuthHardening      = $(ConvertTo-PsBool $SkipAuthHardening)
`$SkipWinRM              = $(ConvertTo-PsBool $SkipWinRM)
`$SkipSmbRpc             = $(ConvertTo-PsBool $SkipSmbRpc)
`$SkipBaselineIndicators = $(ConvertTo-PsBool $SkipBaselineIndicators)
`$SkipBaselineGpoApply   = $(ConvertTo-PsBool $SkipBaselineGpoApply)
`$DisableWinRM           = $(ConvertTo-PsBool $DisableWinRM)
`$WinRmAllowedSourceRange = $rangesLiteral
`$EnforceAesOnlyKerberos = $(ConvertTo-PsBool $EnforceAesOnlyKerberos)
`$DisableLegacyTls       = $(ConvertTo-PsBool $DisableLegacyTls)
`$AsrRuleAction          = $(ConvertTo-PsSingleQuoted $AsrRuleAction)
`$LgpoPath               = $(ConvertTo-PsSingleQuoted $LgpoPath)
`$BaselineGpoBackupPath  = $(ConvertTo-PsSingleQuoted $BaselineGpoBackupPath)
`$ConfirmPreference      = 'None'   # per-item interactive confirm cannot work through Guest Ops; the dry-run default + -Apply + high-risk switches remain the real gates
"@
    return $PayloadTemplate.Replace('#__INJECT_PARAMS__', $inject)
}

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
    try {
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64))
        return ($json | ConvertFrom-Json)
    } catch {
        return $null
    }
}

# Builds the { Server, VMs } inventory once per connected vCenter.
function Get-VMInventory {
    param([object[]]$Servers)
    $inv = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Servers) {
        try {
            $vms = @(Get-VM -Server $s -ErrorAction Stop)
        } catch {
            Write-Log "Get-VM failed on vCenter '$($s.Name)': $($_.Exception.Message)" 'WARN'
            $vms = @()
        }
        $inv.Add([pscustomobject]@{ Server = $s; VMs = $vms }) | Out-Null
    }
    return $inv
}

# Validation steps 1-5 (existence, non-ambiguity, power, Windows, Tools).
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
            Reason = "VM name '$Name' is ambiguous - $($hits.Count) matching VMs ($where). Skipped for safety; rename or target by MoRef."; Detected = "$($hits.Count) matches"; VM = $null; Server = $null }
    }

    $vm  = $hits[0].VM
    $srv = $hits[0].Server

    if ($vm.PowerState -ne 'PoweredOn') {
        return [pscustomobject]@{ Ok = $false; Stage = 'Power state'; ServerName = $srv.Name
            Reason = "VM is not powered on (PowerState=$($vm.PowerState)). Guest operations require a running guest."; Detected = "PowerState=$($vm.PowerState)"; VM = $null; Server = $null }
    }

    $g = $vm.ExtensionData.Guest
    $toolsOk = ($g.ToolsRunningStatus -eq 'guestToolsRunning') -or ($g.ToolsStatus -in 'toolsOk', 'toolsOld')
    if (-not $toolsOk) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VMware Tools'; ServerName = $srv.Name
            Reason = "VMware Tools not installed/running (ToolsStatus=$($g.ToolsStatus), ToolsRunningStatus=$($g.ToolsRunningStatus)). Guest operations unavailable."; Detected = "ToolsStatus=$($g.ToolsStatus)"; VM = $null; Server = $null }
    }

    $isWin = ($g.GuestFamily -eq 'windowsGuest') -or ($g.GuestId -match 'windows') -or ($vm.Guest.OSFullName -match 'Windows')
    if (-not $isWin) {
        return [pscustomobject]@{ Ok = $false; Stage = 'Guest OS'; ServerName = $srv.Name
            Reason = "Guest OS is not Windows (GuestFamily=$($g.GuestFamily), GuestId=$($g.GuestId), OS=$($vm.Guest.OSFullName)). This script only targets Windows Server."; Detected = "$($vm.Guest.OSFullName)"; VM = $null; Server = $null }
    }

    return [pscustomobject]@{ Ok = $true; Stage = 'OK'; ServerName = $srv.Name; Reason = ''; Detected = ''; VM = $vm; Server = $srv }
}

# Validation steps 6-7 (guest auth + PowerShell) via a tiny probe.
function Test-GuestPowerShell {
    param($VM, $Server, [pscredential]$Credential, [int]$ToolsWaitSecs)
    $probe = @'
try {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    $o = [pscustomobject]@{
        PS       = $PSVersionTable.PSVersion.ToString()
        User     = $id.Name
        Elevated = [bool]$pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        Host     = $env:COMPUTERNAME
    }
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
                "Guest authentication failed - VMware Tools rejected the supplied credential. Verify '$($Credential.UserName)' is a valid administrator on this guest. Underlying: $m"; break }
            'Tools are not running|GuestOperationsUnavailable|VMware Tools is not|not installed|VIX' {
                "VMware Tools guest operations unavailable: $m"; break }
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

# Delivers a multi-KB PowerShell payload into the guest and runs it. A single
# Invoke-VMScript call carrying the whole payload fails in some environments with the
# misleading "Could not locate Powershell script interpreter ... not enough
# permissions" error, because -ScriptType Powershell places the script on a command
# line that overruns a guest-operations length limit well below the payload size.
#
# Primary method: Copy-VMGuestFile - one guest file transfer, no length limit.
# Fallback (if Copy-VMGuestFile is unavailable or the vCenter account lacks the
# "Guest Operation Modifications" privilege): write the Base64 into a guest file in
# small pieces, each Invoke-VMScript call only ~ChunkSize+60 chars, then decode
# it in-guest.
#
# Both paths use only VMware guest operations - no WinRM, no PsExec, no SMB, no
# ESXi/guest network path. The staged file is the security-check logic plus the
# injected boolean switches only; it holds no credential and is always deleted.
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
    Write-Log "  [stage] payload is $($PayloadText.Length) chars; delivering to guest as a file (never as one Invoke-VMScript call)." 'INFO'

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
            Write-Log "  [stage] chunks written and decoded in guest." 'INFO'
        }

        # Run the staged payload in a child powershell.exe (execution-policy proof)
        # and return its stdout (the result envelope).
        Write-Log "  [stage] executing staged payload in guest." 'INFO'
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
    param(
        $vCenter, $VMName, $GuestHostname, $IPAddress, $OS,
        $Category, $SecurityCheck, $DetectedValue, $ExpectedValue,
        $Status, $Action, $RequiresReboot, $RiskNote, $Reason
    )
    [PSCustomObject]([ordered]@{
        vCenter        = $vCenter
        VMName         = $VMName
        GuestHostname  = $GuestHostname
        IPAddress      = $IPAddress
        OS             = $OS
        Category       = $Category
        SecurityCheck  = $SecurityCheck
        DetectedValue  = $DetectedValue
        ExpectedValue  = $ExpectedValue
        Status         = $Status
        Action         = $Action
        RequiresReboot = $RequiresReboot
        RiskNote       = $RiskNote
        'Error/Reason' = $Reason
        Timestamp      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

# Map the in-guest status vocabulary to the central status/action vocabulary.
function Get-CentralDisposition {
    param([string]$PayloadStatus)
    switch ($PayloadStatus) {
        'AlreadyCompliant'                     { return @{ Status = 'Compliant';                    Action = 'None (already compliant)' } }
        'DryRun - would apply'                 { return @{ Status = 'Non-Compliant';                Action = 'Would remediate (dry-run)' } }
        'Applied'                              { return @{ Status = 'Remediated';                   Action = 'Remediated in guest' } }
        'Failed'                               { return @{ Status = 'Remediation Failed';           Action = 'Remediation attempt failed' } }
        'Unable to Verify'                     { return @{ Status = 'Unable to Check';              Action = 'None' } }
        'Manual/External Required'             { return @{ Status = 'Manual Verification Required'; Action = 'Manual / external action required' } }
        'Not Applicable'                       { return @{ Status = 'Not Applicable';               Action = 'None' } }
        'Skipped (category)'                   { return @{ Status = 'Skipped';                      Action = 'Category skipped by parameter' } }
        'Skipped (declined at confirm prompt)' { return @{ Status = 'Skipped';                     Action = 'Declined at confirmation' } }
        default                               { return @{ Status = $PayloadStatus;                 Action = '' } }
    }
}

# Reproduces the local RepairWindowsServerSecurityGaps-Local.ps1 on-screen output
# EXACTLY: a Green "=== <host> (<ip>) ===" then one line per item -
#   "[<Status padded to 30>] <Category> > <Item>"
# with the local colour map (AlreadyCompliant/Applied=Green, DryRun*=Yellow,
# Manual*=DarkYellow, Failed=Red, everything else=Gray). $Rows are the RAW rows
# from the in-guest payload, so the Status text is the original vocabulary
# (AlreadyCompliant / Applied / DryRun - would apply / Failed / Manual/External
# Required / Not Applicable / Skipped ...).
function Show-VmResultsLocal {
    param([string]$HostLabel, $Rows)
    Write-Host ""
    Write-Host ("=== {0} ===" -f $HostLabel) -ForegroundColor Green
    foreach ($row in $Rows) {
        $status = [string]$row.Status
        $color = switch -Wildcard ($status) {
            'AlreadyCompliant' { 'Green' }
            'Applied'          { 'Green' }
            'DryRun*'          { 'Yellow' }
            'Manual*'          { 'DarkYellow' }
            'Failed'           { 'Red' }
            default            { 'Gray' }
        }
        Write-Host ("[{0}]" -f $status.PadRight(30)) -ForegroundColor $color -NoNewline
        if ($row.Category) {
            Write-Host (" {0} > {1}" -f $row.Category, $row.Item)
        } else {
            Write-Host (" {0}" -f $row.Item)
        }
    }
}

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows  = New-Object System.Collections.Generic.List[object]
$rebootGlobal = New-Object System.Collections.Generic.List[string]
$summary = [ordered]@{
    Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0
    Remediated = 0; RemediationFailed = 0; Manual = 0; RebootRequired = 0; SkippedFailed = 0
}

$inventory = Get-VMInventory -Servers $connectedServers

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Processing VM '$vmName' ===" 'INFO'

    try {
        $val = Resolve-AndValidateVM -Name $vmName -Inventory $inventory
        if (-not $val.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -SecurityCheck $val.Stage -DetectedValue $val.Detected -ExpectedValue '' `
                -Status 'Unable to Check' -Action 'Skipped - VM not processed' -RequiresReboot $false -RiskNote '' -Reason $val.Reason))
            Write-Log "SKIP '$vmName' [$($val.Stage)]: $($val.Reason)" 'WARN'
            Show-VmResultsLocal -HostLabel "$vmName" -Rows @([pscustomobject]@{ Status = 'Skipped'; Category = ''; Item = $val.Reason })
            continue
        }

        $vm  = $val.VM
        $srv = $val.Server

        $probe = Test-GuestPowerShell -VM $vm -Server $srv -Credential $GuestCredential -ToolsWaitSecs $ToolsWaitSecs
        if (-not $probe.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $vm.Guest.HostName -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest authentication / PowerShell' -DetectedValue '' -ExpectedValue 'PowerShell 5.1+ reachable via VMware Tools' `
                -Status 'Unable to Check' -Action 'Skipped - VM not processed' -RequiresReboot $false -RiskNote '' -Reason $probe.Reason))
            Write-Log "SKIP '$vmName' [probe]: $($probe.Reason)" 'WARN'
            Show-VmResultsLocal -HostLabel "$vmName" -Rows @([pscustomobject]@{ Status = 'Skipped'; Category = ''; Item = $probe.Reason })
            continue
        }

        if ($Apply -and -not $probe.Elevated) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest elevation' -DetectedValue "Elevated=$($probe.Elevated)" -ExpectedValue 'Elevated=True' `
                -Status 'Unable to Check' -Action 'Skipped - not elevated' -RequiresReboot $false -RiskNote '' `
                -Reason 'Guest session returned a non-elevated token; -Apply aborted for this VM to avoid a partial change. Confirm the guest account is a local administrator and that UAC remote-token filtering is not in effect.'))
            Write-Log "SKIP APPLY '$vmName': guest session not elevated." 'WARN'
            Show-VmResultsLocal -HostLabel "$vmName" -Rows @([pscustomobject]@{ Status = 'Skipped'; Category = ''; Item = 'Guest session not elevated - -Apply aborted for this VM.' })
            continue
        }

        # Wrapper-level ShouldProcess: -WhatIf here => dry-run this VM (still reported).
        $effectiveApply = [bool]$Apply
        if ($Apply -and -not $PSCmdlet.ShouldProcess("$vmName @ $($srv.Name)", "Apply Windows Server security remediation inside guest via VMware Tools")) {
            $effectiveApply = $false
            Write-Log "'$vmName': -WhatIf / declined at prompt -> running DRY-RUN for this VM only." 'INFO'
        }

        $payload = New-RemediationPayload -EffectiveApply $effectiveApply
        Write-Log "'$vmName': staging + invoking in-guest payload (Apply=$effectiveApply, user='$($GuestCredential.UserName)')." 'INFO'

        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $payload -ToolsWaitSecs $ToolsWaitSecs

        $envelope = Read-EnvelopeFromScriptOutput -Output $scriptOutput -StartMarker '<<<SECGAP-ENVELOPE-B64>>>' -EndMarker '<<<END-SECGAP-ENVELOPE>>>'
        if ($null -eq $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($scriptOutput) { ($scriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest payload result' -DetectedValue '' -ExpectedValue 'Base64 result envelope' `
                -Status 'Unable to Check' -Action 'Skipped - unparseable result' -RequiresReboot $false -RiskNote '' `
                -Reason "Guest payload returned no parseable result envelope. Output start: $snippet"))
            Write-Log "'$vmName': no parseable envelope returned from guest." 'ERROR'
            Show-VmResultsLocal -HostLabel "$vmName" -Rows @([pscustomobject]@{ Status = 'Skipped'; Category = ''; Item = "Guest payload returned no parseable result. Output start: $snippet" })
            continue
        }

        $meta  = $envelope.Meta
        $ghost = if ($meta.Hostname)  { $meta.Hostname }  else { $vm.Guest.HostName }
        $gip   = if ($meta.IPAddress) { $meta.IPAddress } else { ($vm.Guest.IPAddress -join ',') }
        $gos   = if ($meta.OS)        { $meta.OS }        else { $vm.Guest.OSFullName }

        if ($meta.GuestLog) {
            foreach ($gl in $meta.GuestLog) { Add-Content -LiteralPath $LogPath -Value ("    [guest:$vmName] $gl") }
        }
        if ($meta.FatalError) { Write-Log "'$vmName': guest payload reported a fatal error: $($meta.FatalError)" 'ERROR' }

        $summary.Checked++
        $f = @{ NonCompliant = $false; Remediated = $false; Failed = $false; Manual = $false; Reboot = $false; Any = $false }

        foreach ($r in @($envelope.Rows)) {
            $f.Any = $true
            $disp = Get-CentralDisposition -PayloadStatus $r.Status
            switch ($disp.Status) {
                'Non-Compliant'                { $f.NonCompliant = $true }
                'Remediated'                   { $f.Remediated = $true }
                'Remediation Failed'           { $f.Failed = $true }
                'Manual Verification Required' { $f.Manual = $true }
            }

            $reason = ''
            if ($disp.Status -in 'Unable to Check', 'Remediation Failed', 'Manual Verification Required', 'Skipped') { $reason = $r.Details }

            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category $r.Category -SecurityCheck $r.Item -DetectedValue $r.DetectedValue -ExpectedValue $r.ExpectedValue `
                -Status $disp.Status -Action $disp.Action -RequiresReboot ([bool]$r.RequiresReboot) -RiskNote $r.RiskNote -Reason $reason))
        }

        # On-screen result for this VM, formatted EXACTLY like the local
        # RepairWindowsServerSecurityGaps-Local.ps1 (raw statuses, in addition to CSV).
        Show-VmResultsLocal -HostLabel ("{0} ({1})" -f $ghost, $gip) -Rows @($envelope.Rows)

        # REBOOT list: exactly the local rule - an item is listed only when it was
        # actually Applied AND needs a reboot. The payload already computes this.
        $vmReboots = @($envelope.Meta.RebootRequired | Where-Object { $_ })
        if ($vmReboots.Count -gt 0) {
            $f.Reboot = $true
            foreach ($item in $vmReboots) {
                if ($vmNames.Count -gt 1) { $rebootGlobal.Add(("{0} - {1}" -f $vmName, $item)) | Out-Null }
                else                      { $rebootGlobal.Add([string]$item) | Out-Null }
            }
        }

        if ($f.Failed)       { $summary.RemediationFailed++ }
        if ($f.Remediated)   { $summary.Remediated++ }
        if ($f.NonCompliant) { $summary.NonCompliant++ }
        if ($f.Manual)       { $summary.Manual++ }
        if ($f.Reboot)       { $summary.RebootRequired++ }
        if ($f.Any -and -not $f.NonCompliant -and -not $f.Failed) { $summary.Compliant++ }

        Write-Log ("'$vmName': done. rows={0} nonCompliant={1} remediated={2} failed={3} manual={4} reboot={5}" -f `
            @($envelope.Rows).Count, $f.NonCompliant, $f.Remediated, $f.Failed, $f.Manual, $f.Reboot) 'INFO'
    }
    catch {
        $summary.SkippedFailed++
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        try {
            $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Validation' -SecurityCheck 'Processing' -DetectedValue '' -ExpectedValue '' `
                -Status 'Unable to Check' -Action 'Skipped - exception' -RequiresReboot $false -RiskNote '' -Reason $_.Exception.Message))
            Show-VmResultsLocal -HostLabel "$vmName" -Rows @([pscustomobject]@{ Status = 'Failed'; Category = ''; Item = $_.Exception.Message })
        } catch { }
        continue
    }
}

# =====================================================================================
# 5. Reboot block + central report (formatted like the local script)
# =====================================================================================
if ($rebootGlobal.Count -gt 0) {
    Write-Host ""
    Write-Host "REBOOT REQUIRED to complete these applied changes:" -ForegroundColor Red
    foreach ($item in ($rebootGlobal | Select-Object -Unique)) { Write-Host "  - $item" -ForegroundColor Red }
    Write-Host "This script does NOT reboot the server itself - schedule these during a maintenance window, then re-run the post-remediation check to confirm they took effect." -ForegroundColor Red
}

$centralRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV exported: $CsvPath"

$failedCount = @($centralRows | Where-Object { $_.Status -eq 'Remediation Failed' }).Count
if ($failedCount -gt 0) { Write-Log "$failedCount item row(s) FAILED to apply - see CSV/log for details." 'WARN' }
Write-Log "Run complete. $($centralRows.Count) total item rows. Apply=$Apply."

Write-Host ""
Write-Host "Results: $CsvPath"
Write-Host "Log:     $LogPath"

# Multi-VM roll-up (only shown when more than one VM was targeted, so a single-VM
# run looks identical to the local script).
if ($summary.Total -gt 1) {
    Write-Host ""
    Write-Host "================ SUMMARY (all VMs) ================" -ForegroundColor Green
    Write-Host ("  Total VMs                     : {0}" -f $summary.Total)
    Write-Host ("  Successfully Checked          : {0}" -f $summary.Checked)
    Write-Host ("  Compliant (no gaps found)     : {0}" -f $summary.Compliant)
    Write-Host ("  Non-Compliant (gaps found)    : {0}" -f $summary.NonCompliant) -ForegroundColor $(if ($summary.NonCompliant) { 'Yellow' } else { 'Gray' })
    Write-Host ("  Remediated                    : {0}" -f $summary.Remediated)   -ForegroundColor $(if ($summary.Remediated) { 'Green' } else { 'Gray' })
    Write-Host ("  Remediation Failed            : {0}" -f $summary.RemediationFailed) -ForegroundColor $(if ($summary.RemediationFailed) { 'Red' } else { 'Gray' })
    Write-Host ("  Manual Verification Required  : {0}" -f $summary.Manual) -ForegroundColor DarkYellow
    Write-Host ("  Reboot Required               : {0}" -f $summary.RebootRequired) -ForegroundColor $(if ($summary.RebootRequired) { 'Red' } else { 'Gray' })
    Write-Host ("  Skipped / Failed to process   : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
    Write-Host "=================================================" -ForegroundColor Green
}
