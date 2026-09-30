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
    Optional list of IPs / CIDR / ranges, or Windows Firewall scope keywords
    (LocalSubnet, Any, DNS, ...), to scope the "Windows Remote Management" firewall
    rules to. Without it (and without -RestrictWinRmToLocalSubnet) that step is
    reported "Manual/External Required".

.PARAMETER RestrictWinRmToLocalSubnet
    Opt-in shortcut when the email just says "restrict the listener filters" and no
    specific management subnet is defined: scopes the "Windows Remote Management"
    firewall rules' RemoteAddress to LocalSubnet. Narrower than the default "Any",
    and cannot lock out a same-subnet administrator. Ignored if
    -WinRmAllowedSourceRange is also supplied (that wins).

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
    [switch]  $SkipExtendedBaseline,       # skip the whole "7. Extended Baseline (Cybersecurity Review)" category
    [switch]  $SkipExtendedManualItems,    # in category 7, do not emit the secedit / patch "Manual/External Required" rows
    # --- ISO / Qualys Policy Compliance (iso.xlsx) coverage - category 8 -----------------
    [switch]  $SkipIsoPolicy,                # skip the whole category 8 (Qualys PC control-ID coverage)
    [switch]  $SkipIeHardening,              # within category 8, skip the ~89 Internet Explorer zone/feature registry values
    [switch]  $EnableFirewallProfiles,       # opt-in: turn the Windows Firewall Domain/Private/Public profiles ON (CID 3950-3952)
    [switch]  $DisableNetbios,               # opt-in: set DNSClient EnableNetbios=0 (CID 25358)
    [switch]  $SetSmbMinSmb3,                # opt-in: mandate minimum SMB dialect 3.0.0 server+client (CID 29576/29583)
    [switch]  $RestrictLocalAcctNetworkLogon,# opt-in: LocalAccountTokenFilterPolicy=0 - PtH mitigation (CID 9024)
    [switch]  $EnableUacHardening,           # opt-in: FilterAdministratorToken=1, ConsentPromptBehaviorAdmin=2, EnableVirtualization=1 (CID 2586/2587/3940)
    [switch]  $ApplyAccountPolicy,           # opt-in: set Minimum Password Length / Account Lockout Threshold via secedit (CID 1071/2342)
    [switch]  $EnableServerAsrRules,         # opt-in: enable ASR rule a8f5898e (server webshell) in Block (CID 30456)
    [ValidateRange(8,256)]
    [int]     $MinPasswordLength      = 14,  # target for CID 1071 when -ApplyAccountPolicy is set
    [ValidateRange(1,20)]
    [int]     $AccountLockoutThreshold = 3,  # target for CID 2342 when -ApplyAccountPolicy is set

    [switch]  $DisableWinRM,
    [string[]]$WinRmAllowedSourceRange,
    [switch]  $RestrictWinRmToLocalSubnet,
    [switch]  $EnforceAesOnlyKerberos,
    [switch]  $TreatAllAsWorkgroup,
    [switch]  $DisableLegacyTls,

    [ValidateRange(0, 50)]
    [int]     $CachedLogonsCount = 4,      # category 7: target for Winlogon\CachedLogonsCount (use 0 on a workgroup to clear Qualys QID 90007)
    [string[]]$WinRmFilterRange = @('10.50.10.1-10.50.10.254'),  # category 7: IPv4 range(s) for the WinRM listener IPv4Filter/IPv6Filter. Default = the management subnet; pass -WinRmFilterRange @() to leave it as Manual/External Required.

    [ValidatePattern('^$|^[A-Za-z0-9._-]{1,20}$')]
    [string]  $GuestAccountNewName = 'LocalGuest_Disabled',  # category 7: the built-in Guest account (SID -501) is renamed to this and disabled (QID 105228). Pass -GuestAccountNewName '' to skip and report only.

    [ValidateSet('Enabled', 'AuditMode')]
    [string]  $AsrRuleAction = 'Enabled',

    [string]  $LgpoPath              = 'C:\Tools\LGPO\LGPO.exe',
    [string]  $BaselineGpoBackupPath = 'C:\Tools\Baseline\GPOs',

    [ValidateRange(30, 3600)]
    [int]     $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'
$ScriptBuild = '2026-09-04d-iso-pc-r1'

# =====================================================================================
# 0. Banner
# =====================================================================================
Write-Host ("Invoke-RemoteSecurityGapRemediation.ps1  [build $ScriptBuild]") -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
Write-Host ("If the build above is not '2026-09-04d-iso-pc-r1' you are running an OLD copy - update it.") -ForegroundColor DarkGray
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
$ScriptBuild = '2026-08-17-02-winrm-listener-filter / remote-payload-r1 / iso-pc-r1'

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

    # --- Workgroup / non-domain detection ---------------------------------------
    # Domain-only controls (Machine Identity Isolation, Kerberos encryption types)
    # are reported "Not Applicable" on workgroup servers. Two independent signals:
    # PartOfDomain=$false OR DomainRole in {0,2} (Standalone Workstation/Server).
    # -TreatAllAsWorkgroup forces it on without relying on the query.
    $IsWorkgroup  = $false
    $WorkgroupWhy = ''
    if ($TreatAllAsWorkgroup) {
        $IsWorkgroup  = $true
        $WorkgroupWhy = 'operator passed -TreatAllAsWorkgroup'
    } else {
        try {
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
            $pod  = [bool]$cs.PartOfDomain
            $role = [int]$cs.DomainRole
            if ((-not $pod) -or ($role -in 0, 2)) { $IsWorkgroup = $true; $WorkgroupWhy = "PartOfDomain=$pod; DomainRole=$role" }
        } catch {
            $IsWorkgroup = $false
            $WorkgroupWhy = "domain membership query failed: $($_.Exception.Message)"
        }
    }
    Write-GuestLog "Workgroup=$IsWorkgroup ($WorkgroupWhy)"

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

        # Machine Identity Isolation only has meaning on a domain-joined machine.
        # It binds the Active Directory computer-account credential (the domain
        # secure-channel / Netlogon secret) into the VBS/IUM enclave so it is
        # isolated from LSASS. A workgroup (non-domain-joined) server has no AD
        # machine identity to isolate, so the policy is a no-op: the registry value
        # does not persist and the check can never report compliant. On such a host
        # we report it Not Applicable instead of endlessly "would apply".
        if ($IsWorkgroup) {
            Add-ResultRow -Category 'VBS & Credential Protection' -Item 'Enable Machine Identity Isolation' -Status 'Not Applicable' `
                -Details "NOT APPLICABLE - this server is in a workgroup (not domain-joined). Machine Identity Isolation protects the Active Directory computer-account credential (domain secure-channel / Netlogon secret); a workgroup server has none, so the DeviceGuard\MachineIdentityIsolation policy is a no-op that does not persist. Applicable only after the server is domain-joined. [$WorkgroupWhy]" `
                -DetectedValue "Workgroup ($WorkgroupWhy)" -ExpectedValue 'N/A unless domain-joined'
        } else {
            $p = @{
                Category = 'VBS & Credential Protection'; Item = 'Enable Machine Identity Isolation'
                Description = 'Set HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard MachineIdentityIsolation=2 (Enabled - Enforcement Mode; machine password becomes IUM-bound only, isolated from LSASS). This is the same registry value the "Machine Identity Isolation Configuration" option under the "Turn On Virtualization Based Security" GPO writes.'
                RequiresReboot = $true
                RiskNote = "$vbsCaveat Also: this policy is confirmed in Microsoft's DeviceGuard Policy CSP reference but is not confirmed generally available on every Windows Server 2025 build - if unsupported on this build, the write is a harmless no-op (verify the 'Machine Identity Isolation Configuration' dropdown is present under gpedit.msc if in doubt). Only meaningful on a domain-joined machine."
                ExpectedValue = 'MachineIdentityIsolation=2'
                CurrentStateBlock = { Get-RegValueString 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' 'MachineIdentityIsolation' }
                CheckBlock = { $v = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' -Name MachineIdentityIsolation -ErrorAction SilentlyContinue; $v -and $v.MachineIdentityIsolation -eq 2 }
                ApplyBlock = {
                    New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' -Force -ErrorAction SilentlyContinue | Out-Null
                    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' -Name MachineIdentityIsolation -Value 2 -Type DWord -Force
                }
            }
            Invoke-Remediation @p
        }
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

        if ($IsWorkgroup) {
            Add-ResultRow -Category 'Authentication Security' -Item 'Kerberos Encryption Types - Ensure AES Supported' -Status 'Not Applicable' `
                -Details "NOT APPLICABLE - this server is in a workgroup. There is no Active Directory domain or KDC; authentication is NTLM / local accounts only, so Kerberos is not used. SupportedEncryptionTypes governs domain Kerberos only. [$WorkgroupWhy]" `
                -DetectedValue "Workgroup ($WorkgroupWhy)" -ExpectedValue 'N/A unless domain-joined'
        } elseif ($EnforceAesOnlyKerberos) {
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
                Add-ResultRow -Category 'Remote Management' -Item 'Restrict WinRM Listener to Management Sources' -Status 'Manual/External Required' -Details "Not restricted - left open to any source. Re-run with -RestrictWinRmToLocalSubnet to scope the 'Windows Remote Management' firewall rules to LocalSubnet (satisfies 'restrict the listener filters' and cannot lock out a same-subnet admin), or -WinRmAllowedSourceRange @('10.1.1.0/24') for specific management subnet(s)/host IPs. Not changed automatically because a wrong guess here could cut off legitimate admin access."
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

        if (-not (Get-Command Get-WindowsOptionalFeature -ErrorAction SilentlyContinue)) {
            Add-ResultRow -Category '2025 Security Baseline' -Item 'Remove Legacy Internet Explorer 11 Feature' -Status 'Not Applicable' -Details 'Get-WindowsOptionalFeature not available (e.g. Server Core or restricted session).'
        } elseif ($null -eq (Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction SilentlyContinue)) {
            Add-ResultRow -Category '2025 Security Baseline' -Item 'Remove Legacy Internet Explorer 11 Feature' -Status 'Not Applicable' `
                -Details 'NOT APPLICABLE - the Internet Explorer optional feature does not exist on this Windows Server build; there is nothing to remove.' `
                -DetectedValue 'Feature not present on this build' -ExpectedValue 'N/A - feature absent'
        } else {
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

    # ===========================================================================
    # 7. Extended Baseline (Cybersecurity Review)
    #    Qualys / CS-review findings layered on top of the original 6 areas.
    #    Same dry-run / -Apply / reboot-flag model and result format.
    # ===========================================================================
    if (-not $SkipExtendedBaseline) {
        $CAT = '7. Extended Baseline (Cybersecurity Review)'

        # 7.1 HVCI - Enabled WITH UEFI lock
        $hvciPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'
        Invoke-Remediation -Category $CAT -Item 'HVCI - UEFI lock' `
            -Description "Set $hvciPath Locked=1 (HVCI 'Enabled with UEFI lock'; keeps Enabled=1)" `
            -RequiresReboot $true `
            -RiskNote 'UEFI lock makes HVCI hard to disable later without firmware/physical access (intentional). Only effective once HVCI itself (Enabled=1) is on - see category 1.' `
            -ExpectedValue 'Locked=1 (and Enabled=1)' `
            -CurrentStateBlock { "$(Get-RegValueString $hvciPath 'Enabled'); $(Get-RegValueString $hvciPath 'Locked')" } `
            -CheckBlock { $v = Get-ItemProperty -Path $hvciPath -ErrorAction SilentlyContinue; $v -and $v.Enabled -eq 1 -and $v.Locked -eq 1 } `
            -ApplyBlock {
                New-Item -Path $hvciPath -Force -ErrorAction SilentlyContinue | Out-Null
                Set-ItemProperty -Path $hvciPath -Name Enabled -Value 1 -Type DWord -Force
                Set-ItemProperty -Path $hvciPath -Name Locked  -Value 1 -Type DWord -Force
            }

        # 7.2 Kernel-mode Hardware-enforced Stack Protection - ENFORCEMENT (not audit)
        $kssPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\KernelShadowStacks'
        Invoke-Remediation -Category $CAT -Item 'Kernel-mode HW-enforced Stack Protection - enforcement' `
            -Description "Set $kssPath Enabled=1, AuditMode=0 (enforcement mode)" `
            -RequiresReboot $true `
            -RiskNote 'Enforcement can block old / incompatible kernel drivers that violate shadow-stack rules. Pilot in AuditMode (AuditMode=1) first if driver compatibility is unknown.' `
            -ExpectedValue 'Enabled=1, AuditMode=0' `
            -CurrentStateBlock { "$(Get-RegValueString $kssPath 'Enabled'); $(Get-RegValueString $kssPath 'AuditMode')" } `
            -CheckBlock { $v = Get-ItemProperty -Path $kssPath -ErrorAction SilentlyContinue; $v -and $v.Enabled -eq 1 -and $v.AuditMode -eq 0 } `
            -ApplyBlock {
                New-Item -Path $kssPath -Force -ErrorAction SilentlyContinue | Out-Null
                Set-ItemProperty -Path $kssPath -Name Enabled   -Value 1 -Type DWord -Force
                Set-ItemProperty -Path $kssPath -Name AuditMode  -Value 0 -Type DWord -Force
            }

        # 7.3 LSASS Protected Process (RunAsPPL) - WITH UEFI lock
        $lsaPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA'
        Invoke-Remediation -Category $CAT -Item 'LSASS RunAsPPL - UEFI lock' `
            -Description "Set $lsaPath RunAsPPL=1 (LSA runs as a protected process, UEFI-locked). 2 = enabled without lock." `
            -RequiresReboot $true `
            -RiskNote 'Blocks unsigned / non-Microsoft LSA plugins and SSPs (some SSO / smart-card / EDR agents). Confirm all LSA-integrating agents are WHQL-signed first. UEFI lock is hard to revert.' `
            -ExpectedValue 'RunAsPPL=1' `
            -CurrentStateBlock { Get-RegValueString $lsaPath 'RunAsPPL' } `
            -CheckBlock { $v = Get-ItemProperty -Path $lsaPath -Name RunAsPPL -ErrorAction SilentlyContinue; $v -and $v.RunAsPPL -eq 1 } `
            -ApplyBlock {
                Set-ItemProperty -Path $lsaPath -Name RunAsPPL      -Value 1 -Type DWord -Force
                Set-ItemProperty -Path $lsaPath -Name RunAsPPLBoot  -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
            }

        # 7.4 PowerShell Script Block Logging (+ invocation logging)
        $sblPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'
        Invoke-Remediation -Category $CAT -Item 'PowerShell Script Block (and Invocation) Logging' `
            -Description "Set $sblPath EnableScriptBlockLogging=1, EnableScriptBlockInvocationLogging=1" `
            -RiskNote 'Invocation logging is verbose (start/stop of every script block, event 4104) - size the PowerShell/Operational log and forwarding accordingly. CS review explicitly requires it enabled.' `
            -ExpectedValue 'EnableScriptBlockLogging=1, EnableScriptBlockInvocationLogging=1' `
            -CurrentStateBlock { "$(Get-RegValueString $sblPath 'EnableScriptBlockLogging'); $(Get-RegValueString $sblPath 'EnableScriptBlockInvocationLogging')" } `
            -CheckBlock { $v = Get-ItemProperty -Path $sblPath -ErrorAction SilentlyContinue; $v -and $v.EnableScriptBlockLogging -eq 1 -and $v.EnableScriptBlockInvocationLogging -eq 1 } `
            -ApplyBlock {
                New-Item -Path $sblPath -Force -ErrorAction SilentlyContinue | Out-Null
                Set-ItemProperty -Path $sblPath -Name EnableScriptBlockLogging           -Value 1 -Type DWord -Force
                Set-ItemProperty -Path $sblPath -Name EnableScriptBlockInvocationLogging -Value 1 -Type DWord -Force
            }

        # 7.5 WinRM "Allow remote server management through WinRM" / auto config = Disabled
        $winrmPol = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WinRM\Service'
        Invoke-Remediation -Category $CAT -Item 'WinRM auto-configuration disabled' `
            -Description "Set $winrmPol AllowAutoConfig=0 ('Allow remote server management through WinRM' = Disabled)" `
            -RiskNote 'Stops policy-driven creation of the WinRM listener. If any tooling manages this host over WinRM, confirm it uses another path first.' `
            -ExpectedValue 'AllowAutoConfig=0 (or not configured)' `
            -CurrentStateBlock { Get-RegValueString $winrmPol 'AllowAutoConfig' } `
            -CheckBlock { $v = Get-ItemProperty -Path $winrmPol -Name AllowAutoConfig -ErrorAction SilentlyContinue; (-not $v) -or $v.AllowAutoConfig -eq 0 } `
            -ApplyBlock {
                New-Item -Path $winrmPol -Force -ErrorAction SilentlyContinue | Out-Null
                Set-ItemProperty -Path $winrmPol -Name AllowAutoConfig -Value 0 -Type DWord -Force
            }

        # 7.6 WinRM IPv4Filter / IPv6Filter  (only if -WinRmFilterRange was supplied)
        if ($WinRmFilterRange -and $WinRmFilterRange.Count -gt 0) {
            $filterValue = ($WinRmFilterRange -join ',')
            Invoke-Remediation -Category $CAT -Item 'WinRM listener IPv4/IPv6 filter' `
                -Description "Set $winrmPol IPv4Filter/IPv6Filter = '$filterValue' (currently '*')" `
                -RiskNote 'Wrong range can cut off legitimate WinRM management. Value applies to the listener the next time it is (re)created.' `
                -ExpectedValue "IPv4Filter/IPv6Filter = $filterValue" `
                -CurrentStateBlock { "$(Get-RegValueString $winrmPol 'IPv4Filter'); $(Get-RegValueString $winrmPol 'IPv6Filter')" } `
                -CheckBlock { $v = Get-ItemProperty -Path $winrmPol -ErrorAction SilentlyContinue; $v -and $v.IPv4Filter -eq $filterValue -and $v.IPv6Filter -eq $filterValue } `
                -ApplyBlock {
                    New-Item -Path $winrmPol -Force -ErrorAction SilentlyContinue | Out-Null
                    Set-ItemProperty -Path $winrmPol -Name IPv4Filter -Value $filterValue -Type String -Force
                    Set-ItemProperty -Path $winrmPol -Name IPv6Filter -Value $filterValue -Type String -Force
                }
        } else {
            Add-ResultRow -Category $CAT -Item 'WinRM listener IPv4/IPv6 filter' -Status 'Manual/External Required' `
                -Details "Not changed - no -WinRmFilterRange supplied. To scope the WinRM listener filter, set $winrmPol IPv4Filter and IPv6Filter to your management range, e.g. '10.50.10.1-10.50.10.254' (currently '*'). Re-run with -WinRmFilterRange '<range>' to apply." `
                -DetectedValue "$(Get-RegValueString $winrmPol 'IPv4Filter'); $(Get-RegValueString $winrmPol 'IPv6Filter')" -ExpectedValue 'Scoped to management range'
        }

        # 7.7 Kerberos PKINIT hash algorithms SHA256 / SHA384 / SHA512 = Supported (3)
        $pkBase = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PKINITHashAlgorithms'
        foreach ($alg in 'SHA256', 'SHA384', 'SHA512') {
            $algPath = Join-Path $pkBase $alg
            Invoke-Remediation -Category $CAT -Item "PKINIT hash $alg = Supported" `
                -Description "Set $algPath Support=3 (Supported). 1 = Default." `
                -RiskNote 'Only functionally used with Active Directory / AD CS smart-card (PKINIT) logon. On a workgroup server this is a benchmark-value fix only. Verify via the GPO "Configure hash types allowed for Kerberos PKINIT".' `
                -ExpectedValue 'Support=3' `
                -CurrentStateBlock { Get-RegValueString $algPath 'Support' }.GetNewClosure() `
                -CheckBlock { $v = Get-ItemProperty -Path $algPath -Name Support -ErrorAction SilentlyContinue; $v -and $v.Support -eq 3 }.GetNewClosure() `
                -ApplyBlock {
                    New-Item -Path $algPath -Force -ErrorAction SilentlyContinue | Out-Null
                    Set-ItemProperty -Path $algPath -Name Support -Value 3 -Type DWord -Force
                }.GetNewClosure()
        }

        # 7.8 Windows Ink Workspace - Disabled
        $inkPath = 'HKLM:\SOFTWARE\Policies\Microsoft\WindowsInkWorkspace'
        Invoke-Remediation -Category $CAT -Item 'Windows Ink Workspace disabled' `
            -Description "Set $inkPath AllowWindowsInkWorkspace=0 (Disabled). 1 = on but no access above lock." `
            -ExpectedValue 'AllowWindowsInkWorkspace=0' `
            -CurrentStateBlock { Get-RegValueString $inkPath 'AllowWindowsInkWorkspace' } `
            -CheckBlock { $v = Get-ItemProperty -Path $inkPath -Name AllowWindowsInkWorkspace -ErrorAction SilentlyContinue; $v -and $v.AllowWindowsInkWorkspace -eq 0 } `
            -ApplyBlock {
                New-Item -Path $inkPath -Force -ErrorAction SilentlyContinue | Out-Null
                Set-ItemProperty -Path $inkPath -Name AllowWindowsInkWorkspace -Value 0 -Type DWord -Force
            }

        # 7.9 Advanced audit: "Audit Policy Change" = Success and Failure
        Invoke-Remediation -Category $CAT -Item 'Audit "Audit Policy Change" = Success and Failure' `
            -Description 'auditpol /set /subcategory:"Audit Policy Change" /success:enable /failure:enable' `
            -ExpectedValue 'Success and Failure' `
            -CurrentStateBlock {
                try {
                    $csv = (& auditpol /get /subcategory:"Audit Policy Change" /r 2>$null | ConvertFrom-Csv)
                    $row = $csv | Where-Object { $_.Subcategory -eq 'Audit Policy Change' } | Select-Object -First 1
                    "Inclusion Setting=$($row.'Inclusion Setting')"
                } catch { 'unknown' }
            } `
            -CheckBlock {
                $csv = (& auditpol /get /subcategory:"Audit Policy Change" /r 2>$null | ConvertFrom-Csv)
                $row = $csv | Where-Object { $_.Subcategory -eq 'Audit Policy Change' } | Select-Object -First 1
                $row -and $row.'Inclusion Setting' -eq 'Success and Failure'
            } `
            -ApplyBlock {
                $out = & auditpol /set /subcategory:"Audit Policy Change" /success:enable /failure:enable 2>&1
                if ($LASTEXITCODE -ne 0) { throw "auditpol exited $LASTEXITCODE. Output: $out" }
            }

        # 7.10 Cached logons count
        $wlPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
        Invoke-Remediation -Category $CAT -Item "Cached logons count = $CachedLogonsCount" `
            -Description "Set $wlPath CachedLogonsCount='$CachedLogonsCount' (REG_SZ). Qualys QID 90007; 0 clears it on a workgroup server." `
            -RiskNote 'A value of 0 means no cached domain logons - only matters if this host is ever domain-joined and loses DC connectivity. Fine for workgroup.' `
            -ExpectedValue "CachedLogonsCount=$CachedLogonsCount" `
            -CurrentStateBlock { Get-RegValueString $wlPath 'CachedLogonsCount' } `
            -CheckBlock { $v = Get-ItemProperty -Path $wlPath -Name CachedLogonsCount -ErrorAction SilentlyContinue; $v -and ([string]$v.CachedLogonsCount) -eq ([string]$CachedLogonsCount) } `
            -ApplyBlock { Set-ItemProperty -Path $wlPath -Name CachedLogonsCount -Value ([string]$CachedLogonsCount) -Type String -Force }

        # 7.11 Credential Guard policy value - detect a GPO that disables it
        $dgPol = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'
        $dgVal = (Get-ItemProperty -Path $dgPol -Name LsaCfgFlags -ErrorAction SilentlyContinue).LsaCfgFlags
        $lsaVal = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -ErrorAction SilentlyContinue).LsaCfgFlags
        if ($dgVal -eq 0) {
            Add-ResultRow -Category $CAT -Item 'Credential Guard - policy override' -Status 'Manual/External Required' `
                -Details "A policy is DISABLING Credential Guard: $dgPol\LsaCfgFlags = 0. Change the GPO 'Computer Configuration > Administrative Templates > System > Device Guard > Turn On Virtualization Based Security > Credential Guard Configuration' to 'Enabled with UEFI lock' (or remove that setting), then gpupdate /force and reboot. Workgroup membership does not justify disabling Credential Guard." `
                -DetectedValue "Policies\...\DeviceGuard\LsaCfgFlags=0; Control\LSA\LsaCfgFlags=$([string]$lsaVal)" -ExpectedValue 'LsaCfgFlags=1 (Enabled with UEFI lock)' -RequiresReboot $true
        } else {
            Invoke-Remediation -Category $CAT -Item 'Credential Guard - policy value' `
                -Description "Set $dgPol LsaCfgFlags=1 (Enabled with UEFI lock) so it survives Group Policy refresh" `
                -RequiresReboot $true `
                -RiskNote 'Only meaningful once VBS is available (Secure Boot + virtualization exposed to the VM). UEFI lock is hard to revert.' `
                -ExpectedValue 'LsaCfgFlags=1' `
                -CurrentStateBlock { "Policies=$([string]((Get-ItemProperty -Path $dgPol -Name LsaCfgFlags -ErrorAction SilentlyContinue).LsaCfgFlags)); Control\LSA=$([string]((Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -ErrorAction SilentlyContinue).LsaCfgFlags))" } `
                -CheckBlock { $v = Get-ItemProperty -Path $dgPol -Name LsaCfgFlags -ErrorAction SilentlyContinue; $v -and $v.LsaCfgFlags -eq 1 } `
                -ApplyBlock {
                    New-Item -Path $dgPol -Force -ErrorAction SilentlyContinue | Out-Null
                    Set-ItemProperty -Path $dgPol -Name LsaCfgFlags -Value 1 -Type DWord -Force
                }
        }

        # 7.12 Deny access to this computer from the network (user right) - real remediation via secedit.
        #      Guests (S-1-5-32-546) + Local account (S-1-5-113) + Local account & member of Administrators (S-1-5-114).
        #      This is a NETWORK-logon deny only - it does not touch console or RDP logon.
        Invoke-Remediation -Category $CAT -Item 'Deny access to this computer from the network (user right)' `
            -Description 'secedit: [Privilege Rights] SeDenyNetworkLogonRight = *S-1-5-32-546,*S-1-5-113,*S-1-5-114' `
            -RiskNote 'Denies network logon (SMB / RPC / WinRM) to Guests and ALL local accounts including local admins - the intended anti-lateral-movement control. Does NOT affect console or RDP logon. If a baseline GPO also defines this right, set it there too so a Group Policy refresh does not revert it. Takes effect for new logons.' `
            -ExpectedValue 'Contains S-1-5-32-546, S-1-5-113 and S-1-5-114' `
            -CurrentStateBlock {
                $cfg = Join-Path $env:TEMP ("sgur_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                try {
                    & secedit /export /cfg $cfg /areas USER_RIGHTS /quiet 2>&1 | Out-Null
                    $l = (Select-String -LiteralPath $cfg -Pattern '^\s*SeDenyNetworkLogonRight\s*=' -ErrorAction SilentlyContinue | Select-Object -First 1).Line
                    if ($l) { $l.Trim() } else { 'SeDenyNetworkLogonRight = (not set)' }
                } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
            } `
            -CheckBlock {
                $cfg = Join-Path $env:TEMP ("sgur_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                try {
                    & secedit /export /cfg $cfg /areas USER_RIGHTS /quiet 2>&1 | Out-Null
                    $l = (Select-String -LiteralPath $cfg -Pattern '^\s*SeDenyNetworkLogonRight\s*=' -ErrorAction SilentlyContinue | Select-Object -First 1).Line
                    if (-not $l) { return $false }
                    $ok = $true
                    foreach ($sid in 'S-1-5-32-546', 'S-1-5-113', 'S-1-5-114') { if ($l -notmatch [regex]::Escape($sid)) { $ok = $false } }
                    $ok
                } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
            } `
            -ApplyBlock {
                $inf = Join-Path $env:TEMP ("sgur_{0}.inf" -f [guid]::NewGuid().ToString('N'))
                $sdb = Join-Path $env:TEMP ("sgur_{0}.sdb" -f [guid]::NewGuid().ToString('N'))
                $infBody = @('[Unicode]', 'Unicode=yes', '[Version]', 'signature="$CHICAGO$"', 'Revision=1', '[Privilege Rights]', 'SeDenyNetworkLogonRight = *S-1-5-32-546,*S-1-5-113,*S-1-5-114') -join "`r`n"
                try {
                    Set-Content -LiteralPath $inf -Value $infBody -Encoding Unicode
                    $out = & secedit /configure /db $sdb /cfg $inf /areas USER_RIGHTS /quiet 2>&1
                    if ($LASTEXITCODE) { throw "secedit exited $LASTEXITCODE. $out" }
                } finally { Remove-Item -LiteralPath $inf, $sdb -Force -ErrorAction SilentlyContinue }
            }

        # 7.13 Allow Administrator account lockout - real remediation via secedit ([System Access]).
        Invoke-Remediation -Category $CAT -Item 'Allow Administrator account lockout' `
            -Description 'secedit: [System Access] AllowAdministratorLockout = 1' `
            -RiskNote 'Only has effect once an account lockout threshold > 0 exists (the WS2025 baseline sets 10). Verify with "net accounts". Applies to the built-in Administrator (RID 500).' `
            -ExpectedValue 'AllowAdministratorLockout = 1' `
            -CurrentStateBlock {
                $cfg = Join-Path $env:TEMP ("sgsa_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                try {
                    & secedit /export /cfg $cfg /areas SECURITYPOLICY /quiet 2>&1 | Out-Null
                    $l = (Select-String -LiteralPath $cfg -Pattern '^\s*AllowAdministratorLockout\s*=' -ErrorAction SilentlyContinue | Select-Object -First 1).Line
                    if ($l) { $l.Trim() } else { 'AllowAdministratorLockout = (not set)' }
                } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
            } `
            -CheckBlock {
                $cfg = Join-Path $env:TEMP ("sgsa_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                try {
                    & secedit /export /cfg $cfg /areas SECURITYPOLICY /quiet 2>&1 | Out-Null
                    $l = (Select-String -LiteralPath $cfg -Pattern '^\s*AllowAdministratorLockout\s*=\s*1\s*$' -ErrorAction SilentlyContinue | Select-Object -First 1).Line
                    [bool]$l
                } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
            } `
            -ApplyBlock {
                $inf = Join-Path $env:TEMP ("sgsa_{0}.inf" -f [guid]::NewGuid().ToString('N'))
                $sdb = Join-Path $env:TEMP ("sgsa_{0}.sdb" -f [guid]::NewGuid().ToString('N'))
                $infBody = @('[Unicode]', 'Unicode=yes', '[Version]', 'signature="$CHICAGO$"', 'Revision=1', '[System Access]', 'AllowAdministratorLockout = 1') -join "`r`n"
                try {
                    Set-Content -LiteralPath $inf -Value $infBody -Encoding Unicode
                    $out = & secedit /configure /db $sdb /cfg $inf /areas SECURITYPOLICY /quiet 2>&1
                    if ($LASTEXITCODE) { throw "secedit exited $LASTEXITCODE. $out" }
                } finally { Remove-Item -LiteralPath $inf, $sdb -Force -ErrorAction SilentlyContinue }
            }

        # 7.14 Rename the built-in Guest account (QID 105228) - real remediation only if -GuestAccountNewName was supplied.
        if ($GuestAccountNewName) {
            Invoke-Remediation -Category $CAT -Item 'Rename built-in Guest account (QID 105228)' `
                -Description "Rename the built-in Guest account (SID -501) to '$GuestAccountNewName' and ensure it stays disabled." `
                -RiskNote 'Removes the well-known name. Confirm no local scripts reference the literal name "Guest".' `
                -ExpectedValue "Guest (SID -501) named '$GuestAccountNewName' and disabled" `
                -CurrentStateBlock {
                    $g = Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { $_.SID.Value -like '*-501' } | Select-Object -First 1
                    if ($g) { "Name=$($g.Name); Enabled=$($g.Enabled)" } else { 'SID -501 account not found' }
                } `
                -CheckBlock {
                    $g = Get-LocalUser -ErrorAction Stop | Where-Object { $_.SID.Value -like '*-501' } | Select-Object -First 1
                    $g -and $g.Name -eq $GuestAccountNewName -and -not $g.Enabled
                } `
                -ApplyBlock {
                    $g = Get-LocalUser -ErrorAction Stop | Where-Object { $_.SID.Value -like '*-501' } | Select-Object -First 1
                    if (-not $g) { throw 'Built-in Guest account (SID -501) not found.' }
                    if ($g.Name -ne $GuestAccountNewName) { Rename-LocalUser -Name $g.Name -NewName $GuestAccountNewName }
                    if ((Get-LocalUser -Name $GuestAccountNewName).Enabled) { Disable-LocalUser -Name $GuestAccountNewName }
                }
        } else {
            Add-ResultRow -Category $CAT -Item 'Rename built-in Guest account (QID 105228)' -Status 'Manual/External Required' `
                -Details 'Pass -GuestAccountNewName "<name>" to rename and disable the built-in Guest account automatically. Manual: Rename-LocalUser -Name Guest -NewName "<name>"; Disable-LocalUser -Name "<name>".' `
                -ExpectedValue 'Guest renamed and disabled' -DetectedValue 'Built-in Guest not renamed'
        }

        # Manual / external items (RDP deny-right + patching) - report only.
        if (-not $SkipExtendedManualItems) {
            Add-ResultRow -Category $CAT -Item 'Deny logon through Remote Desktop Services (user right)' -Status 'Manual/External Required' `
                -Details 'EXCLUDED from automatic remediation by design: on an all-workgroup estate SeDenyRemoteInteractiveLogonRight = *S-1-5-32-546,*S-1-5-113 would also block local-admin RDP, and a deny right cannot be granted an exception. Apply manually (Guests only, or the full list once management moves to the VM console).' `
                -ExpectedValue 'Guests (+ Local account if RDP not needed)' -DetectedValue 'Only BUILTIN\Guests assigned'
            Add-ResultRow -Category $CAT -Item 'August 2026 Windows Update (QID 92439)' -Status 'Manual/External Required' `
                -Details 'Install KB5120228 and KB5120233 (build 26100.33158 -> 26100.33222 / 26100.33296). Patch via Windows Update / WSUS, then re-run Get-HotFix as evidence. Out of scope of configuration hardening.' `
                -ExpectedValue 'Build >= 26100.33222' -DetectedValue 'Build 26100.33158'
            Add-ResultRow -Category $CAT -Item 'QID 92446 - Defender EoP zero-day' -Status 'Manual/External Required' `
                -Details 'No Microsoft patch was available at assessment time. Track the MSRC advisory; apply the update as soon as it ships. Keep Defender platform / signatures current in the meantime.' `
                -ExpectedValue 'Patched when available' -DetectedValue 'No patch at assessment time'
            Add-ResultRow -Category $CAT -Item 'Third-party agent vulnerabilities' -Status 'Manual/External Required' `
                -Details 'Upgrade to non-vulnerable versions with the respective owners: Splunk Universal Forwarder (SVD-2026-*), Azure Connected Machine Agent (.NET/Go/gRPC), ServiceNow Agent (Ruby gem/Go/gRPC), Binalyze AIR (Go stdlib/gRPC/crypto).' `
                -ExpectedValue 'Agents on fixed versions' -DetectedValue 'Vulnerable agent builds present'
        }
    } else {
        Add-ResultRow -Category '7. Extended Baseline (Cybersecurity Review)' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipExtendedBaseline was passed'
    }

    # ===========================================================================
    # 8. ISO Policy Compliance (Qualys PC) - Windows Server 2025 WORKGROUP
    #    Consolidated coverage for the Control IDs in iso.xlsx (Qualys Policy
    #    Compliance export). Uses the SAME Add-ResultRow / Invoke-Remediation
    #    contract, status vocabulary (AlreadyCompliant / DryRun - would apply /
    #    Applied / Failed / Manual/External Required / Not Applicable / Skipped)
    #    and reboot-flag model as categories 1-7. No new columns / statuses /
    #    output format. High-impact items stay opt-in behind their own switch.
    # ===========================================================================
    if (-not $SkipIsoPolicy) {
        $ICAT = '8. ISO Policy Compliance (Qualys PC)'

        function Invoke-IsoReg {
            param([int]$Cid,[string]$Title,[string]$Path,[string]$Name,[string]$Type,$Data,
                  [bool]$Reboot = $false,[string]$Risk = '')
            Invoke-Remediation -Category $ICAT -Item ("CID {0} - {1}" -f $Cid, $Title) `
                -Description ("Set {0}\{1} = {2} (REG_{3})" -f $Path, $Name, $Data, $Type) `
                -RequiresReboot $Reboot -RiskNote $Risk `
                -ExpectedValue ("{0} = {1}" -f $Name, $Data) `
                -CurrentStateBlock ({ Get-RegValueString $Path $Name }.GetNewClosure()) `
                -CheckBlock ({
                    $cur = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
                    if ($null -eq $cur) { return $false }
                    if ($Type -eq 'String') { [string]$cur.$Name -eq [string]$Data }
                    else { [int]$cur.$Name -eq [int]$Data }
                }.GetNewClosure()) `
                -ApplyBlock ({
                    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null }
                    $v = if ($Type -eq 'String') { [string]$Data } else { [int]$Data }
                    Set-ItemProperty -Path $Path -Name $Name -Value $v -Type $Type -Force
                }.GetNewClosure())
        }

        # Registry-backed controls (Qualys value path -> value name -> type -> expected).
        $IsoReg = @(
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
        foreach ($it in $IsoReg) {
            if ($it.IE -and $SkipIeHardening) { continue }
            if ($it.Gate) {
                $gateOn = (Get-Variable -Name ($it.Gate.TrimStart('$')) -ValueOnly -ErrorAction SilentlyContinue) -eq $true
                if (-not $gateOn) {
                    Add-ResultRow -Category $ICAT -Item ("CID {0} - {1}" -f $it.Cid, $it.Title) -Status 'Manual/External Required' `
                        -Details ("NOT changed automatically (change-review item). {0} Re-run with -{1} to apply." -f $it.Risk, $it.Gate.TrimStart('$')) `
                        -DetectedValue (Get-RegValueString $it.Path $it.Name) -ExpectedValue ("{0} = {1}" -f $it.Name, $it.Data)
                    continue
                }
            }
            Invoke-IsoReg -Cid $it.Cid -Title $it.Title -Path $it.Path -Name $it.Name -Type $it.Type -Data $it.Data -Reboot $it.Reboot -Risk $it.Risk
        }
        if ($SkipIeHardening) {
            Add-ResultRow -Category $ICAT -Item 'Internet Explorer hardening (89 control IDs)' -Status 'Skipped (category)' `
                -Details '-SkipIeHardening was passed. Re-run without it to remediate the IE zone-lockdown and feature-control registry policy values (HKLM only; no reboot; IE is not installed on WS2025 so production impact is minimal).'
        }

        # --- CID 4501 : advanced audit 'Audit Policy Change' = Success and Failure (auditpol) ---
        Invoke-Remediation -Category $ICAT -Item 'CID 4501 - Audit "Audit Policy Change" = Success and Failure' `
            -Description 'auditpol /set /subcategory:"Audit Policy Change" /success:enable /failure:enable' `
            -ExpectedValue 'Success and Failure' `
            -CurrentStateBlock {
                try {
                    $csv = (& auditpol /get /subcategory:"Audit Policy Change" /r 2>$null | ConvertFrom-Csv)
                    $row = $csv | Where-Object { $_.Subcategory -eq 'Audit Policy Change' } | Select-Object -First 1
                    "Inclusion Setting=$($row.'Inclusion Setting')"
                } catch { 'unknown' }
            } `
            -CheckBlock {
                $csv = (& auditpol /get /subcategory:"Audit Policy Change" /r 2>$null | ConvertFrom-Csv)
                $row = $csv | Where-Object { $_.Subcategory -eq 'Audit Policy Change' } | Select-Object -First 1
                $row -and $row.'Inclusion Setting' -eq 'Success and Failure'
            } `
            -ApplyBlock {
                $o = & auditpol /set /subcategory:"Audit Policy Change" /success:enable /failure:enable 2>&1
                if ($LASTEXITCODE) { throw "auditpol exited $LASTEXITCODE. $o" }
            }

        # --- CID 3950/3951/3952 : Windows Firewall profile state = On  (opt-in: -EnableFirewallProfiles) ---
        foreach ($fp in @(
            @{ Cid=3952; Profile='Domain' }, @{ Cid=3951; Profile='Private' }, @{ Cid=3950; Profile='Public' } )) {
            if (-not $EnableFirewallProfiles) {
                Add-ResultRow -Category $ICAT -Item ("CID {0} - Windows Firewall: Firewall state ({1})" -f $fp.Cid, $fp.Profile) -Status 'Manual/External Required' `
                    -Details "Host firewall OFF for the $($fp.Profile) profile (finding on 10.50.16.32). NOT changed automatically - enabling the firewall with the default inbound-block stance can cut RDP / agent / app traffic. Stage inbound allow rules, then re-run with -EnableFirewallProfiles during a change window (keep VM-console fallback)." `
                    -DetectedValue 'EnableFirewall = 0 (Off)' -ExpectedValue 'On (1)'
                continue
            }
            $fpp = "HKLM:\Software\Policies\Microsoft\WindowsFirewall\$($fp.Profile)Profile"
            Invoke-Remediation -Category $ICAT -Item ("CID {0} - Windows Firewall: Firewall state ({1})" -f $fp.Cid, $fp.Profile) `
                -Description ("Set-NetFirewallProfile -Profile {0} -Enabled True (also $fpp\EnableFirewall=1)" -f $fp.Profile) `
                -RiskNote 'Enabling the firewall can black-hole inbound RDP/agent/app traffic if matching allow rules are absent. Confirmed opt-in via -EnableFirewallProfiles.' `
                -ExpectedValue 'Enabled = True' `
                -CurrentStateBlock ({ try { 'Enabled=' + (Get-NetFirewallProfile -Profile $fp.Profile -ErrorAction Stop).Enabled } catch { Get-RegValueString $fpp 'EnableFirewall' } }.GetNewClosure()) `
                -CheckBlock ({ try { (Get-NetFirewallProfile -Profile $fp.Profile -ErrorAction Stop).Enabled -eq $true } catch { $v = Get-ItemProperty -Path $fpp -Name EnableFirewall -ErrorAction SilentlyContinue; $v -and $v.EnableFirewall -eq 1 } }.GetNewClosure()) `
                -ApplyBlock ({
                    try { Set-NetFirewallProfile -Profile $fp.Profile -Enabled True -ErrorAction Stop }
                    catch {
                        if (-not (Test-Path -LiteralPath $fpp)) { New-Item -Path $fpp -Force -ErrorAction SilentlyContinue | Out-Null }
                        Set-ItemProperty -Path $fpp -Name EnableFirewall -Value 1 -Type DWord -Force
                    }
                }.GetNewClosure())
        }

        # --- CID 29576 / 29583 : mandate minimum SMB dialect 3.0.0  (opt-in: -SetSmbMinSmb3) ---
        foreach ($sd in @(
            @{ Cid=29576; Side='Server'; Key='HKLM:\Software\Policies\Microsoft\Windows\LanmanServer' },
            @{ Cid=29583; Side='Client'; Key='HKLM:\Software\Policies\Microsoft\Windows\LanmanWorkstation' } )) {
            if (-not $SetSmbMinSmb3) {
                Add-ResultRow -Category $ICAT -Item ("CID {0} - Mandate minimum SMB version ({1})" -f $sd.Cid, $sd.Side) -Status 'Manual/External Required' `
                    -Details "NOT changed automatically - clients that only speak SMB 2.0.2/2.1 lose connectivity, and applying the server value can bounce active SMB sessions. Inventory SMB peers (Get-SmbSession / arrays / backup), then re-run with -SetSmbMinSmb3 in a change window." `
                    -DetectedValue 'MinSmb2Dialect = Setting not found' -ExpectedValue 'SMB 3.0.0 (768)'
                continue
            }
            Invoke-Remediation -Category $ICAT -Item ("CID {0} - Mandate minimum SMB version ({1})" -f $sd.Cid, $sd.Side) `
                -Description ("Set-Smb{0}Configuration -MinSmb2Dialect SMB300 (also $($sd.Key)\MinSmb2Dialect=768)" -f $sd.Side) `
                -RiskNote 'Blocks SMB 2.0.2/2.1 peers; server-side apply can drop active SMB sessions. Confirmed opt-in via -SetSmbMinSmb3.' `
                -ExpectedValue 'MinSmb2Dialect = 768 (SMB 3.0.0)' `
                -CurrentStateBlock ({ Get-RegValueString $sd.Key 'MinSmb2Dialect' }.GetNewClosure()) `
                -CheckBlock ({ $v = Get-ItemProperty -Path $sd.Key -Name MinSmb2Dialect -ErrorAction SilentlyContinue; $v -and [int]$v.MinSmb2Dialect -eq 768 }.GetNewClosure()) `
                -ApplyBlock ({
                    try {
                        if ($sd.Side -eq 'Server') { Set-SmbServerConfiguration -MinSmb2Dialect SMB300 -Confirm:$false -ErrorAction Stop }
                        else { Set-SmbClientConfiguration -MinSmb2Dialect SMB300 -Confirm:$false -ErrorAction Stop }
                    } catch {
                        if (-not (Test-Path -LiteralPath $sd.Key)) { New-Item -Path $sd.Key -Force -ErrorAction SilentlyContinue | Out-Null }
                        Set-ItemProperty -Path $sd.Key -Name MinSmb2Dialect -Value 768 -Type DWord -Force
                    }
                }.GetNewClosure())
        }

        # --- CID 1071 / 2342 : account policy via secedit  (opt-in: -ApplyAccountPolicy) ---
        foreach ($ap in @(
            @{ Cid=1071; Key='MinimumPasswordLength'; Target=$MinPasswordLength;      Cmp='ge'; Label='Minimum Password Length' },
            @{ Cid=2342; Key='LockoutBadCount';       Target=$AccountLockoutThreshold; Cmp='le'; Label='Account Lockout Threshold' } )) {
            if (-not $ApplyAccountPolicy) {
                Add-ResultRow -Category $ICAT -Item ("CID {0} - {1}" -f $ap.Cid, $ap.Label) -Status 'Manual/External Required' `
                    -Details ("NOT changed automatically. Set via secedit [System Access] {0} = {1} (or 'net accounts'). Re-run with -ApplyAccountPolicy (and -MinPasswordLength / -AccountLockoutThreshold) once the target values are agreed with the security team - a low lockout threshold plus AllowAdministratorLockout can enable account-lockout DoS." -f $ap.Key, $ap.Target) `
                    -DetectedValue 'see secedit /export /areas SECURITYPOLICY' -ExpectedValue ("{0} {1} {2}" -f $ap.Key, $(if($ap.Cmp -eq 'ge'){'>='}else{'<='}), $ap.Target)
                continue
            }
            Invoke-Remediation -Category $ICAT -Item ("CID {0} - {1}" -f $ap.Cid, $ap.Label) `
                -Description ("secedit /configure [System Access] {0} = {1}" -f $ap.Key, $ap.Target) `
                -RiskNote 'Account-policy change. Confirmed opt-in via -ApplyAccountPolicy.' `
                -ExpectedValue ("{0} {1} {2}" -f $ap.Key, $(if($ap.Cmp -eq 'ge'){'>='}else{'<='}), $ap.Target) `
                -CurrentStateBlock ({
                    $cfg = Join-Path $env:TEMP ("iso_sa_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                    try { & secedit /export /cfg $cfg /areas SECURITYPOLICY /quiet 2>&1 | Out-Null
                        (Select-String -LiteralPath $cfg -Pattern ("^\s*{0}\s*=" -f $ap.Key) | Select-Object -First 1).Line.Trim()
                    } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
                }.GetNewClosure()) `
                -CheckBlock ({
                    $cfg = Join-Path $env:TEMP ("iso_sa_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                    try { & secedit /export /cfg $cfg /areas SECURITYPOLICY /quiet 2>&1 | Out-Null
                        $l = (Select-String -LiteralPath $cfg -Pattern ("^\s*{0}\s*=\s*(\d+)" -f $ap.Key) | Select-Object -First 1)
                        if (-not $l) { return $false }
                        $cur = [int]$l.Matches[0].Groups[1].Value
                        if ($ap.Cmp -eq 'ge') { $cur -ge [int]$ap.Target } else { $cur -le [int]$ap.Target -and $cur -gt 0 }
                    } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
                }.GetNewClosure()) `
                -ApplyBlock ({
                    $inf = Join-Path $env:TEMP ("iso_sa_{0}.inf" -f [guid]::NewGuid().ToString('N'))
                    $sdb = Join-Path $env:TEMP ("iso_sa_{0}.sdb" -f [guid]::NewGuid().ToString('N'))
                    $body = @('[Unicode]','Unicode=yes','[Version]','signature="$CHICAGO$"','Revision=1','[System Access]',("{0} = {1}" -f $ap.Key, [int]$ap.Target)) -join "
"
                    try { Set-Content -LiteralPath $inf -Value $body -Encoding Unicode
                        $o = & secedit /configure /db $sdb /cfg $inf /areas SECURITYPOLICY /quiet 2>&1
                        if ($LASTEXITCODE) { throw "secedit exited $LASTEXITCODE. $o" }
                    } finally { Remove-Item -LiteralPath $inf, $sdb -Force -ErrorAction SilentlyContinue }
                }.GetNewClosure())
        }

        # --- CID 2196 : Deny access to this computer from the network (secedit; same right as category 7.12) ---
        Invoke-Remediation -Category $ICAT -Item 'CID 2196 - Deny access to this computer from the network (user right)' `
            -Description 'secedit: [Privilege Rights] SeDenyNetworkLogonRight = *S-1-5-32-546,*S-1-5-113,*S-1-5-114' `
            -RiskNote 'Network-logon deny only (SMB/RPC/WinRM) for Guests + all local accounts. Does NOT affect console or RDP. Safe on a VMware guest-ops managed estate; verify no app authenticates to this host over SMB with a local account first.' `
            -ExpectedValue 'Contains S-1-5-32-546, S-1-5-113 and S-1-5-114' `
            -CurrentStateBlock {
                $cfg = Join-Path $env:TEMP ("iso_ura_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                try { & secedit /export /cfg $cfg /areas USER_RIGHTS /quiet 2>&1 | Out-Null
                    $l = (Select-String -LiteralPath $cfg -Pattern '^\s*SeDenyNetworkLogonRight\s*=' | Select-Object -First 1).Line
                    if ($l) { $l.Trim() } else { 'SeDenyNetworkLogonRight = (not set)' }
                } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
            } `
            -CheckBlock {
                $cfg = Join-Path $env:TEMP ("iso_ura_{0}.cfg" -f [guid]::NewGuid().ToString('N'))
                try { & secedit /export /cfg $cfg /areas USER_RIGHTS /quiet 2>&1 | Out-Null
                    $l = (Select-String -LiteralPath $cfg -Pattern '^\s*SeDenyNetworkLogonRight\s*=' | Select-Object -First 1).Line
                    if (-not $l) { return $false }
                    $ok = $true
                    foreach ($sid in 'S-1-5-32-546','S-1-5-113','S-1-5-114') { if ($l -notmatch [regex]::Escape($sid)) { $ok = $false } }
                    $ok
                } finally { Remove-Item -LiteralPath $cfg -Force -ErrorAction SilentlyContinue }
            } `
            -ApplyBlock {
                $inf = Join-Path $env:TEMP ("iso_ura_{0}.inf" -f [guid]::NewGuid().ToString('N'))
                $sdb = Join-Path $env:TEMP ("iso_ura_{0}.sdb" -f [guid]::NewGuid().ToString('N'))
                $body = @('[Unicode]','Unicode=yes','[Version]','signature="$CHICAGO$"','Revision=1','[Privilege Rights]','SeDenyNetworkLogonRight = *S-1-5-32-546,*S-1-5-113,*S-1-5-114') -join "
"
                try { Set-Content -LiteralPath $inf -Value $body -Encoding Unicode
                    $o = & secedit /configure /db $sdb /cfg $inf /areas USER_RIGHTS /quiet 2>&1
                    if ($LASTEXITCODE) { throw "secedit exited $LASTEXITCODE. $o" }
                } finally { Remove-Item -LiteralPath $inf, $sdb -Force -ErrorAction SilentlyContinue }
            }

        # --- CID 25357 / 30456 : Attack Surface Reduction rules ---
        foreach ($ar in @(
            @{ Cid=25357; Guid='56a863a9-875e-4185-98a7-b882c64b5ce5'; Name='Block abuse of exploited vulnerable signed drivers'; Optin=$false },
            @{ Cid=30456; Guid='a8f5898e-1dc8-49a9-9878-85004b8a61e6'; Name='ASR rule a8f5898e (server webshell / verify friendly name)'; Optin=$true } )) {
            if ($ar.Optin -and -not $EnableServerAsrRules) {
                Add-ResultRow -Category $ICAT -Item ("CID {0} - {1}" -f $ar.Cid, $ar.Name) -Status 'Manual/External Required' `
                    -Details ("Confirm the rule's friendly name/intent with the security team, then re-run with -EnableServerAsrRules (starts the rule in Block). GUID {0}." -f $ar.Guid) `
                    -DetectedValue 'Rule not configured' -ExpectedValue 'Block (1)'
                continue
            }
            if (-not (Get-Command Add-MpPreference -ErrorAction SilentlyContinue)) {
                Add-ResultRow -Category $ICAT -Item ("CID {0} - {1}" -f $ar.Cid, $ar.Name) -Status 'Manual/External Required' `
                    -Details 'Defender module (Add-MpPreference) not present - configure the ASR rule via the third-party AV / Defender console.' `
                    -DetectedValue 'Defender cmdlets unavailable' -ExpectedValue 'Block (1)'
                continue
            }
            Invoke-Remediation -Category $ICAT -Item ("CID {0} - {1}" -f $ar.Cid, $ar.Name) `
                -Description ("Add-MpPreference -AttackSurfaceReductionRules_Ids {0} -AttackSurfaceReductionRules_Actions Enabled" -f $ar.Guid) `
                -RiskNote 'ASR rules can block legitimate behaviour. Pilot in AuditMode first if the host runs affected workloads (drivers / web roots).' `
                -ExpectedValue 'Rule action = Block (1)' `
                -CurrentStateBlock ({
                    try { $p = Get-MpPreference -ErrorAction Stop
                        $i = [array]::IndexOf(@($p.AttackSurfaceReductionRules_Ids), $ar.Guid)
                        if ($i -ge 0) { "action=$(@($p.AttackSurfaceReductionRules_Actions)[$i])" } else { 'not configured' }
                    } catch { 'unknown' }
                }.GetNewClosure()) `
                -CheckBlock ({
                    $p = Get-MpPreference -ErrorAction Stop
                    $i = [array]::IndexOf(@($p.AttackSurfaceReductionRules_Ids), $ar.Guid)
                    $i -ge 0 -and [int](@($p.AttackSurfaceReductionRules_Actions)[$i]) -eq 1
                }.GetNewClosure()) `
                -ApplyBlock ({ Add-MpPreference -AttackSurfaceReductionRules_Ids $ar.Guid -AttackSurfaceReductionRules_Actions Enabled -ErrorAction Stop }.GetNewClosure())
        }

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 29586, '''Configure hash algorithms for certificate logon (PKINITHashAlgorithmConfigurationEnabled)'' - KDC') -Status 'Not Applicable' `
            -Details 'NOT APPLICABLE - KDC role not installed (not a Domain Controller) Evidence: Get-CimInstance Win32_ComputerSystem (PartOfDomain=False / DomainRole 2) and Get-WindowsFeature AD-Domain-Services (not installed).' -DetectedValue 'PKINITHashAlgorithmConfigurationEnabled = Setting not found' -ExpectedValue 'N/A (would be 1 on a Domain Controller)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 29587, '''Configure hash algorithms for certificate logon (PKINITSHA1)'' - KDC') -Status 'Not Applicable' `
            -Details 'NOT APPLICABLE - KDC role not installed (not a Domain Controller) Evidence: Get-CimInstance Win32_ComputerSystem (PartOfDomain=False / DomainRole 2) and Get-WindowsFeature AD-Domain-Services (not installed).' -DetectedValue 'PKINITSHA1 = Setting not found' -ExpectedValue 'N/A (would be 1 on a Domain Controller)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 29588, '''Configure hash algorithms for certificate logon (PKINITSHA256)'' - KDC') -Status 'Not Applicable' `
            -Details 'NOT APPLICABLE - KDC role not installed (not a Domain Controller) Evidence: Get-CimInstance Win32_ComputerSystem (PartOfDomain=False / DomainRole 2) and Get-WindowsFeature AD-Domain-Services (not installed).' -DetectedValue 'PKINITSHA256 = Setting not found' -ExpectedValue 'N/A (would be 3 on a Domain Controller)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 29589, '''Configure hash algorithms for certificate logon (PKINITSHA384)'' - KDC') -Status 'Not Applicable' `
            -Details 'NOT APPLICABLE - KDC role not installed (not a Domain Controller) Evidence: Get-CimInstance Win32_ComputerSystem (PartOfDomain=False / DomainRole 2) and Get-WindowsFeature AD-Domain-Services (not installed).' -DetectedValue 'PKINITSHA384 = Setting not found' -ExpectedValue 'N/A (would be 3 on a Domain Controller)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 29590, '''Configure hash algorithms for certificate logon (PKINITSHA512)'' - KDC') -Status 'Not Applicable' `
            -Details 'NOT APPLICABLE - KDC role not installed (not a Domain Controller) Evidence: Get-CimInstance Win32_ComputerSystem (PartOfDomain=False / DomainRole 2) and Get-WindowsFeature AD-Domain-Services (not installed).' -DetectedValue 'PKINITSHA512 = Setting not found' -ExpectedValue 'N/A (would be 3 on a Domain Controller)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 29740, '''Turn On VBS (Machine Identity Isolation Policy)''') -Status 'Not Applicable' `
            -Details 'NOT APPLICABLE - Workgroup - no AD machine identity to isolate Evidence: Get-CimInstance Win32_ComputerSystem (PartOfDomain=False / DomainRole 2) and Get-WindowsFeature AD-Domain-Services (not installed).' -DetectedValue 'MachineIdentityIsolation = Key not found' -ExpectedValue 'N/A unless domain-joined (would be 2 = enforcement on a domain member)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 2186, 'Groups/accounts with ''Back up files and directories'' (SeBackupPrivilege)') -Status 'Manual/External Required' `
            -Details 'Either (a) file a Qualys exception documenting Splunk''s operational need, or (b) remove NT SERVICE\SplunkForwarder via secedit and validate Splunk still forwards. Do NOT auto-remove.' -DetectedValue 'BUILTIN\Administrators, NT SERVICE\SplunkForwarder' -ExpectedValue 'Administrators only (Qualys also accepts ''Right not assigned'')'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 2195, 'Groups/accounts with ''Debug Programs'' (SeDebugPrivilege)') -Status 'Manual/External Required' `
            -Details 'Remove the ''ServiceNow Users'' group from SeDebugPrivilege via secedit unless a written justification exists; then re-scan. Treat as change-reviewed manual remediation.' -DetectedValue 'BUILTIN\Administrators, <host>\ServiceNow Users' -ExpectedValue 'Administrators only (Qualys also accepts ''Right not assigned'')'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 2200, 'Groups/accounts with ''Deny log on through Remote Desktop Services'' (SeDenyRemoteInteractiveLogonRight)') -Status 'Manual/External Required' `
            -Details 'MANUAL on this estate: adding ''Local account'' (S-1-5-113) denies RDP to every local account, including the local Administrator used to manage these WORKGROUP servers. Apply only ''Guests'' automatically; apply the full list manually once console/iLO/VM-console access is confirmed as the management path. A deny right cannot be granted an exception.' -DetectedValue 'Right not assigned' -ExpectedValue 'Contains S-1-5-113 (Local account). CIS: Guests + Local account.'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 2392, 'Groups/accounts with ''Manage auditing and security log'' (SeSecurityPrivilege)') -Status 'Manual/External Required' `
            -Details 'Preferred: remove SeSecurityPrivilege from SplunkForwarder and add the Splunk service account to BUILTIN\Event Log Readers (read-only). Otherwise document a Qualys exception. Do NOT auto-remediate.' -DetectedValue 'BUILTIN\Administrators, NT SERVICE\SplunkForwarder' -ExpectedValue 'Administrators only (Qualys also accepts ''Right not assigned'')'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 2642, 'Groups/accounts with ''Impersonate a client after authentication'' (SeImpersonatePrivilege)') -Status 'Manual/External Required' `
            -Details 'If IIS is legitimately installed on 10.50.16.32, file a Qualys exception (IIS_IUSRS + SeImpersonate is by design). The ''RESTRICTED SERVICES\PrintSpoolerService'' entry matches the \bSERVICE$ regex and is compliant - the finding is likely a scanner display artefact; re-scan after the other user-rights fixes. No automatic change.' -DetectedValue 'BUILTIN\Administrators, BUILTIN\IIS_IUSRS (host .32), NT AUTHORITY\{LOCAL SERVICE, NETWORK SERVICE, SERVICE}, RESTRICTED SERVICES\PrintSpoolerService' -ExpectedValue 'Administrators, LOCAL SERVICE, NETWORK SERVICE, SERVICE (regex allows the \bSERVICE$ pattern which already covers the spooler service SID)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 25350, '''Allow Custom SSPs and APs to be loaded into LSASS'' (AllowCustomSSPsAPs)') -Status 'Manual/External Required' `
            -Details 'Decision required - do NOT auto-remediate. If hardened stance (0) is chosen and Qualys policy insists on 1, file a documented exception referencing RunAsPPL + Credential Guard as compensating controls.' -DetectedValue 'AllowCustomSSPsAPs = Not Configured' -ExpectedValue '1 per Qualys policy - VALIDATE; recommended hardened value is 0 unless a custom SSP is required'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 30461, '''Allow network protection on Windows Server'' (AllowNetworkProtectionOnWinServer)') -Status 'Manual/External Required' `
            -Details 'Decision required. If confirmed: New-Item + Set-ItemProperty ''...\Network Protection'' AllowNetworkProtectionOnWinServer 0 (DWORD). Recommended alternative: set to 1 and set EnableNetworkProtection=1 (Block) and file a Qualys exception.' -DetectedValue 'AllowNetworkProtectionOnWinServer = Key not found' -ExpectedValue '0 (per Qualys policy) - VALIDATE against your Defender standard'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 10821, 'IE (per-user) - Turn on the auto-complete feature for user names and passwords on forms (FormSuggest Passwords = no)') -Status 'Manual/External Required' `
            -Details 'PER-USER (HKEY_USERS) setting - not a single machine value. For each loaded hive under HKEY_USERS and for HKU\.DEFAULT (and every profile hive mounted from NTUSER.DAT): New-Item -Force ''<hive>\Software\Policies\Microsoft\Internet Explorer\Main'' ; Set-ItemProperty ... ''FormSuggest Passwords'' ''no'' (REG_SZ). Best delivered as a per-user logon script / Active Setup rather than a one-shot machine script.' -DetectedValue 'FormSuggest Passwords: ''Setting not found'' for several profile SIDs (some profiles already = ''no'')' -ExpectedValue 'no for every user hive (Qualys regex no$)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 17333, 'IE (per-user) - Prompt me to save passwords (FormSuggest PW Ask = no)') -Status 'Manual/External Required' `
            -Details 'PER-USER (HKEY_USERS) setting - not a single machine value. For each loaded hive under HKEY_USERS and for HKU\.DEFAULT (and every profile hive mounted from NTUSER.DAT): New-Item -Force ''<hive>\Software\Policies\Microsoft\Internet Explorer\Main'' ; Set-ItemProperty ... ''FormSuggest PW Ask'' ''no'' (REG_SZ). Best delivered as a per-user logon script / Active Setup rather than a one-shot machine script.' -DetectedValue 'FormSuggest PW Ask: ''Setting not found'' for several profile SIDs (some profiles already = ''no'')' -ExpectedValue 'no for every user hive (Qualys regex no$)'

        Add-ResultRow -Category $ICAT -Item ('CID {0} - {1}' -f 30467, 'IE (per-user) - Turn on the auto-complete feature ... (Control Panel\FormSuggest Passwords = 1, i.e. policy-locked)') -Status 'Manual/External Required' `
            -Details 'PER-USER (HKEY_USERS) setting - not a single machine value. For each loaded hive under HKEY_USERS and for HKU\.DEFAULT (and every profile hive mounted from NTUSER.DAT): New-Item -Force ''<hive>\Software\Policies\Microsoft\Internet Explorer\Control Panel'' ; Set-ItemProperty ... ''FormSuggest Passwords'' ''1'' (REG_SZ). Best delivered as a per-user logon script / Active Setup rather than a one-shot machine script.' -DetectedValue 'FormSuggest Passwords: ''Setting not found'' for several profile SIDs (some profiles already = ''1'')' -ExpectedValue '1 for every user hive (Qualys regex .+:1)'
    } else {
        Add-ResultRow -Category '8. ISO Policy Compliance (Qualys PC)' -Item 'All items' -Status 'Skipped (category)' -Details '-SkipIsoPolicy was passed'
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

    # Windows Firewall keyword scopes that are valid -RemoteAddress values in
    # addition to IPs / CIDR / ranges.
    $fwKeywords = 'LocalSubnet', 'Any', 'DNS', 'DHCP', 'WINS', 'DefaultGateway', 'Intranet', 'Internet', 'PlayToDevice', 'RemoteCorpNetwork'

    $effectiveRange = @()
    if ($WinRmAllowedSourceRange -and $WinRmAllowedSourceRange.Count -gt 0) {
        $effectiveRange = $WinRmAllowedSourceRange
    } elseif ($RestrictWinRmToLocalSubnet) {
        # "restrict the listener filters" with no specific management subnet given:
        # scope the WinRM firewall rules to LocalSubnet - narrower than 'Any', and
        # cannot lock out a same-subnet administrator.
        $effectiveRange = @('LocalSubnet')
    }

    if ($effectiveRange.Count -gt 0) {
        $parts = foreach ($entry in $effectiveRange) {
            if ($fwKeywords -contains $entry) {
                ConvertTo-PsSingleQuoted $entry
            } else {
                $clean = ($entry -replace '[^0-9A-Fa-f\.:/\-]', '')
                if ($clean -ne $entry -or [string]::IsNullOrWhiteSpace($clean)) {
                    throw "WinRmAllowedSourceRange entry '$entry' is neither a firewall keyword ($($fwKeywords -join ', ')) nor a plain IP/CIDR/range - refused."
                }
                ConvertTo-PsSingleQuoted $clean
            }
        }
        $rangesLiteral = '@(' + ($parts -join ',') + ')'
    } else {
        $rangesLiteral = '@()'
    }

    # Category 7: sanitise -WinRmFilterRange (IPv4 addresses / ranges only) into a literal.
    $winRmFilterLiteral = '@()'
    if ($WinRmFilterRange -and $WinRmFilterRange.Count -gt 0) {
        $fparts = foreach ($entry in $WinRmFilterRange) {
            $clean = ($entry -replace '[^0-9A-Fa-f\.:/\-]', '')
            if ($clean -ne $entry -or [string]::IsNullOrWhiteSpace($clean)) {
                throw "WinRmFilterRange entry '$entry' contains characters outside [0-9A-Fa-f.:/-] - refused."
            }
            ConvertTo-PsSingleQuoted $clean
        }
        $winRmFilterLiteral = '@(' + ($fparts -join ',') + ')'
    }

    # Category 7: sanitise -GuestAccountNewName (letters/digits/._- only, <=20) into a literal.
    $guestNameLiteral = "''"
    if ($GuestAccountNewName) {
        if ($GuestAccountNewName -notmatch '^[A-Za-z0-9._-]{1,20}$') {
            throw "GuestAccountNewName '$GuestAccountNewName' must be 1-20 chars of letters, digits, dot, underscore or hyphen - refused."
        }
        $guestNameLiteral = ConvertTo-PsSingleQuoted $GuestAccountNewName
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
`$SkipExtendedBaseline   = $(ConvertTo-PsBool $SkipExtendedBaseline)
`$SkipExtendedManualItems = $(ConvertTo-PsBool $SkipExtendedManualItems)
`$SkipIsoPolicy          = $(ConvertTo-PsBool $SkipIsoPolicy)
`$SkipIeHardening        = $(ConvertTo-PsBool $SkipIeHardening)
`$EnableFirewallProfiles = $(ConvertTo-PsBool $EnableFirewallProfiles)
`$DisableNetbios         = $(ConvertTo-PsBool $DisableNetbios)
`$SetSmbMinSmb3          = $(ConvertTo-PsBool $SetSmbMinSmb3)
`$RestrictLocalAcctNetworkLogon = $(ConvertTo-PsBool $RestrictLocalAcctNetworkLogon)
`$EnableUacHardening     = $(ConvertTo-PsBool $EnableUacHardening)
`$ApplyAccountPolicy     = $(ConvertTo-PsBool $ApplyAccountPolicy)
`$EnableServerAsrRules   = $(ConvertTo-PsBool $EnableServerAsrRules)
`$MinPasswordLength      = $([int]$MinPasswordLength)
`$AccountLockoutThreshold = $([int]$AccountLockoutThreshold)
`$CachedLogonsCount      = $([int]$CachedLogonsCount)
`$WinRmFilterRange       = $winRmFilterLiteral
`$GuestAccountNewName    = $guestNameLiteral
`$DisableWinRM           = $(ConvertTo-PsBool $DisableWinRM)
`$WinRmAllowedSourceRange = $rangesLiteral
`$EnforceAesOnlyKerberos = $(ConvertTo-PsBool $EnforceAesOnlyKerberos)
`$TreatAllAsWorkgroup    = $(ConvertTo-PsBool $TreatAllAsWorkgroup)
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
