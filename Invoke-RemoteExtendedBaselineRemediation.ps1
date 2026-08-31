#Requires -Version 5.1
<#
.SYNOPSIS
    Companion to Invoke-RemoteSecurityGapRemediation.ps1.

    Adds the extra Microsoft Security Baseline controls raised by the Cybersecurity
    team's Qualys review (SF-UFM-APP01 master remediation reference), using the SAME
    delivery model, safety model and result format as the other remote scripts:
      * Central PowerCLI orchestration, per-VM, over VMware Tools Guest Operations
        (no WinRM / PsExec / SMB / RDP / guest network path).
      * DRY-RUN by default. Nothing changes without -Apply.
      * Same per-VM validation battery and per-item statuses
        (AlreadyCompliant / DryRun - would apply / Applied / Failed /
         Manual/External Required / Not Applicable), the same
        "=== <host> (<ip>) ===" + "[Status] Category > Item" on-screen layout,
        the same "REBOOT REQUIRED" block, and the same central CSV columns.
      * Never reboots. Items that need a reboot are applied and listed.

    CONTROLS HANDLED AUTOMATICALLY (registry / auditpol - safely scriptable):
        1  HVCI - UEFI lock                 (DeviceGuard\...\HVCI\Locked = 1)          reboot
        2  Kernel-mode HW Stack Protection  (DeviceGuard\...\KernelShadowStacks)       reboot
        3  LSASS RunAsPPL - UEFI lock       (Control\LSA\RunAsPPL = 1)                 reboot
        4  PowerShell Script Block Logging  (+ Invocation logging)
        5  WinRM "Allow auto configuration" (Policies\...\WinRM\Service\AllowAutoConfig = 0)
        6  WinRM IPv4Filter / IPv6Filter    (only if -WinRmFilterRange supplied)
        7  Kerberos PKINIT hash SHA256/384/512 = Supported (3)
        8  Windows Ink Workspace            (AllowWindowsInkWorkspace = 0)
        9  Audit "Audit Policy Change"      = Success and Failure
        10 Cached logons count             (Winlogon\CachedLogonsCount, -CachedLogonsCount)
        11 Credential Guard policy value    (detects a GPO that disables it)

    CONTROLS REPORTED "Manual/External Required" (Local Security Policy / patching -
    NOT auto-applied here, exact fix text is in each result row):
        - Deny access to this computer from the network      (user-rights, secedit)
        - Deny logon through Remote Desktop Services          (user-rights, secedit)
        - Allow Administrator account lockout                 (account policy, secedit)
        - Rename built-in Guest account                       (secedit)
        - Unused local accounts (s.l.jganzon.m, s.l.samohamed.m)   (manual review)
        - August 2026 Windows update (KB5120228 / KB5120233)  (Windows Update)
        - QID 92446 Defender EoP zero-day                     (no patch yet - monitor)
        - Splunk UF / Azure Arc / ServiceNow / Binalyze agent vulns  (agent owners)

.PARAMETER Apply
    Master switch. Without it nothing is changed - assessment only.

.PARAMETER VMListPath / CredentialPath / OutputPath / ToolsWaitSecs
    Same meaning as the other remote scripts. Defaults: C:\temp\vmlist.txt,
    C:\temp\wincred.xml, .\SecurityGap_Reports, 180.

.PARAMETER WinRmFilterRange
    One or more IPv4 ranges (e.g. '10.50.10.1-10.50.10.254') / addresses used to set
    the WinRM listener IPv4Filter / IPv6Filter. If omitted, that item is reported
    Manual/External Required (the value is not guessed).

.PARAMETER CachedLogonsCount
    Target for Winlogon\CachedLogonsCount. Default 4 (CIS). Use 0 on workgroup
    servers to fully clear Qualys QID 90007.

.PARAMETER SkipManualItems
    Do not emit the "Manual/External Required" rows for the secedit / patch items
    (keeps the report to just the auto-remediated controls).

.EXAMPLE
    .\Invoke-RemoteExtendedBaselineRemediation.ps1
    .\Invoke-RemoteExtendedBaselineRemediation.ps1 -Apply
    .\Invoke-RemoteExtendedBaselineRemediation.ps1 -Apply -CachedLogonsCount 0 -WinRmFilterRange '10.50.10.1-10.50.10.254'

.NOTES
    Prerequisites: PowerCLI imported and already Connect-VIServer'd; the vCenter
    account holds the three "Guest Operation ..." privileges. This script performs
    no vSphere write. Run Invoke-RemotePostRemediationCheck.ps1 after a reboot to
    re-verify, and collect gpresult / secedit / auditpol evidence for CS.
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [switch]   $Apply,

    [string]   $VMListPath     = 'C:\temp\vmlist.txt',
    [string]   $CredentialPath = 'C:\temp\wincred.xml',
    [string]   $OutputPath     = (Join-Path $PSScriptRoot 'SecurityGap_Reports'),

    [string[]] $WinRmFilterRange,

    [ValidateRange(0, 50)]
    [int]      $CachedLogonsCount = 4,

    [switch]   $SkipManualItems,

    [ValidateRange(30, 3600)]
    [int]      $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'
$ScriptBuild = '2026-08-31a-extended-baseline'

Write-Host ("Invoke-RemoteExtendedBaselineRemediation.ps1  [build $ScriptBuild]") -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
if ($Apply) {
    Write-Host "*** -Apply IS SET: validated guests WILL have configuration changed. Never snapshots/reboots/alters any VM setting." -ForegroundColor Red
} else {
    Write-Host "DRY-RUN MODE (default) - no guest will be changed. Pass -Apply to remediate." -ForegroundColor Yellow
}

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
Write-Host ("Target VMs: {0}" -f $vmNames.Count) -ForegroundColor Gray

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
$CsvName  = if ($Apply) { "SecurityGap_ExtBaseline_Remediation_$RunStamp.csv" } else { "SecurityGap_ExtBaseline_PreCheck_$RunStamp.csv" }
$CsvPath  = Join-Path $OutputPath $CsvName
$LogPath  = Join-Path $OutputPath "SecurityGap_ExtBaseline_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Run started. Build=$ScriptBuild. Script=$PSCommandPath. Apply=$Apply. VMs=$($vmNames.Count). vCenter(s)=$(($connectedServers | ForEach-Object { $_.Name }) -join ', ')."

# =====================================================================================
# 2. In-guest payload  (#__INJECT_PARAMS__ is replaced with literal assignments;
#    it never carries a secret). Same Invoke-Remediation / Add-ResultRow contract
#    and same envelope format as Invoke-RemoteSecurityGapRemediation.ps1.
# =====================================================================================
$PayloadTemplate = @'
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param()

#__INJECT_PARAMS__

$ProgressPreference = 'SilentlyContinue'
$ScriptBuild = 'ext-baseline-remote-r1'

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
    try { $IPAddresses = ([System.Net.Dns]::GetHostAddresses($ComputerName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString } catch { $IPAddresses = 'Unknown' }
}
try { $OSCaption = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Caption } catch { $OSCaption = 'Unknown' }

function Add-ResultRow {
    param([string]$Category, [string]$Item, [string]$Status, [string]$Details = '',
          [bool]$RequiresReboot = $false, [string]$RiskNote = '',
          [string]$DetectedValue = '', [string]$ExpectedValue = '')
    $Results.Add([PSCustomObject]@{
        Category = $Category; Item = $Item; Status = $Status; Details = $Details
        DetectedValue = $DetectedValue; ExpectedValue = $ExpectedValue
        RequiresReboot = $RequiresReboot; RiskNote = $RiskNote
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }) | Out-Null
    if ($RequiresReboot -and $Status -eq 'Applied') { $RebootRequiredItems.Add($Item) | Out-Null }
}

function Invoke-Remediation {
    param(
        [string]$Category, [string]$Item, [string]$Description,
        [scriptblock]$CheckBlock, [scriptblock]$ApplyBlock,
        [bool]$RequiresReboot = $false, [string]$RiskNote = '',
        [scriptblock]$CurrentStateBlock, [string]$ExpectedValue = ''
    )
    $detected = ''
    if ($CurrentStateBlock) { try { $detected = [string](& $CurrentStateBlock) } catch { $detected = "(unreadable: $($_.Exception.Message))" } }
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

$CAT = '7. Extended Baseline (Cybersecurity Review)'
$fatal = $null
try {
    try {
        $IsElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $IsElevated = $null }
    if ($Apply -and $IsElevated -ne $true) {
        throw "Refusing to -Apply inside guest: session is not confirmed elevated (Administrator). No changes made."
    }
    Write-GuestLog "Start. $ComputerName ($IPAddresses). Apply=$Apply. Elevated=$IsElevated"

    # -------------------------------------------------------------------------
    # 1. HVCI - Enabled WITH UEFI lock
    # -------------------------------------------------------------------------
    $hvciPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'
    Invoke-Remediation -Category $CAT -Item 'HVCI - UEFI lock' `
        -Description "Set $hvciPath Locked=1 (HVCI 'Enabled with UEFI lock'; keeps Enabled=1)" `
        -RequiresReboot $true `
        -RiskNote 'UEFI lock makes HVCI hard to disable later without firmware/physical access (intentional). Only effective once HVCI itself (Enabled=1) is on - see category 1 of the main script.' `
        -ExpectedValue 'Locked=1 (and Enabled=1)' `
        -CurrentStateBlock { "$(Get-RegValueString $hvciPath 'Enabled'); $(Get-RegValueString $hvciPath 'Locked')" } `
        -CheckBlock { $v = Get-ItemProperty -Path $hvciPath -ErrorAction SilentlyContinue; $v -and $v.Enabled -eq 1 -and $v.Locked -eq 1 } `
        -ApplyBlock {
            New-Item -Path $hvciPath -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path $hvciPath -Name Enabled -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $hvciPath -Name Locked  -Value 1 -Type DWord -Force
        }

    # -------------------------------------------------------------------------
    # 2. Kernel-mode Hardware-enforced Stack Protection - ENFORCEMENT (not audit)
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # 3. LSASS Protected Process (RunAsPPL) - WITH UEFI lock
    # -------------------------------------------------------------------------
    $lsaPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA'
    Invoke-Remediation -Category $CAT -Item 'LSASS RunAsPPL - UEFI lock' `
        -Description "Set $lsaPath RunAsPPL=1 (LSA runs as a protected process, UEFI-locked). 2 = enabled without lock." `
        -RequiresReboot $true `
        -RiskNote 'Blocks unsigned / non-Microsoft LSA plugins and SSPs (some SSO / smart-card / EDR agents). Confirm all LSA-integrating agents are WHQL-signed first. UEFI lock is hard to revert.' `
        -ExpectedValue 'RunAsPPL=1' `
        -CurrentStateBlock { Get-RegValueString $lsaPath 'RunAsPPL' } `
        -CheckBlock { $v = Get-ItemProperty -Path $lsaPath -Name RunAsPPL -ErrorAction SilentlyContinue; $v -and $v.RunAsPPL -eq 1 } `
        -ApplyBlock {
            Set-ItemProperty -Path $lsaPath -Name RunAsPPL     -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $lsaPath -Name RunAsPPLBoot  -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
        }

    # -------------------------------------------------------------------------
    # 4. PowerShell Script Block Logging (+ invocation logging)
    # -------------------------------------------------------------------------
    $sblPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\PowerShell\ScriptBlockLogging'
    Invoke-Remediation -Category $CAT -Item 'PowerShell Script Block (and Invocation) Logging' `
        -Description "Set $sblPath EnableScriptBlockLogging=1, EnableScriptBlockInvocationLogging=1" `
        -RiskNote 'Invocation logging is verbose (start/stop of every script block) and can generate large volumes of event 4104 - size the PowerShell/Operational log and forwarding accordingly. CS review explicitly requires it enabled.' `
        -ExpectedValue 'EnableScriptBlockLogging=1, EnableScriptBlockInvocationLogging=1' `
        -CurrentStateBlock { "$(Get-RegValueString $sblPath 'EnableScriptBlockLogging'); $(Get-RegValueString $sblPath 'EnableScriptBlockInvocationLogging')" } `
        -CheckBlock { $v = Get-ItemProperty -Path $sblPath -ErrorAction SilentlyContinue; $v -and $v.EnableScriptBlockLogging -eq 1 -and $v.EnableScriptBlockInvocationLogging -eq 1 } `
        -ApplyBlock {
            New-Item -Path $sblPath -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path $sblPath -Name EnableScriptBlockLogging           -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $sblPath -Name EnableScriptBlockInvocationLogging -Value 1 -Type DWord -Force
        }

    # -------------------------------------------------------------------------
    # 5. WinRM - "Allow remote server management through WinRM" / auto config = Disabled
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # 6. WinRM IPv4Filter / IPv6Filter  (only if a range was supplied)
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # 7. Kerberos PKINIT hash algorithms SHA256 / SHA384 / SHA512 = Supported (3)
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # 8. Windows Ink Workspace - Disabled
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # 9. Advanced audit: "Audit Policy Change" = Success and Failure
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # 10. Cached logons count
    # -------------------------------------------------------------------------
    $wlPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'
    Invoke-Remediation -Category $CAT -Item "Cached logons count = $CachedLogonsCount" `
        -Description "Set $wlPath CachedLogonsCount='$CachedLogonsCount' (REG_SZ). Qualys QID 90007; 0 clears it on a workgroup server." `
        -RiskNote 'A value of 0 means no cached domain logons - only matters if this host is ever domain-joined and loses DC connectivity. Fine for workgroup.' `
        -ExpectedValue "CachedLogonsCount=$CachedLogonsCount" `
        -CurrentStateBlock { Get-RegValueString $wlPath 'CachedLogonsCount' } `
        -CheckBlock { $v = Get-ItemProperty -Path $wlPath -Name CachedLogonsCount -ErrorAction SilentlyContinue; $v -and ([string]$v.CachedLogonsCount) -eq ([string]$CachedLogonsCount) } `
        -ApplyBlock { Set-ItemProperty -Path $wlPath -Name CachedLogonsCount -Value ([string]$CachedLogonsCount) -Type String -Force }

    # -------------------------------------------------------------------------
    # 11. Credential Guard policy value - detect a GPO that disables it
    # -------------------------------------------------------------------------
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

    # -------------------------------------------------------------------------
    # Manual / external items (Local Security Policy + patching) - report only.
    # -------------------------------------------------------------------------
    if (-not $SkipManualItems) {
        Add-ResultRow -Category $CAT -Item 'Deny access to this computer from the network (user right)' -Status 'Manual/External Required' `
            -Details 'secedit: [Privilege Rights] SeDenyNetworkLogonRight = *S-1-5-32-546,*S-1-5-113,*S-1-5-114  (Guests, Local account, Local account and member of Administrators group). Export with "secedit /export /cfg secpol.txt", edit that line, apply with "secedit /configure /db secedit.sdb /cfg secpol.txt /areas USER_RIGHTS".' `
            -ExpectedValue 'Guests + Local account + Local account & Administrators' -DetectedValue 'Only BUILTIN\Guests assigned'
        Add-ResultRow -Category $CAT -Item 'Deny logon through Remote Desktop Services (user right)' -Status 'Manual/External Required' `
            -Details 'secedit: [Privilege Rights] SeDenyRemoteInteractiveLogonRight = *S-1-5-32-546,*S-1-5-113  (Guests, Local account). Same export/edit/apply as above.' `
            -ExpectedValue 'Guests + Local account' -DetectedValue 'Only BUILTIN\Guests assigned'
        Add-ResultRow -Category $CAT -Item 'Allow Administrator account lockout' -Status 'Manual/External Required' `
            -Details 'secedit: [System Access] AllowAdministratorLockout = 1, and set an account lockout threshold (e.g. "net accounts /lockoutthreshold:10 /lockoutduration:15 /lockoutwindow:15"). Provide secpol.txt as evidence.' `
            -ExpectedValue 'Enabled (1)' -DetectedValue 'Not Applicable / not configured'
        Add-ResultRow -Category $CAT -Item 'Rename built-in Guest account (QID 105228)' -Status 'Manual/External Required' `
            -Details 'secedit: [System Access] NewGuestName = "<non-obvious name>"  (or: Rename-LocalUser -Name Guest -NewName "<name>"). The Guest account should also stay Disabled.' `
            -ExpectedValue 'Guest renamed' -DetectedValue 'Built-in Guest not renamed'
        Add-ResultRow -Category $CAT -Item 'Unused local accounts' -Status 'Manual/External Required' `
            -Details 'Review and disable/remove unused local accounts flagged by Qualys (e.g. s.l.jganzon.m, s.l.samohamed.m): Disable-LocalUser -Name "<name>" after confirming with the account owner.' `
            -ExpectedValue 'No stale local accounts' -DetectedValue 's.l.jganzon.m, s.l.samohamed.m'
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

} catch {
    $fatal = $_.Exception.Message
    Write-GuestLog "FATAL: $fatal" 'ERROR'
}

$meta = [PSCustomObject]@{
    Hostname = $ComputerName; IPAddress = $IPAddresses; OS = $OSCaption
    Apply = [bool]$Apply; Elevated = $IsElevated; Build = $ScriptBuild
    RebootRequired = @(($RebootRequiredItems | Select-Object -Unique))
    GuestLog = $GuestLog.ToArray(); FatalError = $fatal
}
$envelope = [PSCustomObject]@{ Schema = 'secgap-extbaseline-1'; Meta = $meta; Rows = $Results.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<SECGAP-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-SECGAP-ENVELOPE>>>'
'@

# =====================================================================================
# 3. Admin-side helpers  (same as the other remote scripts)
# =====================================================================================
function ConvertTo-PsBool { param([bool]$Value) if ($Value) { '$true' } else { '$false' } }
function ConvertTo-PsSingleQuoted {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return "''" }
    if ($Value -match "[`r`n]") { throw "Refusing to inject a value containing a newline: '$Value'" }
    "'" + ($Value -replace "'", "''") + "'"
}

function New-Payload {
    param([bool]$EffectiveApply)
    if ($WinRmFilterRange -and $WinRmFilterRange.Count -gt 0) {
        $parts = foreach ($entry in $WinRmFilterRange) {
            $clean = ($entry -replace '[^0-9A-Fa-f\.:/\-]', '')
            if ($clean -ne $entry -or [string]::IsNullOrWhiteSpace($clean)) {
                throw "WinRmFilterRange entry '$entry' contains characters outside [0-9A-Fa-f.:/-] - refused."
            }
            ConvertTo-PsSingleQuoted $clean
        }
        $rangeLiteral = '@(' + ($parts -join ',') + ')'
    } else {
        $rangeLiteral = '@()'
    }
    $inject = @"
`$Apply             = $(ConvertTo-PsBool $EffectiveApply)
`$WinRmFilterRange  = $rangeLiteral
`$CachedLogonsCount = $([int]$CachedLogonsCount)
`$SkipManualItems   = $(ConvertTo-PsBool $SkipManualItems)
`$ConfirmPreference = 'None'
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

# Deliver a multi-KB payload as a FILE (Copy-VMGuestFile), with a small-chunk
# fallback. See Invoke-RemoteSecurityGapRemediation.ps1 for the rationale.
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
        'Skipped (declined at confirm prompt)' { return @{ Status = 'Skipped';                     Action = 'Declined at confirmation' } }
        default                               { return @{ Status = $PayloadStatus;                 Action = '' } }
    }
}

# Same on-screen layout as Invoke-RemoteSecurityGapRemediation.ps1.
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
        if ($row.Category) { Write-Host (" {0} > {1}" -f $row.Category, $row.Item) }
        else               { Write-Host (" {0}" -f $row.Item) }
    }
}

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows  = New-Object System.Collections.Generic.List[object]
$rebootGlobal = New-Object System.Collections.Generic.List[string]
$summary = [ordered]@{ Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0; Remediated = 0; RemediationFailed = 0; Manual = 0; RebootRequired = 0; SkippedFailed = 0 }
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
        $vm = $val.VM; $srv = $val.Server

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
                -Reason 'Guest session returned a non-elevated token; -Apply aborted for this VM to avoid a partial change.'))
            Write-Log "SKIP APPLY '$vmName': guest session not elevated." 'WARN'
            Show-VmResultsLocal -HostLabel "$vmName" -Rows @([pscustomobject]@{ Status = 'Skipped'; Category = ''; Item = 'Guest session not elevated - -Apply aborted for this VM.' })
            continue
        }

        $effectiveApply = [bool]$Apply
        if ($Apply -and -not $PSCmdlet.ShouldProcess("$vmName @ $($srv.Name)", "Apply extended-baseline remediation inside guest via VMware Tools")) {
            $effectiveApply = $false
            Write-Log "'$vmName': -WhatIf / declined -> running DRY-RUN for this VM only." 'INFO'
        }

        $payload = New-Payload -EffectiveApply $effectiveApply
        Write-Log "'$vmName': staging + invoking payload (Apply=$effectiveApply, user='$($GuestCredential.UserName)')." 'INFO'
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $payload -ToolsWaitSecs $ToolsWaitSecs

        $envelope = Read-EnvelopeFromScriptOutput -Output $scriptOutput -StartMarker '<<<SECGAP-ENVELOPE-B64>>>' -EndMarker '<<<END-SECGAP-ENVELOPE>>>'
        if ($null -eq $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($scriptOutput) { ($scriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Validation' -SecurityCheck 'Guest payload result' -DetectedValue '' -ExpectedValue 'Base64 result envelope' `
                -Status 'Unable to Check' -Action 'Skipped - unparseable result' -RequiresReboot $false -RiskNote '' -Reason "Guest payload returned no parseable result. Output start: $snippet"))
            Write-Log "'$vmName': no parseable envelope returned." 'ERROR'
            Show-VmResultsLocal -HostLabel "$vmName" -Rows @([pscustomobject]@{ Status = 'Skipped'; Category = ''; Item = "Guest payload returned no parseable result. Output start: $snippet" })
            continue
        }

        $meta  = $envelope.Meta
        $ghost = if ($meta.Hostname)  { $meta.Hostname }  else { $vm.Guest.HostName }
        $gip   = if ($meta.IPAddress) { $meta.IPAddress } else { ($vm.Guest.IPAddress -join ',') }
        $gos   = if ($meta.OS)        { $meta.OS }        else { $vm.Guest.OSFullName }
        if ($meta.GuestLog) { foreach ($gl in $meta.GuestLog) { Add-Content -LiteralPath $LogPath -Value ("    [guest:$vmName] $gl") } }
        if ($meta.FatalError) { Write-Log "'$vmName': guest payload reported a fatal error: $($meta.FatalError)" 'ERROR' }

        $summary.Checked++
        $f = @{ NonCompliant = $false; Remediated = $false; Failed = $false; Manual = $false; Any = $false }
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

        Show-VmResultsLocal -HostLabel ("{0} ({1})" -f $ghost, $gip) -Rows @($envelope.Rows)

        $vmReboots = @($envelope.Meta.RebootRequired | Where-Object { $_ })
        if ($vmReboots.Count -gt 0) {
            $summary.RebootRequired++
            foreach ($item in $vmReboots) {
                if ($vmNames.Count -gt 1) { $rebootGlobal.Add(("{0} - {1}" -f $vmName, $item)) | Out-Null }
                else                      { $rebootGlobal.Add([string]$item) | Out-Null }
            }
        }
        if ($f.Failed)       { $summary.RemediationFailed++ }
        if ($f.Remediated)   { $summary.Remediated++ }
        if ($f.NonCompliant) { $summary.NonCompliant++ }
        if ($f.Manual)       { $summary.Manual++ }
        if ($f.Any -and -not $f.NonCompliant -and -not $f.Failed) { $summary.Compliant++ }
        Write-Log "'$vmName': done. rows=$(@($envelope.Rows).Count) nonCompliant=$($f.NonCompliant) remediated=$($f.Remediated) failed=$($f.Failed) manual=$($f.Manual)" 'INFO'
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
# 5. Reboot block + central report
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
