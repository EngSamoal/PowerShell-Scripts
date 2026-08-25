#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Server Security Audit - READ-ONLY, runs entirely locally on the server being checked.
    No PowerCLI, no vCenter, no VMware Tools, no WinRM to another host. Makes zero configuration
    changes.

.DESCRIPTION
    Copy this single file onto EACH target server (RDP clipboard/file copy, a network share, USB,
    or paste its contents into a new .ps1 in an elevated PowerShell/ISE session on that server) and
    run it there, in an elevated PowerShell session, once per server:

        AQ-UFM-APP01, AQ-UFM-DB01, SF-UFM-APP01, SF-UFM-DB01

    Each run produces its own CSV/log (named with that computer's hostname), covering:
      1. VBS & Credential Protection (Credential Guard, HVCI/Memory Integrity, Machine Identity
         Isolation)
      2. Microsoft Defender (Attack Surface Reduction rules, real-time/cloud protection, tamper
         protection)
      3. Authentication Security (WDigest, Kerberos supported encryption types)
      4. Remote Management (WinRM service, listeners, hardening settings)
      5. SMB & RPC Security (SMB1 auditing, SMB signing/encryption, Printer RPC / PrintNightmare
         mitigations, Point and Print restrictions)
      6. Windows Server 2025 Security Baseline indicators (SmartScreen, IE ActiveX/Scripting/
         Protected Mode zone policy, Schannel SSL/TLS protocol state, legacy IE feature state)

    Every check is a read-only Get-*/Test-Path/registry-read operation - nothing here Sets/News/
    Removes/Starts/Stops/Restarts/Enables/Disables anything. No service is touched, no registry
    value is written, no GPO is modified, and no reboot is triggered by this script under any code
    path. If you want one combined report afterward, merge the four per-server CSVs yourself, e.g.:
        Get-ChildItem .\SecurityAudit_Reports\SecurityAudit_*.csv | Import-Csv | Export-Csv Combined.csv -NoTypeInformation

    Status values used throughout:
      Enabled / Disabled  - the setting was read directly from this machine and its state
                             confirmed.
      Not Configured      - no explicit value is present; the OS/domain default is in effect.
      Not Applicable      - the check does not apply here (feature/cmdlet absent on this OS
                             build/SKU, etc.).
      Unable to Verify     - the check could not be completed (cmdlet missing, access denied,
                             not running elevated, etc.) - this is NOT the same as
                             Disabled/non-compliant, and is never treated as a negative finding.

    Several checks are explicitly flagged RequiresAdditionalEvidence = $true in the output. These
    are registry-level indicators only (most of the "2025 Security Baseline Indicators" category,
    plus the Kerberos SHA-2/next-gen-crypto note). A registry value can be overridden by an
    unlinked/stale GPO, WMI filtering, or local vs. domain policy precedence - these checks are NOT
    proof the corresponding GPO is applied. Confirm with effective GPO/RSOP evidence
    (gpresult /h <file>.html or Microsoft's Policy Analyzer against the official baseline GPO
    backup) before treating them as authoritative. "Machine Identity Isolation" is reported as
    Unable to Verify on every run, on purpose: no single well-documented registry/WMI indicator for
    that specific (newer) capability was identified at the time this script was written, and
    guessing at one risks reporting a false Enabled/Disabled result.

.PARAMETER OutputPath
    Folder for the CSV/XLSX + log file. Created if it doesn't exist. Defaults to
    .\SecurityAudit_Reports next to this script.

.PARAMETER ExportExcel
    Also export an .xlsx (in addition to the CSV, which is always written) if the ImportExcel
    module is installed. Silently falls back to CSV-only with a log note if the module is missing.

.EXAMPLE
    # Run locally, elevated, on AQ-UFM-APP01:
    .\Get-WindowsServerSecurityAudit-Local.ps1
    # Repeat on AQ-UFM-DB01, SF-UFM-APP01, SF-UFM-DB01.

.NOTES
    This is a standalone counterpart to Get-WindowsServerSecurityAudit.ps1 (the PowerCLI/
    Invoke-VMScript version). It exists because Invoke-VMScript's guest-operations API has a
    practical size ceiling for embedded script text on some vSphere/Tools builds, which surfaced
    as a misleading "not enough permissions" error for the full ~30-40KB check payload used there -
    even with a confirmed local administrator account and a pre-flight probe proving basic guest
    execution worked. Running this file directly on each server sidesteps that transport entirely:
    there is no vCenter, no VMware Tools, no guest-operations API involved - just PowerShell
    reading its own local registry and WMI/CIM classes.
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot 'SecurityAudit_Reports'),
    [switch]$ExportExcel
)

$ProgressPreference = 'SilentlyContinue'
$ScriptBuild = '2026-08-17-01-local-standalone'
Write-Host "Get-WindowsServerSecurityAudit-Local.ps1 - build $ScriptBuild" -ForegroundColor Magenta
Write-Host "READ-ONLY ASSESSMENT - no configuration changes, restarts, or GPO/registry/service/Defender/WinRM writes are made by this script." -ForegroundColor Yellow

try {
    $IsElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
} catch {
    $IsElevated = $null
}
if ($IsElevated -eq $false) {
    Write-Host "WARNING: this PowerShell session is not running elevated (Run as Administrator). Several checks (Defender, Device Guard/VBS, some HKLM values) require elevation to read accurately and will report 'Unable to Verify' rather than a real Disabled/absent result if access is denied. Close this session and re-run from an elevated PowerShell/ISE window for a complete report." -ForegroundColor Yellow
} elseif ($null -eq $IsElevated) {
    Write-Host "WARNING: could not determine whether this session is elevated. If checks below unexpectedly show 'Unable to Verify', re-run from an elevated (Run as Administrator) PowerShell/ISE window." -ForegroundColor Yellow
}

$ComputerName = $env:COMPUTERNAME
try {
    $IPAddresses = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
        Where-Object { $_.InterfaceAlias -notmatch 'Loopback' -and $_.IPAddress -notmatch '^169\.254\.' }).IPAddress -join ', '
} catch {
    $IPAddresses = $null
}
if ([string]::IsNullOrWhiteSpace($IPAddresses)) {
    try {
        $IPAddresses = ([System.Net.Dns]::GetHostAddresses($ComputerName) |
            Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
            Select-Object -First 1).IPAddressToString
    } catch {
        $IPAddresses = 'Unknown'
    }
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogPath  = Join-Path $OutputPath "SecurityAudit_${ComputerName}_$RunStamp.log"
$CsvPath  = Join-Path $OutputPath "SecurityAudit_${ComputerName}_$RunStamp.csv"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

$Results = New-Object System.Collections.Generic.List[object]

function New-Check {
    param(
        [string]$Category,
        [string]$Check,
        [string]$Status,
        [string]$Evidence = "",
        [string]$Interpretation = "",
        [bool]$RequiresAdditionalEvidence = $false,
        [string]$EvidenceGuidance = ""
    )
    $Results.Add([PSCustomObject]@{
        ComputerName                = $ComputerName
        IPAddress                   = $IPAddresses
        Category                    = $Category
        Check                       = $Check
        Status                      = $Status
        Evidence                    = $Evidence
        Interpretation              = $Interpretation
        RequiresAdditionalEvidence  = $RequiresAdditionalEvidence
        EvidenceGuidance            = $EvidenceGuidance
        Timestamp                   = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }) | Out-Null
}

Write-Log "Run started. Build $ScriptBuild. Computer: $ComputerName ($IPAddresses). Elevated: $IsElevated"

Write-Host ""
Write-Host "Status legend:" -ForegroundColor Cyan
Write-Host "  Enabled / Disabled          - setting was read directly from this machine and confirmed."
Write-Host "  Not Configured              - no explicit value present; OS/domain default applies."
Write-Host "  Not Applicable              - check does not apply (feature absent, etc.)."
Write-Host "  Unable to Verify            - check could not be completed - NOT the same as Disabled."
Write-Host "  RequiresAdditionalEvidence  - indicator only; confirm with effective GPO/RSOP before treating as authoritative."
Write-Host ""

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $p = @{
        Category       = "Host Info"
        Check          = "Operating System"
        Status         = "Info"
        Evidence       = "$($os.Caption) Build $($os.BuildNumber)"
        Interpretation = "Guest operating system reported by WMI - for reference only."
    }
    New-Check @p
} catch {
    New-Check -Category "Host Info" -Check "Operating System" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

# ===========================================================================
# 1. VBS & Credential Protection
# ===========================================================================
try {
    $dg = Get-CimInstance -Namespace "root\Microsoft\Windows\DeviceGuard" -ClassName Win32_DeviceGuard -ErrorAction Stop

    $vbsStatus = $dg.VirtualizationBasedSecurityStatus
    $vbsText = switch ($vbsStatus) {
        2 { "Enabled (Running)" }
        1 { "Configured but NOT running (check hardware/firmware/Hyper-V requirements)" }
        0 { "Not Enabled" }
        default { "Unknown ($vbsStatus)" }
    }
    $p = @{
        Category       = "VBS & Credential Protection"
        Check          = "Virtualization Based Security (VBS)"
        Status         = $vbsText
        Evidence       = "VirtualizationBasedSecurityStatus=$vbsStatus"
        Interpretation = "0=Not enabled, 1=Enabled in policy but not actually running (hardware/firmware may not support it), 2=Enabled and running."
    }
    New-Check @p

    $configured = @($dg.SecurityServicesConfigured)
    $running    = @($dg.SecurityServicesRunning)

    $cgConfigured = $configured -contains 1
    $cgRunning    = $running -contains 1
    $cgStatus = if ($cgRunning) { "Enabled (Running)" } elseif ($cgConfigured) { "Configured but NOT Running" } else { "Not Configured" }
    $p = @{
        Category       = "VBS & Credential Protection"
        Check          = "Credential Guard"
        Status         = $cgStatus
        Evidence       = "SecurityServicesConfigured=[$($configured -join ',')] SecurityServicesRunning=[$($running -join ',')]"
        Interpretation = "Isolates LSASS secrets in a VBS container to block credential-theft tools (e.g. Mimikatz-style attacks). 'Configured but not running' usually means a pending reboot or unmet hardware/firmware requirements."
    }
    New-Check @p

    $hvciConfigured = $configured -contains 2
    $hvciRunning    = $running -contains 2
    $hvciStatus = if ($hvciRunning) { "Enabled (Running)" } elseif ($hvciConfigured) { "Configured but NOT Running" } else { "Not Configured" }
    $p = @{
        Category       = "VBS & Credential Protection"
        Check          = "HVCI / Memory Integrity"
        Status         = $hvciStatus
        Evidence       = "SecurityServicesConfigured=[$($configured -join ',')] SecurityServicesRunning=[$($running -join ',')]"
        Interpretation = "Hypervisor-protected Code Integrity ensures only signed, trusted code runs in kernel mode. 'Configured but not running' usually means a pending reboot or unmet hardware/firmware requirements."
    }
    New-Check @p
} catch {
    New-Check -Category "VBS & Credential Protection" -Check "Credential Guard" -Status "Unable to Verify" -Evidence $_.Exception.Message -Interpretation "Could not query Win32_DeviceGuard (root\Microsoft\Windows\DeviceGuard) - class may be unavailable on this OS build/SKU, or this session isn't elevated."
    New-Check -Category "VBS & Credential Protection" -Check "HVCI / Memory Integrity" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

New-Check -Category "VBS & Credential Protection" -Check "Machine Identity Isolation" -Status "Unable to Verify" `
    -Evidence "No single well-documented registry/WMI indicator identified for this setting at time of writing." `
    -Interpretation "This is a newer Windows security capability; this script deliberately does not guess at an unverified indicator rather than report a false Enabled/Disabled result." `
    -RequiresAdditionalEvidence $true `
    -EvidenceGuidance "Confirm current implementation details against up-to-date Microsoft Windows Server 2025 documentation, then extend this script once confirmed."

# ===========================================================================
# 2. Microsoft Defender
# ===========================================================================
$mpAvailable = [bool](Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)
if (-not $mpAvailable) {
    New-Check -Category "Microsoft Defender" -Check "Defender Cmdlets Available" -Status "Not Applicable" `
        -Evidence "Get-MpComputerStatus / Get-MpPreference not found." `
        -Interpretation "Defender PowerShell module is not present - this can mean Defender is disabled/uninstalled, or a third-party AV has replaced it. Verify manually via the Windows Security app or the AV vendor's console."
} else {
    try {
        $mpStatus = Get-MpComputerStatus -ErrorAction Stop
        $p = @{
            Category       = "Microsoft Defender"
            Check          = "Real-Time Protection"
            Status         = $(if ($mpStatus.RealTimeProtectionEnabled) { "Enabled" } else { "Disabled" })
            Evidence       = "RealTimeProtectionEnabled=$($mpStatus.RealTimeProtectionEnabled); AMServiceEnabled=$($mpStatus.AMServiceEnabled); AntivirusEnabled=$($mpStatus.AntivirusEnabled)"
            Interpretation = "Whether Defender's on-access/real-time scanning engine is currently active."
        }
        New-Check @p
        $p = @{
            Category       = "Microsoft Defender"
            Check          = "Tamper Protection"
            Status         = $(if ($mpStatus.IsTamperProtected) { "Enabled" } else { "Disabled" })
            Evidence       = "IsTamperProtected=$($mpStatus.IsTamperProtected)"
            Interpretation = "Blocks unauthorized changes to Defender security settings, including via admin-level scripts or direct registry edits."
        }
        New-Check @p
    } catch {
        New-Check -Category "Microsoft Defender" -Check "Real-Time Protection" -Status "Unable to Verify" -Evidence $_.Exception.Message
    }

    try {
        $mpPref = Get-MpPreference -ErrorAction Stop

        $asrIds     = @($mpPref.AttackSurfaceReductionRules_Ids)
        $asrActions = @($mpPref.AttackSurfaceReductionRules_Actions)
        if ($asrIds.Count -eq 0) {
            New-Check -Category "Microsoft Defender" -Check "Attack Surface Reduction (ASR) Rules" -Status "Not Configured" `
                -Evidence "No ASR rule IDs present in Get-MpPreference." `
                -Interpretation "No ASR rules are configured locally or via policy on this host."
        } else {
            $pairs = for ($i = 0; $i -lt $asrIds.Count; $i++) {
                $actionText = switch ([int]$asrActions[$i]) {
                    0 { "NotConfigured" }
                    1 { "Block" }
                    2 { "Audit" }
                    6 { "Warn" }
                    default { "Unknown($($asrActions[$i]))" }
                }
                "$($asrIds[$i])=$actionText"
            }
            $blockCount = @($asrActions | Where-Object { [int]$_ -eq 1 }).Count
            $auditCount = @($asrActions | Where-Object { [int]$_ -eq 2 }).Count
            $overall = if ($blockCount -gt 0) { "Enabled (Enforced) - $blockCount rule(s) Block, $auditCount Audit" }
                       elseif ($auditCount -gt 0) { "Audit Mode Only - $auditCount rule(s)" }
                       else { "Not Configured" }
            $p = @{
                Category       = "Microsoft Defender"
                Check          = "Attack Surface Reduction (ASR) Rules"
                Status         = $overall
                Evidence       = ($pairs -join "; ")
                Interpretation = "ASR rules block common malware/exploit behaviors. Block=enforced, Audit=logged only (not enforced), Warn=user can bypass. Reflects Defender's merged effective policy."
            }
            New-Check @p
        }

        $mapsText = switch ($mpPref.MAPSReporting) {
            0 { "Disabled" }
            1 { "Basic" }
            2 { "Advanced" }
            default { "Unknown($($mpPref.MAPSReporting))" }
        }
        $p = @{
            Category       = "Microsoft Defender"
            Check          = "Cloud-Delivered Protection (MAPS Reporting)"
            Status         = $mapsText
            Evidence       = "MAPSReporting=$($mpPref.MAPSReporting); SubmitSamplesConsent=$($mpPref.SubmitSamplesConsent); CloudBlockLevel=$($mpPref.CloudBlockLevel)"
            Interpretation = "Controls whether Defender submits telemetry/sample data to Microsoft's cloud for enhanced detection."
        }
        New-Check @p
    } catch {
        New-Check -Category "Microsoft Defender" -Check "Attack Surface Reduction (ASR) Rules" -Status "Unable to Verify" -Evidence $_.Exception.Message
    }
}

# ===========================================================================
# 3. Authentication Security
# ===========================================================================
try {
    $wdigestPath = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest"
    $wdigestVal = $null
    if (Test-Path $wdigestPath) {
        $prop = Get-ItemProperty -Path $wdigestPath -Name UseLogonCredential -ErrorAction SilentlyContinue
        if ($prop) { $wdigestVal = $prop.UseLogonCredential }
    }
    if ($null -eq $wdigestVal) {
        New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Not Configured" `
            -Evidence "Registry value absent." `
            -Interpretation "No explicit setting present. Windows 8.1/Server 2012 R2+ ships with WDigest credential caching off by default, but an explicit Disabled (0) value is recommended so the posture is enforced rather than relying on the OS default."
    } elseif ($wdigestVal -eq 1) {
        New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Enabled" `
            -Evidence "UseLogonCredential=1" `
            -Interpretation "INSECURE: WDigest caches reversible/plaintext-equivalent credentials in LSASS memory, a common credential-theft target. Should be Disabled unless a specific legacy application requires it."
    } else {
        New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Disabled" -Evidence "UseLogonCredential=$wdigestVal"
    }
} catch {
    New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

try {
    $kerbPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters"
    $encVal = $null
    if (Test-Path $kerbPath) {
        $prop = Get-ItemProperty -Path $kerbPath -Name SupportedEncryptionTypes -ErrorAction SilentlyContinue
        if ($prop) { $encVal = $prop.SupportedEncryptionTypes }
    }
    if ($null -eq $encVal) {
        New-Check -Category "Authentication Security" -Check "Kerberos Supported Encryption Types" -Status "Not Configured" `
            -Evidence "SupportedEncryptionTypes registry value absent - OS/domain default negotiation applies." `
            -Interpretation "No explicit local policy restricting Kerberos encryption types." `
            -RequiresAdditionalEvidence $true `
            -EvidenceGuidance "Confirm effective value via 'Network security: Configure encryption types allowed for Kerberos' in effective GPO (gpresult /h) - this can be domain-policy driven without a local registry value."
    } else {
        $bits = [int]$encVal
        $types = @()
        if ($bits -band 0x1)  { $types += "DES-CBC-CRC (legacy, weak)" }
        if ($bits -band 0x2)  { $types += "DES-CBC-MD5 (legacy, weak)" }
        if ($bits -band 0x4)  { $types += "RC4-HMAC-MD5 (legacy)" }
        if ($bits -band 0x8)  { $types += "AES128-CTS-HMAC-SHA1-96" }
        if ($bits -band 0x10) { $types += "AES256-CTS-HMAC-SHA1-96" }
        New-Check -Category "Authentication Security" -Check "Kerberos Supported Encryption Types" -Status "Configured" `
            -Evidence "SupportedEncryptionTypes=0x$($bits.ToString('X')) -> $($types -join ', ')" `
            -Interpretation "This value governs the classic AES128/AES256-SHA1 (and legacy RC4/DES) Kerberos suites. Microsoft's newer AES-SHA2 'next-generation crypto' suites (the SHA256/SHA384/SHA512-based Kerberos support referenced in the Windows Server 2025 baseline) were not confirmed to be controlled by this same value as of this script's authoring - do NOT treat this check as proof SHA256/384/512 Kerberos support is enabled or disabled." `
            -RequiresAdditionalEvidence $true `
            -EvidenceGuidance "Confirm SHA-2 Kerberos suite support/enforcement via current Microsoft KDC/Windows Server 2025 documentation, 'klist' on an authenticated session, or effective GPO/RSOP - do not rely on this registry check alone for that specific claim."
    }
} catch {
    New-Check -Category "Authentication Security" -Check "Kerberos Supported Encryption Types" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

# ===========================================================================
# 4. Remote Management (WinRM)
# ===========================================================================
try {
    $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'" -ErrorAction Stop
    New-Check -Category "Remote Management" -Check "WinRM Service" -Status "$($svc.State) (StartMode=$($svc.StartMode))" `
        -Evidence "State=$($svc.State); StartMode=$($svc.StartMode)" `
        -Interpretation "Current run state and boot-time start mode of the WinRM service, queried locally and read-only - no service control action performed."
} catch {
    New-Check -Category "Remote Management" -Check "WinRM Service" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

try {
    $listenerRoot = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Listener"
    if (Test-Path $listenerRoot) {
        $listenerKeys = Get-ChildItem -Path $listenerRoot -ErrorAction SilentlyContinue
        if (-not $listenerKeys -or $listenerKeys.Count -eq 0) {
            New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status "Not Configured" -Evidence "No listener subkeys under $listenerRoot"
        } else {
            $listenerInfo = foreach ($lk in $listenerKeys) {
                $lp = Get-ItemProperty -Path $lk.PSPath -ErrorAction SilentlyContinue
                "Transport=$($lp.Transport) Port=$($lp.Port) Enabled=$($lp.Enabled) ListeningOn=$($lp.ListeningOn -join '/')"
            }
            $httpCount = @($listenerInfo | Where-Object { $_ -match 'Transport=HTTP ' }).Count
            $status = if ($httpCount -gt 0) { "Configured - includes unencrypted HTTP listener(s)" } else { "Configured - HTTPS only" }
            New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status $status `
                -Evidence ($listenerInfo -join " | ") `
                -Interpretation "Read directly from the registry (does not require the WinRM service to be running). An HTTP (unencrypted) listener is generally a hardening gap versus HTTPS-only."
        }
    } else {
        New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status "Not Configured" -Evidence "$listenerRoot not present."
    }
} catch {
    New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

try {
    $svcCfgPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service"
    if (Test-Path $svcCfgPath) {
        $cfg = Get-ItemProperty -Path $svcCfgPath -ErrorAction SilentlyContinue
        New-Check -Category "Remote Management" -Check "WinRM Hardening (Unencrypted/Basic Auth)" -Status "Configured" `
            -Evidence "AllowUnencrypted=$($cfg.AllowUnencrypted); auth_Basic=$($cfg.auth_Basic); auth_Kerberos=$($cfg.auth_Kerberos); auth_Negotiate=$($cfg.auth_Negotiate); auth_CredSSP=$($cfg.auth_CredSSP)" `
            -Interpretation "0/absent generally means the more secure default; 1 for AllowUnencrypted or auth_Basic indicates a weaker/legacy configuration that should be reviewed."
    } else {
        New-Check -Category "Remote Management" -Check "WinRM Hardening (Unencrypted/Basic Auth)" -Status "Not Configured" -Evidence "$svcCfgPath not present (WinRM likely never configured)."
    }
} catch {
    New-Check -Category "Remote Management" -Check "WinRM Hardening (Unencrypted/Basic Auth)" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

# ===========================================================================
# 5. SMB & RPC Security
# ===========================================================================
try {
    if (Get-Command Get-SmbServerConfiguration -ErrorAction SilentlyContinue) {
        $smb = Get-SmbServerConfiguration -ErrorAction Stop
        $auditSmb1 = if ($smb.PSObject.Properties.Name -contains 'AuditSmb1Access') { $smb.AuditSmb1Access } else { $null }
        $auditStatus = if ($null -eq $auditSmb1) { "Not Applicable (property not present on this OS build)" } elseif ($auditSmb1) { "Enabled" } else { "Disabled" }
        New-Check -Category "SMB & RPC Security" -Check "SMB1 Access Auditing" -Status $auditStatus `
            -Evidence "AuditSmb1Access=$auditSmb1; EnableSMB1Protocol=$($smb.EnableSMB1Protocol)" `
            -Interpretation "AuditSmb1Access logs any client still attempting legacy SMB1 connections, to support safely disabling SMB1. Requires Windows Server 2022+ for this property."
        New-Check -Category "SMB & RPC Security" -Check "SMB Encryption / Signing" -Status "Info" `
            -Evidence "EncryptData=$($smb.EncryptData); RejectUnencryptedAccess=$($smb.RejectUnencryptedAccess); RequireSecuritySignature=$($smb.RequireSecuritySignature); EnableSecuritySignature=$($smb.EnableSecuritySignature)" `
            -Interpretation "Current SMB server encryption and signing enforcement settings."
    } else {
        New-Check -Category "SMB & RPC Security" -Check "SMB1 Access Auditing" -Status "Unable to Verify" -Evidence "Get-SmbServerConfiguration cmdlet not available (SMB module missing)."
    }
} catch {
    New-Check -Category "SMB & RPC Security" -Check "SMB1 Access Auditing" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

try {
    $printPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print"
    $rpcVal = $null
    if (Test-Path $printPath) {
        $prop = Get-ItemProperty -Path $printPath -Name RpcAuthnLevelPrivacyEnabled -ErrorAction SilentlyContinue
        if ($prop) { $rpcVal = $prop.RpcAuthnLevelPrivacyEnabled }
    }
    $rpcStatus = if ($null -eq $rpcVal) { "Not Configured (default applies - confirm this OS build's default)" } elseif ($rpcVal -eq 1) { "Enabled (Hardened - RPC packet privacy required)" } else { "Disabled" }
    New-Check -Category "SMB & RPC Security" -Check "Printer RPC Packet Privacy (PrintNightmare mitigation)" -Status $rpcStatus `
        -Evidence "RpcAuthnLevelPrivacyEnabled=$rpcVal" `
        -Interpretation "Requires RPC connections to the Print Spooler to use packet privacy (encryption+signing) - part of the CVE-2021-34527 (PrintNightmare) mitigation set."
} catch {
    New-Check -Category "SMB & RPC Security" -Check "Printer RPC Packet Privacy (PrintNightmare mitigation)" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

try {
    $papPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint"
    if (Test-Path $papPath) {
        $pap = Get-ItemProperty -Path $papPath -ErrorAction SilentlyContinue
        New-Check -Category "SMB & RPC Security" -Check "Point and Print Restrictions" -Status "Configured" `
            -Evidence "NoWarningNoElevationOnInstall=$($pap.NoWarningNoElevationOnInstall); UpdatePromptSettings=$($pap.UpdatePromptSettings); RestrictDriverInstallationToAdministrators=$($pap.RestrictDriverInstallationToAdministrators)" `
            -Interpretation "RestrictDriverInstallationToAdministrators=1 is the key PrintNightmare hardening setting (blocks non-admins from installing print drivers). NoWarningNoElevationOnInstall=1 is the opposite (a weakening setting) and should be 0/absent."
    } else {
        New-Check -Category "SMB & RPC Security" -Check "Point and Print Restrictions" -Status "Not Configured" -Evidence "$papPath not present - policy not applied, OS defaults apply."
    }
} catch {
    New-Check -Category "SMB & RPC Security" -Check "Point and Print Restrictions" -Status "Unable to Verify" -Evidence $_.Exception.Message
}

# ===========================================================================
# 6. Windows Server 2025 Security Baseline (indicator checks only)
# ===========================================================================
$baselineNote = "Proxy indicator only - does NOT by itself confirm the Microsoft Windows Server 2025 Security Baseline GPO is applied. Confirm with 'gpresult /h <file>.html', RSOP, or Microsoft Policy Analyzer against the official baseline GPO backup."

try {
    $ssPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
    $ssVal = $null
    if (Test-Path $ssPath) {
        $prop = Get-ItemProperty -Path $ssPath -Name EnableSmartScreen -ErrorAction SilentlyContinue
        if ($prop) { $ssVal = $prop.EnableSmartScreen }
    }
    $ssStatus = if ($null -eq $ssVal) { "Not Configured" } elseif ($ssVal -eq 0) { "Disabled" } else { "Enabled" }
    New-Check -Category "2025 Security Baseline Indicators" -Check "SmartScreen (Explorer policy)" -Status $ssStatus `
        -Evidence "EnableSmartScreen=$ssVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true `
        -EvidenceGuidance "Compare against baseline GPO backup / gpresult."
} catch {
    New-Check -Category "2025 Security Baseline Indicators" -Check "SmartScreen (Explorer policy)" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true
}

try {
    $zonePath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
    if (Test-Path $zonePath) {
        $zone = Get-ItemProperty -Path $zonePath -ErrorAction SilentlyContinue
        $activeXVal   = $zone.'1200'
        $scriptingVal = $zone.'1400'
        $protectedVal = $zone.'2500'
        $decode = { param($v) if ($null -eq $v) { "Not Configured" } elseif ($v -eq 0) { "Enable" } elseif ($v -eq 1) { "Prompt" } elseif ($v -eq 3) { "Disable" } else { "Value=$v" } }
        New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - Run ActiveX Controls" -Status (& $decode $activeXVal) -Evidence "1200=$activeXVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true
        New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - Active Scripting" -Status (& $decode $scriptingVal) -Evidence "1400=$scriptingVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true
        $protStatus = if ($null -eq $protectedVal) { "Not Configured" } elseif ($protectedVal -eq 0) { "Enabled (Protected Mode ON)" } else { "Disabled (Protected Mode OFF)" }
        New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - Protected Mode" -Status $protStatus -Evidence "2500=$protectedVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true
    } else {
        New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - ActiveX/Scripting/Protected Mode" -Status "Not Configured" -Evidence "$zonePath not present." -Interpretation $baselineNote -RequiresAdditionalEvidence $true
    }
} catch {
    New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - ActiveX/Scripting/Protected Mode" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true
}

try {
    $schannelBase = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols"
    $protoList = @('SSL 2.0','SSL 3.0','TLS 1.0','TLS 1.1','TLS 1.2','TLS 1.3')
    $protoEvidence = foreach ($proto in $protoList) {
        $serverKey = Join-Path (Join-Path $schannelBase $proto) 'Server'
        if (Test-Path $serverKey) {
            $sp = Get-ItemProperty -Path $serverKey -ErrorAction SilentlyContinue
            $enabled = if ($sp.PSObject.Properties.Name -contains 'Enabled') { $sp.Enabled } else { $null }
            $disabledByDefault = if ($sp.PSObject.Properties.Name -contains 'DisabledByDefault') { $sp.DisabledByDefault } else { $null }
            "$proto[Enabled=$enabled;DisabledByDefault=$disabledByDefault]"
        } else {
            "$proto[Not Configured - OS default]"
        }
    }
    $status = if ($protoEvidence -match 'SSL [23]\.0\[Enabled=1') { "Weak protocol explicitly enabled (SSL 2.0/3.0)" }
              elseif ($protoEvidence -match 'TLS 1\.[01]\[Enabled=1') { "Legacy TLS 1.0/1.1 explicitly enabled" }
              else { "No legacy SSL/TLS explicitly enabled via registry (OS default in effect where not configured)" }
    New-Check -Category "2025 Security Baseline Indicators" -Check "SSL/TLS Protocol Configuration (Schannel)" -Status $status `
        -Evidence ($protoEvidence -join " | ") -Interpretation $baselineNote -RequiresAdditionalEvidence $true `
        -EvidenceGuidance "Registry-absent entries fall back to the OS-default protocol set for this build, which varies by Windows Server version/patch level - confirm actual negotiated protocols with a TLS scan tool if certainty is required."
} catch {
    New-Check -Category "2025 Security Baseline Indicators" -Check "SSL/TLS Protocol Configuration (Schannel)" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true
}

try {
    if (Get-Command Get-WindowsOptionalFeature -ErrorAction SilentlyContinue) {
        $ieFeature = Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction Stop
        New-Check -Category "2025 Security Baseline Indicators" -Check "Legacy Internet Explorer 11 Feature State" -Status "$($ieFeature.State)" `
            -Evidence "FeatureName=Internet-Explorer-Optional-amd64; State=$($ieFeature.State)" -Interpretation $baselineNote -RequiresAdditionalEvidence $true
    } else {
        New-Check -Category "2025 Security Baseline Indicators" -Check "Legacy Internet Explorer 11 Feature State" -Status "Unable to Verify" `
            -Evidence "Get-WindowsOptionalFeature not available (e.g. Server Core or restricted session)." -RequiresAdditionalEvidence $true
    }
} catch {
    New-Check -Category "2025 Security Baseline Indicators" -Check "Legacy Internet Explorer 11 Feature State" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true
}

New-Check -Category "2025 Security Baseline Indicators" -Check "Overall Baseline Applied?" -Status "Not Determinable From Registry Alone" `
    -Evidence "$($Results.Count) individual indicator checks captured above." `
    -Interpretation "The full Windows Server 2025 Security Baseline covers hundreds of settings across many areas. The indicators above are a representative subset only." `
    -RequiresAdditionalEvidence $true `
    -EvidenceGuidance "Run 'gpresult /h report.html' on the server, or use Microsoft's free Policy Analyzer tool to diff effective local policy against the official baseline GPO backup, for an authoritative answer."

# ---------------------------------------------------------------------------------------------
# On-screen summary + export
# ---------------------------------------------------------------------------------------------
$quickChecks = @(
    'Credential Guard','HVCI / Memory Integrity','Machine Identity Isolation',
    'Attack Surface Reduction (ASR) Rules','Real-Time Protection','Tamper Protection',
    'WDigest (UseLogonCredential)','Kerberos Supported Encryption Types',
    'WinRM Service','WinRM Listeners',
    'SMB1 Access Auditing','Printer RPC Packet Privacy (PrintNightmare mitigation)',
    'Overall Baseline Applied?'
)

Write-Host ""
Write-Host "=== $ComputerName ($IPAddresses) ===" -ForegroundColor Green
foreach ($qc in $quickChecks) {
    $row = $Results | Where-Object { $_.Check -eq $qc } | Select-Object -First 1
    if ($row) {
        Write-Host ("  {0,-55} : {1}" -f $qc, $row.Status)
    }
}

$Results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV exported: $CsvPath"

if ($ExportExcel) {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        $xlsxPath = Join-Path $OutputPath "SecurityAudit_${ComputerName}_$RunStamp.xlsx"
        try {
            Import-Module ImportExcel -ErrorAction Stop
            $Results | Export-Excel -Path $xlsxPath -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName 'SecurityAudit' -ErrorAction Stop
            Write-Log "Excel exported: $xlsxPath"
        } catch {
            Write-Log "Excel export failed ($($_.Exception.Message)) - CSV is still available at $CsvPath" 'WARN'
        }
    } else {
        Write-Log "-ExportExcel requested but the ImportExcel module is not installed. Install with 'Install-Module ImportExcel' or use the CSV output at $CsvPath." 'WARN'
    }
}

Write-Log "Run complete. $($Results.Count) total check rows for $ComputerName."
Write-Host ""
Write-Host "Results: $CsvPath"
Write-Host "Log:     $LogPath"
