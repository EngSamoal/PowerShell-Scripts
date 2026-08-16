#Requires -Version 5.1
<#
.SYNOPSIS
    Windows Server Security Audit - READ-ONLY status/configuration check via VMware Tools /
    Invoke-VMScript. Makes zero configuration changes.

.DESCRIPTION
    Reports current guest-OS security configuration for a list of VMs, covering:
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

    Every guest-OS check runs entirely through VMware Tools via Invoke-VMScript - WinRM/PowerShell
    Remoting is never used to reach the guest. Every cmdlet used inside the guest script is a
    read-only Get-*/Test-Path/registry-read operation; nothing is Set, New, Remove, Start, Stop,
    Restart, or Enable/Disable. No service is touched, no registry value is written, no GPO is
    modified, and no reboot is triggered by this script under any code path.

    Status values used throughout:
      Enabled / Disabled  - the setting was read directly from the guest and its state confirmed.
      Not Configured      - no explicit value is present; the OS/domain default is in effect.
      Not Applicable      - the check does not apply here (VM powered off, feature/cmdlet absent
                             on this OS build/SKU, etc.).
      Unable to Verify     - the check could not be completed (Tools not running, guest auth
                             failure, cmdlet missing, access denied, etc.) - this is NOT the same
                             as Disabled/non-compliant, and is never treated as a negative finding.

    Several checks are explicitly flagged RequiresAdditionalEvidence = $true in the output. These
    are registry-level indicators only (most of the "2025 Security Baseline Indicators" category,
    plus the Kerberos SHA-2/next-gen-crypto note). A registry value can be overridden by an
    unlinked/stale GPO, WMI filtering, or local vs. domain policy precedence - these checks are NOT
    proof the corresponding GPO is applied. Confirm with effective GPO/RSOP evidence
    (gpresult /h <file>.html or Microsoft's Policy Analyzer against the official baseline GPO
    backup) before treating them as authoritative. "Machine Identity Isolation" is reported as
    Unable to Verify on every run, on purpose: no single well-documented registry/WMI indicator for
    that specific (newer) capability was identified at the time this script was written, and
    guessing at one risks reporting a false Enabled/Disabled result. Extend the relevant check once
    Microsoft publishes a confirmed indicator, rather than trusting a guess.

.PARAMETER Targets
    Array of objects with Name (VM name exactly as it appears in vCenter inventory) and IPAddress
    (informational only - Invoke-VMScript addresses the VM object, not the IP). Defaults to the
    four servers requested for this audit.

.PARAMETER GuestCredential
    Credential used for guest-OS authentication via Invoke-VMScript (needs local admin or
    equivalent read access inside each guest). Optional - if not supplied, the script tries to
    load one from -GuestCredentialPath, and only prompts interactively if that file isn't found.

.PARAMETER GuestCredentialPath
    Path to a credential previously saved with Export-Clixml (e.g.
    Get-Credential | Export-Clixml C:\temp\guestcred.xml). Defaults to C:\temp\guestcred.xml.
    Only used when -GuestCredential isn't passed explicitly. Note: Export-Clixml encrypts the
    password with Windows DPAPI tied to the user account + machine that created it, so this only
    works when run as that same user on that same machine.

.PARAMETER OutputPath
    Folder for the CSV/XLSX + log file. Created if it doesn't exist. Defaults to
    .\SecurityAudit_Reports next to this script.

.PARAMETER ExportExcel
    Also export an .xlsx (in addition to the CSV, which is always written) if the ImportExcel
    module is installed. Silently falls back to CSV-only with a log note if the module is missing.

.EXAMPLE
    # Already connected: Connect-VIServer aq-vc.aq.local ; Connect-VIServer sf-vc.sixflags.local
    # Loads the credential automatically from C:\temp\guestcred.xml
    .\Get-WindowsServerSecurityAudit.ps1

.EXAMPLE
    .\Get-WindowsServerSecurityAudit.ps1 -GuestCredential (Get-Credential) -Targets @([PSCustomObject]@{Name='AQ-UFM-APP01';IPAddress='10.28.9.46'}) -ExportExcel
#>

[CmdletBinding()]
param(
    [object[]]$Targets = @(
        [PSCustomObject]@{ Name = 'AQ-UFM-APP01'; IPAddress = '10.28.9.46' }
        [PSCustomObject]@{ Name = 'AQ-UFM-DB01';  IPAddress = '10.28.10.14' }
        [PSCustomObject]@{ Name = 'SF-UFM-APP01'; IPAddress = '10.50.16.35' }
        [PSCustomObject]@{ Name = 'SF-UFM-DB01';  IPAddress = '10.50.18.28' }
    ),

    [System.Management.Automation.PSCredential]
    $GuestCredential,

    [string]$GuestCredentialPath = 'C:\temp\guestcred.xml',

    [string]$OutputPath = (Join-Path $PSScriptRoot 'SecurityAudit_Reports'),

    # Absolute path to powershell.exe INSIDE each guest. Guest checks run via -ScriptType Bat
    # calling this path directly rather than -ScriptType Powershell, because some VMware Tools
    # versions fail to auto-detect PowerShell as a known interpreter (surfaces as "Could not
    # locate 'Powershell' script interpreter... probably you do not have enough permissions" even
    # for a fully-privileged account). Default is the standard Windows Server location; override
    # only if a target guest has PowerShell installed somewhere non-standard.
    [string]$GuestPowerShellPath = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',

    [switch]$ExportExcel
)

# Bump this on every change so it's always possible to confirm which version of the script
# produced a given run - printed first thing at startup and written into the log file.
$ScriptBuild = '2026-08-16-04-no-backticks'
Write-Host "Get-WindowsServerSecurityAudit.ps1 - build $ScriptBuild" -ForegroundColor Magenta
Write-Host "READ-ONLY ASSESSMENT - no configuration changes, restarts, or GPO/registry/service/Defender/WinRM writes are made by this script." -ForegroundColor Yellow

if (-not (Get-Module -ListAvailable -Name VMware.VimAutomation.Core)) {
    throw "VMware PowerCLI (VMware.VimAutomation.Core) module not found. Install PowerCLI before running this script."
}
if (-not $global:DefaultVIServers -or $global:DefaultVIServers.Count -eq 0) {
    throw "No connected vCenter session found. Run Connect-VIServer <vcenter> first, then re-run this script."
}

if (-not $PSBoundParameters.ContainsKey('GuestCredential')) {
    # Try the given path, then the same filename with/without a .xml extension, so it doesn't
    # matter which exact form was used when the credential was originally exported.
    $candidatePaths = @(
        $GuestCredentialPath,
        $(if ($GuestCredentialPath -notmatch '\.xml$') { "$GuestCredentialPath.xml" }),
        $(if ($GuestCredentialPath -match '\.xml$') { $GuestCredentialPath -replace '\.xml$', '' })
    ) | Where-Object { $_ } | Select-Object -Unique

    $foundPath = $candidatePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($foundPath) {
        try {
            $GuestCredential = Import-Clixml -Path $foundPath -ErrorAction Stop
            Write-Host "Loaded guest credential from $foundPath" -ForegroundColor Cyan
        } catch {
            throw "Found a credential file at '$foundPath' but could not import it (it may have been exported by a different user account or on a different machine - Export-Clixml credentials only decrypt for the same user+machine that created them). Underlying error: $($_.Exception.Message)"
        }
    } else {
        Write-Host "No saved credential found at '$GuestCredentialPath' - prompting interactively." -ForegroundColor Yellow
        $GuestCredential = Get-Credential -Message 'Guest OS credential for Invoke-VMScript (local/domain admin on the target VMs)'
    }
}

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$LogPath  = Join-Path $OutputPath "SecurityAudit_$RunStamp.log"
$CsvPath  = Join-Path $OutputPath "SecurityAudit_$RunStamp.csv"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

function Add-ResultRow {
    param(
        [System.Collections.Generic.List[object]]$Target,
        [string]$VMName,
        [string]$IPAddress,
        [string]$Category,
        [string]$Check,
        [string]$Status,
        [string]$Evidence = '',
        [string]$Interpretation = '',
        [bool]$RequiresAdditionalEvidence = $false,
        [string]$EvidenceGuidance = ''
    )
    $Target.Add([PSCustomObject]@{
        VMName                     = $VMName
        IPAddress                  = $IPAddress
        Category                   = $Category
        Check                      = $Check
        Status                     = $Status
        Evidence                   = $Evidence
        Interpretation             = $Interpretation
        RequiresAdditionalEvidence = $RequiresAdditionalEvidence
        EvidenceGuidance           = $EvidenceGuidance
        Timestamp                  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

Write-Log "Run started. Build $ScriptBuild. Targets: $($Targets.Name -join ', ')"

Write-Host ""
Write-Host "Status legend:" -ForegroundColor Cyan
Write-Host "  Enabled / Disabled          - setting was read directly from the guest and confirmed."
Write-Host "  Not Configured              - no explicit value present; OS/domain default applies."
Write-Host "  Not Applicable              - check does not apply (VM off, feature absent, etc.)."
Write-Host "  Unable to Verify            - check could not be completed - NOT the same as Disabled."
Write-Host "  RequiresAdditionalEvidence  - indicator only; confirm with effective GPO/RSOP before treating as authoritative."
Write-Host ""

# ---------------------------------------------------------------------------------------------
# Guest-side script: everything below runs INSIDE the target VM via Invoke-VMScript. Every
# operation is a read (Get-*/Test-Path/registry read) - nothing here Sets/News/Removes/
# Starts/Stops/Restarts anything. No VM name/credential text is interpolated into this string,
# so it is identical for every target.
# ---------------------------------------------------------------------------------------------
$GuestScript = @'
$ProgressPreference = "SilentlyContinue"
$results = New-Object System.Collections.Generic.List[object]

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
    [PSCustomObject]@{
        Category                   = $Category
        Check                      = $Check
        Status                     = $Status
        Evidence                   = $Evidence
        Interpretation             = $Interpretation
        RequiresAdditionalEvidence = $RequiresAdditionalEvidence
        EvidenceGuidance           = $EvidenceGuidance
    }
}

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $results.Add((New-Check -Category "Host Info" -Check "Operating System" -Status "Info" -Evidence "$($os.Caption) Build $($os.BuildNumber)" -Interpretation "Guest operating system reported by WMI - for reference only."))
} catch {
    $results.Add((New-Check -Category "Host Info" -Check "Operating System" -Status "Unable to Verify" -Evidence $_.Exception.Message))
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
    $results.Add((New-Check -Category "VBS & Credential Protection" -Check "Virtualization Based Security (VBS)" -Status $vbsText -Evidence "VirtualizationBasedSecurityStatus=$vbsStatus" -Interpretation "0=Not enabled, 1=Enabled in policy but not actually running (hardware/firmware may not support it), 2=Enabled and running."))

    $configured = @($dg.SecurityServicesConfigured)
    $running    = @($dg.SecurityServicesRunning)

    $cgConfigured = $configured -contains 1
    $cgRunning    = $running -contains 1
    $cgStatus = if ($cgRunning) { "Enabled (Running)" } elseif ($cgConfigured) { "Configured but NOT Running" } else { "Not Configured" }
    $results.Add((New-Check -Category "VBS & Credential Protection" -Check "Credential Guard" -Status $cgStatus -Evidence "SecurityServicesConfigured=[$($configured -join ',')] SecurityServicesRunning=[$($running -join ',')]" -Interpretation "Isolates LSASS secrets in a VBS container to block credential-theft tools (e.g. Mimikatz-style attacks). 'Configured but not running' usually means a pending reboot or unmet hardware/firmware requirements."))

    $hvciConfigured = $configured -contains 2
    $hvciRunning    = $running -contains 2
    $hvciStatus = if ($hvciRunning) { "Enabled (Running)" } elseif ($hvciConfigured) { "Configured but NOT Running" } else { "Not Configured" }
    $results.Add((New-Check -Category "VBS & Credential Protection" -Check "HVCI / Memory Integrity" -Status $hvciStatus -Evidence "SecurityServicesConfigured=[$($configured -join ',')] SecurityServicesRunning=[$($running -join ',')]" -Interpretation "Hypervisor-protected Code Integrity ensures only signed, trusted code runs in kernel mode. 'Configured but not running' usually means a pending reboot or unmet hardware/firmware requirements."))
} catch {
    $results.Add((New-Check -Category "VBS & Credential Protection" -Check "Credential Guard" -Status "Unable to Verify" -Evidence $_.Exception.Message -Interpretation "Could not query Win32_DeviceGuard (root\Microsoft\Windows\DeviceGuard) - class may be unavailable on this OS build/SKU."))
    $results.Add((New-Check -Category "VBS & Credential Protection" -Check "HVCI / Memory Integrity" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

$results.Add((New-Check -Category "VBS & Credential Protection" -Check "Machine Identity Isolation" -Status "Unable to Verify" -Evidence "No single well-documented registry/WMI indicator identified for this setting at time of writing." -Interpretation "This is a newer Windows security capability; this script deliberately does not guess at an unverified indicator rather than report a false Enabled/Disabled result." -RequiresAdditionalEvidence $true -EvidenceGuidance "Confirm current implementation details against up-to-date Microsoft Windows Server 2025 documentation, then extend this script once confirmed."))

# ===========================================================================
# 2. Microsoft Defender
# ===========================================================================
$mpAvailable = [bool](Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)
if (-not $mpAvailable) {
    $results.Add((New-Check -Category "Microsoft Defender" -Check "Defender Cmdlets Available" -Status "Not Applicable" -Evidence "Get-MpComputerStatus / Get-MpPreference not found." -Interpretation "Defender PowerShell module is not present - this can mean Defender is disabled/uninstalled, or a third-party AV has replaced it. Verify manually via the Windows Security app or the AV vendor's console."))
} else {
    try {
        $mpStatus = Get-MpComputerStatus -ErrorAction Stop
        $results.Add((New-Check -Category "Microsoft Defender" -Check "Real-Time Protection" -Status $(if ($mpStatus.RealTimeProtectionEnabled) { "Enabled" } else { "Disabled" }) -Evidence "RealTimeProtectionEnabled=$($mpStatus.RealTimeProtectionEnabled); AMServiceEnabled=$($mpStatus.AMServiceEnabled); AntivirusEnabled=$($mpStatus.AntivirusEnabled)" -Interpretation "Whether Defender's on-access/real-time scanning engine is currently active."))
        $results.Add((New-Check -Category "Microsoft Defender" -Check "Tamper Protection" -Status $(if ($mpStatus.IsTamperProtected) { "Enabled" } else { "Disabled" }) -Evidence "IsTamperProtected=$($mpStatus.IsTamperProtected)" -Interpretation "Blocks unauthorized changes to Defender security settings, including via admin-level scripts or direct registry edits."))
    } catch {
        $results.Add((New-Check -Category "Microsoft Defender" -Check "Real-Time Protection" -Status "Unable to Verify" -Evidence $_.Exception.Message))
    }

    try {
        $mpPref = Get-MpPreference -ErrorAction Stop

        $asrIds     = @($mpPref.AttackSurfaceReductionRules_Ids)
        $asrActions = @($mpPref.AttackSurfaceReductionRules_Actions)
        if ($asrIds.Count -eq 0) {
            $results.Add((New-Check -Category "Microsoft Defender" -Check "Attack Surface Reduction (ASR) Rules" -Status "Not Configured" -Evidence "No ASR rule IDs present in Get-MpPreference." -Interpretation "No ASR rules are configured locally or via policy on this host."))
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
            $results.Add((New-Check -Category "Microsoft Defender" -Check "Attack Surface Reduction (ASR) Rules" -Status $overall -Evidence ($pairs -join "; ") -Interpretation "ASR rules block common malware/exploit behaviors. Block=enforced, Audit=logged only (not enforced), Warn=user can bypass. Reflects Defender's merged effective policy."))
        }

        $mapsText = switch ($mpPref.MAPSReporting) {
            0 { "Disabled" }
            1 { "Basic" }
            2 { "Advanced" }
            default { "Unknown($($mpPref.MAPSReporting))" }
        }
        $results.Add((New-Check -Category "Microsoft Defender" -Check "Cloud-Delivered Protection (MAPS Reporting)" -Status $mapsText -Evidence "MAPSReporting=$($mpPref.MAPSReporting); SubmitSamplesConsent=$($mpPref.SubmitSamplesConsent); CloudBlockLevel=$($mpPref.CloudBlockLevel)" -Interpretation "Controls whether Defender submits telemetry/sample data to Microsoft's cloud for enhanced detection."))
    } catch {
        $results.Add((New-Check -Category "Microsoft Defender" -Check "Attack Surface Reduction (ASR) Rules" -Status "Unable to Verify" -Evidence $_.Exception.Message))
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
        $results.Add((New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Not Configured" -Evidence "Registry value absent." -Interpretation "No explicit setting present. Windows 8.1/Server 2012 R2+ ships with WDigest credential caching off by default, but an explicit Disabled (0) value is recommended so the posture is enforced rather than relying on the OS default."))
    } elseif ($wdigestVal -eq 1) {
        $results.Add((New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Enabled" -Evidence "UseLogonCredential=1" -Interpretation "INSECURE: WDigest caches reversible/plaintext-equivalent credentials in LSASS memory, a common credential-theft target. Should be Disabled unless a specific legacy application requires it."))
    } else {
        $results.Add((New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Disabled" -Evidence "UseLogonCredential=$wdigestVal"))
    }
} catch {
    $results.Add((New-Check -Category "Authentication Security" -Check "WDigest (UseLogonCredential)" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

try {
    $kerbPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters"
    $encVal = $null
    if (Test-Path $kerbPath) {
        $prop = Get-ItemProperty -Path $kerbPath -Name SupportedEncryptionTypes -ErrorAction SilentlyContinue
        if ($prop) { $encVal = $prop.SupportedEncryptionTypes }
    }
    if ($null -eq $encVal) {
        $results.Add((New-Check -Category "Authentication Security" -Check "Kerberos Supported Encryption Types" -Status "Not Configured" -Evidence "SupportedEncryptionTypes registry value absent - OS/domain default negotiation applies." -Interpretation "No explicit local policy restricting Kerberos encryption types." -RequiresAdditionalEvidence $true -EvidenceGuidance "Confirm effective value via 'Network security: Configure encryption types allowed for Kerberos' in effective GPO (gpresult /h) - this can be domain-policy driven without a local registry value."))
    } else {
        $bits = [int]$encVal
        $types = @()
        if ($bits -band 0x1)  { $types += "DES-CBC-CRC (legacy, weak)" }
        if ($bits -band 0x2)  { $types += "DES-CBC-MD5 (legacy, weak)" }
        if ($bits -band 0x4)  { $types += "RC4-HMAC-MD5 (legacy)" }
        if ($bits -band 0x8)  { $types += "AES128-CTS-HMAC-SHA1-96" }
        if ($bits -band 0x10) { $types += "AES256-CTS-HMAC-SHA1-96" }
        $results.Add((New-Check -Category "Authentication Security" -Check "Kerberos Supported Encryption Types" -Status "Configured" -Evidence "SupportedEncryptionTypes=0x$($bits.ToString('X')) -> $($types -join ', ')" -Interpretation "This value governs the classic AES128/AES256-SHA1 (and legacy RC4/DES) Kerberos suites. Microsoft's newer AES-SHA2 'next-generation crypto' suites (the SHA256/SHA384/SHA512-based Kerberos support referenced in the Windows Server 2025 baseline) were not confirmed to be controlled by this same value as of this script's authoring - do NOT treat this check as proof SHA256/384/512 Kerberos support is enabled or disabled." -RequiresAdditionalEvidence $true -EvidenceGuidance "Confirm SHA-2 Kerberos suite support/enforcement via current Microsoft KDC/Windows Server 2025 documentation, 'klist' on an authenticated session, or effective GPO/RSOP - do not rely on this registry check alone for that specific claim."))
    }
} catch {
    $results.Add((New-Check -Category "Authentication Security" -Check "Kerberos Supported Encryption Types" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

# ===========================================================================
# 4. Remote Management (WinRM)
# ===========================================================================
try {
    $svc = Get-CimInstance -ClassName Win32_Service -Filter "Name='WinRM'" -ErrorAction Stop
    $results.Add((New-Check -Category "Remote Management" -Check "WinRM Service" -Status "$($svc.State) (StartMode=$($svc.StartMode))" -Evidence "State=$($svc.State); StartMode=$($svc.StartMode)" -Interpretation "Current run state and boot-time start mode of the WinRM service, queried locally and read-only - no service control action performed."))
} catch {
    $results.Add((New-Check -Category "Remote Management" -Check "WinRM Service" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

try {
    $listenerRoot = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Listener"
    if (Test-Path $listenerRoot) {
        $listenerKeys = Get-ChildItem -Path $listenerRoot -ErrorAction SilentlyContinue
        if (-not $listenerKeys -or $listenerKeys.Count -eq 0) {
            $results.Add((New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status "Not Configured" -Evidence "No listener subkeys under $listenerRoot"))
        } else {
            $listenerInfo = foreach ($lk in $listenerKeys) {
                $p = Get-ItemProperty -Path $lk.PSPath -ErrorAction SilentlyContinue
                "Transport=$($p.Transport) Port=$($p.Port) Enabled=$($p.Enabled) ListeningOn=$($p.ListeningOn -join '/')"
            }
            $httpCount = @($listenerInfo | Where-Object { $_ -match 'Transport=HTTP ' }).Count
            $status = if ($httpCount -gt 0) { "Configured - includes unencrypted HTTP listener(s)" } else { "Configured - HTTPS only" }
            $results.Add((New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status $status -Evidence ($listenerInfo -join " | ") -Interpretation "Read directly from the registry (does not require the WinRM service to be running). An HTTP (unencrypted) listener is generally a hardening gap versus HTTPS-only."))
        }
    } else {
        $results.Add((New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status "Not Configured" -Evidence "$listenerRoot not present."))
    }
} catch {
    $results.Add((New-Check -Category "Remote Management" -Check "WinRM Listeners" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

try {
    $svcCfgPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service"
    if (Test-Path $svcCfgPath) {
        $cfg = Get-ItemProperty -Path $svcCfgPath -ErrorAction SilentlyContinue
        $results.Add((New-Check -Category "Remote Management" -Check "WinRM Hardening (Unencrypted/Basic Auth)" -Status "Configured" -Evidence "AllowUnencrypted=$($cfg.AllowUnencrypted); auth_Basic=$($cfg.auth_Basic); auth_Kerberos=$($cfg.auth_Kerberos); auth_Negotiate=$($cfg.auth_Negotiate); auth_CredSSP=$($cfg.auth_CredSSP)" -Interpretation "0/absent generally means the more secure default; 1 for AllowUnencrypted or auth_Basic indicates a weaker/legacy configuration that should be reviewed."))
    } else {
        $results.Add((New-Check -Category "Remote Management" -Check "WinRM Hardening (Unencrypted/Basic Auth)" -Status "Not Configured" -Evidence "$svcCfgPath not present (WinRM likely never configured)."))
    }
} catch {
    $results.Add((New-Check -Category "Remote Management" -Check "WinRM Hardening (Unencrypted/Basic Auth)" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

# ===========================================================================
# 5. SMB & RPC Security
# ===========================================================================
try {
    if (Get-Command Get-SmbServerConfiguration -ErrorAction SilentlyContinue) {
        $smb = Get-SmbServerConfiguration -ErrorAction Stop
        $auditSmb1 = if ($smb.PSObject.Properties.Name -contains 'AuditSmb1Access') { $smb.AuditSmb1Access } else { $null }
        $auditStatus = if ($null -eq $auditSmb1) { "Not Applicable (property not present on this OS build)" } elseif ($auditSmb1) { "Enabled" } else { "Disabled" }
        $results.Add((New-Check -Category "SMB & RPC Security" -Check "SMB1 Access Auditing" -Status $auditStatus -Evidence "AuditSmb1Access=$auditSmb1; EnableSMB1Protocol=$($smb.EnableSMB1Protocol)" -Interpretation "AuditSmb1Access logs any client still attempting legacy SMB1 connections, to support safely disabling SMB1. Requires Windows Server 2022+ for this property."))
        $results.Add((New-Check -Category "SMB & RPC Security" -Check "SMB Encryption / Signing" -Status "Info" -Evidence "EncryptData=$($smb.EncryptData); RejectUnencryptedAccess=$($smb.RejectUnencryptedAccess); RequireSecuritySignature=$($smb.RequireSecuritySignature); EnableSecuritySignature=$($smb.EnableSecuritySignature)" -Interpretation "Current SMB server encryption and signing enforcement settings."))
    } else {
        $results.Add((New-Check -Category "SMB & RPC Security" -Check "SMB1 Access Auditing" -Status "Unable to Verify" -Evidence "Get-SmbServerConfiguration cmdlet not available (SMB module missing)."))
    }
} catch {
    $results.Add((New-Check -Category "SMB & RPC Security" -Check "SMB1 Access Auditing" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

try {
    $printPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print"
    $rpcVal = $null
    if (Test-Path $printPath) {
        $prop = Get-ItemProperty -Path $printPath -Name RpcAuthnLevelPrivacyEnabled -ErrorAction SilentlyContinue
        if ($prop) { $rpcVal = $prop.RpcAuthnLevelPrivacyEnabled }
    }
    $rpcStatus = if ($null -eq $rpcVal) { "Not Configured (default applies - confirm this OS build's default)" } elseif ($rpcVal -eq 1) { "Enabled (Hardened - RPC packet privacy required)" } else { "Disabled" }
    $results.Add((New-Check -Category "SMB & RPC Security" -Check "Printer RPC Packet Privacy (PrintNightmare mitigation)" -Status $rpcStatus -Evidence "RpcAuthnLevelPrivacyEnabled=$rpcVal" -Interpretation "Requires RPC connections to the Print Spooler to use packet privacy (encryption+signing) - part of the CVE-2021-34527 (PrintNightmare) mitigation set."))
} catch {
    $results.Add((New-Check -Category "SMB & RPC Security" -Check "Printer RPC Packet Privacy (PrintNightmare mitigation)" -Status "Unable to Verify" -Evidence $_.Exception.Message))
}

try {
    $papPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint"
    if (Test-Path $papPath) {
        $pap = Get-ItemProperty -Path $papPath -ErrorAction SilentlyContinue
        $results.Add((New-Check -Category "SMB & RPC Security" -Check "Point and Print Restrictions" -Status "Configured" -Evidence "NoWarningNoElevationOnInstall=$($pap.NoWarningNoElevationOnInstall); UpdatePromptSettings=$($pap.UpdatePromptSettings); RestrictDriverInstallationToAdministrators=$($pap.RestrictDriverInstallationToAdministrators)" -Interpretation "RestrictDriverInstallationToAdministrators=1 is the key PrintNightmare hardening setting (blocks non-admins from installing print drivers). NoWarningNoElevationOnInstall=1 is the opposite (a weakening setting) and should be 0/absent."))
    } else {
        $results.Add((New-Check -Category "SMB & RPC Security" -Check "Point and Print Restrictions" -Status "Not Configured" -Evidence "$papPath not present - policy not applied, OS defaults apply."))
    }
} catch {
    $results.Add((New-Check -Category "SMB & RPC Security" -Check "Point and Print Restrictions" -Status "Unable to Verify" -Evidence $_.Exception.Message))
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
    $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "SmartScreen (Explorer policy)" -Status $ssStatus -Evidence "EnableSmartScreen=$ssVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true -EvidenceGuidance "Compare against baseline GPO backup / gpresult."))
} catch {
    $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "SmartScreen (Explorer policy)" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true))
}

try {
    $zonePath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3"
    if (Test-Path $zonePath) {
        $zone = Get-ItemProperty -Path $zonePath -ErrorAction SilentlyContinue
        $activeXVal   = $zone.'1200'
        $scriptingVal = $zone.'1400'
        $protectedVal = $zone.'2500'
        $decode = { param($v) if ($null -eq $v) { "Not Configured" } elseif ($v -eq 0) { "Enable" } elseif ($v -eq 1) { "Prompt" } elseif ($v -eq 3) { "Disable" } else { "Value=$v" } }
        $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - Run ActiveX Controls" -Status (& $decode $activeXVal) -Evidence "1200=$activeXVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true))
        $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - Active Scripting" -Status (& $decode $scriptingVal) -Evidence "1400=$scriptingVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true))
        $protStatus = if ($null -eq $protectedVal) { "Not Configured" } elseif ($protectedVal -eq 0) { "Enabled (Protected Mode ON)" } else { "Disabled (Protected Mode OFF)" }
        $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - Protected Mode" -Status $protStatus -Evidence "2500=$protectedVal" -Interpretation $baselineNote -RequiresAdditionalEvidence $true))
    } else {
        $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - ActiveX/Scripting/Protected Mode" -Status "Not Configured" -Evidence "$zonePath not present." -Interpretation $baselineNote -RequiresAdditionalEvidence $true))
    }
} catch {
    $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "IE Zone (Internet) - ActiveX/Scripting/Protected Mode" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true))
}

try {
    $schannelBase = "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols"
    $protoList = @('SSL 2.0','SSL 3.0','TLS 1.0','TLS 1.1','TLS 1.2','TLS 1.3')
    $protoEvidence = foreach ($proto in $protoList) {
        $serverKey = Join-Path (Join-Path $schannelBase $proto) 'Server'
        if (Test-Path $serverKey) {
            $p = Get-ItemProperty -Path $serverKey -ErrorAction SilentlyContinue
            $enabled = if ($p.PSObject.Properties.Name -contains 'Enabled') { $p.Enabled } else { $null }
            $disabledByDefault = if ($p.PSObject.Properties.Name -contains 'DisabledByDefault') { $p.DisabledByDefault } else { $null }
            "$proto[Enabled=$enabled;DisabledByDefault=$disabledByDefault]"
        } else {
            "$proto[Not Configured - OS default]"
        }
    }
    $status = if ($protoEvidence -match 'SSL [23]\.0\[Enabled=1') { "Weak protocol explicitly enabled (SSL 2.0/3.0)" }
              elseif ($protoEvidence -match 'TLS 1\.[01]\[Enabled=1') { "Legacy TLS 1.0/1.1 explicitly enabled" }
              else { "No legacy SSL/TLS explicitly enabled via registry (OS default in effect where not configured)" }
    $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "SSL/TLS Protocol Configuration (Schannel)" -Status $status -Evidence ($protoEvidence -join " | ") -Interpretation $baselineNote -RequiresAdditionalEvidence $true -EvidenceGuidance "Registry-absent entries fall back to the OS-default protocol set for this build, which varies by Windows Server version/patch level - confirm actual negotiated protocols with a TLS scan tool if certainty is required."))
} catch {
    $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "SSL/TLS Protocol Configuration (Schannel)" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true))
}

try {
    if (Get-Command Get-WindowsOptionalFeature -ErrorAction SilentlyContinue) {
        $ieFeature = Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction Stop
        $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "Legacy Internet Explorer 11 Feature State" -Status "$($ieFeature.State)" -Evidence "FeatureName=Internet-Explorer-Optional-amd64; State=$($ieFeature.State)" -Interpretation $baselineNote -RequiresAdditionalEvidence $true))
    } else {
        $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "Legacy Internet Explorer 11 Feature State" -Status "Unable to Verify" -Evidence "Get-WindowsOptionalFeature not available (e.g. Server Core or restricted session)." -RequiresAdditionalEvidence $true))
    }
} catch {
    $results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "Legacy Internet Explorer 11 Feature State" -Status "Unable to Verify" -Evidence $_.Exception.Message -RequiresAdditionalEvidence $true))
}

$results.Add((New-Check -Category "2025 Security Baseline Indicators" -Check "Overall Baseline Applied?" -Status "Not Determinable From Registry Alone" -Evidence "$($results.Count) individual indicator checks captured above." -Interpretation "The full Windows Server 2025 Security Baseline covers hundreds of settings across many areas. The indicators above are a representative subset only." -RequiresAdditionalEvidence $true -EvidenceGuidance "Run 'gpresult /h report.html' on the server, or use Microsoft's free Policy Analyzer tool to diff effective local policy against the official baseline GPO backup, for an authoritative answer."))

$json = $results | ConvertTo-Json -Depth 6 -Compress
Write-Output $json
'@

# ---------------------------------------------------------------------------------------------
# Bat-wrapper staging: -ScriptType Powershell relies on VMware Tools auto-detecting where
# powershell.exe lives inside the guest, and on some Tools versions that detection is simply
# broken (confirmed against these targets: identical failure for both a confirmed local-admin
# account and the built-in Administrator, while -ScriptType Bat succeeded and found
# powershell.exe at the expected path every time). Using -ScriptType Bat to invoke
# -GuestPowerShellPath directly sidesteps that broken detection entirely.
#
# $GuestScript above is too large to pass as a single inline command (Windows command-line
# length limits), so the Bat wrapper below stages it into a small temp file in the guest's own
# %TEMP% first, runs it, then deletes that same file in the same call - nothing persists on the
# guest after Invoke-VMScript returns. Content is base64 (UTF-16LE, matching PowerShell's
# -EncodedCommand convention) purely so it's safe to write via `echo` without cmd.exe
# misinterpreting any special characters in the script text - the base64 alphabet contains none.
$EncodedGuestScript = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($GuestScript))
$ChunkSize = 4000
$GuestScriptChunks = for ($i = 0; $i -lt $EncodedGuestScript.Length; $i += $ChunkSize) {
    $EncodedGuestScript.Substring($i, [Math]::Min($ChunkSize, $EncodedGuestScript.Length - $i))
}

function New-GuestBatWrapper {
    param([string[]]$EncodedChunks, [string]$PowerShellExePath)

    $tempFileName = "secaudit_$([guid]::NewGuid().ToString('N')).b64"
    # `$p`/`$b64`/etc. below must stay literal (escaped) so they land in the loader as PowerShell
    # variable names; only $tempFileName should be expanded here, into the loader's file path.
    $loaderSource = @"
`$p = Join-Path `$env:TEMP '$tempFileName'
`$b64 = Get-Content -Raw -Path `$p
`$bytes = [Convert]::FromBase64String(`$b64)
`$code = [Text.Encoding]::Unicode.GetString(`$bytes)
Invoke-Expression `$code
"@
    $encodedLoader = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($loaderSource))

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('@echo off')
    for ($i = 0; $i -lt $EncodedChunks.Count; $i++) {
        $redirect = if ($i -eq 0) { '>' } else { '>>' }
        $lines.Add("echo $($EncodedChunks[$i]) $redirect `"%TEMP%\$tempFileName`"")
    }
    $lines.Add("`"$PowerShellExePath`" -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encodedLoader")
    $lines.Add("del `"%TEMP%\$tempFileName`" 2>nul")
    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------------------------
$AllResults = New-Object System.Collections.Generic.List[object]
$quickChecks = @(
    'Credential Guard','HVCI / Memory Integrity','Machine Identity Isolation',
    'Attack Surface Reduction (ASR) Rules','Real-Time Protection','Tamper Protection',
    'WDigest (UseLogonCredential)','Kerberos Supported Encryption Types',
    'WinRM Service','WinRM Listeners',
    'SMB1 Access Auditing','Printer RPC Packet Privacy (PrintNightmare mitigation)',
    'Overall Baseline Applied?'
)

foreach ($target in $Targets) {
    $vmName = $target.Name
    $vmIp   = $target.IPAddress
    Write-Log "Processing $vmName ($vmIp)..."
    $rowsForVm = New-Object System.Collections.Generic.List[object]

    $vm = $null
    try {
        $vm = Get-VM -Name $vmName -ErrorAction Stop
        if ($vm -is [array]) {
            Write-Log "Multiple VMs named '$vmName' found across connected vCenters - using the first match." 'WARN'
            $vm = $vm[0]
        }
    } catch {
        Write-Log "VM '$vmName' not found in any connected vCenter: $($_.Exception.Message)" 'ERROR'
        Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'VM Found in vCenter' -Status 'Unable to Verify' -Evidence $_.Exception.Message -Interpretation 'The VM name could not be resolved on any currently connected vCenter Server.'
        $AllResults.AddRange($rowsForVm)
        continue
    }

    Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'VM Found in vCenter' -Status 'Enabled' -Evidence "PowerState=$($vm.PowerState)"

    if ($vm.PowerState -ne 'PoweredOn') {
        Write-Log "$vmName is $($vm.PowerState) - skipping guest-OS checks." 'WARN'
        Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'Power State' -Status 'Not Applicable' -Evidence "PowerState=$($vm.PowerState)" -Interpretation 'Guest-OS checks require the VM to be powered on and running VMware Tools; all remaining checks are Not Applicable for this run.'
        $AllResults.AddRange($rowsForVm)
        continue
    }

    $toolsStatus = $vm.ExtensionData.Guest.ToolsRunningStatus
    if ($toolsStatus -ne 'guestToolsRunning') {
        Write-Log "$vmName VMware Tools not running (status: $toolsStatus) - cannot run guest checks." 'WARN'
        Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'VMware Tools Running' -Status 'Unable to Verify' -Evidence "ToolsRunningStatus=$toolsStatus" -Interpretation 'Invoke-VMScript requires a running VMware Tools service inside the guest. All guest-OS checks are Unable to Verify for this run.'
        $AllResults.AddRange($rowsForVm)
        continue
    }
    Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'VMware Tools Running' -Status 'Enabled' -Evidence "ToolsRunningStatus=$toolsStatus"

    $rawOutput = $null
    try {
        $batWrapper = New-GuestBatWrapper -EncodedChunks $GuestScriptChunks -PowerShellExePath $GuestPowerShellPath
        $invokeResult = Invoke-VMScript -VM $vm -ScriptType Bat -ScriptText $batWrapper -GuestCredential $GuestCredential -ErrorAction Stop
        $rawOutput = $invokeResult.ScriptOutput
    } catch {
        $msg = $_.Exception.Message
        $authFailure       = $msg -match 'login|credential|password|authenticat'
        $interpreterMissing = $msg -match 'locate.*script interpreter|not have enough permissions'
        Write-Log "Invoke-VMScript failed for $vmName - $msg" 'ERROR'
        $interpretation =
            if ($interpreterMissing) {
                "VMware Tools could not run the Bat wrapper itself (before it ever reached PowerShell) - this is a step below the interpreter-detection issue this script already works around, so something more fundamental with guest operations on this VM is failing. Confirm -GuestPowerShellPath ($GuestPowerShellPath) actually exists on $vmName, and that VMware Tools is fully installed/running (not just the minimal/legacy driver set)."
            } elseif ($authFailure) {
                'Guest authentication failed - verify -GuestCredential/-GuestCredentialPath has a valid username and password for this VM.'
            } else {
                'Invoke-VMScript could not run inside the guest - see Evidence for the underlying error.'
            }
        Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'Guest Script Execution' -Status 'Unable to Verify' -Evidence $msg -Interpretation $interpretation
        $AllResults.AddRange($rowsForVm)
        continue
    }

    if ([string]::IsNullOrWhiteSpace($rawOutput)) {
        Write-Log "$vmName returned no output from the guest script." 'ERROR'
        Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'Guest Script Execution' -Status 'Unable to Verify' -Evidence 'Guest script returned empty output.'
        $AllResults.AddRange($rowsForVm)
        continue
    }

    try {
        $parsed = $rawOutput | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-Log "$vmName - failed to parse guest script JSON output: $($_.Exception.Message)" 'ERROR'
        $truncated = $rawOutput.Substring(0, [Math]::Min(500, $rawOutput.Length))
        Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category 'VM Connectivity' -Check 'Guest Script Execution' -Status 'Unable to Verify' -Evidence "JSON parse failure. Raw output (truncated): $truncated"
        $AllResults.AddRange($rowsForVm)
        continue
    }

    foreach ($item in $parsed) {
        Add-ResultRow -Target $rowsForVm -VMName $vmName -IPAddress $vmIp -Category $item.Category -Check $item.Check -Status $item.Status -Evidence $item.Evidence -Interpretation $item.Interpretation -RequiresAdditionalEvidence ([bool]$item.RequiresAdditionalEvidence) -EvidenceGuidance $item.EvidenceGuidance
    }

    $AllResults.AddRange($rowsForVm)

    Write-Host ""
    Write-Host "=== $vmName ($vmIp) ===" -ForegroundColor Green
    # Manually formatted (not Format-Table -AutoSize): AutoSize depends on detecting the
    # console width, which returns -1/unknown under redirected output, scheduled tasks, and
    # other non-interactive hosts - it silently prints nothing in those cases.
    $quickRows = $rowsForVm | Where-Object { $quickChecks -contains $_.Check }
    foreach ($qc in $quickChecks) {
        $row = $quickRows | Where-Object { $_.Check -eq $qc } | Select-Object -First 1
        if ($row) {
            Write-Host ("  {0,-55} : {1}" -f $qc, $row.Status)
        }
    }
    Write-Log "$vmName processed - $($rowsForVm.Count) checks recorded."
}

# ---------------------------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------------------------
$AllResults | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV exported: $CsvPath"

if ($ExportExcel) {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        $xlsxPath = Join-Path $OutputPath "SecurityAudit_$RunStamp.xlsx"
        try {
            Import-Module ImportExcel -ErrorAction Stop
            $AllResults | Export-Excel -Path $xlsxPath -AutoSize -FreezeTopRow -BoldTopRow -WorksheetName 'SecurityAudit' -ErrorAction Stop
            Write-Log "Excel exported: $xlsxPath"
        } catch {
            Write-Log "Excel export failed ($($_.Exception.Message)) - CSV is still available at $CsvPath" 'WARN'
        }
    } else {
        Write-Log "-ExportExcel requested but the ImportExcel module is not installed. Install with 'Install-Module ImportExcel' or use the CSV output at $CsvPath." 'WARN'
    }
}

Write-Log "Run complete. $($AllResults.Count) total check rows across $($Targets.Count) target VM(s)."
Write-Host ""
Write-Host "Results: $CsvPath"
Write-Host "Log:     $LogPath"
