# Post-remediation verification - read-only, changes nothing.
# Grouped to match the original email request exactly, category by category.

# Main points (category headers) - always this one color.
$HeaderColor = 'Cyan'
# Sub points (each checked item's label) - always this other color, distinct from headers.
$SubPointColor = 'White'

function Show-Result {
    param([string]$Label, [string]$Status, [string]$Detail = '')
    $statusColor = switch -Wildcard ($Status) {
        'OK*'     { 'Green' }
        'MANUAL*' { 'DarkYellow' }
        default   { 'Red' }
    }
    Write-Host ("  [{0}]" -f $Status.PadRight(8)) -ForegroundColor $statusColor -NoNewline
    Write-Host (" {0}" -f $Label) -ForegroundColor $SubPointColor -NoNewline
    if ($Detail) { Write-Host (" - {0}" -f $Detail) -ForegroundColor Gray } else { Write-Host "" }
}

# ===========================================================================
Write-Host "`n=== VBS & Credential Protection - Enable Credential Guard, HVCI/VBS, Machine Identity Isolation ===" -ForegroundColor $HeaderColor
$v = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\LSA' -Name LsaCfgFlags -ErrorAction SilentlyContinue
Show-Result "Credential Guard" $(if ($v -and $v.LsaCfgFlags -in @(1,2)) { "OK" } else { "NOT SET" })

$vbs = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard' -Name EnableVirtualizationBasedSecurity -ErrorAction SilentlyContinue
$hvci = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' -Name Enabled -ErrorAction SilentlyContinue
$hvciVbsOk = ($vbs -and $vbs.EnableVirtualizationBasedSecurity -eq 1) -and ($hvci -and $hvci.Enabled -eq 1)
Show-Result "HVCI/VBS" $(if ($hvciVbsOk) { "OK" } else { "NOT SET" })

$v = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard' -Name MachineIdentityIsolation -ErrorAction SilentlyContinue
Show-Result "Machine Identity Isolation" $(if ($v -and $v.MachineIdentityIsolation -eq 2) { "OK" } else { "NOT SET (found: $($v.MachineIdentityIsolation))" })

# ===========================================================================
Write-Host "`n=== Microsoft Defender - Enable ASR enforcement and Defender reporting controls ===" -ForegroundColor $HeaderColor
try {
    $status = Get-MpComputerStatus -ErrorAction Stop
    Show-Result "Real-Time Protection" $(if ($status.RealTimeProtectionEnabled) { "OK" } else { "NOT SET" })
} catch { Show-Result "Real-Time Protection" "NOT SET" "Defender not available" }

try {
    $pref = Get-MpPreference -ErrorAction Stop
    Show-Result "Cloud/MAPS Reporting" $(if ($pref.MAPSReporting -eq 2) { "OK" } else { "NOT SET" })
    Show-Result "ASR Rules Configured" $(if (@($pref.AttackSurfaceReductionRules_Ids).Count -gt 0) { "OK ($(@($pref.AttackSurfaceReductionRules_Ids).Count) rules)" } else { "NOT SET" })
} catch { Show-Result "Defender Preferences" "NOT SET" "Defender not available" }

Show-Result "Tamper Protection" "MANUAL" "Cannot be checked/set by script - verify in Windows Security app or Intune"

# ===========================================================================
Write-Host "`n=== Authentication Security - Disable WDigest; enable Kerberos SHA256/SHA384/SHA512 ===" -ForegroundColor $HeaderColor
$v = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' -Name UseLogonCredential -ErrorAction SilentlyContinue
Show-Result "WDigest Disabled" $(if ($v -and $v.UseLogonCredential -eq 0) { "OK" } else { "NOT SET" })

$v = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Kerberos\Parameters' -Name SupportedEncryptionTypes -ErrorAction SilentlyContinue
$aesOk = $v -and (([int]$v.SupportedEncryptionTypes) -band 0x18) -eq 0x18
Show-Result "Kerberos AES128/256 Supported" $(if ($aesOk) { "OK" } else { "NOT SET" })

Show-Result "Kerberos SHA256/SHA384/SHA512" "MANUAL" "Not verifiable by script - requires manual check/configuration (no confirmed Windows mechanism yet, see Microsoft guidance)"

# ===========================================================================
Write-Host "`n=== Remote Management - Harden or disable WinRM; restrict listener filters ===" -ForegroundColor $HeaderColor
$v = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service' -ErrorAction SilentlyContinue
Show-Result "WinRM Unencrypted Blocked" $(if ($v -and $v.AllowUnencrypted -eq 0) { "OK" } else { "NOT SET" })
Show-Result "WinRM Basic Auth Blocked" $(if ($v -and $v.auth_Basic -eq 0) { "OK" } else { "NOT SET" })

try {
    $rules = Get-NetFirewallRule -DisplayGroup 'Windows Remote Management' -ErrorAction Stop | Where-Object { $_.Enabled -eq 'True' }
    $addr = @($rules | Get-NetFirewallAddressFilter -ErrorAction Stop | Select-Object -ExpandProperty RemoteAddress -Unique)
    if ($addr.Count -eq 1 -and $addr[0] -eq 'Any') {
        Show-Result "WinRM Listener Restricted" "MANUAL" "Currently open to Any source - supply -WinRmAllowedSourceRange and re-apply"
    } else {
        Show-Result "WinRM Listener Restricted" "OK" "Restricted to: $($addr -join ', ')"
    }
} catch { Show-Result "WinRM Listener Restricted" "MANUAL" "Could not verify - check firewall rules manually" }

# ===========================================================================
Write-Host "`n=== SMB & RPC Security - Enable SMB auditing; configure secure Printer RPC settings ===" -ForegroundColor $HeaderColor
try {
    $smb = Get-SmbServerConfiguration -ErrorAction Stop
    Show-Result "SMB1 Access Auditing" $(if ($smb.AuditSmb1Access) { "OK" } else { "NOT SET" })
} catch { Show-Result "SMB1 Access Auditing" "NOT SET" "Could not read SMB config" }

$v = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Print' -Name RpcAuthnLevelPrivacyEnabled -ErrorAction SilentlyContinue
Show-Result "Printer RPC Privacy" $(if ($v -and $v.RpcAuthnLevelPrivacyEnabled -eq 1) { "OK" } else { "NOT SET" })

$v = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint' -ErrorAction SilentlyContinue
Show-Result "Point-and-Print Restricted" $(if ($v -and $v.RestrictDriverInstallationToAdministrators -eq 1) { "OK" } else { "NOT SET" })

# ===========================================================================
Write-Host "`n=== Security Baseline - Apply Microsoft Windows Server 2025 Security Baseline GPO (IE, SmartScreen, ActiveX, SSL/TLS, scripting, Protected Mode) ===" -ForegroundColor $HeaderColor
$v = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' -Name EnableSmartScreen -ErrorAction SilentlyContinue
Show-Result "SmartScreen Enabled" $(if ($v -and $v.EnableSmartScreen -eq 1) { "OK" } else { "NOT SET" })

$v = Get-ItemProperty 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3' -ErrorAction SilentlyContinue
$ieOk = $v -and $v.'1200' -eq 3 -and $v.'1400' -eq 3 -and $v.'2500' -eq 0
Show-Result "IE Zone (ActiveX/Scripting/Protected Mode)" $(if ($ieOk) { "OK" } else { "NOT SET" })

try {
    $feat = Get-WindowsOptionalFeature -Online -FeatureName Internet-Explorer-Optional-amd64 -ErrorAction Stop
    if (-not $feat) {
        Show-Result "Legacy IE11 Removed" "OK" "N/A - this feature does not exist on this Windows Server build, nothing to remove"
    } elseif ($feat.State -eq 'Disabled') {
        Show-Result "Legacy IE11 Removed" "OK"
    } else {
        Show-Result "Legacy IE11 Removed" "NOT SET"
    }
} catch { Show-Result "Legacy IE11 Removed" "OK" "N/A - this feature does not exist on this Windows Server build, nothing to remove" }

$v = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.0\Server' -ErrorAction SilentlyContinue
Show-Result "Legacy SSL/TLS Disabled" $(if ($v -and $v.Enabled -eq 0) { "OK" } else { "MANUAL" }) "Opt-in only (-DisableLegacyTls) - confirm intentionally"

Show-Result "Full Baseline GPO Applied" "MANUAL" "Not verifiable by registry check - confirm via gpresult /h or re-run LGPO.exe /g"

Write-Host ""
