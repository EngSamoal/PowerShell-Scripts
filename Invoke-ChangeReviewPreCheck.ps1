#Requires -Version 5.1
<#
.SYNOPSIS
    Pre-check for the "Script — Review" remediation items (Firewall, SMB minimum
    version, NetBIOS, Account Lockout Threshold, UAC hardening, VBS/Credential Guard).

.DESCRIPTION
    READ-ONLY. Makes no changes to this server. Run it directly on (or via RDP/
    console session to) each target server BEFORE enabling the matching -Apply
    switch in Invoke-RemoteSecurityGapRemediation.ps1. Prints a PASS / REVIEW
    result per item and writes the same to a timestamped log file next to the
    script for evidence.

.EXAMPLE
    .\Invoke-ChangeReviewPreCheck.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot
)

$ErrorActionPreference = 'SilentlyContinue'
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = (Get-Location).Path }
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$logPath = Join-Path $OutputPath "PreCheck_$($env:COMPUTERNAME)_$stamp.log"

function Write-Section { param($Title)
    ""; ("=" * 70); "  $Title"; ("=" * 70) | ForEach-Object {
        Write-Host $_ -ForegroundColor Cyan
        Add-Content -LiteralPath $logPath -Value $_
    }
}
function Write-Line { param($Text, $Color = 'Gray')
    Write-Host $Text -ForegroundColor $Color
    Add-Content -LiteralPath $logPath -Value $Text
}
function Write-Result { param($Label, $Status, $Detail)
    $color = switch ($Status) { 'PASS' {'Green'} 'REVIEW' {'Yellow'} default {'Red'} }
    $line = "  [{0}] {1}" -f $Status.PadRight(6), $Label
    Write-Host $line -ForegroundColor $color
    if ($Detail) { Write-Line "        $Detail" 'Gray' }
    Add-Content -LiteralPath $logPath -Value $line
}

Write-Host "Pre-Check for change-reviewed remediation items — $($env:COMPUTERNAME)" -ForegroundColor Magenta
Write-Host "Log: $logPath" -ForegroundColor DarkGray
Add-Content -LiteralPath $logPath -Value "Pre-Check for change-reviewed remediation items — $($env:COMPUTERNAME) — $(Get-Date)"

$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isElevated) {
    Write-Host "NOTE: not running as Administrator — the SMB session and firewall checks below may show 'Access is denied'. Re-run from an elevated PowerShell for complete results." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
Write-Section "1. Windows Firewall — inbound allow rules (before enabling firewall)"
try {
    $rules = Get-NetFirewallRule -Enabled True -Direction Inbound -ErrorAction Stop |
        Where-Object { $_.Action -eq 'Allow' }
    $withPorts = $rules | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue |
        Where-Object { $_.LocalPort -and $_.LocalPort -ne 'Any' }
    Write-Line ("  {0} enabled inbound allow rule(s) found." -f $rules.Count)
    $rdp = $rules | Where-Object { $_.DisplayName -match 'Remote Desktop' }
    Write-Result 'RDP (3389) allow rule present' $(if ($rdp) {'PASS'} else {'REVIEW'}) $(if (-not $rdp) {'No enabled RDP rule found — add one before enabling the firewall, or you may lose remote access.'})
    Write-Line "  Sample of allowed ports (first 15):"
    ($withPorts | Select-Object -First 15 DisplayName, LocalPort | Format-Table -AutoSize | Out-String).Trim() -split "`n" | ForEach-Object { Write-Line "    $_" }
    Write-Result 'Review complete' 'REVIEW' 'Confirm the list above covers RDP + your monitoring/EDR agent + application ports before running -EnableFirewallProfiles.'
} catch { Write-Result 'Firewall rule check' 'REVIEW' "Could not query firewall rules: $($_.Exception.Message)" }

# ---------------------------------------------------------------------------
Write-Section "2. SMB minimum version — active sessions and their dialect"
try {
    $sessions = Get-SmbSession -ErrorAction Stop
    if (-not $sessions) {
        Write-Result 'Active SMB sessions' 'PASS' 'No active SMB sessions right now — safe to check again right before applying, but also review scheduled/off-hours jobs (backup, NAS) separately.'
    } else {
        $old = $sessions | Where-Object { $_.Dialect -in '2.0.2','2.1' }
        $sessions | Select-Object ClientComputerName, Dialect, NumOpens | Format-Table -AutoSize | Out-String -Width 200 | ForEach-Object { $_ -split "`n" } | ForEach-Object { Write-Line "    $_" }
        Write-Result 'Clients on SMB 2.0.2/2.1' $(if ($old) {'REVIEW'} else {'PASS'}) $(if ($old) {"$($old.Count) client(s) still on an older dialect — they will lose SMB access if minimum is raised to 3.0.0."})
    }
} catch { Write-Result 'SMB session check' 'REVIEW' "Could not query SMB sessions (module may be unavailable): $($_.Exception.Message)" }

# ---------------------------------------------------------------------------
Write-Section "3. NetBIOS — is this host resolvable without it (DNS)"
try {
    $dns = Resolve-DnsName -Name $env:COMPUTERNAME -ErrorAction Stop
    Write-Result 'DNS resolves this hostname' 'PASS' ("Resolved to: {0}" -f (($dns | Select-Object -ExpandProperty IPAddress) -join ', '))
} catch {
    Write-Result 'DNS resolves this hostname' 'REVIEW' 'DNS lookup failed — if anything reaches this server by NetBIOS/short name only, disabling NetBIOS will break it. Verify DNS records exist first.'
}
try {
    $nbt = & nbtstat -c 2>$null
    Write-Line "  nbtstat -c (local NetBIOS name cache):"
    ($nbt | Select-Object -First 10) | ForEach-Object { Write-Line "    $_" }
} catch {}

# ---------------------------------------------------------------------------
Write-Section "4. Account Lockout Threshold — current policy + local accounts"
try {
    $na = & net accounts 2>$null
    $na | ForEach-Object { Write-Line "    $_" }
    Write-Result 'Current policy captured' 'REVIEW' 'Manually confirm with the app/service owners which local or service accounts log in frequently enough to risk hitting a lower threshold (e.g. Splunk, ServiceNow, backup agents) before lowering it.'
    $locked = Get-LocalUser -ErrorAction SilentlyContinue | Where-Object { -not $_.Enabled -or $_.Name -eq 'Administrator' }
    if ($locked) { $locked | Select-Object Name, Enabled, PasswordExpires | Format-Table -AutoSize | Out-String | ForEach-Object { $_ -split "`n" } | ForEach-Object { Write-Line "    $_" } }
} catch { Write-Result 'Account policy check' 'REVIEW' "Could not read account policy: $($_.Exception.Message)" }

# ---------------------------------------------------------------------------
Write-Section "5. UAC hardening — scheduled tasks/services that rely on silent elevation"
try {
    $tasks = Get-ScheduledTask -ErrorAction Stop | Where-Object {
        $_.Principal.UserId -match 'Administrator|SYSTEM' -or $_.Principal.RunLevel -eq 'Highest'
    }
    Write-Result 'Elevated scheduled tasks found' $(if ($tasks) {'REVIEW'} else {'PASS'}) $(if ($tasks) {"$($tasks.Count) task(s) run elevated/as Administrator — confirm they don't depend on silent (no-prompt) elevation before enabling UAC hardening."})
    $tasks | Select-Object -First 15 TaskName, @{n='User';e={$_.Principal.UserId}}, @{n='RunLevel';e={$_.Principal.RunLevel}} |
        Format-Table -AutoSize | Out-String | ForEach-Object { $_ -split "`n" } | ForEach-Object { Write-Line "    $_" }
} catch { Write-Result 'Scheduled task check' 'REVIEW' "Could not query scheduled tasks: $($_.Exception.Message)" }

# ---------------------------------------------------------------------------
Write-Section "6. VBS & Credential Guard — platform prerequisites"
try {
    $sb = Confirm-SecureBootUEFI -ErrorAction Stop
    Write-Result 'Secure Boot enabled' $(if ($sb) {'PASS'} else {'REVIEW'}) $(if (-not $sb) {'Secure Boot is OFF — enable it on the VM in vCenter before running the VBS/Credential Guard items, or they will stay "configured, not running".'})
} catch {
    Write-Result 'Secure Boot enabled' 'REVIEW' 'Could not determine Secure Boot state (not UEFI, or access denied) — verify the VM firmware is set to EFI with Secure Boot in vCenter.'
}
try {
    $ci = Get-ComputerInfo -ErrorAction Stop
    Write-Result 'Firmware virtualization enabled'      $(if ($ci.HyperVRequirementVirtualizationFirmwareEnabled) {'PASS'} else {'REVIEW'}) ''
    Write-Result 'Second Level Address Translation'     $(if ($ci.HyperVRequirementSecondLevelAddressTranslation) {'PASS'} else {'REVIEW'}) ''
    Write-Result 'Data Execution Prevention available'  $(if ($ci.HyperVRequirementDataExecutionPreventionAvailable) {'PASS'} else {'REVIEW'}) ''
    Write-Line "  If any of the above show REVIEW, expose Secure Boot + hardware virtualization to this VM in vCenter first — VBS cannot start without them."
} catch { Write-Result 'Hyper-V/VBS platform check' 'REVIEW' "Could not query platform requirements: $($_.Exception.Message)" }

# ---------------------------------------------------------------------------
Write-Section "Done"
Write-Host "Review every [REVIEW] line above with the relevant owner before enabling that item's -Apply switch." -ForegroundColor Yellow
Write-Host "Full log saved to: $logPath" -ForegroundColor DarkGray
