<#
.SYNOPSIS
    Qualys VA / Policy Compliance remediation for Windows Server 2025 guests through VMware Tools
    Guest Operations (Copy-VMGuestFile + Invoke-VMScript). No WinRM dependency.

.DESCRIPTION
    Built from CHG0002363 (Compliance_Report + VA_Report, scan 2026-10-07/08).
    Default run is AUDIT ONLY - nothing is changed in the guest. Add -Apply to remediate.

    Tiers
      Low     : applied with -Apply.
      Medium  : detected always, applied only with -Apply -IncludeMediumRisk.
      PerUser : HKU policy values for every local profile, applied only with -Apply -IncludeUserProfiles.
    Domain-scoped controls (DC-only audit subcategories, firewall Domain profile logging, picture password
    for domain users, DC authentication to unlock) are reported as not applicable on a WORKGROUP server and
    left unchanged. -IncludeDomainScoped sets the harmless ones anyway if Security rejects the N/A request.
    High-risk / owner-decision findings (LAPS, Kerberos etypes, LocalAccountTokenFilterPolicy,
    DisableDomainCreds, SAM RPC password change, Network Protection block, RDP clipboard, WU policy)
    are intentionally NOT in this script. See the HTML action plan.

    The script never reboots a server.

    Targets are read from C:\temp\vmlist.txt (one vCenter VM name per line) unless -VMName or
    -VMListPath is given.

.EXAMPLE
    # 1) Assessment of every VM in C:\temp\vmlist.txt (no change)
    .\Invoke-QualysRemediation.ps1
    # 2) Everything on one server first
    .\Invoke-QualysRemediation.ps1 -VMName <first VM> -Apply -IncludeUserProfiles -IncludeMediumRisk
    # 3) Then all servers in the list (already-compliant settings are left alone)
    .\Invoke-QualysRemediation.ps1 -Apply -IncludeUserProfiles -IncludeMediumRisk
#>
[CmdletBinding()]
param(
    [string[]]$VMName,
    [string]$VMListPath = 'C:\temp\vmlist.txt',
    [Parameter(Mandatory)][pscredential]$GuestCredential,
    [switch]$Apply,
    [switch]$IncludeMediumRisk,
    [switch]$IncludeUserProfiles,
    [switch]$IncludeDomainScoped,
    [string]$LogDir = (Join-Path $PSScriptRoot 'Logs')
)

$ErrorActionPreference = 'Stop'

# Targets: -VMName wins; otherwise one vCenter VM name per line in -VMListPath (blank lines and # comments ignored)
if (-not $VMName) {
    if (-not (Test-Path -LiteralPath $VMListPath)) { throw "No -VMName given and VM list file not found: $VMListPath" }
    $VMName = @(Get-Content -LiteralPath $VMListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -Unique)
    if ($VMName.Count -eq 0) { throw "VM list file is empty: $VMListPath" }
}

$GuestPayload = @'
param([string]$Mode = 'Audit', [string]$Tiers = 'Low', [int]$UserProfiles = 0, [int]$DomainScoped = 0, [string]$OutFile)
$ErrorActionPreference = 'Stop'
$apply   = ($Mode -eq 'Apply')
$tierSet = $Tiers -split ','
$rows    = New-Object System.Collections.Generic.List[object]
$Spec    = New-Object System.Collections.Generic.List[object]
$work    = Join-Path $env:TEMP ('qrem_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work -Force | Out-Null

function Add-Row($Ids, $Grp, $Item, $Old, $Expected, $New, $Status, $Note) {
    $rows.Add([pscustomobject]@{
        Computer = $env:COMPUTERNAME; ControlIds = "$Ids"; Group = "$Grp"; Item = "$Item"
        OldValue = "$Old"; Expected = "$Expected"; NewValue = "$New"; Status = "$Status"; Note = "$Note"
        Time = (Get-Date).ToString('s')
    })
}
function qR($Ids, $Tier, $Grp, $Path, $Name, $Kind, $Value, $Ok = 'eq', $Reboot = $false, $Cond = '') {
    $Spec.Add(@{T='Reg'; Ids=$Ids; Tier=$Tier; Grp=$Grp; Path=$Path; Name=$Name; Kind=$Kind; Value=$Value; Ok=$Ok; Reboot=$Reboot; Cond=$Cond}) }
function qA($Ids, $Tier, $Grp, $Guid, $Name, $Need, $Cond = '') { $Spec.Add(@{T='Audit'; Ids=$Ids; Tier=$Tier; Grp=$Grp; Guid=$Guid; Name=$Name; Need=$Need; Cond=$Cond}) }
function qU($Ids, $Tier, $Grp, $Right, $Sid)       { $Spec.Add(@{T='Right'; Ids=$Ids; Tier=$Tier; Grp=$Grp; Right=$Right; Sid=$Sid}) }
function qS($Ids, $Tier, $Grp, $Key, $Value, $Ok)  { $Spec.Add(@{T='SecPol'; Ids=$Ids; Tier=$Tier; Grp=$Grp; Key=$Key; Value=$Value; Ok=$Ok}) }
function qUH($Ids, $Tier, $Grp, $Path, $Name)      { $Spec.Add(@{T='User'; Ids=$Ids; Tier=$Tier; Grp=$Grp; Path=$Path; Name=$Name; Value=1}) }

#SPEC-BEGIN
$P  = 'SOFTWARE\Policies\Microsoft'
$CV = 'SOFTWARE\Microsoft\Windows\CurrentVersion\Policies'
$FW = "$P\WindowsFirewall"
$WD = "$P\Windows Defender"

# G03 Interactive logon / session
qR 1190 Low G03 "$CV\System" DontDisplayLastUserName DWord 1
qR 1376 Low G03 "$CV\System" DisableCAD DWord 0
qR 1196 Low G03 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' ScreenSaverGracePeriod String '0' 'le:5'
qS 4469 Low G03 ForceLogoffWhenHourExpire 1 'eq'

# G04 SMB server hardening (no protocol / signing-requirement change)
qR 1199  Low G04 'SYSTEM\CurrentControlSet\Services\LanManServer\Parameters' AutoDisconnect DWord 15 'range:1-15'
qR 1381  Low G04 'SYSTEM\CurrentControlSet\Services\LanManServer\Parameters' EnableSecuritySignature DWord 1
qR 29573 Low G04 "$P\Windows\LanmanServer" EnableAuthRateLimiter DWord 1
qR 29577 Low G04 "$P\Windows\LanmanServer" InvalidAuthenticationDelayTimeInMs DWord 2000 'ge:2000'
qR 29574 Low G04 "$P\Windows\Bowser" EnableMailslots DWord 0
qR 29581 Low G04 "$P\Windows\NetworkProvider" EnableMailslots DWord 0

# G05 Network / RPC / printing / internet-communication policies
qR 1513  Low G05 "$P\Windows NT\Rpc" EnableAuthEpResolution DWord 1
qR 2607  Low G05 "$P\Windows\Network Connections" NC_ShowSharedAccessUI DWord 0
qR 2616  Low G05 "$P\Windows\Network Connections" NC_AllowNetBridge_NLA DWord 0
qR 3899  Low G05 "$P\Windows NT\Terminal Services" fAllowToGetHelp DWord 0
qR 3920  Low G05 "$CV\Explorer" NoWebServices DWord 1
qR 3922  Low G05 "$P\Windows NT\Printers" DisableWebPnPDownload DWord 1
qR 4143  Low G05 "$P\Windows\Windows Error Reporting" Disabled DWord 1
qR 4742  Low G05 'SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters' TcpMaxDataRetransmissions DWord 3 'eq' $true
qR 5267  Low G05 'SYSTEM\CurrentControlSet\Control\Lsa\pku2u' AllowOnlineID DWord 0
qR 27616 Low G05 "$P\Windows\Group Policy\{827D319E-6EAC-11D2-A4EA-00C04F79F83A}" NoBackgroundPolicy DWord 0
qR 30171 Low G05 "$P\Windows NT\DNSClient" DisableIPv6DefaultDnsServers DWord 1
qR 31687 Low G05 "$P\Windows NT\Printers\IPP" RequireIpps DWord 1

# G06 Privacy / cloud content / sign-in experience (machine policies)
qR 8274  Low G06 "$P\Windows\System" EnableSmartScreen DWord 1
qR 8399  Low G06 "$P\Windows\System" DisableLockScreenAppNotifications DWord 1
qR 9388  Low G35 "$P\Windows\System" BlockDomainPicturePassword DWord 1 'eq' $false DomainOnly
qR 8420  Low G06 "$P\Windows\LocationAndSensors" DisableLocation DWord 1
qR 8425  Low G06 "$P\Windows\CredUI" DisablePasswordReveal DWord 1
qR 9402  Low G06 "$P\Windows\AdvertisingInfo" DisabledByGroupPolicy DWord 1
qR 9404  Low G06 "$P\Windows\OneDrive" DisableFileSyncNGSC DWord 1
qR 10016 Low G06 "$P\Windows NT\CurrentVersion\Software Protection Platform" NoGenTicket DWord 1
qR 10082 Low G06 "$P\Windows\TabletPC" PreventHandwritingDataSharing DWord 1
qR 10085 Low G06 "$P\Control Panel\International" BlockUserInputMethodsForSignIn DWord 1
qR 10086 Low G06 "$P\Windows\WDI\{9c5a40da-b965-4fc3-8781-88dd50a6299d}" ScenarioExecutionEnabled DWord 0
qR 11034 Low G06 "$P\Windows\Device Metadata" PreventDeviceMetadataFromNetwork DWord 1
qR 11205 Low G06 "$P\Power\PowerSettings\f15576e8-98b7-4186-b944-eafa664402d9" DCSettingIndex DWord 0
qR 11206 Low G06 "$P\Power\PowerSettings\f15576e8-98b7-4186-b944-eafa664402d9" ACSettingIndex DWord 0
qR 13940 Low G06 "$CV\Explorer" AllowOnlineTips DWord 0
qR 13941 Low G06 "$P\Windows\Messaging" AllowMessageSync DWord 0
qR 13942 Low G06 "$P\PushToInstall" DisablePushToInstall DWord 1
qR 13943 Low G06 "$P\Windows\Windows Search" AllowCloudSearch DWord 0
qR 14878 Low G06 "$P\Windows\CurrentVersion\PushNotifications" NoCloudApplicationNotification DWord 1
qR 23130 Low G06 "$P\Windows\DataCollection" EnableOneSettingsAuditing DWord 1

# G07 App Installer / winget
qR 25900 Low G07 "$P\Windows\AppInstaller" EnableAppInstaller DWord 0
qR 25901 Low G07 "$P\Windows\AppInstaller" EnableExperimentalFeatures DWord 0
qR 25903 Low G07 "$P\Windows\AppInstaller" EnableMSAppInstallerProtocol DWord 0
qR 30182 Low G07 "$P\Windows\AppInstaller" EnableLocalArchiveMalwareScanOverride DWord 0
qR 30183 Low G07 "$P\Windows\AppInstaller" EnableBypassCertificatePinningForMicrosoftStore DWord 0
qR 30184 Low G07 "$P\Windows\AppInstaller" EnableWindowsPackageManagerCommandLineInterfaces DWord 0

# G08 Event log size / process command-line auditing
qR 9014 Low G08 "$P\Windows\EventLog\Setup" MaxSize DWord 32768 'ge:32768'
qR 9440 Low G08 "$CV\System\Audit" ProcessCreationIncludeCmdLine_Enabled DWord 1

# G09 Advanced audit policy (only the missing Success/Failure flag is added)
# G33 = subcategories that only generate events on domain controllers: skipped on WORKGROUP unless -IncludeDomainScoped
qA 4473 Low G09 '{0CCE9213-69AE-11D9-BED3-505054503030}' 'IPsec Driver' SF
qA 4476 Low G09 '{0CCE9216-69AE-11D9-BED3-505054503030}' 'Logoff' S
qA 4503 Low G09 '{0CCE9231-69AE-11D9-BED3-505054503030}' 'Authorization Policy Change' S
qA 4508 Low G33 '{0CCE9236-69AE-11D9-BED3-505054503030}' 'Computer Account Management' S DomainOnly
qA 4510 Low G33 '{0CCE9238-69AE-11D9-BED3-505054503030}' 'Distribution Group Management' S DomainOnly
qA 4511 Low G09 '{0CCE9239-69AE-11D9-BED3-505054503030}' 'Application Group Management' SF
qA 4512 Low G09 '{0CCE923A-69AE-11D9-BED3-505054503030}' 'Other Account Management Events' S
qA 4513 Low G33 '{0CCE923B-69AE-11D9-BED3-505054503030}' 'Directory Service Access' F DomainOnly
qA 4514 Low G33 '{0CCE923C-69AE-11D9-BED3-505054503030}' 'Directory Service Changes' S DomainOnly
qA 4518 Low G33 '{0CCE9240-69AE-11D9-BED3-505054503030}' 'Kerberos Service Ticket Operations' SF DomainOnly
qA 4522 Low G33 '{0CCE9242-69AE-11D9-BED3-505054503030}' 'Kerberos Authentication Service' SF DomainOnly

# G10 Windows Firewall logging (logging only - no rule or profile state change)
# G34 = Domain profile, never active on a WORKGROUP server: skipped unless -IncludeDomainScoped
qR 1524 Low G34 "$FW\DomainProfile\Logging"  LogDroppedPackets DWord 1 'eq' $false DomainOnly
qR 1525 Low G34 "$FW\DomainProfile\Logging"  LogFilePath String '%SystemRoot%\System32\logfiles\firewall\domainfw.log' 'any' $false DomainOnly
qR 1526 Low G34 "$FW\DomainProfile\Logging"  LogFileSize DWord 16384 'ge:16384' $false DomainOnly
qR 1527 Low G34 "$FW\DomainProfile\Logging"  LogSuccessfulConnections DWord 1 'eq' $false DomainOnly
qR 8161 Low G10 "$FW\PrivateProfile\Logging" LogFilePath String '%SystemRoot%\System32\logfiles\firewall\privatefw.log'
qR 8162 Low G10 "$FW\PrivateProfile\Logging" LogSuccessfulConnections DWord 1
qR 8163 Low G10 "$FW\PrivateProfile\Logging" LogDroppedPackets DWord 1
qR 8165 Low G10 "$FW\PublicProfile\Logging"  LogDroppedPackets DWord 1
qR 8166 Low G10 "$FW\PublicProfile\Logging"  LogFilePath String '%SystemRoot%\System32\logfiles\firewall\publicfw.log'
qR 8167 Low G10 "$FW\PublicProfile\Logging"  LogSuccessfulConnections DWord 1
qR 8168 Low G10 "$FW\PublicProfile\Logging"  LogFileSize DWord 16384 'ge:16384'

# G11 User rights - add BUILTIN\Guests (S-1-5-32-546) only; existing members preserved
qU 2196 Low G11 SeDenyNetworkLogonRight '*S-1-5-32-546'
qU 2198 Low G11 SeDenyServiceLogonRight '*S-1-5-32-546'
qU 2200 Low G11 SeDenyRemoteInteractiveLogonRight '*S-1-5-32-546'

# G12 Microsoft Defender policy baseline (reporting / audit-mode / non-blocking)
qR 9469  Low G12 "$WD\Reporting" DisableGenericRePorts DWord 1
qR 9526  Low G12 "$WD\Scan" CheckForSignaturesBeforeRunningScan DWord 1
qR 9537  Low G12 "$WD\Scan" DisableEmailScanning DWord 0
qR 30180 Low G12 "$WD\Scan" DaysUntilAggressiveCatchupQuickScan DWord 7
qR 29592 Low G12 "$WD" HideExclusionsFromLocalUsers DWord 1
qR 29593 Low G12 "$WD\NIS" EnableConvertWarnToBlock DWord 1
qR 29594 Low G12 "$WD\Real-Time Protection" OobeEnableRtpAndSigUpdate DWord 1
qR 29595 Low G12 "$WD\Reporting" EnableDynamicSignatureDroppedEventReporting DWord 1
qR 29746 Low G12 "$WD\Features" PassiveRemediation DWord 1
qR 30179 Low G12 "$WD\Remediation\Behavioral Network Blocks\Remote Encryption Protection" RemoteEncryptionProtectionAggressiveness DWord 1 'in:1,2'
qR 31004 Low G12 "$WD\Remediation\Behavioral Network Blocks\Remote Encryption Protection" RemoteEncryptionProtectionConfiguredState DWord 2 'in:1,2'
qR 30187 Low G12 "$WD\Remediation\Behavioral Network Blocks\Brute Force Protection" BruteForceProtectionAggressiveness DWord 1 'in:1,2'
qR 30188 Low G12 "$WD\Remediation\Behavioral Network Blocks\Brute Force Protection" BruteForceProtectionConfiguredState DWord 2 'in:1,2'
qR 31121 Low G12 "$WD\Windows Defender Exploit Guard\ASR" ExploitGuard_ASR_Rules DWord 1
qR 31121 Low G12 "$WD\Windows Defender Exploit Guard\ASR\Rules" '01443614-cd74-433a-b99e-2ecdc07bfc25' String '2' 'in:1,2'

# G14 (low part) allow minimum password length policy above 14
qR 19057 Low G14 'SYSTEM\CurrentControlSet\Control\SAM' RelaxMinimumPasswordLengthLimits DWord 1

# ---------------- MEDIUM (only with -IncludeMediumRisk) ----------------
# G01 CVE-2013-3900 strict Authenticode padding check (VA QID 378332)
qR 'QID378332' Medium G01 'SOFTWARE\Microsoft\Cryptography\Wintrust\Config' EnableCertPaddingCheck DWord 1
qR 'QID378332' Medium G01 'SOFTWARE\Wow6432Node\Microsoft\Cryptography\Wintrust\Config' EnableCertPaddingCheck DWord 1
# G13 Defender - behaviour changing
qR 19065 Medium G13 "$WD\MpEngine" EnableFileHashComputation DWord 1
qR 29596 Medium G13 "$WD\Scan" QuickScanIncludeExclusions DWord 1
qR 9472  Medium G13 "$WD\Scan" ScanParameters DWord 1 'in:1,2'
qR 9492  Medium G13 "$WD\Scan" ScheduleTime DWord 120 'ge:1'
qR 9571  Medium G13 "$WD\Scan" ScheduleQuickScanTime DWord 120 'ge:1'
qR 32072 Medium G13 "$WD\Windows Defender Exploit Guard\ASR\Rules" '33ddedf1-c6e0-47cb-833e-de6133960387' String '1'
# G14 Account policy
qS 1071  Medium G14 MinimumPasswordLength 14 'ge:14'
qS 26502 Medium G14 AllowAdministratorLockout 1 'eq'
# G15..G22
qR 2587  Medium G15 "$CV\System" ConsentPromptBehaviorAdmin DWord 2 'in:1,2'
qR 5264  Medium G16 'SYSTEM\CurrentControlSet\Services\LanManServer\Parameters' SMBServerNameHardeningLevel DWord 1 'in:1,2'
qR 5268  Medium G17 "$P\Windows\WinRM\Service\WinRS" AllowRemoteShellAccess DWord 0
qR 10028 Medium G18 "$P\Windows\PowerShell\Transcription" EnableTranscripting DWord 1
qR 12026 Medium G19 "$P\Windows\DataCollection" DisableEnterpriseAuthProxy DWord 1
qR 31693 Medium G19 'SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings' DisableProxyAuthenticationSchemes DWord 256 'in:256,287'
qR 21711 Medium G20 "$P\Windows NT\Printers" RegisterSpoolerRemoteRpcEndPoint DWord 2
qR 10404 Medium G21 "$P\Windows NT\Terminal Services" UserAuthentication DWord 1 'eq' $false 'Rdp:UserAuthentication:1'
qR 10431 Medium G21 "$P\Windows NT\Terminal Services" SecurityLayer DWord 2 'eq' $false 'Rdp:SecurityLayer:2'
qR 1377  Medium G22 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' ForceUnlockLogon DWord 1 'eq' $false 'Domain'

# ---------------- PER-USER (only with -IncludeUserProfiles) ----------------
qUH 9827  PerUser G23 'Software\Policies\Microsoft\Assistance\Client\1.0' NoImplicitFeedback
qUH 9830  PerUser G23 'Software\Microsoft\Windows\CurrentVersion\Policies\Explorer' NoInplaceSharing
qUH 9836  PerUser G23 'Software\Policies\Microsoft\WindowsMediaPlayer' PreventCodecDownload
qUH 9842  PerUser G23 'Software\Policies\Microsoft\Windows\CurrentVersion\PushNotifications' NoToastApplicationNotificationOnLockScreen
qUH 11203 PerUser G23 'Software\Policies\Microsoft\Windows\CloudContent' DisableThirdPartySuggestions
qUH 23138 PerUser G23 'Software\Policies\Microsoft\Windows\CloudContent' DisableSpotlightCollectionOnDesktop
#SPEC-END

function Test-Ok($Cur, $Ok, $Value) {
    if ($null -eq $Cur) { return $false }
    $s = "$Cur".Trim(); $n = 0L
    if ($Ok -eq 'eq')  { return ($s -eq "$Value") }
    if ($Ok -eq 'any') { return ($s.Length -gt 0) }
    if ($Ok -match '^ge:(\d+)$')        { $a = [long]$Matches[1]; return ([long]::TryParse($s, [ref]$n) -and $n -ge $a) }
    if ($Ok -match '^le:(\d+)$')        { $a = [long]$Matches[1]; return ([long]::TryParse($s, [ref]$n) -and $n -le $a) }
    if ($Ok -match '^range:(\d+)-(\d+)$') { $a = [long]$Matches[1]; $b = [long]$Matches[2]; return ([long]::TryParse($s, [ref]$n) -and $n -ge $a -and $n -le $b) }
    if ($Ok -match '^in:(.+)$')         { return (($Matches[1] -split ',') -contains $s) }
    throw "Unknown comparison '$Ok'"
}
function Get-Reg($Path, $Name) {
    $k = "HKLM:\$Path"
    if (-not (Test-Path -LiteralPath $k)) { return $null }
    $key = Get-Item -LiteralPath $k
    if ($key.GetValueNames() -notcontains $Name) { return $null }
    return $key.GetValue($Name, $null, 'DoNotExpandEnvironmentNames')
}
function Show($v) { if ($null -eq $v) { '(not set)' } else { "$v" } }
function Get-Gate($Tier) {
    if (-not $apply) { return 'audit' }
    if ($Tier -eq 'PerUser') { if ($UserProfiles -eq 1) { return 'go' } else { return 'tier' } }
    if ($tierSet -contains $Tier) { return 'go' }
    return 'tier'
}
function Export-SecPol($Area) {
    $f = Join-Path $work "$Area.inf"
    $o = & secedit.exe /export /cfg $f /areas $Area /quiet
    if ($LASTEXITCODE -ne 0) { throw "secedit /export $Area failed (exit $LASTEXITCODE): $o" }
    $h = @{}
    foreach ($l in (Get-Content -LiteralPath $f)) { if ($l -match '^\s*([A-Za-z]\w*)\s*=\s*(.*)$') { $h[$Matches[1]] = $Matches[2].Trim() } }
    return $h
}
function Import-SecPol($Area, $Section, $Key, $Value) {
    $inf = Join-Path $work ("set_$Key.inf"); $db = Join-Path $work ("set_$Key.sdb")
    $txt = "[Unicode]`r`nUnicode=yes`r`n[Version]`r`nsignature=`"`$CHICAGO`$`"`r`nRevision=1`r`n[$Section]`r`n$Key = $Value`r`n"
    Set-Content -LiteralPath $inf -Value $txt -Encoding Unicode
    $o = & secedit.exe /configure /db $db /cfg $inf /areas $Area /quiet
    if ($LASTEXITCODE -ne 0) { throw "secedit /configure $Key failed (exit $LASTEXITCODE): $o" }
}
function Get-AuditFlags($Guid) {
    $o = & auditpol.exe /get /subcategory:"$Guid" /r
    if ($LASTEXITCODE -ne 0) { throw "auditpol /get $Guid failed (exit $LASTEXITCODE): $o" }
    $line = @($o | Where-Object { $_ -match [regex]::Escape($Guid) })[0]
    if (-not $line) { throw "auditpol returned no row for $Guid" }
    return ($line -split ',')[4]
}

$elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$cs = Get-CimInstance Win32_ComputerSystem
$os = Get-CimInstance Win32_OperatingSystem
$isDomain = [bool]$cs.PartOfDomain
$naWg = (-not $isDomain) -and ($DomainScoped -ne 1)

try {
    # ---- context / evidence rows ----
    $ubr = Get-Reg 'SOFTWARE\Microsoft\Windows NT\CurrentVersion' 'UBR'
    Add-Row '' INFO 'OS' '' '' "$($os.Caption) build $($os.BuildNumber).$ubr" 'Info' "ProductType=$($os.ProductType) (2=Domain Controller, 3=Server)"
    Add-Row '' INFO 'Domain membership' '' '' "PartOfDomain=$isDomain; Domain/Workgroup=$($cs.Domain)" 'Info' 'Domain GPO may override local settings'
    Add-Row '' INFO 'Elevated guest session' '' 'True' "$elevated" 'Info' ''
    $ntds = Get-Service -Name NTDS -ErrorAction SilentlyContinue
    Add-Row '19058,30197' G32 'NTDS service (AD DS role)' '' 'absent on non-DC' $(if ($ntds) { "present ($($ntds.Status))" } else { 'absent' }) 'Info' 'Evidence for DC-only controls'
    $vdll = 'C:\Program Files\Veeam\Endpoint Backup\Veeam.Backup.Agent.dll'
    if (Test-Path -LiteralPath $vdll) {
        $vv = (Get-Item -LiteralPath $vdll).VersionInfo.FileVersion
        $vst = 'Manual Review'; try { if ([version]$vv -ge [version]'6.3.2.1302') { $vst = 'Already Compliant' } } catch { $vst = 'Manual Review' }
        Add-Row 'QID386127' G02 'Veeam.Backup.Agent.dll version' $vv '>= 6.3.2.1302 (Veeam KB4771)' $vv $vst 'Patch via Veeam B&R / agent installer - not scripted'
    } else { Add-Row 'QID386127' G02 'Veeam.Backup.Agent.dll' '(file not found)' '>= 6.3.2.1302' '' 'Info' 'Agent not found at default path' }
    # third-party / agent versions reported by Qualys VA (evidence only - patching is not scripted)
    $verChecks = @(
        @('G40', 'PowerShell 7 pwsh.exe', 'C:\Program Files\PowerShell\7\pwsh.exe', '7.5.10'),
        @('G45', 'Mozilla Firefox firefox.exe', 'C:\Program Files\Mozilla Firefox\firefox.exe', '157.0'),
        @('G44', 'Splunk Universal Forwarder splunkd.exe', 'C:\Program Files\SplunkUniversalForwarder\bin\splunkd.exe', ''),
        @('G43', 'Binalyze AIR agent AIR.exe', 'C:\Program Files (x86)\Binalyze\AIR\agent\AIR.exe', ''))
    foreach ($vc in $verChecks) {
        if (Test-Path -LiteralPath $vc[2]) {
            $fv = (Get-Item -LiteralPath $vc[2]).VersionInfo.FileVersion
            $st = 'Info'
            if ($vc[3]) { $st = 'Manual Review'; try { if ([version](($fv -split '[^\d\.]')[0]) -ge [version]$vc[3]) { $st = 'Already Compliant' } } catch { $st = 'Manual Review' } }
            Add-Row '' $vc[0] $vc[1] $fv $(if ($vc[3]) { ">= $($vc[3])" } else { 'vendor current build' }) $fv $st 'Patch through the product owner - not scripted'
        }
    }
    $arc = Get-Reg 'SOFTWARE\Microsoft\Azure Connected Machine Agent' 'Version'
    if ($arc) { $st = 'Manual Review'; try { if ([version]$arc -ge [version]'1.63') { $st = 'Already Compliant' } } catch { $st = 'Manual Review' }
        Add-Row '' G41 'Azure Connected Machine Agent version' $arc '>= 1.63' $arc $st 'Patch via Microsoft Update / agent MSI - not scripted' }
    foreach ($uk in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall') {
        if (Test-Path -LiteralPath $uk) { foreach ($sk in (Get-ChildItem -LiteralPath $uk)) { $dn = $sk.GetValue('DisplayName'); if ($dn -and $dn -match '^Java') { Add-Row '' G46 "Installed: $dn" $sk.GetValue('DisplayVersion') 'Java 8 >= 8u503 (8.0.5030.x) or removed' $sk.GetValue('DisplayVersion') 'Info' "$($sk.GetValue('InstallLocation'))" } } }
    }
    try { foreach ($fp in (Get-NetFirewallProfile -PolicyStore ActiveStore)) { Add-Row '' G36 "Firewall profile $($fp.Name) (effective)" '' 'Enabled = True' "Enabled=$($fp.Enabled)" 'Info' 'Firewall state is never changed by this script' } }
    catch { Add-Row '' G36 'Firewall profiles' '' '' 'Get-NetFirewallProfile unavailable' 'Info' $_.Exception.Message }
    try { $ncp = (Get-NetConnectionProfile | ForEach-Object { "$($_.InterfaceAlias)=$($_.NetworkCategory)" }) -join '; '; Add-Row '' G34 'Active network category' '' 'Private/Public (never DomainAuthenticated on WORKGROUP)' $ncp 'Info' '' }
    catch { Add-Row '' G34 'Active network category' '' '' 'Get-NetConnectionProfile unavailable' 'Info' $_.Exception.Message }
    try { $mp = Get-MpComputerStatus; Add-Row '' INFO 'Microsoft Defender' '' '' "AMRunningMode=$($mp.AMRunningMode); RTP=$($mp.RealTimeProtectionEnabled); Tamper=$($mp.IsTamperProtected)" 'Info' 'Defender policy values only take effect when Defender AV is installed' }
    catch { Add-Row '' INFO 'Microsoft Defender' '' '' 'Get-MpComputerStatus unavailable' 'Info' $_.Exception.Message }

    $lgpo = Join-Path $env:SystemRoot 'System32\GroupPolicy'
    $acsv = Join-Path $lgpo 'Machine\Microsoft\Windows NT\Audit\audit.csv'
    if (Test-Path -LiteralPath $acsv) { $al = @(Get-Content -LiteralPath $acsv | Where-Object { $_ -match '\{0cce' }); Add-Row '' LGPO 'Local GPO audit file (audit.csv)' '' '' "present, $($al.Count) subcategories defined" 'Info' 'auditpol changes for subcategories NOT in this file are wiped at policy refresh - set them in gpedit.msc > Advanced Audit Policy Configuration' }
    else { Add-Row '' LGPO 'Local GPO audit file (audit.csv)' '' '' 'absent' 'Info' '' }
    Add-Row '' LGPO 'SCENoApplyLegacyAuditPolicy' '' '1' (Show (Get-Reg 'SYSTEM\CurrentControlSet\Control\Lsa' 'SCENoApplyLegacyAuditPolicy')) 'Info' 'Must be 1 or legacy category auditing overrides subcategory settings'
    $rpol = Join-Path $lgpo 'Machine\Registry.pol'
    if (Test-Path -LiteralPath $rpol) {
        $ptxt = [Text.Encoding]::Unicode.GetString([IO.File]::ReadAllBytes($rpol))
        $hits = @($Spec | Where-Object { $_.T -eq 'Reg' -and $ptxt.IndexOf($_.Name, [StringComparison]::OrdinalIgnoreCase) -ge 0 } | ForEach-Object { "$($_.Ids):$($_.Name)" })
        Add-Row '' LGPO 'Local GPO Registry.pol' '' '' "present; scripted values also defined in it: $(if ($hits) { $hits -join ', ' } else { 'none' })" 'Info' 'Values listed here are re-applied from the local GPO - change them in gpedit.msc or they revert'
    } else { Add-Row '' LGPO 'Local GPO Registry.pol' '' '' 'absent' 'Info' '' }

    if ($apply -and -not $elevated) { throw 'Guest session is not elevated - no change attempted. Use the built-in Administrator or an account that receives a full token through VMware Tools.' }

    $rights = $null; $secpol = $null

    foreach ($it in $Spec) {
        $gate = Get-Gate $it.Tier
        try {
            switch ($it.T) {
                'Reg' {
                    $item = "HKLM\$($it.Path)\$($it.Name)"
                    $old = Get-Reg $it.Path $it.Name
                    if (Test-Ok $old $it.Ok $it.Value) { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value (Show $old) 'Already Compliant' ''; break }
                    if ($it.Cond -eq 'DomainOnly' -and $naWg) { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value '' 'Manual Review' 'Not applicable: WORKGROUP server (domain-scoped control) - not changed; request N/A. Use -IncludeDomainScoped to set it anyway'; break }
                    if ($it.Cond -eq 'Domain' -and -not $isDomain) { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value '' 'Manual Review' 'Not applicable: WORKGROUP server - not changed; request N/A'; break }
                    if ($gate -eq 'audit') { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value '' 'Non-Compliant (AuditOnly)' "$($it.Tier) tier"; break }
                    if ($gate -eq 'tier')  { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value '' 'Manual Review' "$($it.Tier)-risk item not selected"; break }
                    if ($it.Cond -eq 'Domain' -and -not $isDomain) { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value '' 'Manual Review' 'Server is not domain-joined: control not meaningful - request exception'; break }
                    if ($it.Cond -match '^Rdp:(\w+):(\d+)$') {
                        $eff = Get-Reg 'SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' $Matches[1]
                        if ("$eff" -ne $Matches[2]) { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value '' 'Manual Review' "Effective RDP-Tcp $($Matches[1])=$(Show $eff): enforcing would change live RDP behaviour - owner approval needed"; break }
                    }
                    $k = "HKLM:\$($it.Path)"
                    if (-not (Test-Path -LiteralPath $k)) { New-Item -Path $k -Force | Out-Null }
                    New-ItemProperty -LiteralPath $k -Name $it.Name -PropertyType $it.Kind -Value $it.Value -Force | Out-Null
                    $new = Get-Reg $it.Path $it.Name
                    if (-not (Test-Ok $new $it.Ok $it.Value)) { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value (Show $new) 'Failed' 'Value did not verify after write' }
                    elseif ($it.Reboot) { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value (Show $new) 'Reboot Required' 'Set and verified; effective after restart' }
                    else { Add-Row $it.Ids $it.Grp $item (Show $old) $it.Value (Show $new) 'Remediated' '' }
                }
                'Audit' {
                    $item = "auditpol: $($it.Name)"
                    if (-not $elevated) { Add-Row $it.Ids $it.Grp $item '' $it.Need '' 'Manual Review' 'auditpol needs an elevated session'; break }
                    $old = Get-AuditFlags $it.Guid
                    if (@('No Auditing','Success','Failure','Success and Failure') -notcontains $old) { Add-Row $it.Ids $it.Grp $item $old $it.Need '' 'Manual Review' 'Unrecognised (non-English?) auditpol output'; break }
                    $hasS = $old -match 'Success'; $hasF = $old -match 'Failure'
                    $needS = $it.Need -match 'S'; $needF = $it.Need -match 'F'
                    if ((-not $needS -or $hasS) -and (-not $needF -or $hasF)) { Add-Row $it.Ids $it.Grp $item $old $it.Need $old 'Already Compliant' ''; break }
                    if ($it.Cond -eq 'DomainOnly' -and $naWg) { Add-Row $it.Ids $it.Grp $item $old $it.Need '' 'Manual Review' 'Not applicable: WORKGROUP server (subcategory only generates events on domain controllers) - not changed; request N/A. Use -IncludeDomainScoped to set it anyway'; break }
                    if ($gate -eq 'audit') { Add-Row $it.Ids $it.Grp $item $old $it.Need '' 'Non-Compliant (AuditOnly)' "$($it.Tier) tier"; break }
                    if ($gate -eq 'tier')  { Add-Row $it.Ids $it.Grp $item $old $it.Need '' 'Manual Review' 'Tier not selected'; break }
                    $a = @('/set', "/subcategory:$($it.Guid)")
                    if ($needS -and -not $hasS) { $a += '/success:enable' }
                    if ($needF -and -not $hasF) { $a += '/failure:enable' }
                    $o = & auditpol.exe @a
                    if ($LASTEXITCODE -ne 0) { throw "auditpol /set failed (exit $LASTEXITCODE): $o" }
                    $new = Get-AuditFlags $it.Guid
                    $ok = ((-not $needS) -or ($new -match 'Success')) -and ((-not $needF) -or ($new -match 'Failure'))
                    Add-Row $it.Ids $it.Grp $item $old $it.Need $new $(if ($ok) { 'Remediated' } else { 'Failed' }) ''
                }
                'Right' {
                    $item = "User right $($it.Right)"
                    if (-not $elevated) { Add-Row $it.Ids $it.Grp $item '' "contains $($it.Sid)" '' 'Manual Review' 'secedit needs an elevated session'; break }
                    if ($null -eq $rights) { $rights = Export-SecPol 'USER_RIGHTS' }
                    $old = $rights[$it.Right]
                    $members = @(); if ($old) { $members = @($old -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                    if ($members -contains $it.Sid -or $members -contains 'Guests') { Add-Row $it.Ids $it.Grp $item (Show $old) "contains $($it.Sid)" (Show $old) 'Already Compliant' ''; break }
                    if ($gate -eq 'audit') { Add-Row $it.Ids $it.Grp $item (Show $old) "contains $($it.Sid)" '' 'Non-Compliant (AuditOnly)' "$($it.Tier) tier"; break }
                    if ($gate -eq 'tier')  { Add-Row $it.Ids $it.Grp $item (Show $old) "contains $($it.Sid)" '' 'Manual Review' 'Tier not selected'; break }
                    Import-SecPol 'USER_RIGHTS' 'Privilege Rights' $it.Right ((@($members) + $it.Sid) -join ',')
                    $rights = Export-SecPol 'USER_RIGHTS'
                    $new = $rights[$it.Right]
                    $ok = $new -and (@($new -split ',' | ForEach-Object { $_.Trim() }) -contains $it.Sid)
                    Add-Row $it.Ids $it.Grp $item (Show $old) "contains $($it.Sid)" (Show $new) $(if ($ok) { 'Remediated' } else { 'Failed' }) 'Existing members preserved'
                }
                'SecPol' {
                    $item = "Security policy [System Access] $($it.Key)"
                    if (-not $elevated) { Add-Row $it.Ids $it.Grp $item '' $it.Value '' 'Manual Review' 'secedit needs an elevated session'; break }
                    if ($null -eq $secpol) { $secpol = Export-SecPol 'SECURITYPOLICY' }
                    if (-not $secpol.ContainsKey($it.Key)) { Add-Row $it.Ids $it.Grp $item '(not exported)' $it.Value '' 'Manual Review' 'Setting not present in secedit export - Requires validation'; break }
                    $old = $secpol[$it.Key]
                    if (Test-Ok $old $it.Ok $it.Value) { Add-Row $it.Ids $it.Grp $item $old $it.Value $old 'Already Compliant' ''; break }
                    if ($gate -eq 'audit') { Add-Row $it.Ids $it.Grp $item $old $it.Value '' 'Non-Compliant (AuditOnly)' "$($it.Tier) tier"; break }
                    if ($gate -eq 'tier')  { Add-Row $it.Ids $it.Grp $item $old $it.Value '' 'Manual Review' "$($it.Tier)-risk item not selected"; break }
                    Import-SecPol 'SECURITYPOLICY' 'System Access' $it.Key $it.Value
                    $secpol = Export-SecPol 'SECURITYPOLICY'
                    $new = $secpol[$it.Key]
                    Add-Row $it.Ids $it.Grp $item $old $it.Value $new $(if (Test-Ok $new $it.Ok $it.Value) { 'Remediated' } else { 'Failed' }) $(if ($isDomain) { 'Domain-joined: a domain GPO may re-apply its own value' } else { '' })
                }
                'User' { }   # handled per profile below
            }
        } catch { Add-Row $it.Ids $it.Grp "$($it.T): $($it.Name)$($it.Key)$($it.Right)" '' '' '' 'Failed' $_.Exception.Message }
    }

    # ---- per-user policy values (HKU) ----
    $uSpec = @($Spec | Where-Object { $_.T -eq 'User' })
    $gateU = Get-Gate 'PerUser'
    $HKU = [Microsoft.Win32.Registry]::Users
    $targets = New-Object System.Collections.Generic.List[object]
    $plKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    foreach ($pk in (Get-ChildItem -LiteralPath $plKey | Where-Object { $_.PSChildName -match '^S-1-5-21-' })) {
        $targets.Add(@{ Label = $pk.PSChildName; Sid = $pk.PSChildName; Hive = (Join-Path ($pk.GetValue('ProfileImagePath')) 'NTUSER.DAT') })
    }
    $defDir = (Get-Item -LiteralPath $plKey).GetValue('Default')
    if ($defDir) { $targets.Add(@{ Label = 'Default profile (new users)'; Sid = $null; Hive = (Join-Path $defDir 'NTUSER.DAT') }) }

    foreach ($t in $targets) {
        $root = $null; $tmp = $false
        try {
            if ($t.Sid -and ($HKU.GetSubKeyNames() -contains $t.Sid)) { $root = $t.Sid }
            elseif (-not ($gateU -eq 'go' -or ((-not $apply) -and $UserProfiles -eq 1))) {
                Add-Row (($uSpec | ForEach-Object { $_.Ids }) -join ';') G23 "HKU profile $($t.Label)" '(hive not loaded)' '1' '' 'Manual Review' 'Not evaluated: this user is logged off. Add -IncludeUserProfiles (also works without -Apply) to load the profile and verify it'
                continue
            }
            else {
                if (-not (Test-Path -LiteralPath $t.Hive)) { Add-Row '' G23 "HKU profile $($t.Label)" '' '' '' 'Manual Review' "NTUSER.DAT not found: $($t.Hive)"; continue }
                $o = & reg.exe load 'HKU\QREM_TMP' $t.Hive
                if ($LASTEXITCODE -ne 0) { throw "reg load failed (exit $LASTEXITCODE): $o" }
                $root = 'QREM_TMP'; $tmp = $true
            }
            foreach ($u in $uSpec) {
                $item = "HKU\$($t.Label)\$($u.Path)\$($u.Name)"
                $k = $HKU.OpenSubKey("$root\$($u.Path)", $false)
                $old = $null; if ($k) { $old = $k.GetValue($u.Name, $null); $k.Close() }
                if ("$old" -eq '1') { Add-Row $u.Ids G23 $item '1' '1' '1' 'Already Compliant' ''; continue }
                if ($gateU -eq 'audit') { Add-Row $u.Ids G23 $item (Show $old) '1' '' 'Non-Compliant (AuditOnly)' 'PerUser tier'; continue }
                if ($gateU -eq 'tier')  { Add-Row $u.Ids G23 $item (Show $old) '1' '' 'Manual Review' 'Per-user items not selected (-IncludeUserProfiles)'; continue }
                $w = $HKU.CreateSubKey("$root\$($u.Path)")
                $w.SetValue($u.Name, 1, [Microsoft.Win32.RegistryValueKind]::DWord); $w.Close()
                $v = $HKU.OpenSubKey("$root\$($u.Path)", $false); $new = $v.GetValue($u.Name, $null); $v.Close()
                Add-Row $u.Ids G23 $item (Show $old) '1' (Show $new) $(if ("$new" -eq '1') { 'Remediated' } else { 'Failed' }) ''
            }
        } catch { Add-Row '' G23 "HKU profile $($t.Label)" '' '' '' 'Failed' $_.Exception.Message }
        finally {
            if ($tmp) {
                [gc]::Collect(); [gc]::WaitForPendingFinalizers()
                $o = & reg.exe unload 'HKU\QREM_TMP'
                if ($LASTEXITCODE -ne 0) { Add-Row '' G23 "HKU profile $($t.Label)" '' '' '' 'Failed' "reg unload failed (exit $LASTEXITCODE): $o - hive stays loaded until reboot" }
            }
        }
    }
}
catch { Add-Row '' FATAL 'Guest payload' '' '' '' 'Failed' $_.Exception.Message }
finally {
    $rows | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $OutFile -Encoding UTF8
    Remove-Item -LiteralPath $work -Recurse -Force
}
'@

# ------------------------------- host side -------------------------------
if (-not (Get-Command Invoke-VMScript -ErrorAction SilentlyContinue)) { throw 'VMware PowerCLI (Invoke-VMScript) is not available in this session.' }
if (-not $global:DefaultVIServers -or @($global:DefaultVIServers | Where-Object IsConnected).Count -eq 0) { throw 'Not connected to vCenter. Run Connect-VIServer -Server <vCenter IP> first.' }

New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
$stamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$mode    = if ($Apply) { 'Apply' } else { 'Audit' }
$tiers   = if ($IncludeMediumRisk) { 'Low,Medium' } else { 'Low' }
$userFlg = if ($IncludeUserProfiles) { 1 } else { 0 }
$domFlg  = if ($IncludeDomainScoped) { 1 } else { 0 }
$logFile = Join-Path $LogDir "QualysRemediation_${mode}_$stamp.log"
$csvFile = Join-Path $LogDir "QualysRemediation_${mode}_$stamp.csv"
$payloadLocal = Join-Path $LogDir "qrem_payload_$stamp.ps1"
Set-Content -LiteralPath $payloadLocal -Value $GuestPayload -Encoding UTF8

function Write-Log([string]$Message, [string]$Level = 'INFO') {
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Level, $Message
    Add-Content -LiteralPath $logFile -Value $line
    Write-Host $line
}

$guestDir = 'C:\Windows\Temp\QRem'
$all = New-Object System.Collections.Generic.List[object]
Write-Log "Mode=$mode Tiers=$tiers UserProfiles=$userFlg Targets=$($VMName -join ',') Operator=$env:USERNAME"

foreach ($name in $VMName) {
    try {
        $vm = Get-VM -Name $name
        if (@($vm).Count -ne 1) { throw "Expected exactly one VM named '$name', found $(@($vm).Count)." }
        if ($vm.PowerState -ne 'PoweredOn') { throw "VM is $($vm.PowerState)." }
        if ($vm.ExtensionData.Guest.ToolsRunningStatus -ne 'guestToolsRunning') { throw "VMware Tools not running ($($vm.ExtensionData.Guest.ToolsRunningStatus))." }
        Write-Log "[$name] copying payload"
        Copy-VMGuestFile -VM $vm -Source $payloadLocal -Destination "$guestDir\" -LocalToGuest -GuestCredential $GuestCredential -Force
        $guestPs  = "$guestDir\$(Split-Path $payloadLocal -Leaf)"
        $guestOut = "$guestDir\result_$stamp.json"
        $cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$guestPs`" -Mode $mode -Tiers $tiers -UserProfiles $userFlg -DomainScoped $domFlg -OutFile `"$guestOut`""
        Write-Log "[$name] executing in guest ($mode)"
        $r = Invoke-VMScript -VM $vm -GuestCredential $GuestCredential -ScriptType Bat -ScriptText $cmd
        if ($r.ExitCode -ne 0) { Write-Log "[$name] guest exit code $($r.ExitCode): $($r.ScriptOutput)" 'WARN' }
        $localOut = Join-Path $LogDir "result_${name}_$stamp.json"
        Copy-VMGuestFile -VM $vm -Source $guestOut -Destination $localOut -GuestToLocal -GuestCredential $GuestCredential -Force
        # Windows PowerShell 5.1 returns the JSON array as ONE object from the pipeline - enumerate it explicitly
        $parsed = ConvertFrom-Json -InputObject (Get-Content -LiteralPath $localOut -Raw)
        $res = @($parsed | ForEach-Object { $_ })
        if ($res.Count -eq 0) { throw "Guest returned no result rows ($localOut)." }
        foreach ($x in $res) {
            $x | Add-Member -NotePropertyName VM -NotePropertyValue $name -Force
            $all.Add($x)
            Write-Log ("[{0}] {1,-26} {2} | old={3} | new={4} | {5}" -f $name, $x.Status, $x.Item, $x.OldValue, $x.NewValue, $x.Note) $(if ($x.Status -eq 'Failed') { 'ERROR' } else { 'INFO' })
        }
        $null = Invoke-VMScript -VM $vm -GuestCredential $GuestCredential -ScriptType Bat -ScriptText "rmdir /s /q `"$guestDir`""
    }
    catch {
        Write-Log "[$name] $($_.Exception.Message)" 'ERROR'
        $all.Add([pscustomobject]@{ VM=$name; Computer=''; ControlIds=''; Group='HOST'; Item='PowerCLI execution'; OldValue=''; Expected=''; NewValue=''; Status='Failed'; Note=$_.Exception.Message; Time=(Get-Date).ToString('s') })
    }
}

Remove-Item -LiteralPath $payloadLocal -Force
$all | Select-Object VM, Computer, Group, ControlIds, Item, OldValue, Expected, NewValue, Status, Note, Time | Export-Csv -LiteralPath $csvFile -NoTypeInformation -Encoding UTF8
Write-Log "Results: $csvFile"
$all | Group-Object VM, Status | Sort-Object Name | Select-Object Count, Name | Format-Table -AutoSize | Out-String | Write-Host
if (@($all | Where-Object Status -eq 'Reboot Required').Count) { Write-Log 'One or more settings need a restart to take effect. This script does not reboot.' 'WARN' }
