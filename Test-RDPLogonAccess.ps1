<#
.SYNOPSIS
    Read-only RDP logon health check for one local account on a WORKGROUP server.
    Prints PASS / FAIL / WARN / INFO for every check, saves a CSV, and can compare the result
    against the CSV from a server where RDP still works.

.DESCRIPTION
    Built for "Authentication failed" from BeyondTrust PRA / plain RDP after monthly patching or
    hardening. Run it ON THE TARGET SERVER, in an elevated PowerShell, with the account the staff /
    Jump Item uses.

    SAFE: the script changes nothing. The only file it writes besides its own CSV is a temporary
    secedit export (to read user rights), which is deleted straight away.

    Checks
      Account     : exists, enabled, locked out, password expired, must change password, account
                    expired, logon hours, Remote Desktop Users / Administrators / Guests membership
      User rights : Allow / Deny log on through Remote Desktop Services,
                    Access this computer from the network / Deny (NLA needs the network-logon right)
      RDP service : RDP enabled, TermService running, port listening, firewall rule, NLA,
                    security layer, encryption level, RDP certificate expiry
      Auth/crypto : NTLM restrictions, LmCompatibilityLevel, TLS 1.2 server enabled, legacy TLS,
                    Credential Guard, LSA protection, lockout and password-age policy
      Evidence    : failed logons (4625) for the account with the decoded reason, last successful
                    RDP logon, last patches, pending reboot

.PARAMETER UserName
    Local account to test (the one that fails). Default: prompts.

.PARAMETER CompareTo
    Path of a CSV this script produced on a server where RDP works. Differences are printed.

.PARAMETER OutputFolder
    Where the CSV is saved. Default C:\temp (created if missing).

.PARAMETER Hours
    How far back to read failed logon events. Default 24.

.EXAMPLE
    # 1) On a WORKING server (same account name if it exists there)
    .\Test-RDPLogonAccess.ps1 -UserName staffuser
    # 2) Copy that CSV to the BROKEN server, then run there with -CompareTo
    .\Test-RDPLogonAccess.ps1 -UserName staffuser -CompareTo C:\temp\RDPCheck_GOODSRV_staffuser.csv
#>
[CmdletBinding()]
param(
    [string] $UserName,
    [string] $CompareTo,
    [string] $OutputFolder = 'C:\temp',
    [ValidateRange(1, 720)] [int] $Hours = 24
)

$ScriptBuild = 'Test-RDPLogonAccess build 2026-11-08.1'
Write-Host $ScriptBuild -ForegroundColor Cyan

$ErrorActionPreference = 'Stop'
$Results = New-Object System.Collections.Generic.List[object]

function Add-Check {
    param([string]$Area, [string]$Check, [ValidateSet('PASS', 'FAIL', 'WARN', 'INFO')][string]$Result,
          [string]$Value, [string]$Fix = '')
    $Results.Add([pscustomobject]@{
        Area = $Area; Check = $Check; Result = $Result; Value = $Value; Fix = $Fix
    })
}

function Get-Reg {
    param([string]$Path, [string]$Name)
    try { (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name } catch { $null }
}

function Show-Value { param($v) if ($null -eq $v -or "$v" -eq '') { '(not set)' } else { "$v" } }

# --- Pre-flight -----------------------------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
    IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host 'Run this in an ELEVATED PowerShell (Run as administrator). User rights and the Security log need it.' -ForegroundColor Red
    return
}
if (-not $UserName) { $UserName = Read-Host 'Local account to test (the one that fails)' }
$UserName = ($UserName -replace '^\.\\', '' -replace '^.*\\', '').Trim()
$Computer = $env:COMPUTERNAME
Write-Host ("Server: {0}   Account: {1}   Time: {2}`n" -f $Computer, $UserName, (Get-Date -Format 'yyyy-MM-dd HH:mm'))

$cs = Get-CimInstance Win32_ComputerSystem
Add-Check 'Server' 'Domain / workgroup' 'INFO' ("{0} = {1}" -f $(if ($cs.PartOfDomain) { 'Domain' } else { 'Workgroup' }), $cs.Domain)

# --- Account --------------------------------------------------------------------------------------
$user = $null
try { $user = Get-LocalUser -Name $UserName -ErrorAction Stop } catch { }
$userSids = @()

if (-not $user) {
    Add-Check 'Account' 'Account exists' 'FAIL' "No local account '$UserName'" 'Check the spelling in the Jump Item, or create the account.'
} else {
    Add-Check 'Account' 'Account exists' 'PASS' $user.SID.Value
    $userSids += $user.SID.Value

    if ($user.Enabled) { Add-Check 'Account' 'Account enabled' 'PASS' 'Enabled' }
    else { Add-Check 'Account' 'Account enabled' 'FAIL' 'Disabled' "Enable-LocalUser $UserName" }

    $wmiUser = Get-CimInstance Win32_UserAccount -Filter ("LocalAccount=True AND Name='{0}'" -f ($UserName -replace "'", "''"))
    if ($wmiUser -and $wmiUser.Lockout) {
        Add-Check 'Account' 'Not locked out' 'FAIL' 'LOCKED OUT' "net user $UserName /active:yes   (then find what keeps sending the wrong password)"
    } else { Add-Check 'Account' 'Not locked out' 'PASS' 'Not locked' }

    if ($null -eq $user.PasswordLastSet) {
        Add-Check 'Account' 'No "must change password at next logon"' 'FAIL' 'Must change password at next logon' `
            'NLA/PRA cannot show a change-password screen. Set a new password: Set-LocalUser, or untick the option in lusrmgr.msc.'
    } else { Add-Check 'Account' 'No "must change password at next logon"' 'PASS' ("Password last set {0:yyyy-MM-dd}" -f $user.PasswordLastSet) }

    if ($user.PasswordExpires -and $user.PasswordExpires -lt (Get-Date)) {
        Add-Check 'Account' 'Password not expired' 'FAIL' ("Expired {0:yyyy-MM-dd}" -f $user.PasswordExpires) `
            "Reset the password (and update it in PRA/Vault). Optional: Set-LocalUser $UserName -PasswordNeverExpires `$true"
    } else { Add-Check 'Account' 'Password not expired' 'PASS' $(if ($user.PasswordExpires) { "Expires {0:yyyy-MM-dd}" -f $user.PasswordExpires } else { 'Never expires' }) }

    if ($user.AccountExpires -and $user.AccountExpires -lt (Get-Date)) {
        Add-Check 'Account' 'Account not expired' 'FAIL' ("Expired {0:yyyy-MM-dd}" -f $user.AccountExpires) "Set-LocalUser $UserName -AccountNeverExpires"
    } else { Add-Check 'Account' 'Account not expired' 'PASS' 'Not expired' }

    $netUser = (net user $UserName 2>$null) -join "`n"
    $hoursLine = ($netUser -split "`n" | Where-Object { $_ -match '^Logon hours allowed' }) -replace '^Logon hours allowed\s*', ''
    if ($hoursLine -and $hoursLine.Trim() -ne 'All') {
        Add-Check 'Account' 'Logon hours' 'WARN' $hoursLine.Trim() "Logon hours are restricted. net user $UserName /times:all"
    } else { Add-Check 'Account' 'Logon hours' 'PASS' 'All' }
}

# Local group membership (ADSI copes with orphaned SIDs that break Get-LocalGroupMember)
$memberOf = @()
if ($user) {
    foreach ($g in Get-LocalGroup) {
        try {
            $adsi = [ADSI]("WinNT://./{0},group" -f $g.Name)
            foreach ($m in @($adsi.psbase.Invoke('Members'))) {
                $path = $m.GetType().InvokeMember('ADsPath', 'GetProperty', $null, $m, $null)
                if ($path -match ("/{0}/{1}$" -f [regex]::Escape($Computer), [regex]::Escape($UserName))) {
                    $memberOf += $g; break
                }
            }
        } catch { }
    }
    $userSids += $memberOf | ForEach-Object { $_.SID.Value }
    Add-Check 'Account' 'Local groups' 'INFO' $(if ($memberOf) { ($memberOf.Name -join ', ') } else { '(none)' })

    $isRdpUser = $memberOf | Where-Object { $_.SID.Value -in 'S-1-5-32-555', 'S-1-5-32-544' }
    if ($isRdpUser) { Add-Check 'Account' 'In Remote Desktop Users or Administrators' 'PASS' ($isRdpUser.Name -join ', ') }
    else { Add-Check 'Account' 'In Remote Desktop Users or Administrators' 'FAIL' 'Not a member' "Add-LocalGroupMember -SID S-1-5-32-555 -Member $UserName" }

    if ($memberOf | Where-Object { $_.SID.Value -eq 'S-1-5-32-546' }) {
        Add-Check 'Account' 'Not in Guests' 'FAIL' 'Member of Guests' "Guests is in 'Deny log on through RDS' after hardening. Remove-LocalGroupMember -SID S-1-5-32-546 -Member $UserName"
    } else { Add-Check 'Account' 'Not in Guests' 'PASS' 'Not a member' }

    # Well-known SIDs every local logon carries
    $userSids += 'S-1-1-0', 'S-1-5-11', 'S-1-5-113', 'S-1-5-2', 'S-1-5-14', 'S-1-5-4'
    if ($memberOf | Where-Object { $_.SID.Value -eq 'S-1-5-32-544' }) { $userSids += 'S-1-5-114' }
}

# --- User rights (secedit export, read-only) ------------------------------------------------------
$tmp = Join-Path $env:TEMP ("rdpcheck_{0}.inf" -f [guid]::NewGuid())
$rights = @{}
try {
    secedit /export /areas USER_RIGHTS /cfg $tmp /quiet | Out-Null
    foreach ($line in Get-Content $tmp) {
        if ($line -match '^(Se\w+)\s*=\s*(.*)$') {
            # Copy the captures first: the -like test below must not clobber $Matches
            $rightName = $Matches[1]; $entries = $Matches[2]
            $rights[$rightName] = @($entries -split ',' | ForEach-Object {
                $e = $_.Trim().TrimStart('*')
                if ($e -like 'S-1-*') { $e } else {
                    try { ([Security.Principal.NTAccount]$e).Translate([Security.Principal.SecurityIdentifier]).Value } catch { $e }
                }
            })
        }
    }
} catch {
    Add-Check 'User rights' 'Read user rights' 'WARN' "secedit export failed: $($_.Exception.Message)"
} finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }

function Show-Sids { param($sids)
    ($sids | ForEach-Object {
        $sid = $_
        try { ([Security.Principal.SecurityIdentifier]$sid).Translate([Security.Principal.NTAccount]).Value } catch { $sid }
    }) -join ', '
}

if ($user -and $rights.Count) {
    $pairs = @(
        @{ Allow = 'SeRemoteInteractiveLogonRight'; Deny = 'SeDenyRemoteInteractiveLogonRight'; Label = 'Log on through Remote Desktop Services' },
        @{ Allow = 'SeNetworkLogonRight'; Deny = 'SeDenyNetworkLogonRight'; Label = 'Access this computer from the network (needed for NLA)' }
    )
    foreach ($p in $pairs) {
        $allow = @($rights[$p.Allow]); $deny = @($rights[$p.Deny])
        $hitAllow = $allow | Where-Object { $_ -in $userSids }
        $hitDeny = $deny | Where-Object { $_ -in $userSids }
        if ($hitAllow) { Add-Check 'User rights' ("Allow: " + $p.Label) 'PASS' ("Granted via " + (Show-Sids $hitAllow)) }
        else { Add-Check 'User rights' ("Allow: " + $p.Label) 'FAIL' ("Granted only to: " + (Show-Sids $allow)) "secpol.msc > Local Policies > User Rights Assignment: add the account or its group." }
        if ($hitDeny) { Add-Check 'User rights' ("Deny: " + $p.Label) 'FAIL' ("DENIED via " + (Show-Sids $hitDeny)) "secpol.msc > User Rights Assignment > Deny ...: remove that entry (a deny always wins over an allow)." }
        else { Add-Check 'User rights' ("Deny: " + $p.Label) 'PASS' $(if ($deny) { "Deny list: " + (Show-Sids $deny) } else { 'Deny list empty' }) }
    }
}

# --- RDP service ----------------------------------------------------------------------------------
$tsKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
$tcpKey = "$tsKey\WinStations\RDP-Tcp"
$polKey = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'

$deny = Get-Reg $polKey 'fDenyTSConnections'; if ($null -eq $deny) { $deny = Get-Reg $tsKey 'fDenyTSConnections' }
if ($deny -eq 0) { Add-Check 'RDP' 'RDP enabled' 'PASS' 'fDenyTSConnections=0' }
else { Add-Check 'RDP' 'RDP enabled' 'FAIL' "fDenyTSConnections=$deny" 'Enable Remote Desktop in System Properties (or the GPO).' }

$svc = Get-Service TermService
if ($svc.Status -eq 'Running') { Add-Check 'RDP' 'TermService running' 'PASS' 'Running' }
else { Add-Check 'RDP' 'TermService running' 'FAIL' $svc.Status 'Start-Service TermService' }

$port = Get-Reg $tcpKey 'PortNumber'; if (-not $port) { $port = 3389 }
$listen = Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction SilentlyContinue
if ($listen) { Add-Check 'RDP' "Listening on port $port" 'PASS' 'Listening' }
else { Add-Check 'RDP' "Listening on port $port" 'FAIL' 'Not listening' 'Restart TermService; check RDP-Tcp listener.' }

try {
    $fw = Get-NetFirewallPortFilter -Protocol TCP -ErrorAction Stop | Where-Object { $_.LocalPort -contains "$port" } |
        Get-NetFirewallRule | Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' -and $_.Action -eq 'Allow' }
    if ($fw) { Add-Check 'RDP' 'Firewall allows RDP inbound' 'PASS' (($fw.DisplayName | Select-Object -Unique) -join '; ') }
    else { Add-Check 'RDP' 'Firewall allows RDP inbound' 'WARN' "No enabled inbound allow rule for TCP $port" 'Enable the "Remote Desktop" firewall rules.' }
} catch { Add-Check 'RDP' 'Firewall allows RDP inbound' 'INFO' 'Could not read firewall rules' }

$nla = Get-Reg $polKey 'UserAuthentication'; $nlaSrc = 'GPO'
if ($null -eq $nla) { $nla = Get-Reg $tcpKey 'UserAuthentication'; $nlaSrc = 'local' }
Add-Check 'RDP' 'NLA required' 'INFO' ("{0} ({1})" -f $(if ($nla -eq 1) { 'Yes' } else { 'No' }), $nlaSrc) `
    'If this differs from the working server, it was changed by patching/hardening (Qualys Rdp:UserAuthentication).'

$sl = Get-Reg $polKey 'SecurityLayer'; $slSrc = 'GPO'
if ($null -eq $sl) { $sl = Get-Reg $tcpKey 'SecurityLayer'; $slSrc = 'local' }
$slName = switch ($sl) { 0 { 'RDP' } 1 { 'Negotiate' } 2 { 'TLS (SSL)' } default { Show-Value $sl } }
Add-Check 'RDP' 'Security layer' 'INFO' ("{0} ({1})" -f $slName, $slSrc) 'If this differs from the working server, it was changed by hardening (Qualys Rdp:SecurityLayer).'

$enc = Get-Reg $polKey 'MinEncryptionLevel'; if ($null -eq $enc) { $enc = Get-Reg $tcpKey 'MinEncryptionLevel' }
Add-Check 'RDP' 'Min encryption level' 'INFO' (Show-Value $enc)

try {
    $ts = Get-CimInstance -Namespace root\cimv2\TerminalServices -ClassName Win32_TSGeneralSetting -Filter "TerminalName='RDP-Tcp'"
    $thumb = $ts.SSLCertificateSHA1Hash
    $cert = Get-ChildItem 'Cert:\LocalMachine\Remote Desktop', 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1
    if (-not $cert) { Add-Check 'RDP' 'RDP certificate valid' 'WARN' "Certificate $thumb not found" 'Delete the stale cert from Cert:\LocalMachine\Remote Desktop and restart TermService to regenerate.' }
    elseif ($cert.NotAfter -lt (Get-Date)) { Add-Check 'RDP' 'RDP certificate valid' 'FAIL' ("Expired {0:yyyy-MM-dd}" -f $cert.NotAfter) 'Delete the expired cert from Cert:\LocalMachine\Remote Desktop and restart TermService to regenerate.' }
    else { Add-Check 'RDP' 'RDP certificate valid' 'PASS' ("Valid until {0:yyyy-MM-dd}" -f $cert.NotAfter) }
} catch { Add-Check 'RDP' 'RDP certificate valid' 'INFO' 'Could not read RDP certificate' }

# --- Authentication / crypto ----------------------------------------------------------------------
$lsa = 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa'
$msv = "$lsa\MSV1_0"

$rin = Get-Reg $msv 'RestrictReceivingNTLMTraffic'
if ($rin -eq 2) { Add-Check 'Auth' 'Incoming NTLM allowed' 'FAIL' 'RestrictReceivingNTLMTraffic=2 (deny all)' 'Workgroup RDP needs NTLM. Set "Network security: Restrict NTLM: Incoming NTLM traffic" to Allow all.' }
else { Add-Check 'Auth' 'Incoming NTLM allowed' 'PASS' ("RestrictReceivingNTLMTraffic={0}" -f (Show-Value $rin)) }

Add-Check 'Auth' 'LmCompatibilityLevel' 'INFO' (Show-Value (Get-Reg $lsa 'LmCompatibilityLevel'))
Add-Check 'Auth' 'NTLMMinServerSec' 'INFO' (Show-Value (Get-Reg $msv 'NTLMMinServerSec'))
Add-Check 'Auth' 'Credential Guard (LsaCfgFlags)' 'INFO' (Show-Value (Get-Reg $lsa 'LsaCfgFlags'))
Add-Check 'Auth' 'LSA protection (RunAsPPL)' 'INFO' (Show-Value (Get-Reg $lsa 'RunAsPPL'))

$sch = 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols'
if ((Get-Reg "$sch\TLS 1.2\Server" 'Enabled') -eq 0) { Add-Check 'Auth' 'TLS 1.2 server enabled' 'FAIL' 'Disabled' "Set $sch\TLS 1.2\Server Enabled=1 and reboot." }
else { Add-Check 'Auth' 'TLS 1.2 server enabled' 'PASS' 'Enabled' }
$legacy = foreach ($p in 'TLS 1.0', 'TLS 1.1') { "{0}={1}" -f $p, $(if ((Get-Reg "$sch\$p\Server" 'Enabled') -eq 0) { 'off' } else { 'on' }) }
Add-Check 'Auth' 'Legacy TLS (server)' 'INFO' ($legacy -join ', ') 'If off here but on at the working server, an old Jumpoint/RDP client may still need it.'

$na = (net accounts 2>$null)
$lockLine = ($na | Where-Object { $_ -match '^Lockout threshold' }) -replace '^Lockout threshold:\s*', ''
$maxAge = ($na | Where-Object { $_ -match '^Maximum password age' }) -replace '^Maximum password age \(days\):\s*', ''
Add-Check 'Policy' 'Account lockout threshold' 'INFO' (Show-Value $lockLine)
Add-Check 'Policy' 'Maximum password age (days)' 'INFO' (Show-Value $maxAge)

# --- Evidence: event log --------------------------------------------------------------------------
$codes = @{
    'C0000064' = 'User does not exist'; 'C000006A' = 'Wrong password'; 'C0000234' = 'Account locked out'
    'C0000072' = 'Account disabled'; 'C000006F' = 'Outside allowed logon hours'; 'C0000070' = 'Workstation restriction'
    'C0000193' = 'Account expired'; 'C0000071' = 'Password expired'; 'C0000224' = 'Must change password at next logon'
    'C000015B' = 'Logon type not granted (User Rights Assignment / Deny right)'; 'C000006E' = 'Account restriction'
    'C0000133' = 'Clock skew'; 'C0000413' = 'Authentication firewall'; 'C000006D' = 'Bad user name or authentication info (with SubStatus 0: NTLM version/security mismatch, not a wrong password)'
}
$ms = $Hours * 3600000
$xp = "*[System[EventID=4625 and TimeCreated[timediff(@SystemTime) <= $ms]]] and *[EventData[Data[@Name='TargetUserName']='$UserName']]"
try {
    $fails = @(Get-WinEvent -LogName Security -FilterXPath $xp -ErrorAction Stop)
} catch { $fails = @() }
if ($fails.Count -eq 0) {
    Add-Check 'Evidence' "Failed logons (4625), last $Hours h" 'INFO' 'None found for this account' `
        'If staff still get "Authentication failed", the failure is before Windows logon (NLA/TLS) or the Jump Item sends a different user name. Try one connection and re-run.'
} else {
    $x = ([xml]$fails[0].ToXml()).Event.EventData.Data
    $d = @{}; foreach ($n in $x) { $d[$n.Name] = $n.'#text' }
    $sub = ("{0}" -f $d['SubStatus']).ToUpper() -replace '^0X', ''
    if ($sub -eq '' -or $sub -eq '0' -or $sub -eq '00000000') { $sub = ("{0}" -f $d['Status']).ToUpper() -replace '^0X', '' }
    $reason = if ($codes.ContainsKey($sub)) { $codes[$sub] } else { 'Unknown code' }
    Add-Check 'Evidence' "Failed logons (4625), last $Hours h" 'FAIL' `
        ("{0} failures. Latest {1:yyyy-MM-dd HH:mm}: 0x{2} = {3} | Status {4} SubStatus {5} | LogonType {6} | Auth {7} ({8}) | From {9} {10}" -f
            $fails.Count, $fails[0].TimeCreated, $sub, $reason, $d['Status'], $d['SubStatus'], $d['LogonType'],
            $d['AuthenticationPackageName'], (Show-Value $d['LmPackageName']), $d['WorkstationName'], $d['IpAddress']) `
        'This is the reason Windows rejected the logon - fix the matching FAIL above.'
}

$xpOk = "*[System[EventID=4624 and TimeCreated[timediff(@SystemTime) <= 2592000000]]] and *[EventData[Data[@Name='TargetUserName']='$UserName'] and EventData[Data[@Name='LogonType']='10']]"
try { $ok = Get-WinEvent -LogName Security -FilterXPath $xpOk -MaxEvents 1 -ErrorAction Stop } catch { $ok = $null }
Add-Check 'Evidence' 'Last successful RDP logon (30 days)' 'INFO' $(if ($ok) { "{0:yyyy-MM-dd HH:mm}" -f $ok.TimeCreated } else { 'None in 30 days' })

$hf = Get-HotFix -ErrorAction SilentlyContinue | Where-Object InstalledOn | Sort-Object InstalledOn -Descending | Select-Object -First 3
Add-Check 'Server' 'Last patches' 'INFO' (($hf | ForEach-Object { "{0} ({1:yyyy-MM-dd})" -f $_.HotFixID, $_.InstalledOn }) -join ', ')

$pending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
           (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')
$boot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
if ($pending) { Add-Check 'Server' 'No pending reboot' 'WARN' ("Reboot pending. Last boot {0:yyyy-MM-dd HH:mm}" -f $boot) 'Reboot to finish patching, then test again.' }
else { Add-Check 'Server' 'No pending reboot' 'PASS' ("Last boot {0:yyyy-MM-dd HH:mm}" -f $boot) }

# --- Output ---------------------------------------------------------------------------------------
$color = @{ PASS = 'Green'; FAIL = 'Red'; WARN = 'Yellow'; INFO = 'Gray' }
foreach ($r in $Results) {
    Write-Host ("[{0}] " -f $r.Result) -ForegroundColor $color[$r.Result] -NoNewline
    Write-Host ("{0,-12} {1,-55} {2}" -f $r.Area, $r.Check, $r.Value)
}

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$csv = Join-Path $OutputFolder ("RDPCheck_{0}_{1}.csv" -f $Computer, $UserName)
$Results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

$bad = @($Results | Where-Object { $_.Result -in 'FAIL', 'WARN' })
Write-Host "`n================ RESULT ================" -ForegroundColor Cyan
if ($bad.Count -eq 0) {
    Write-Host 'OVERALL: PASS - no blocking problem found for this account.' -ForegroundColor Green
} else {
    Write-Host ("OVERALL: FAILED - {0} problem(s). Likely root cause(s):" -f $bad.Count) -ForegroundColor Red
    foreach ($b in $bad) {
        Write-Host (" - [{0}] {1}: {2}" -f $b.Result, $b.Check, $b.Value) -ForegroundColor $color[$b.Result]
        if ($b.Fix) { Write-Host ("     Fix: {0}" -f $b.Fix) }
    }
}

if ($CompareTo) {
    if (-not (Test-Path $CompareTo)) {
        Write-Host "`nCompare file not found: $CompareTo" -ForegroundColor Yellow
    } else {
        $good = Import-Csv $CompareTo
        $ignore = 'Last patches', 'Last successful RDP logon (30 days)', 'No pending reboot', 'Account exists'
        Write-Host ("`n====== DIFFERENCES vs WORKING SERVER ({0}) ======" -f (Split-Path $CompareTo -Leaf)) -ForegroundColor Cyan
        $diffs = 0
        foreach ($r in $Results) {
            if ($r.Check -in $ignore -or $r.Check -like 'Failed logons*') { continue }
            $g = $good | Where-Object { $_.Check -eq $r.Check } | Select-Object -First 1
            if (-not $g) { continue }
            if ($g.Result -ne $r.Result -or $g.Value -ne $r.Value) {
                $diffs++
                Write-Host (" * {0}" -f $r.Check) -ForegroundColor Yellow
                Write-Host ("     This server   : [{0}] {1}" -f $r.Result, $r.Value)
                Write-Host ("     Working server: [{0}] {1}" -f $g.Result, $g.Value)
            }
        }
        if ($diffs -eq 0) { Write-Host 'No setting differences found.' -ForegroundColor Green }
    }
}

Write-Host "`nSaved: $csv" -ForegroundColor Cyan
Write-Host 'Nothing was changed on this server.' -ForegroundColor Cyan
