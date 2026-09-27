#Requires -Version 5.1
<#
.SYNOPSIS
    READ-ONLY AD reachability diagnostics for the migration pre-flight failure
    "Invoke-PreFlightCheck error: Source domain could not be reached."

.DESCRIPTION
    Makes NO changes to the device, DNS, secure channel or AD.
    The only thing written is the transcript log (OutDir).
    Note: 'nltest /dsgetdc ... /force' just bypasses the DC-locator cache and re-runs discovery;
    it does not change any configuration.

    Run on the PILOT DEVICE in an elevated PowerShell. Then re-run as SYSTEM
    (migration agent context), e.g.:  psexec -s -i powershell.exe -ExecutionPolicy Bypass -File <this.ps1>

.EXAMPLE
    .\Invoke-ADReachabilityCheck-ReadOnly.ps1
#>
[CmdletBinding()]
param(
    [string]  $Domain  = 'amccinemasksa.net',
    [string]  $NetBIOS = 'AMCCINEMASKSA',
    [string[]]$DCs     = @('10.7.62.5','10.7.7.13','10.7.7.14','10.7.30.8'),
    [string]  $OutDir  = $PSScriptRoot
)

if (-not $OutDir) { $OutDir = $env:TEMP }
# TCP ports: DNS, Kerberos, RPC EPM, NetBIOS-SSN, LDAP, SMB, Kpasswd, LDAPS, GC, GC-SSL, ADWS (Get-AD* cmdlets)
$Ports = @(53,88,135,139,389,445,464,636,3268,3269,9389)

$Log = Join-Path $OutDir ("ADCheck_{0}_{1}.txt" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'))
Start-Transcript -Path $Log | Out-Null

$Summary = New-Object System.Collections.Generic.List[object]
function Section($t) { Write-Host ""; Write-Host ("===== {0} =====" -f $t) -ForegroundColor Cyan }
function Run($label, [scriptblock]$sb) {
    Write-Host ("--- {0}" -f $label) -ForegroundColor Yellow
    try { & $sb 2>&1 | Out-String -Width 200 | Write-Host }
    catch { Write-Host ("ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red }
}
function Result($area, $check, [bool]$pass, $detail, $ifFail) {
    $Summary.Add([pscustomobject]@{
        Area = $area; Check = $check; Result = $(if ($pass) {'PASS'} else {'FAIL'})
        Detail = $detail; IfFail = $(if ($pass) {''} else {$ifFail})
    })
}
function NlTest([string[]]$a) {
    $out = & nltest.exe @a 2>&1 | Out-String
    [pscustomobject]@{ Ok = ($LASTEXITCODE -eq 0); Out = $out }
}

# ---------------------------------------------------------------- 1. CONTEXT
Section '1. CONTEXT'
Run 'Running as'   { whoami; whoami /groups | Select-String 'S-1-5-18|Domain Admins|BUILTIN\\Administrators' }
Run 'Hostname'     { hostname }
$cs = Get-CimInstance Win32_ComputerSystem
Run 'Domain membership' { $cs | Select-Object Name,Domain,PartOfDomain,DomainRole | Format-List }
Result 'Client' 'Device is domain-joined to source domain' ($cs.PartOfDomain -and $cs.Domain -eq $Domain) `
    ("PartOfDomain={0}, Domain={1}" -f $cs.PartOfDomain, $cs.Domain) 'Device is not (or no longer) joined to the source domain - migration tool cannot reach it by design.'
Run 'dsregcmd'      { dsregcmd /status | Select-String 'DomainJoined','AzureAdJoined','EnterpriseJoined','DomainName','TenantName','DeviceId' }
Run 'ipconfig /all' { ipconfig /all }
Run 'DNS servers'   { Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.ServerAddresses } | Format-Table InterfaceAlias,ServerAddresses -AutoSize }
Run 'DNS suffix / search list' { Get-DnsClientGlobalSetting | Format-List SuffixSearchList,UseSuffixSearchList; Get-DnsClient | Where-Object ConnectionSpecificSuffix | Format-Table InterfaceAlias,ConnectionSpecificSuffix -AutoSize }
Run 'Routes (default)' { Get-NetRoute -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue | Format-Table InterfaceAlias,NextHop,RouteMetric -AutoSize }
$prof = @(Get-NetConnectionProfile)
Run 'Network profile' { $prof | Format-Table InterfaceAlias,Name,NetworkCategory,IPv4Connectivity -AutoSize }
$domAuth = @($prof | Where-Object { $_.NetworkCategory -eq 'DomainAuthenticated' }).Count -gt 0
Result 'Client' 'NLA network profile = DomainAuthenticated' $domAuth (($prof | ForEach-Object { "$($_.InterfaceAlias)=$($_.NetworkCategory)" }) -join '; ') `
    'NLA could not reach a DC via LDAP - points to DC-locator/network/firewall (UDP/TCP 389) issue.'
Run 'WinHTTP proxy' { netsh winhttp show proxy }

# ---------------------------------------------------------------- 2. TIME
Section '2. TIME (Kerberos needs < 5 min skew)'
Run 'w32tm status' { w32tm /query /status }
Run 'w32tm source' { w32tm /query /source }
foreach ($d in $DCs) { Run ("Offset vs {0}" -f $d) { w32tm /stripchart /computer:$d /samples:2 /dataonly } }

# ---------------------------------------------------------------- 3. SECURE CHANNEL
Section '3. SECURE CHANNEL (read-only - no -Repair)'
$sc = $false
try { $sc = Test-ComputerSecureChannel -Verbose -ErrorAction Stop } catch { Write-Host ("ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red }
Write-Host ("Test-ComputerSecureChannel = {0}" -f $sc)
Result 'Trust' 'Test-ComputerSecureChannel' $sc "$sc" 'Broken machine trust OR no DC reachable. If DC locator/ports PASS but this FAILS -> trust issue (repair via remediation script).'
$scq = NlTest @("/sc_query:$Domain")
Write-Host $scq.Out
Result 'Trust' 'nltest /sc_query' $scq.Ok (($scq.Out -split "`r?`n" | Select-String 'Trusted DC Name|Connection Status') -join ' | ') 'Secure channel not established - see status code (e.g. 1311 no logon servers, 5 access denied, 1355 domain not found).'
$scv = NlTest @("/sc_verify:$Domain")
Write-Host $scv.Out

# ---------------------------------------------------------------- 4. DC LOCATOR
Section '4. DC LOCATOR'
$site = NlTest @('/dsgetsite')
Write-Host $site.Out
$siteName = if ($site.Ok) { ($site.Out -split "`r?`n")[0].Trim() } else { '' }
Result 'DC Locator' 'Client AD site resolved' $site.Ok $siteName 'Client subnet not mapped to an AD site (NO_CLIENT_SITE) - add subnet in AD Sites & Services.'

$locTests = [ordered]@{
    'default'  = @("/dsgetdc:$Domain")
    'writable' = @("/dsgetdc:$Domain", '/writable', '/force')
    'gc'       = @("/dsgetdc:$Domain", '/gc', '/force')
    'pdc'      = @("/dsgetdc:$Domain", '/pdc', '/force')
    'kdc'      = @("/dsgetdc:$Domain", '/kdc', '/force')
    'netbios'  = @("/dsgetdc:$NetBIOS", '/force')
}
foreach ($k in $locTests.Keys) {
    Write-Host ("--- dsgetdc {0}" -f $k) -ForegroundColor Yellow
    $r = NlTest $locTests[$k]
    Write-Host $r.Out
    $dcLine = ($r.Out -split "`r?`n" | Select-String 'DC: ') -join ''
    Result 'DC Locator' "nltest /dsgetdc $k" $r.Ok $dcLine.Trim() 'DC locator failed for this role - check SRV records, UDP/TCP 389 and 88 to the returned DCs.'
    if ($k -eq 'default' -and $r.Ok) {
        $isWritable = $r.Out -match 'WRITABLE'
        Result 'RODC' 'Default DC is writable (not RODC)' $isWritable $dcLine.Trim() `
            'Client is using an RODC. Migration tools usually need a WRITABLE DC/GC - see /writable result and ports to writable DCs.'
    }
}
Run 'nltest /dclist' { nltest /dclist:$Domain }

# ---------------------------------------------------------------- 5. DNS SRV
Section '5. DNS SRV RECORDS'
$srv = @(
    "_ldap._tcp.dc._msdcs.$Domain",
    "_kerberos._tcp.dc._msdcs.$Domain",
    "_gc._tcp.$Domain",
    "_ldap._tcp.pdc._msdcs.$Domain",
    "_kerberos._udp.$Domain",
    "_kpasswd._tcp.$Domain"
)
if ($siteName) { $srv += "_ldap._tcp.$siteName._sites.dc._msdcs.$Domain"; $srv += "_ldap._tcp.$siteName._sites.gc._msdcs.$Domain" }

$dcNames = @()
foreach ($rec in $srv) {
    $ok = $false; $targets = ''
    try {
        $ans = @(Resolve-DnsName $rec -Type SRV -ErrorAction Stop | Where-Object { $_.Type -eq 'SRV' })
        $ans | Format-Table Name,NameTarget,Port,Priority,Weight -AutoSize | Out-String -Width 200 | Write-Host
        $ok = $ans.Count -gt 0
        $targets = ($ans.NameTarget | Sort-Object -Unique) -join ', '
        $dcNames += $ans.NameTarget
    } catch { Write-Host ("{0}: {1}" -f $rec, $_.Exception.Message) -ForegroundColor Red }
    Result 'DNS' "SRV $rec" $ok $targets 'SRV missing on client DNS - DC netlogon DNS registration or DNS server (RODC/forwarder) issue.'
}
foreach ($rec in $srv[0..2]) {
    Run ("SRV via 10.7.7.13 (compare) {0}" -f $rec) { Resolve-DnsName $rec -Type SRV -Server 10.7.7.13 -ErrorAction Stop | Where-Object Type -eq 'SRV' | Format-Table Name,NameTarget,Port -AutoSize }
}

# Add DCs discovered via SRV to the port test list
$dcNames = @($dcNames | Sort-Object -Unique)
$DcMap = [ordered]@{}
foreach ($ip in $DCs) { $DcMap[$ip] = '(static list)' }
foreach ($n in $dcNames) {
    try {
        foreach ($a in @(Resolve-DnsName $n -Type A -ErrorAction Stop | Where-Object Type -eq 'A')) {
            if ($DcMap.Contains($a.IPAddress)) { $DcMap[$a.IPAddress] = $n } else { $DcMap[$a.IPAddress] = $n }
        }
    } catch { Write-Host ("A record for {0} failed: {1}" -f $n, $_.Exception.Message) -ForegroundColor Red }
}
Run 'DCs to be tested' { $DcMap.GetEnumerator() | Format-Table @{n='IP';e={$_.Key}}, @{n='Name';e={$_.Value}} -AutoSize }

# ---------------------------------------------------------------- 6. PORTS
Section '6. TCP PORTS (2s timeout)'
$portRes = foreach ($d in $DcMap.Keys) {
    foreach ($p in $Ports) {
        $ok = $false
        $c = New-Object System.Net.Sockets.TcpClient
        try {
            $iar = $c.BeginConnect($d, $p, $null, $null)
            if ($iar.AsyncWaitHandle.WaitOne(2000, $false)) { $c.EndConnect($iar); $ok = $true }
        } catch { $ok = $false } finally { $c.Close() }
        [pscustomobject]@{ DC = $d; Name = $DcMap[$d]; Port = $p; Open = $ok }
    }
}
$portRes | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
foreach ($g in ($portRes | Group-Object DC)) {
    $closed = @($g.Group | Where-Object { -not $_.Open } | ForEach-Object Port)
    Result 'Network' ("Ports to {0} ({1})" -f $g.Name, $DcMap[$g.Name]) ($closed.Count -eq 0) `
        $(if ($closed.Count) { 'Closed: ' + ($closed -join ',') } else { 'All open' }) 'Firewall/routing blocks these ports to this DC. 9389 closed = Get-AD* (ADWS) fails even if LDAP works.'
}
Write-Host 'NOTE: UDP 53/88/123/389 and dynamic RPC (49152-65535) cannot be verified with a TCP socket.' -ForegroundColor DarkYellow
Write-Host '      Use PortQry from a jump box:  portqry -n <DC> -e 135   /   portqry -n <DC> -p udp -e 389' -ForegroundColor DarkYellow

# ---------------------------------------------------------------- 7. FUNCTIONAL
Section '7. FUNCTIONAL LDAP / GC / SYSVOL / KERBEROS'
$ldapOk = $false
try { $h = ([ADSI]"LDAP://$Domain/RootDSE").dnsHostName; Write-Host "LDAP RootDSE (domain) -> $h"; $ldapOk = [bool]$h } catch { Write-Host ("ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red }
Result 'AD' "LDAP bind to $Domain" $ldapOk "$h" 'LDAP to domain failed - network/firewall or DC issue.'
foreach ($d in $DcMap.Keys) { Run ("LDAP RootDSE ({0})" -f $d) { $r = [ADSI]"LDAP://$d/RootDSE"; "{0}  writable={1}" -f $r.dnsHostName, ($r.supportedCapabilities -notcontains '1.2.840.113556.1.4.1920') } }

$gdOk = $false
try {
    $ctx = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $Domain)
    $dom = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($ctx)
    $dom | Select-Object Name,PdcRoleOwner,DomainMode | Format-List | Out-String | Write-Host
    $gdOk = $true
} catch { Write-Host ("GetDomain ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red }
Result 'AD' 'Domain::GetDomain()' $gdOk '' 'Typical API used by migration pre-flight - failure here reproduces "Source domain could not be reached".'

$gcdOk = $false
try { $cd = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain(); Write-Host "GetComputerDomain -> $($cd.Name)"; $gcdOk = $true }
catch { Write-Host ("GetComputerDomain ERROR: {0}" -f $_.Exception.Message) -ForegroundColor Red }
Result 'AD' 'Domain::GetComputerDomain()' $gcdOk '' 'Machine-context domain lookup failed (common when run as SYSTEM with broken trust or unreachable DC).'

if (Get-Module -ListAvailable ActiveDirectory) {
    Run 'Get-ADDomain (ADWS 9389)' { Get-ADDomain -Server $Domain -ErrorAction Stop | Select-Object DNSRoot,PDCEmulator,RIDMaster | Format-List }
} else { Write-Host 'ActiveDirectory module not installed - Get-ADDomain skipped.' }

Run 'GC bind' { ([ADSI]"GC://$Domain").Path }
$sysvol = Test-Path "\\$Domain\SYSVOL" -ErrorAction SilentlyContinue
Result 'AD' "SYSVOL \\$Domain\SYSVOL" $sysvol '' 'SMB 445 / DFS referral / Kerberos to DC failing.'
Run 'Kerberos tickets (current user)' { klist }
Run 'Kerberos tickets (SYSTEM / machine)' { klist -li 0x3e7 }
Run 'Kerberos TGT test (read-only request)' { klist get krbtgt/$Domain }

# ---------------------------------------------------------------- 8. LOGS
Section '8. EVENT LOGS (last 3 days)'
$start = (Get-Date).AddDays(-3)
Run 'NETLOGON events'       { Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='NETLOGON';StartTime=$start} -MaxEvents 20 -ErrorAction Stop | Format-List TimeCreated,Id,Message }
Run 'Kerberos/Time/GPO/Trust' { Get-WinEvent -FilterHashtable @{LogName='System';Id=4,1053,1054,1129,3210,5719,5722,5723,5805;StartTime=$start} -MaxEvents 30 -ErrorAction Stop | Format-List TimeCreated,ProviderName,Id,Message }
Run 'Time-Service events'   { Get-WinEvent -FilterHashtable @{LogName='System';ProviderName='Microsoft-Windows-Time-Service';StartTime=$start} -MaxEvents 10 -ErrorAction Stop | Format-List TimeCreated,Id,Message }
Run 'netlogon.log'          { Select-String -Path "$env:windir\debug\netlogon.log" -Pattern 'NO_CLIENT_SITE','failed','0xC000' -ErrorAction Stop | Select-Object -Last 20 }

# ---------------------------------------------------------------- 9. MIGRATION AGENT
Section '9. MIGRATION AGENT SERVICE'
Run 'Agent services' { Get-CimInstance Win32_Service | Where-Object { $_.Name -match 'agent|migrat|quest|binary|powersyncpro' -or $_.DisplayName -match 'agent|migrat|quest|binary|powersyncpro' } | Format-Table Name,DisplayName,StartName,State -AutoSize }

# ---------------------------------------------------------------- SUMMARY
Section 'SUMMARY'
$Summary | Format-Table Area,Check,Result,Detail -AutoSize -Wrap | Out-String -Width 250 | Write-Host
$fails = @($Summary | Where-Object Result -eq 'FAIL')
if ($fails.Count) {
    Write-Host 'FAILED CHECKS - interpretation:' -ForegroundColor Red
    foreach ($f in $fails) { Write-Host (" [{0}] {1}`n     -> {2}" -f $f.Area, $f.Check, $f.IfFail) -ForegroundColor Red }
} else { Write-Host 'All summarized checks passed. If pre-flight still fails, the issue is in the migration tool config/context (account, domain name, credentials).' -ForegroundColor Green }

Stop-Transcript | Out-Null
Write-Host ""
Write-Host ("Log saved: {0}" -f $Log) -ForegroundColor Green
