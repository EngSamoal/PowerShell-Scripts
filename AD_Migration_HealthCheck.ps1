<#
.SYNOPSIS
    Read-only Active Directory health check for the Azure -> GCP domain controller migration.

.DESCRIPTION
    Runs PASS / FAIL / WARN / UNABLE checks against every DC in the current domain (or the DCs you
    pass in) and writes a CSV report. Makes NO configuration changes to AD, DNS, DFSR or the DCs.
    The only write is the CSV report in -OutputFolder.

    Checks per DC: TCP reachability of AD ports, LDAPS certificate handshake, DC role/GC flag,
    inbound replication (Get-ADReplicationPartnerMetadata), dcdiag, SRV/A record resolution,
    SYSVOL/NETLOGON share access, DFSR SYSVOL state, time offset, Kerberos ticket retrieval,
    and recent error events in Directory Service / DNS Server / DFS Replication / System logs.
    Local checks: secure channel and Group Policy errors on the machine running the script.

.REQUIREMENTS
    - Run in an elevated PowerShell session as a Domain Admin (read-only use) on a DC or an
      admin server with RSAT: AD DS tools, DNS Server tools (dcdiag/repadmin come with RSAT AD DS).
    - Remote Event Log (RPC) and WinRM/CIM to each DC for the event-log and DFSR checks; if
      blocked, those checks report UNABLE with the reason instead of failing silently.
    - Output parsing of klist/w32tm/dcdiag assumes an English OS display language.

.EXAMPLE
    .\AD_Migration_HealthCheck.ps1
    .\AD_Migration_HealthCheck.ps1 -DomainControllers gcpdc01.corp.example.com,azdc01.corp.example.com
#>
[CmdletBinding()]
param(
    [string[]]$DomainControllers,
    [string]$OutputFolder = (Join-Path $env:USERPROFILE 'Documents\ADHealth'),
    [int]$MaxReplicationAgeMinutes = 180,   # set to ~2x your inter-site replication interval
    [double]$MaxTimeOffsetSeconds = 2,
    [int]$EventLookbackHours = 24
)

$ScriptBuild = 'AD-Migration-HealthCheck 2026-11-04.1'
Write-Host $ScriptBuild -ForegroundColor Cyan
$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()

function Add-Result {
    param([string]$Dc, [string]$Check, [ValidateSet('PASS','FAIL','WARN','UNABLE')][string]$Status, [string]$Detail)
    $results.Add([pscustomobject]@{
        Time = (Get-Date -Format 's'); DC = $Dc; Check = $Check; Status = $Status; Detail = $Detail })
}

function Test-TcpPort {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 3000)
    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $ar = $client.BeginConnect($HostName, $Port, $null, $null)
        if ($ar.AsyncWaitHandle.WaitOne($TimeoutMs) -and $client.Connected) { $client.EndConnect($ar); return $true }
        return $false
    } catch { return $false } finally { $client.Close() }
}

function Test-Ldaps {
    param([string]$HostName)
    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $ar = $client.BeginConnect($HostName, 636, $null, $null)
        if (-not ($ar.AsyncWaitHandle.WaitOne(3000) -and $client.Connected)) { return 'TCP 636 not reachable' }
        $client.EndConnect($ar)
        $script:ldapsErrors = 'None'
        $cb = [System.Net.Security.RemoteCertificateValidationCallback]{
            param($s, $cert, $chain, $errors) $script:ldapsErrors = $errors.ToString(); return $true }
        $ssl = [System.Net.Security.SslStream]::new($client.GetStream(), $false, $cb)
        $ssl.AuthenticateAsClient($HostName)
        $c = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($ssl.RemoteCertificate)
        $ssl.Dispose()
        return "OK|Subject=$($c.Subject); Expires=$($c.NotAfter.ToString('yyyy-MM-dd')); ValidationErrors=$script:ldapsErrors"
    } catch { return "TLS handshake failed: $($_.Exception.Message)" } finally { $client.Close() }
}

# ---------------- Domain / forest context ----------------
try {
    Import-Module ActiveDirectory
    $domain = Get-ADDomain
    $forest = Get-ADForest
} catch {
    Write-Error "Cannot load the ActiveDirectory module or contact the domain: $($_.Exception.Message)"
    return
}
$dnsDomain = $domain.DNSRoot
if (-not $DomainControllers) {
    $DomainControllers = @((Get-ADDomainController -Filter * -Server $domain.PDCEmulator).HostName)
}
$allDcObjects = @(Get-ADDomainController -Filter * -Server $domain.PDCEmulator)

Add-Result 'FOREST' 'Functional levels' 'PASS' "Forest=$($forest.ForestMode); Domain=$($domain.DomainMode)"
Add-Result 'FOREST' 'FSMO holders' 'PASS' ("Schema=$($forest.SchemaMaster); DomainNaming=$($forest.DomainNamingMaster); " +
    "PDC=$($domain.PDCEmulator); RID=$($domain.RIDMaster); Infrastructure=$($domain.InfrastructureMaster)")

foreach ($role in @($forest.SchemaMaster, $forest.DomainNamingMaster, $domain.PDCEmulator, $domain.RIDMaster, $domain.InfrastructureMaster)) {
    if ($allDcObjects.HostName -notcontains $role) {
        Add-Result 'FOREST' 'FSMO holder exists' 'FAIL' "Role holder '$role' is not a current DC object"
    }
}

$writableDcs = @($allDcObjects | Where-Object { -not $_.IsReadOnly })
$gcs = @($allDcObjects | Where-Object { $_.IsGlobalCatalog })
Add-Result 'FOREST' 'Writable DC count' ($(if ($writableDcs.Count -ge 2) {'PASS'} else {'FAIL'})) "$($writableDcs.Count) writable DC(s)"
Add-Result 'FOREST' 'GC count' ($(if ($gcs.Count -ge 2) {'PASS'} else {'FAIL'})) "$($gcs.Count) Global Catalog(s)"

foreach ($site in @($allDcObjects | Group-Object Site)) {
    Add-Result 'FOREST' "DCs in site '$($site.Name)'" 'PASS' (($site.Group.HostName) -join ', ')
}

try {
    $srv = @(Resolve-DnsName "_ldap._tcp.dc._msdcs.$dnsDomain" -Type SRV -ErrorAction Stop | Where-Object Type -eq 'SRV')
    $missing = @($allDcObjects.HostName | Where-Object { $srv.NameTarget -notcontains $_ })
    if ($missing.Count -eq 0) { Add-Result 'FOREST' 'DC SRV records (local resolver)' 'PASS' "$($srv.Count) SRV record(s); all DCs registered" }
    else { Add-Result 'FOREST' 'DC SRV records (local resolver)' 'FAIL' "Missing SRV for: $($missing -join ', ')" }
} catch { Add-Result 'FOREST' 'DC SRV records (local resolver)' 'FAIL' $_.Exception.Message }

# ---------------- Per-DC checks ----------------
$ports = [ordered]@{ 53='DNS'; 88='Kerberos'; 135='RPC-EPM'; 389='LDAP'; 445='SMB'; 464='Kpasswd'; 636='LDAPS'; 3268='GC'; 3269='GC-SSL'; 9389='ADWS' }
$since = (Get-Date).AddHours(-$EventLookbackHours)

foreach ($dc in $DomainControllers) {
    Write-Host "Checking $dc ..." -ForegroundColor Yellow
    $dcObj = $allDcObjects | Where-Object HostName -eq $dc

    # Ports
    $closed = @(foreach ($kv in $ports.GetEnumerator()) { if (-not (Test-TcpPort $dc $kv.Key)) { "$($kv.Key)/$($kv.Value)" } })
    if ($closed.Count -eq 0) { Add-Result $dc 'AD TCP ports' 'PASS' (($ports.Keys) -join ',') }
    else { Add-Result $dc 'AD TCP ports' 'FAIL' "Unreachable: $($closed -join ', ') (UDP 53/88/123/389/464 and RPC 49152-65535 not tested here)" }

    # LDAPS
    $l = Test-Ldaps $dc
    if ($l -like 'OK|*') {
        $st = if ($l -match 'ValidationErrors=None') { 'PASS' } else { 'WARN' }
        Add-Result $dc 'LDAPS handshake' $st $l.Substring(3)
    } else { Add-Result $dc 'LDAPS handshake' 'WARN' "$l (FAIL only if LDAPS is required in your environment)" }

    # Role
    if ($dcObj) {
        Add-Result $dc 'DC object' 'PASS' "Site=$($dcObj.Site); GC=$($dcObj.IsGlobalCatalog); RODC=$($dcObj.IsReadOnly); OS=$($dcObj.OperatingSystem); IP=$($dcObj.IPv4Address)"
    } else { Add-Result $dc 'DC object' 'FAIL' 'Not found as a DC in this domain' }

    # Replication (inbound, all partitions)
    try {
        $meta = @(Get-ADReplicationPartnerMetadata -Target $dc -Partition * -ErrorAction Stop)
        $bad = @($meta | Where-Object {
            $_.LastReplicationResult -ne 0 -or $_.ConsecutiveReplicationFailures -gt 0 -or
            -not $_.LastReplicationSuccess -or ((Get-Date) - $_.LastReplicationSuccess).TotalMinutes -gt $MaxReplicationAgeMinutes })
        if ($meta.Count -eq 0) { Add-Result $dc 'Inbound replication' 'FAIL' 'No replication partners returned' }
        elseif ($bad.Count -eq 0) {
            $oldest = ($meta | Measure-Object -Property LastReplicationSuccess -Minimum).Minimum
            Add-Result $dc 'Inbound replication' 'PASS' "$($meta.Count) partner/partition links OK; oldest success $oldest"
        } else {
            $d = ($bad | ForEach-Object { "$(($_.Partner -split ',')[1]) $($_.Partition) result=$($_.LastReplicationResult) fails=$($_.ConsecutiveReplicationFailures) last=$($_.LastReplicationSuccess)" }) -join ' | '
            Add-Result $dc 'Inbound replication' 'FAIL' $d
        }
    } catch { Add-Result $dc 'Inbound replication' 'UNABLE' $_.Exception.Message }

    # dcdiag (quiet mode prints only problems)
    try {
        $dd = (& dcdiag.exe /s:$dc /q 2>&1 | Out-String)
        $failed = @([regex]::Matches($dd, 'failed test (\S+)') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
        if ($failed.Count -eq 0) { Add-Result $dc 'dcdiag' 'PASS' 'No failed tests' }
        else { Add-Result $dc 'dcdiag' 'FAIL' "Failed: $($failed -join ', ') - run 'dcdiag /s:$dc /v' for detail" }
    } catch { Add-Result $dc 'dcdiag' 'UNABLE' $_.Exception.Message }

    # DNS served by this DC
    try {
        $srvDc = @(Resolve-DnsName "_ldap._tcp.dc._msdcs.$dnsDomain" -Type SRV -Server $dc -DnsOnly -ErrorAction Stop | Where-Object Type -eq 'SRV')
        $missing = @($allDcObjects.HostName | Where-Object { $srvDc.NameTarget -notcontains $_ })
        $aOk = @(Resolve-DnsName $dc -Type A -Server $dc -DnsOnly -ErrorAction Stop).Count -gt 0
        if ($missing.Count -eq 0 -and $aOk) { Add-Result $dc 'DNS (queried on this DC)' 'PASS' "All $($allDcObjects.Count) DC SRV records present; own A record resolves" }
        else { Add-Result $dc 'DNS (queried on this DC)' 'FAIL' "Missing SRV: $($missing -join ', '); A record OK=$aOk" }
    } catch { Add-Result $dc 'DNS (queried on this DC)' 'FAIL' "Query to $dc failed: $($_.Exception.Message)" }

    # SYSVOL / NETLOGON
    $sysOk = Test-Path "\\$dc\SYSVOL\$dnsDomain\Policies"
    $netOk = Test-Path "\\$dc\NETLOGON"
    Add-Result $dc 'SYSVOL/NETLOGON shares' ($(if ($sysOk -and $netOk) {'PASS'} else {'FAIL'})) "SYSVOL=$sysOk; NETLOGON=$netOk"

    # DFSR SYSVOL state (4 = Normal)
    try {
        $rf = Get-CimInstance -ComputerName $dc -Namespace 'root/microsoftdfs' -ClassName DfsrReplicatedFolderInfo `
              -Filter "ReplicatedFolderName='SYSVOL Share'" -OperationTimeoutSec 30 -ErrorAction Stop
        $map = @{0='Uninitialized';1='Initialized';2='Initial Sync';3='Auto Recovery';4='Normal';5='In Error'}
        Add-Result $dc 'DFSR SYSVOL state' ($(if ($rf.State -eq 4) {'PASS'} else {'FAIL'})) "State=$($rf.State) ($($map[[int]$rf.State]))"
    } catch { Add-Result $dc 'DFSR SYSVOL state' 'UNABLE' "CIM/WinRM query failed (or SYSVOL still uses FRS): $($_.Exception.Message)" }

    # Time offset relative to the machine running this script
    try {
        $tc = (& w32tm.exe /stripchart /computer:$dc /samples:1 /dataonly 2>&1 | Out-String)
        if ($tc -match '([+-]\d+\.\d+)s') {
            $off = [math]::Abs([double]$Matches[1])
            Add-Result $dc 'Time offset vs this host' ($(if ($off -le $MaxTimeOffsetSeconds) {'PASS'} else {'FAIL'})) "Offset=$($Matches[1])s (Kerberos hard limit 300s; target <= $MaxTimeOffsetSeconds s)"
        } else { Add-Result $dc 'Time offset vs this host' 'UNABLE' ($tc.Trim() -replace '\s+',' ') }
    } catch { Add-Result $dc 'Time offset vs this host' 'UNABLE' $_.Exception.Message }

    # Kerberos service ticket for this DC
    try {
        $kl = (& klist.exe get "ldap/$dc" 2>&1 | Out-String)
        if ($kl -match 'retrieved successfully') { Add-Result $dc 'Kerberos ticket (ldap/<dc>)' 'PASS' 'Service ticket obtained' }
        else { Add-Result $dc 'Kerberos ticket (ldap/<dc>)' 'FAIL' ($kl.Trim() -replace '\s+',' ') }
    } catch { Add-Result $dc 'Kerberos ticket (ldap/<dc>)' 'UNABLE' $_.Exception.Message }

    # Event logs (errors/critical in lookback window)
    foreach ($log in 'Directory Service','DNS Server','DFS Replication','System') {
        try {
            $ev = @(Get-WinEvent -ComputerName $dc -FilterHashtable @{ LogName = $log; Level = 1,2; StartTime = $since } -ErrorAction Stop)
            $top = ($ev | Group-Object Id | Sort-Object Count -Descending | Select-Object -First 5 | ForEach-Object { "ID $($_.Name) x$($_.Count)" }) -join ', '
            Add-Result $dc "Errors: $log (${EventLookbackHours}h)" 'WARN' "$($ev.Count) error(s): $top - review each ID"
        } catch {
            if ($_.Exception.Message -match 'No events were found') { Add-Result $dc "Errors: $log (${EventLookbackHours}h)" 'PASS' 'None' }
            else { Add-Result $dc "Errors: $log (${EventLookbackHours}h)" 'UNABLE' $_.Exception.Message }
        }
    }
}

# ---------------- Local machine checks ----------------
try {
    $sc = Test-ComputerSecureChannel -ErrorAction Stop   # read-only (no -Repair)
    Add-Result $env:COMPUTERNAME 'Secure channel (local)' ($(if ($sc) {'PASS'} else {'FAIL'})) "Test-ComputerSecureChannel=$sc"
} catch { Add-Result $env:COMPUTERNAME 'Secure channel (local)' 'UNABLE' $_.Exception.Message }

try {
    $gp = @(Get-WinEvent -FilterHashtable @{ LogName = 'Microsoft-Windows-GroupPolicy/Operational'; Level = 1,2; StartTime = $since } -ErrorAction Stop)
    Add-Result $env:COMPUTERNAME 'Group Policy errors (local)' 'FAIL' "$($gp.Count) error(s); IDs: $((($gp.Id | Select-Object -Unique) -join ','))"
} catch {
    if ($_.Exception.Message -match 'No events were found') { Add-Result $env:COMPUTERNAME 'Group Policy errors (local)' 'PASS' 'None' }
    else { Add-Result $env:COMPUTERNAME 'Group Policy errors (local)' 'UNABLE' $_.Exception.Message }
}

# ---------------- Output ----------------
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$csv = Join-Path $OutputFolder ("ADHealth_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))
$results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

Write-Host ($results | Format-Table DC, Check, Status, Detail -AutoSize -Wrap | Out-String -Width 250)
$fail = @($results | Where-Object Status -eq 'FAIL').Count
$warn = @($results | Where-Object Status -eq 'WARN').Count
$unab = @($results | Where-Object Status -eq 'UNABLE').Count
$overall = if ($fail -gt 0) { 'FAIL' } elseif ($unab -gt 0 -or $warn -gt 0) { 'REVIEW' } else { 'PASS' }
Write-Host ("OVERALL: {0}  (FAIL={1} WARN={2} UNABLE={3})  Report: {4}" -f $overall, $fail, $warn, $unab, $csv) -ForegroundColor ($(if ($overall -eq 'PASS') {'Green'} elseif ($overall -eq 'FAIL') {'Red'} else {'Yellow'}))
