#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    CHANGE script - remediation actions for "Source domain could not be reached".
    Run ONLY after Invoke-ADReachabilityCheck-ReadOnly.ps1 has confirmed the root cause.

.DESCRIPTION
    Nothing runs unless you pass -Action. Every action asks for confirmation
    (ConfirmImpact=High). Use -WhatIf to preview. Actions run in the order given.

    Action               Fixes                                   Impact
    -------------------  --------------------------------------  -------------------------------------------
    FlushDns             Stale/negative DNS cache                 Low  - clears resolver cache
    RegisterDns          Client A/PTR record missing              Low  - re-registers client in DNS
    PurgeKerberos        Stale tickets (after time/trust fix)     Low  - user + SYSTEM tickets re-requested
    ResyncTime           Clock skew > 5 min (KRB_AP_ERR_SKEW)     Low  - w32tm resync
    RestartNetlogon      DC locator stuck / cached bad DC         Low  - brief Netlogon restart
    ResetSecureChannel   Secure channel pinned to bad DC/RODC     Low  - re-binds channel to -DCName (no pwd reset)
    GpUpdate             Refresh policy after connectivity fix    Low
    SetDnsServers        Wrong DNS servers on NIC                 MED  - backs up current DNS, then sets -DnsServers
    RepairSecureChannel  BROKEN TRUST confirmed (sc_query fails   HIGH - resets machine account password on DC.
                         while DC locator + ports PASS)                  Needs domain creds. Last resort.

.EXAMPLE
    .\Invoke-ADReachabilityRemediation.ps1 -Action FlushDns,PurgeKerberos -WhatIf
.EXAMPLE
    .\Invoke-ADReachabilityRemediation.ps1 -Action ResetSecureChannel -DCName DC01.amccinemasksa.net
.EXAMPLE
    .\Invoke-ADReachabilityRemediation.ps1 -Action SetDnsServers -InterfaceAlias Ethernet -DnsServers 10.7.7.13,10.7.7.14
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('FlushDns','RegisterDns','PurgeKerberos','ResyncTime','RestartNetlogon',
                 'ResetSecureChannel','GpUpdate','SetDnsServers','RepairSecureChannel')]
    [string[]]$Action,

    [string]  $Domain = 'amccinemasksa.net',
    [string]  $DCName,                 # WRITABLE DC FQDN - for ResetSecureChannel / RepairSecureChannel
    [string]  $InterfaceAlias,         # for SetDnsServers
    [string[]]$DnsServers,             # for SetDnsServers
    [string]  $OutDir = $PSScriptRoot
)

if (-not $OutDir) { $OutDir = $env:TEMP }
$Log = Join-Path $OutDir ("ADRemediation_{0}_{1}.txt" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'))
Start-Transcript -Path $Log | Out-Null
function Step($t) { Write-Host ""; Write-Host ("===== {0} =====" -f $t) -ForegroundColor Cyan }

foreach ($a in $Action) {
    Step $a
    switch ($a) {
        'FlushDns' {
            if ($PSCmdlet.ShouldProcess('DNS resolver cache', 'ipconfig /flushdns')) { ipconfig /flushdns }
        }
        'RegisterDns' {
            if ($PSCmdlet.ShouldProcess('Client DNS registration', 'ipconfig /registerdns')) { ipconfig /registerdns }
        }
        'PurgeKerberos' {
            if ($PSCmdlet.ShouldProcess('Kerberos tickets (current user + SYSTEM 0x3e7)', 'klist purge')) {
                klist purge
                klist -li 0x3e7 purge
            }
        }
        'ResyncTime' {
            if ($PSCmdlet.ShouldProcess('Windows Time', 'w32tm /resync /force')) {
                w32tm /query /source
                w32tm /resync /force
                w32tm /query /status
            }
        }
        'RestartNetlogon' {
            if ($PSCmdlet.ShouldProcess('Netlogon service', 'Restart-Service')) {
                Restart-Service Netlogon -Force
                nltest /dsgetdc:$Domain /force
            }
        }
        'ResetSecureChannel' {
            if (-not $DCName) { Write-Host 'Skipped: -DCName (writable DC FQDN) required.' -ForegroundColor Red; break }
            if ($PSCmdlet.ShouldProcess("Secure channel $Domain", "nltest /sc_reset:$Domain\$DCName")) {
                nltest /sc_reset:$Domain\$DCName
                nltest /sc_query:$Domain
            }
        }
        'GpUpdate' {
            if ($PSCmdlet.ShouldProcess('Group Policy', 'gpupdate /force')) { gpupdate /force }
        }
        'SetDnsServers' {
            if (-not $InterfaceAlias -or -not $DnsServers) { Write-Host 'Skipped: -InterfaceAlias and -DnsServers required.' -ForegroundColor Red; break }
            $cur = (Get-DnsClientServerAddress -InterfaceAlias $InterfaceAlias -AddressFamily IPv4).ServerAddresses
            $bak = Join-Path $OutDir ("DNSBackup_{0}_{1}.txt" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'))
            "$InterfaceAlias : $($cur -join ',')" | Set-Content $bak
            Write-Host ("Current DNS: {0}  (backup: {1})" -f ($cur -join ','), $bak)
            Write-Host ("Rollback: Set-DnsClientServerAddress -InterfaceAlias '{0}' -ServerAddresses {1}" -f $InterfaceAlias, ($cur -join ',')) -ForegroundColor DarkYellow
            if ($PSCmdlet.ShouldProcess($InterfaceAlias, "Set DNS servers to $($DnsServers -join ',')")) {
                Set-DnsClientServerAddress -InterfaceAlias $InterfaceAlias -ServerAddresses $DnsServers
                ipconfig /flushdns
                Resolve-DnsName "_ldap._tcp.dc._msdcs.$Domain" -Type SRV | Format-Table Name,NameTarget,Port -AutoSize
            }
        }
        'RepairSecureChannel' {
            if (-not $DCName) { Write-Host 'Skipped: -DCName (writable DC FQDN) required.' -ForegroundColor Red; break }
            Write-Host 'WARNING: resets the computer account password against the DC. Use only when trust is confirmed broken.' -ForegroundColor Red
            if ($PSCmdlet.ShouldProcess("Computer account $env:COMPUTERNAME on $DCName", 'Test-ComputerSecureChannel -Repair')) {
                $cred = Get-Credential -Message "Domain account with rights to reset $env:COMPUTERNAME"
                Test-ComputerSecureChannel -Repair -Server $DCName -Credential $cred -Verbose
                Test-ComputerSecureChannel -Verbose
            }
        }
    }
}

Stop-Transcript | Out-Null
Write-Host ""
Write-Host ("Log saved: {0}. Re-run Invoke-ADReachabilityCheck-ReadOnly.ps1 to validate." -f $Log) -ForegroundColor Green
