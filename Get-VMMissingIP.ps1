# Get-VMMissingIP.ps1
# Finds IP addresses for VMs / templates where VMware Tools could not report one.
# Read-only: vCenter is only queried, nothing is changed.
# Input : C:\temp\vmlist.txt  (one VM or template name per line)
# Output: C:\temp\vm_ip_results.csv  + table on screen
# Run   : .\Get-VMMissingIP.ps1
param(
    [string]$ListFile = 'C:\temp\vmlist.txt',            # <-- VM / template names, one per line
    [string]$OutFile  = 'C:\temp\vm_ip_results.csv'      # <-- results
)

#region ======================= CONFIGURATION (edit here) =======================
$VCenters    = @('10.7.53.6')    # used only if you are not already connected (Connect-VIServer)
$DhcpServers = @()               # optional, e.g. @('10.7.1.10') - Windows DHCP servers to search by MAC
$PingCheck   = $true             # ping each IP found to show if it answers
#endregion ======================================================================

$ErrorActionPreference = 'Stop'

function Get-IPv4([object[]]$Values) {
    # Routable IPv4 addresses only (no 169.254 / 127 / 0.0.0.0)
    foreach ($v in @($Values)) {
        $a = $null
        if ($v -and [Net.IPAddress]::TryParse([string]$v, [ref]$a) -and $a.AddressFamily -eq 'InterNetwork' -and
            [string]$v -notlike '169.254.*' -and [string]$v -notlike '127.*' -and [string]$v -ne '0.0.0.0') { [string]$v }
    }
}

function Resolve-DnsIPv4([string]$Name) {
    if (-not $Name) { return @() }
    try { return @(Get-IPv4 @([Net.Dns]::GetHostAddresses($Name) | ForEach-Object { $_.ToString() })) } catch { return @() }
}

function Test-Ping([string]$IP) {
    try { return ((New-Object Net.NetworkInformation.Ping).Send($IP, 1000).Status -eq 'Success') } catch { return $false }
}

function Resolve-VMIp($View, $ArpTable, $DhcpTable) {
    # Returns @(IP, Source) using the first method that works
    $guest = @(Get-IPv4 (@($View.Guest.IpAddress) + @($View.Guest.Net | ForEach-Object { $_.IpAddress })))
    if ($guest.Count) { return @((($guest | Select-Object -Unique) -join ', '), 'vCenter guest info (last known)') }

    foreach ($n in @($View.Name, $View.Guest.HostName) | Where-Object { $_ } | Select-Object -Unique) {
        $dns = @(Resolve-DnsIPv4 $n)
        if ($dns.Count) { return @(($dns -join ', '), "DNS ($n)") }
    }

    $notes = @(Get-IPv4 @([regex]::Matches([string]$View.Config.Annotation, '\b\d{1,3}(\.\d{1,3}){3}\b') | ForEach-Object { $_.Value }))
    if ($notes.Count) { return @(($notes -join ', '), 'vCenter Notes') }

    foreach ($mac in (Get-VMMacs $View)) {
        $key = $mac.ToUpper().Replace(':', '-')
        if ($ArpTable.ContainsKey($key)) { return @($ArpTable[$key], 'ARP table (MAC match)') }
        if ($DhcpTable.ContainsKey($key)) { return @($DhcpTable[$key], 'DHCP lease (MAC match)') }
    }
    return @('', '')
}

function Get-VMMacs($View) {
    @($View.Config.Hardware.Device | Where-Object { $_.MacAddress } | ForEach-Object { $_.MacAddress })
}

function Get-VMPortGroups($View) {
    (@($View.Config.Hardware.Device | Where-Object { $_.MacAddress } | ForEach-Object { $_.DeviceInfo.Summary }) -join '; ')
}

#region ---------- Main ----------
if (-not (Test-Path $ListFile)) { Write-Host "List file not found: $ListFile" -ForegroundColor Red; return }
$names = @(Get-Content $ListFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notlike '#*' } | Select-Object -Unique)
Write-Host "Names in list: $($names.Count)"

try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue }
catch { Write-Host 'VMware PowerCLI is not installed.' -ForegroundColor Red; return }

$servers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($servers.Count -eq 0) {
    foreach ($vc in $VCenters) {
        try { $servers += Connect-VIServer -Server $vc -ErrorAction Stop }
        catch { Write-Host "Cannot connect to $vc : $($_.Exception.Message)" -ForegroundColor Red }
    }
}
if ($servers.Count -eq 0) { Write-Host 'No vCenter connection.' -ForegroundColor Red; return }
Write-Host "Using vCenter: $(($servers | ForEach-Object { $_.Name }) -join ', ')"

# All VMs and templates, read once
$props = 'Name', 'Config.Template', 'Config.Annotation', 'Config.Hardware.Device', 'Runtime.PowerState',
         'Guest.IpAddress', 'Guest.Net', 'Guest.HostName', 'Guest.ToolsRunningStatus', 'Guest.ToolsVersionStatus2'
$byName = @{}
foreach ($srv in $servers) {
    foreach ($v in @(Get-View -Server $srv -ViewType VirtualMachine -Property $props)) {
        $k = $v.Name.ToLower()
        if (-not $byName.ContainsKey($k)) { $byName[$k] = New-Object System.Collections.Generic.List[object] }
        $byName[$k].Add(@{ View = $v; VCenter = $srv.Name })
    }
}

# ARP table of this machine (MAC -> IP; only for VMs on the same network)
$arp = @{}
try {
    foreach ($n in @(Get-NetNeighbor -AddressFamily IPv4 -ErrorAction Stop | Where-Object { $_.LinkLayerAddress -and $_.LinkLayerAddress -ne '00-00-00-00-00-00' })) {
        $arp[$n.LinkLayerAddress.ToUpper()] = $n.IPAddress
    }
} catch { }

# DHCP leases (optional)
$dhcp = @{}
foreach ($ds in $DhcpServers) {
    try {
        foreach ($scope in @(Get-DhcpServerv4Scope -ComputerName $ds -ErrorAction Stop)) {
            foreach ($l in @(Get-DhcpServerv4Lease -ComputerName $ds -ScopeId $scope.ScopeId -ErrorAction Stop)) {
                if ($l.ClientId) { $dhcp[$l.ClientId.ToUpper()] = [string]$l.IPAddress }
            }
        }
    }
    catch { Write-Host "DHCP server $ds not readable: $($_.Exception.Message)" -ForegroundColor Yellow }
}

$results = foreach ($name in $names) {
    $hits = $byName[$name.ToLower()]
    if (-not $hits) {
        $dns = @(Resolve-DnsIPv4 $name)
        [pscustomobject]@{ 'VM NAME' = $name; Type = 'NOT FOUND in vCenter'; 'Power Status' = ''; 'VM Tools' = ''
            IP = ($dns -join ', '); Source = $(if ($dns.Count) { "DNS ($name)" } else { '' }); Ping = ''; MAC = ''; Network = ''; vCenter = '' }
        continue
    }
    foreach ($h in $hits) {
        $v = $h.View
        try {
            $r = Resolve-VMIp $v $arp $dhcp
            $ip = $r[0]; $firstIp = ($ip -split ',')[0].Trim()
            [pscustomobject]@{
                'VM NAME'      = $v.Name
                Type           = $(if ($v.Config.Template) { 'Template' } else { 'VM' })
                'Power Status' = $(if ($v.Config.Template) { '' } else { [string]$v.Runtime.PowerState })
                'VM Tools'     = [string]$v.Guest.ToolsRunningStatus
                IP             = $ip
                Source         = $(if ($ip) { $r[1] } else { 'Not found' })
                Ping           = $(if ($PingCheck -and $firstIp) { if (Test-Ping $firstIp) { 'Yes' } else { 'No' } } else { '' })
                MAC            = ((Get-VMMacs $v) -join ', ')
                Network        = (Get-VMPortGroups $v)
                vCenter        = $h.VCenter
            }
        }
        catch {
            [pscustomobject]@{ 'VM NAME' = $v.Name; Type = 'ERROR'; 'Power Status' = ''; 'VM Tools' = ''; IP = ''
                Source = $_.Exception.Message; Ping = ''; MAC = ''; Network = ''; vCenter = $h.VCenter }
        }
    }
}

$results | Export-Csv $OutFile -NoTypeInformation -Encoding UTF8
$results | Format-Table 'VM NAME', Type, 'Power Status', IP, Source, Ping -AutoSize | Out-String -Width 220 | Write-Host
$found = @($results | Where-Object { $_.IP }).Count
Write-Host ("IP found for {0} of {1}.  Results: {2}" -f $found, @($results).Count, $OutFile) -ForegroundColor Green
Write-Host 'Note: DNS / Notes / ARP results are not reported by the VM itself - check them (Ping column helps).' -ForegroundColor Yellow
#endregion
