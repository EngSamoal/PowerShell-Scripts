$dc   = 'sncgd-aad-app01'
$zone = 'amccinemasksa.net'
$subnets = @(
 'Riyadh Gallery|IT Sales/Service|10.7.62.0/24','Riyadh Gallery|Security|10.7.63.0/24','Riyadh Gallery|Sight & Sound|10.7.64.0/24',
 'Riyadh Gallery|QubeWire|10.6.64.0/24','Riyadh Gallery|ICC Vlan|10.77.62.0/24','Riyadh Gallery|Services|192.168.62.0/24',
 'Panorama|IT Sales/Service|10.7.53.0/24','Panorama|Security|10.7.54.0/24','Panorama|Sight & Sound|10.7.55.0/24',
 'Panorama|QubeWire|10.6.55.0/24','Panorama|ICC Vlan|10.77.53.0/24','Panorama|Services|192.168.80.0/24',
 'Stars Avenue|IT Sales/Service|10.7.68.0/24','Stars Avenue|Security|10.7.69.0/24','Stars Avenue|Sight & Sound|10.7.70.0/24',
 'Stars Avenue|QubeWire|10.6.70.0/24','Stars Avenue|ICC Vlan|10.77.68.0/24','Stars Avenue|Services|192.168.70.0/24',
 'Ajdan Walk|IT Sales/Service|10.7.74.0/24','Ajdan Walk|Security|10.7.75.0/24','Ajdan Walk|Sight & Sound|10.7.76.0/24',
 'Ajdan Walk|QubeWire|10.6.76.0/24','Ajdan Walk|ICC Vlan|10.77.74.0/24','Ajdan Walk|Services|192.168.74.0/24')

function ToNum($ip){ $b=([Net.IPAddress]$ip).GetAddressBytes(); [Array]::Reverse($b); [BitConverter]::ToUInt32($b,0) }
function InNet($ip,$cidr){ $n,$m = $cidr -split '/'; $mask = [uint32]([math]::Pow(2,32) - [math]::Pow(2,32-[int]$m)); ((ToNum $ip) -band $mask) -eq ((ToNum $n) -band $mask) }

# Domain computer accounts (name -> OS, last logon)
$ad = @{}
Get-ADComputer -Filter * -Server $dc -Properties OperatingSystem,LastLogonDate | % { $ad[$_.Name.ToUpper()] = $_ }

# All A records in DNS
$recs = Get-DnsServerResourceRecord -ComputerName $dc -ZoneName $zone -RRType A |
        ? { $_.HostName -notmatch '^(@|DomainDnsZones|ForestDnsZones)$' -and $_.HostName -notmatch '\.' } |
        % { [pscustomobject]@{ Name=$_.HostName; IP=$_.RecordData.IPv4Address.IPAddressToString; Seen=$_.Timestamp } }

# Full machine list per subnet
$all = foreach ($s in $subnets) {
  $site,$type,$cidr = $s -split '\|'
  foreach ($r in ($recs | ? { InNet $_.IP $cidr })) {
    $c = $ad[$r.Name.ToUpper()]
    [pscustomobject]@{
      Site=$site; Type=$type; Subnet=$cidr; Name=$r.Name; IP=$r.IP
      Domain= if($c){'Domain'}else{'Non-domain'}
      OS= if($c){$c.OperatingSystem}else{''}
      LastLogon= if($c -and $c.LastLogonDate){$c.LastLogonDate.ToString('yyyy-MM-dd')}else{''}
      DnsSeen= if($r.Seen){$r.Seen.ToString('yyyy-MM-dd')}else{'static'}
    }
  }
}

# 1) Per subnet summary
$summary = foreach ($s in $subnets) {
  $site,$type,$cidr = $s -split '\|'
  $m = @($all | ? Subnet -eq $cidr)
  [pscustomobject]@{ Site=$site; Type=$type; Subnet=$cidr; Total=$m.Count
    Domain=@($m | ? Domain -eq 'Domain').Count; NonDomain=@($m | ? Domain -eq 'Non-domain').Count }
}
Write-Host "`n=== 1) SUBNETS AND MACHINE COUNTS ===" -ForegroundColor Cyan
$summary | Format-Table -AutoSize | Out-String -Width 200 | Write-Host

# 2) Per site summary
Write-Host "=== 2) PER SITE ===" -ForegroundColor Cyan
$summary | Group-Object Site | Sort-Object Name | % { [pscustomobject]@{ Site=$_.Name; Subnets=$_.Count
  SubnetsWithMachines=@($_.Group | ? Total -gt 0).Count; Machines=[int]($_.Group | Measure-Object Total -Sum).Sum
  Domain=[int]($_.Group | Measure-Object Domain -Sum).Sum; NonDomain=[int]($_.Group | Measure-Object NonDomain -Sum).Sum } } |
  Format-Table -AutoSize | Out-String -Width 200 | Write-Host

# 3) Non-domain machines
Write-Host "=== 3) NON-DOMAIN MACHINES (in DNS, no AD computer account) ===" -ForegroundColor Cyan
$all | ? Domain -eq 'Non-domain' | Sort-Object Site,Subnet,@{e={ToNum $_.IP}} | Format-Table Site,Type,Subnet,Name,IP,DnsSeen -AutoSize | Out-String -Width 200 | Write-Host

# 4) Full list -> CSV (all machines with IPs)
$path = "$env:USERPROFILE\Desktop\AD_Subnet_Inventory.csv"
$all | Sort-Object Site,Subnet,@{e={ToNum $_.IP}} | Export-Csv $path -NoTypeInformation
$sumPath = "$env:USERPROFILE\Desktop\AD_Subnet_Summary.csv"
$summary | Export-Csv $sumPath -NoTypeInformation
Write-Host "Full machine list (every IP): $path" -ForegroundColor Yellow
Write-Host "Subnet summary:               $sumPath" -ForegroundColor Yellow
