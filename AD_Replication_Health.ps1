# AD Replication Health - Azure DC + GCP DC1 + GCP DC2
# Creates ONE disabled test user to prove end-to-end replication, then deletes it.
$dcs  = @('p-amcsccl-adc','sncgd-aad-app01','sncgd-aad-app02')
$src  = 'sncgd-aad-app01'          # DC where the test change is made
$wait = 90                         # seconds to wait for the change to replicate
$res  = [System.Collections.Generic.List[object]]::new()
function Add-R($dc,$check,$ok,$detail){ $res.Add([pscustomobject]@{DC=$dc; Check=$check; Result=$(if($ok){'PASS'}else{'FAIL'}); Detail=$detail}) }

Write-Host "AD replication health check - $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor Cyan

# 1) Inbound replication partners and failures
foreach ($dc in $dcs) {
  try {
    $m = @(Get-ADReplicationPartnerMetadata -Target $dc -Partition * -ErrorAction Stop)
    $from = ($m | % { ($_.Partner -split ',')[1] -replace 'CN=' } | Sort-Object -Unique) -join ', '
    $fail = @($m | ? { $_.LastReplicationResult -ne 0 }).Count
    Add-R $dc 'Inbound replication' ($fail -eq 0 -and $m.Count -gt 0) "from: $from | failures: $fail"
  } catch { Add-R $dc 'Inbound replication' $false $_.Exception.Message }
}

# 2) SYSVOL (DFSR) state - 4 = Normal
foreach ($dc in $dcs) {
  try {
    $st = (Get-CimInstance -ComputerName $dc -Namespace root/microsoftdfs -ClassName DfsrReplicatedFolderInfo -Filter "ReplicatedFolderName='SYSVOL Share'" -ErrorAction Stop).State
    Add-R $dc 'SYSVOL (DFSR) state' ($st -eq 4) "State=$st (4 = Normal)"
  } catch { Add-R $dc 'SYSVOL (DFSR) state' $false "Could not query: $($_.Exception.Message)" }
}

# 3) End-to-end test: create on $src, check it arrives on the others, then delete
$name = "ReplTest-$(Get-Date -Format 'yyyyMMddHHmmss')"
try {
  New-ADUser -Name $name -SamAccountName $name.Substring(0,20) -Enabled $false -Server $src -ErrorAction Stop
  Write-Host "Test user '$name' created on $src - waiting $wait seconds..." -ForegroundColor Yellow
  Start-Sleep -Seconds $wait
  foreach ($dc in ($dcs | ? { $_ -ne $src })) {
    $found = [bool](Get-ADUser -Filter "Name -eq '$name'" -Server $dc -ErrorAction SilentlyContinue)
    Add-R $dc "Test change from $src arrived" $found $(if($found){'Received'}else{"Not received after $wait s"})
  }
} catch { Add-R $src 'End-to-end test' $false $_.Exception.Message }
finally { Get-ADUser -Filter "Name -eq '$name'" -Server $src -ErrorAction SilentlyContinue | Remove-ADUser -Server $src -Confirm:$false -ErrorAction SilentlyContinue }

# Output
$res | Format-Table DC,Check,Result,Detail -AutoSize -Wrap | Out-String -Width 220 | Write-Host
$f = @($res | ? Result -eq 'FAIL').Count
if ($f -eq 0) { Write-Host "OVERALL: PASS - replication is healthy on all DCs" -ForegroundColor Green }
else          { Write-Host "OVERALL: FAIL - $f check(s) failed, see above" -ForegroundColor Red }
$res | Export-Csv "$env:USERPROFILE\Desktop\AD_Replication_Health.csv" -NoTypeInformation
