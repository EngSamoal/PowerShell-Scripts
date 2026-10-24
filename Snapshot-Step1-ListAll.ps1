# Snapshot-Step1-ListAll.ps1
# STEP 1 of 3 - Lists every VM snapshot on all vCenters you are connected to.
# Read-only: nothing is changed in vCenter.
#
# Before running : Connect-VIServer <vcenter1>, <vcenter2> ...
# Run            : .\Snapshot-Step1-ListAll.ps1
# Output         : list on screen + C:\temp\SnapshotCleanup\1-AllSnapshots.csv
# Next step      : .\Snapshot-Step2-Filter.ps1
param(
    [string]$OutputFolder = 'C:\temp\SnapshotCleanup'
)

$ScriptBuild = 'Snapshot-Step1-ListAll 2026.10.24-r4'
Write-Host $ScriptBuild -ForegroundColor Cyan
$ErrorActionPreference = 'Stop'

if (-not (Get-Command Get-Snapshot -ErrorAction SilentlyContinue)) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue }
    catch { Write-Host 'VMware PowerCLI is not installed.' -ForegroundColor Red; return }
}
$servers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($servers.Count -eq 0) { Write-Host 'Not connected to any vCenter. Run Connect-VIServer first.' -ForegroundColor Red; return }
Write-Host ("Connected vCenters: {0}" -f (($servers | ForEach-Object { $_.Name }) -join ', '))

if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$outCsv = Join-Path $OutputFolder '1-AllSnapshots.csv'

$all = New-Object System.Collections.ArrayList
$errors = 0
foreach ($srv in $servers) {
    Write-Host "Reading snapshots from $($srv.Name) ..." -ForegroundColor Yellow
    try { $vms = @(Get-VM -Server $srv -ErrorAction Stop) }
    catch { Write-Host "  Cannot list VMs on $($srv.Name): $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
    $n = 0
    foreach ($vm in $vms) {
        $n++
        if ($n % 50 -eq 0) { Write-Progress -Activity "Scanning $($srv.Name)" -Status "$n / $($vms.Count) VMs" -PercentComplete ($n * 100 / $vms.Count) }
        try { $snaps = @(Get-Snapshot -VM $vm -Server $srv -ErrorAction Stop) }
        catch { Write-Host "  Cannot read snapshots of $($vm.Name): $($_.Exception.Message)" -ForegroundColor Red; $errors++; continue }
        foreach ($s in $snaps) {
            $parentName = ''
            if ($s.Parent) { $parentName = $s.Parent.Name }
            [void]$all.Add([pscustomobject]@{
                VCenter      = $srv.Name
                VM           = $vm.Name
                SnapshotName = [string]$s.Name
                Description  = [string]$s.Description
                Created      = $s.Created.ToString('yyyy-MM-dd HH:mm')
                AgeDays      = [math]::Round(((Get-Date) - $s.Created).TotalDays, 1)
                SizeGB       = [math]::Round([double]$s.SizeGB, 2)
                ParentName   = $parentName
                HasChildren  = (@($s.Children).Count -gt 0)
                PowerState   = [string]$vm.PowerState
                VMId         = $vm.Id
                SnapshotId   = $s.Id
            })
        }
    }
    Write-Progress -Activity "Scanning $($srv.Name)" -Completed
}

$all = @($all | Sort-Object Created)
$all | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
if ($all.Count -gt 0) { $all | Format-Table VCenter, VM, SnapshotName, Description, Created, SizeGB -AutoSize -Wrap | Out-String -Width 250 | Write-Host }
$totalGB = [math]::Round(($all | Measure-Object SizeGB -Sum).Sum + 0, 2)
Write-Host ''
Write-Host ("Snapshots found: {0}  ({1} GB)" -f $all.Count, $totalGB) -ForegroundColor Green
Write-Host "Saved to: $outCsv" -ForegroundColor Green
if ($errors -gt 0) { Write-Host "WARNING: $errors VM(s)/vCenter(s) could not be read - see red messages above." -ForegroundColor Red }
Write-Host ''
Write-Host "NEXT: .\Snapshot-Step2-Filter.ps1 -NameKeyword '<word>' -DescriptionKeyword '<word>'" -ForegroundColor Cyan
