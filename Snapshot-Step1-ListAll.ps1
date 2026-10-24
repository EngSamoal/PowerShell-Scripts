# Snapshot-Step1-ListAll.ps1
# STEP 1 of 3 - Lists every VM snapshot on the vCenters you are connected to.
# Read-only: nothing is changed in vCenter.
#
# Before running : Connect-VIServer <vcenter1>, <vcenter2> ...
# Run            : .\Snapshot-Step1-ListAll.ps1
# Next step      : .\Snapshot-Step2-Filter.ps1 -NameKeyword '<word>' -DescriptionKeyword '<word>'

$ScriptBuild = 'Snapshot-Step1-ListAll 2026.10.24-r6'
Write-Host $ScriptBuild -ForegroundColor Cyan
$ErrorActionPreference = 'Continue'   # a VM that cannot be read is reported, the rest is still listed

Get-VM | Get-Snapshot | Select-Object VM,
    @{N='SnapshotName';E={$_.Name}},
    Description,
    Created,
    @{N='SizeGB';E={[math]::Round($_.SizeGB,2)}},
    @{N='Parent';E={$_.ParentSnapshot.Name}} | Format-Table -AutoSize
