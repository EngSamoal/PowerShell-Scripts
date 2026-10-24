# Snapshot-Step1-ListAll.ps1
# STEP 1 of 3 - Lists every VM snapshot on the vCenters you are connected to.
# Read-only: nothing is changed in vCenter.
#
# Before running : Connect-VIServer <vcenter1>, <vcenter2> ...
# Run            : .\Snapshot-Step1-ListAll.ps1
# Next step      : .\Snapshot-Step2-Filter.ps1 -NameKeyword '<word>' -DescriptionKeyword '<word>'

$ScriptBuild = 'Snapshot-Step1-ListAll 2026.10.24-r7'
Write-Host $ScriptBuild -ForegroundColor Cyan
$ErrorActionPreference = 'Continue'   # a VM that cannot be read is reported, the rest is still listed
# Full lines: nothing is cut off at the window edge; line breaks inside a description show as ' / '

Get-VM | Get-Snapshot | Select-Object VM,
    @{N='SnapshotName';E={$_.Name}},
    @{N='Description';E={($_.Description -replace '\s*[\r\n]+\s*', ' / ').Trim()}},
    Created,
    @{N='SizeGB';E={[math]::Round($_.SizeGB,2)}},
    @{N='Parent';E={$_.ParentSnapshot.Name}} | Format-Table -AutoSize | Out-String -Width 4096 | Write-Host
