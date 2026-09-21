<#
.SYNOPSIS
    Removes VM snapshots whose name matches or contains the given text.

.EXAMPLE
    ./Remove-VMSnapshots-PowerCli.ps1 -vCenter vcenter01.domain.local -NameMatch "Pre-Patch"
#>

param(
    [Parameter(Mandatory)]
    [string]$vCenter,

    [Parameter(Mandatory)]
    [string]$NameMatch
)

Connect-VIServer -Server $vCenter | Out-Null

$snapshots = Get-VM | Get-Snapshot | Where-Object { $_.Name -like "*$NameMatch*" }

if (-not $snapshots) {
    Write-Host "No snapshots found matching '$NameMatch'."
} else {
    $snapshots | Select-Object VM, Name, Created | Format-Table -AutoSize

    foreach ($snap in $snapshots) {
        Remove-Snapshot -Snapshot $snap -Confirm:$true
    }
}

Disconnect-VIServer -Server $vCenter -Confirm:$false
