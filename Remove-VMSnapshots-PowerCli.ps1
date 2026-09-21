<#
.SYNOPSIS
    Removes VM snapshots whose name matches or contains the given text.
    Assumes an active PowerCLI session (Connect-VIServer already run).

.EXAMPLE
    ./Remove-VMSnapshots-PowerCli.ps1 -NameMatch "Pre-Patch"
#>

param(
    [Parameter(Mandatory)]
    [string]$NameMatch
)

$snapshots = Get-VM | Get-Snapshot | Where-Object { $_.Name -like "*$NameMatch*" }

if (-not $snapshots) {
    Write-Host "No snapshots found matching '$NameMatch'."
} else {
    $snapshots | Select-Object VM, Name, Created | Format-Table -AutoSize

    foreach ($snap in $snapshots) {
        Remove-Snapshot -Snapshot $snap -Confirm:$true
    }
}
