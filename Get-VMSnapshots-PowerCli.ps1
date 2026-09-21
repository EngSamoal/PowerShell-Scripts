<#
.SYNOPSIS
    Lists all VM snapshots that currently exist in vCenter.

.EXAMPLE
    ./Get-VMSnapshots-PowerCli.ps1 -vCenter vcenter01.domain.local
#>

param(
    [Parameter(Mandatory)]
    [string]$vCenter
)

Connect-VIServer -Server $vCenter | Out-Null

Get-VM | Get-Snapshot | Select-Object VM, Name, Created, SizeGB, Description | Sort-Object VM, Created | Format-Table -AutoSize

Disconnect-VIServer -Server $vCenter -Confirm:$false
