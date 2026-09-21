<#
.SYNOPSIS
    Lists all VM snapshots that currently exist in vCenter.
    Assumes an active PowerCLI session (Connect-VIServer already run).

.EXAMPLE
    ./Get-VMSnapshots-PowerCli.ps1
#>

Get-VM | Get-Snapshot | Select-Object VM, Name, Created, SizeGB, Description | Sort-Object VM, Created | Format-Table -AutoSize
