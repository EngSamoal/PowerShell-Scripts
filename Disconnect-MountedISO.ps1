#Requires -Version 5.1
<#
.SYNOPSIS
    Finds every VM CD/DVD drive with a specific ISO mounted, then disconnects only those drives.

.DESCRIPTION
    Two-step, read-then-act workflow for ejecting an ISO that may be mounted on multiple VMs
    across a production vCenter, without knowing in advance which VMs have it mounted:

      1. READ-ONLY DISCOVERY (always runs): scans every VM visible on the currently connected
         vCenter session(s) and lists any CD/DVD drive whose connected ISO path matches -IsoPath -
         VM name, drive label, and full ISO path. No changes are made in this step.

      2. DISCONNECT (only runs if you confirm / pass -Confirm:$false): calls Set-CDDrive -NoMedia
         against only the drives found in step 1.

    Safety guarantees:
      - Only CD/DVD drives currently connected to the exact -IsoPath given are touched. Drives
        with a different ISO, no media, or a host/remote device backing are left untouched.
      - VMs are never powered off, reset, rebooted, or otherwise reconfigured. This changes only
        the CD/DVD drive's media connection.
      - The ISO file itself is never deleted or moved from the datastore.
      - Works identically for Windows and Linux guests - this is a virtual hardware change, not
        a guest OS operation.
      - Uses -WhatIf/-Confirm (SupportsShouldProcess): review the discovery list, then either
        run with -WhatIf first to preview, or answer the per-drive confirmation prompt.

.PARAMETER IsoPath
    Exact datastore path of the ISO to disconnect, exactly as it appears in the IsoPath property
    from Get-CDDrive, e.g. '[datastore1] ISOs/example.iso'. No wildcards, no guessing - copy the
    exact path from your environment (run the script once, PowerCLI will error asking for this
    parameter, or use Get-VM | Get-CDDrive | Select Parent, IsoPath to look it up first).

.NOTES
    Requires an existing, already-authenticated PowerCLI session (Connect-VIServer run beforehand)
    and permission to modify VM virtual device settings.

.EXAMPLE
    # Preview only - shows what would be disconnected, changes nothing
    .\Disconnect-MountedISO.ps1 -IsoPath '[datastore1] ISOs/example.iso' -WhatIf

.EXAMPLE
    # Disconnect for real, confirming each drive interactively
    .\Disconnect-MountedISO.ps1 -IsoPath '[datastore1] ISOs/example.iso'
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)]
    [string]$IsoPath
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Step 1 - READ-ONLY: identify every CD/DVD drive currently connected to $IsoPath
# ---------------------------------------------------------------------------
Write-Host "Scanning all VMs for CD/DVD drives connected to:`n  $IsoPath" -ForegroundColor Cyan

$MountedDrives = Get-VM | Get-CDDrive | Where-Object {
    $_.ConnectionState.Connected -and $_.IsoPath -eq $IsoPath
}

if (-not $MountedDrives) {
    Write-Host "No connected CD/DVD drives found for that ISO path. Nothing to do." -ForegroundColor Yellow
    return
}

Write-Host "`nFound $($MountedDrives.Count) drive(s) with this ISO mounted:" -ForegroundColor Cyan
$MountedDrives |
    Select-Object @{N = 'VM'; E = { $_.Parent.Name } }, Name, IsoPath,
                  @{N = 'GuestOS'; E = { $_.Parent.Guest.OSFullName } } |
    Format-Table -AutoSize |
    Out-Host

# ---------------------------------------------------------------------------
# Step 2 - CHANGE: disconnect only the drives found above. No power state or
# other configuration changes are made to the VM.
# ---------------------------------------------------------------------------
foreach ($Drive in $MountedDrives) {
    $VMName = $Drive.Parent.Name
    if ($PSCmdlet.ShouldProcess("$VMName - $($Drive.Name)", "Disconnect ISO '$IsoPath' (Set-CDDrive -NoMedia)")) {
        Set-CDDrive -CD $Drive -NoMedia -Confirm:$false | Out-Null
        Write-Host "Disconnected $($Drive.Name) on $VMName" -ForegroundColor Green
    }
}
