<#
.SYNOPSIS
    Clones every VM listed in a text file (one VM name per line) using PowerCLI.

.DESCRIPTION
    For each VM name in the list, a full clone is created named <VMName><CloneSuffix>
    (default suffix: "_clone"). By default each clone lands on the same host/cluster,
    datastore and VM folder as its source VM; any of these can be overridden.

    - Blank lines and lines starting with '#' in the list are ignored.
    - VMs that aren't found, or whose clone name already exists, are skipped (not failed).
    - Powered-on VMs can be cloned (vCenter clones them live); the clone is left powered off
      unless -PowerOnClone is used.
    - A CSV results log is written next to the VM list (C:\temp\VMClone_Results_<timestamp>.csv).
    - Use -WhatIf for a dry run that only shows what would be cloned.
    - If you are already connected to vCenter (Connect-VIServer), that session is reused and
      left connected. Otherwise the script connects to -vCenterServer and disconnects at the end.

.EXAMPLE
    .\Clone-VMsFromList-PowerCli.ps1
    (Uses your existing Connect-VIServer session if you are already connected.)

.EXAMPLE
    .\Clone-VMsFromList-PowerCli.ps1 -vCenterServer 10.50.10.10 -CloneSuffix '_bak_2026' -Datastore 'DS-Backup01' -DiskStorageFormat Thin

.EXAMPLE
    .\Clone-VMsFromList-PowerCli.ps1 -vCenterServer 10.50.10.10 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$vCenterServer = '10.50.10.10',
    [string]$VMListPath    = 'C:\temp\vmlist.txt',
    [string]$CloneSuffix   = '_clone',

    # Optional overrides - leave empty to use the source VM's own location
    [string]$Datastore     = '',   # target datastore or datastore cluster name
    [string]$Cluster       = '',   # target cluster name (a host in it is chosen automatically)
    [string]$VMHost        = '',   # target ESXi host name (takes precedence over -Cluster)
    [string]$Folder        = '',   # target VM folder name

    [ValidateSet('Thin', 'Thick', 'EagerZeroedThick', 'SameAsSource')]
    [string]$DiskStorageFormat = 'SameAsSource',

    [switch]$PowerOnClone
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Read the VM list
# ---------------------------------------------------------------------------
if (-not (Test-Path -Path $VMListPath)) {
    Write-Host "VM list not found: $VMListPath" -ForegroundColor Red
    return
}

$vmNames = @(Get-Content -Path $VMListPath |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -ne '' -and -not $_.StartsWith('#') } |
    Select-Object -Unique)

if ($vmNames.Count -eq 0) {
    Write-Host "No VM names found in $VMListPath" -ForegroundColor Yellow
    return
}

Write-Host "Found $($vmNames.Count) VM name(s) in $VMListPath" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# Connect to vCenter
# ---------------------------------------------------------------------------
$connectedHere = $false
if ($global:DefaultVIServer -and $global:DefaultVIServer.IsConnected) {
    $viConnection = $global:DefaultVIServer
    Write-Host "Using existing vCenter connection: $($viConnection.Name)" -ForegroundColor Cyan
} else {
    $viConnection  = Connect-VIServer -Server $vCenterServer
    $connectedHere = $true
}

$timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
$resultPath = Join-Path -Path (Split-Path -Path $VMListPath -Parent) -ChildPath "VMClone_Results_$timestamp.csv"
$results    = New-Object System.Collections.Generic.List[object]

function Add-Result {
    param($SourceVM, $CloneName, $Status, $Details)
    $results.Add([pscustomobject]@{
        Time      = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        SourceVM  = $SourceVM
        CloneName = $CloneName
        Status    = $Status
        Details   = $Details
    })
}

try {
    $i = 0
    foreach ($vmName in $vmNames) {
        $i++
        $cloneName = "$vmName$CloneSuffix"
        Write-Host "`n[$i/$($vmNames.Count)] === $vmName -> $cloneName ===" -ForegroundColor Cyan

        try {
            # --- Source VM ---
            $sourceVM = @(Get-VM -Name $vmName -ErrorAction SilentlyContinue | Where-Object { $_ })
            if ($sourceVM.Count -eq 0) {
                Write-Host "SKIPPED: VM '$vmName' not found in vCenter" -ForegroundColor Yellow
                Add-Result $vmName $cloneName 'Skipped' 'Source VM not found'
                continue
            }
            if ($sourceVM.Count -gt 1) {
                Write-Host "SKIPPED: more than one VM named '$vmName' - clone it manually" -ForegroundColor Yellow
                Add-Result $vmName $cloneName 'Skipped' "Ambiguous name: $($sourceVM.Count) VMs match"
                continue
            }
            $sourceVM = $sourceVM[0]

            # --- Clone name must not already exist ---
            if (Get-VM -Name $cloneName -ErrorAction SilentlyContinue) {
                Write-Host "SKIPPED: a VM named '$cloneName' already exists" -ForegroundColor Yellow
                Add-Result $vmName $cloneName 'Skipped' 'Clone name already exists'
                continue
            }

            # --- Target location (default = same as source) ---
            $cloneParams = @{
                Name        = $cloneName
                VM          = $sourceVM
                ErrorAction = 'Stop'
            }

            if ($VMHost) {
                $cloneParams['VMHost'] = Get-VMHost -Name $VMHost
            } elseif ($Cluster) {
                $cloneParams['ResourcePool'] = Get-Cluster -Name $Cluster
            } else {
                $cloneParams['VMHost'] = $sourceVM.VMHost
            }

            if ($Datastore) {
                $cloneParams['Datastore'] = Get-DatastoreCluster -Name $Datastore -ErrorAction SilentlyContinue
                if (-not $cloneParams['Datastore']) {
                    $cloneParams['Datastore'] = Get-Datastore -Name $Datastore
                }
            } else {
                # Datastore holding the source VM's config (.vmx) file
                $vmxDatastoreName = $sourceVM.ExtensionData.Config.Files.VmPathName.Split(']')[0].TrimStart('[')
                $cloneParams['Datastore'] = Get-Datastore -Name $vmxDatastoreName
            }

            if ($Folder) {
                $cloneParams['Location'] = Get-Folder -Name $Folder -Type VM
            } else {
                $cloneParams['Location'] = $sourceVM.Folder
            }

            if ($DiskStorageFormat -ne 'SameAsSource') {
                $cloneParams['DiskStorageFormat'] = $DiskStorageFormat
            }

            $targetDesc = "host/cluster: $(if ($cloneParams['VMHost']) { $cloneParams['VMHost'].Name } else { $cloneParams['ResourcePool'].Name }), " +
                          "datastore: $($cloneParams['Datastore'].Name), folder: $($cloneParams['Location'].Name)"
            Write-Host "Source power state: $($sourceVM.PowerState)"
            Write-Host "Target -> $targetDesc"

            if (-not $PSCmdlet.ShouldProcess($vmName, "Clone to '$cloneName' ($targetDesc)")) {
                Add-Result $vmName $cloneName 'WhatIf' $targetDesc
                continue
            }

            # --- Clone ---
            $start    = Get-Date
            $newVM    = New-VM @cloneParams
            $duration = [math]::Round(((Get-Date) - $start).TotalMinutes, 1)
            Write-Host "SUCCESS: '$cloneName' created in $duration min" -ForegroundColor Green

            if ($PowerOnClone) {
                Start-VM -VM $newVM -Confirm:$false | Out-Null
                Write-Host "Powered on '$cloneName'" -ForegroundColor Green
            }

            Add-Result $vmName $cloneName 'Success' "$targetDesc; $duration min"
        } catch {
            Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
            Add-Result $vmName $cloneName 'Failed' $_.Exception.Message
        }
    }
} finally {
    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $resultPath -NoTypeInformation -WhatIf:$false
        Write-Host "`nResults saved to: $resultPath" -ForegroundColor Cyan
    }

    Write-Host "`n===== Summary =====" -ForegroundColor Cyan
    $results | Group-Object -Property Status | ForEach-Object {
        Write-Host ("{0,-8}: {1}" -f $_.Name, $_.Count)
    }

    # Only disconnect if this script opened the connection
    if ($connectedHere) {
        Disconnect-VIServer -Server $viConnection -Confirm:$false
    }
}
